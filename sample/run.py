"""Combine CSV/XLSX sales files into auditable daily reports."""
from __future__ import annotations

import argparse
import csv
import logging
import re
import sys
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CONFIG = ROOT / "config" / "columns.yml"
LOG = logging.getLogger("sales_report")


def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def load_config(path: Path) -> dict[str, Any]:
    with path.open(encoding="utf-8") as stream:
        config = yaml.safe_load(stream) or {}
    if not isinstance(config.get("columns"), dict) or not config["columns"]:
        raise ValueError("Config must define a non-empty 'columns' mapping")
    return config


def map_headers(headers: list[Any], config: dict[str, Any]) -> dict[str, int]:
    aliases: dict[str, str] = {}
    for canonical, variants in config["columns"].items():
        for variant in [canonical, *variants]:
            key = normalize_header(variant)
            if key in aliases and aliases[key] != canonical:
                raise ValueError(f"Ambiguous header alias: {variant}")
            aliases[key] = canonical
    mapped: dict[str, int] = {}
    for index, header in enumerate(headers):
        canonical = aliases.get(normalize_header(header))
        if canonical:
            if canonical in mapped:
                raise ValueError(f"Duplicate source columns map to '{canonical}'")
            mapped[canonical] = index
    return mapped


def read_records(path: Path) -> list[tuple[int, dict[str, Any]]]:
    if path.suffix.lower() == ".csv":
        with path.open(newline="", encoding="utf-8-sig") as stream:
            rows = list(csv.reader(stream))
        if not rows:
            return []
        headers = rows[0]
        return [(number, dict(zip(headers, row))) for number, row in enumerate(rows[1:], 2)]
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        # The first sheet is the declared data sheet; unrelated extra sheets are ignored.
        sheet = workbook[workbook.sheetnames[0]]
        iterator = sheet.iter_rows(values_only=True)
        headers = list(next(iterator, ()))
        return [(number, dict(zip(headers, row))) for number, row in enumerate(iterator, 2)]
    finally:
        workbook.close()


def parse_date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d", "%b %d, %Y", "%d-%b-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError("invalid or missing date")


def parse_money(value: Any) -> Decimal:
    if value is None or str(value).strip() == "":
        raise ValueError("missing amount")
    cleaned = str(value).strip().replace(",", "").replace("$", "").replace("€", "").replace("£", "")
    cleaned = re.sub(r"^\((.*)\)$", r"-\1", cleaned)
    try:
        amount = Decimal(cleaned)
    except InvalidOperation as exc:
        raise ValueError("invalid amount") from exc
    if not amount.is_finite():
        raise ValueError("invalid amount")
    return amount.quantize(Decimal("0.01"))


def parse_quantity(value: Any) -> int:
    try:
        number = Decimal(str(value).strip())
    except (InvalidOperation, AttributeError):
        raise ValueError("invalid or missing quantity")
    if not number.is_finite() or number != number.to_integral_value() or number <= 0:
        raise ValueError("quantity must be a positive whole number")
    return int(number)


def clean_row(raw: dict[str, Any], mapped: dict[str, int], config: dict[str, Any]) -> tuple[dict[str, Any] | None, str | None]:
    required = config.get("required", list(config["columns"]))
    # Raw records retain source labels; mapped indexes refer to the original header order.
    headers = mapped.get("_headers", [])
    values = {name: raw.get(headers[index]) for name, index in mapped.items() if name != "_headers"}
    for name in required:
        if values.get(name) is None or str(values.get(name)).strip() == "":
            return None, f"missing required field: {name}"
    try:
        normalized = {
            "date": parse_date(values.get("date")),
            "product": str(values.get("product")).strip(),
            "category": str(values.get("category")).strip(),
            "quantity": parse_quantity(values.get("quantity")),
            "unit_price": parse_money(values.get("unit_price")),
            "amount": parse_money(values.get("amount")),
        }
    except ValueError as exc:
        return None, str(exc)
    if not normalized["product"] or not normalized["category"]:
        return None, "product and category must not be blank"
    expected = (normalized["quantity"] * normalized["unit_price"]).quantize(Decimal("0.01"))
    if normalized["amount"] != expected:
        return None, f"amount mismatch: expected {expected:.2f} from quantity x unit_price"
    return normalized, None


def discover_files(input_dir: Path) -> list[Path]:
    return sorted(path for path in input_dir.iterdir() if path.is_file() and path.suffix.lower() in {".csv", ".xlsx"} and not path.name.startswith("~$"))


def process(input_dir: Path, output_dir: Path, config_path: Path = DEFAULT_CONFIG) -> dict[str, Any]:
    config = load_config(config_path)
    output_dir.mkdir(parents=True, exist_ok=True)
    accepted: list[dict[str, Any]] = []
    rejects: list[dict[str, Any]] = []
    seen: set[tuple[Any, ...]] = set()
    files = discover_files(input_dir)
    if not files:
        raise ValueError(f"No .csv or .xlsx input files found in {input_dir}")
    for path in files:
        records = read_records(path)
        if not records:
            LOG.warning("Skipping empty input file: %s", path.name)
            continue
        headers = list(records[0][1].keys())
        try:
            mapped = map_headers(headers, config)
        except ValueError as exc:
            for line, _ in records:
                rejects.append({"source_file": path.name, "source_row": line, "reason": str(exc)})
            continue
        mapped["_headers"] = headers
        required = config.get("required", list(config["columns"]))
        absent = [field for field in required if field not in mapped]
        if absent:
            reason = "missing required columns: " + ", ".join(absent)
            rejects.extend({"source_file": path.name, "source_row": line, "reason": reason} for line, _ in records)
            continue
        for line, raw in records:
            row, reason = clean_row(raw, mapped, config)
            if reason:
                rejects.append({"source_file": path.name, "source_row": line, "reason": reason})
                continue
            key = tuple(row[field] for field in ("date", "product", "category", "quantity", "unit_price", "amount"))
            if key in seen:
                rejects.append({"source_file": path.name, "source_row": line, "reason": "duplicate row"})
                continue
            seen.add(key)
            accepted.append(row)
    write_outputs(output_dir, accepted, rejects)
    LOG.info("Processed %d files: %d accepted, %d rejected", len(files), len(accepted), len(rejects))
    return {"files": len(files), "accepted": accepted, "rejects": rejects}


def write_outputs(output_dir: Path, rows: list[dict[str, Any]], rejects: list[dict[str, Any]]) -> None:
    with (output_dir / "rejects.csv").open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=["source_file", "source_row", "reason"])
        writer.writeheader()
        writer.writerows(rejects)

    daily: dict[date, dict[str, Any]] = defaultdict(lambda: {"orders": 0, "units": 0, "revenue": Decimal("0.00")})
    pivot: dict[tuple[date, str], Decimal] = defaultdict(lambda: Decimal("0.00"))
    for row in rows:
        bucket = daily[row["date"]]
        bucket["orders"] += 1
        bucket["units"] += row["quantity"]
        bucket["revenue"] += row["amount"]
        pivot[(row["date"], row["category"])] += row["amount"]
    wb = Workbook()
    summary = wb.active
    summary.title = "Daily Summary"
    summary.append(["Date", "Orders", "Units", "Revenue"])
    for day in sorted(daily):
        summary.append([day, daily[day]["orders"], daily[day]["units"], float(daily[day]["revenue"])])
    detail = wb.create_sheet("Daily by Category")
    detail.append(["Date", "Category", "Revenue"])
    for (day, category), amount in sorted(pivot.items()):
        detail.append([day, category, float(amount)])
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        for column in ws.columns:
            width = min(max(max(len(str(cell.value or "")) for cell in column) + 2, 12), 28)
            ws.column_dimensions[get_column_letter(column[0].column)].width = width
        for cell in ws.iter_rows(min_row=2):
            for item in cell:
                if item.column == 1:
                    item.number_format = "yyyy-mm-dd"
                if (ws.title == "Daily Summary" and item.column == 4) or (ws.title == "Daily by Category" and item.column == 3):
                    item.number_format = '$#,##0.00;[Red]-$#,##0.00'
    wb.save(output_dir / "report.xlsx")

    lines = ["# Daily Sales Report", "", f"Accepted rows: **{len(rows)}**", f"Rejected rows: **{len(rejects)}**", "", "## Daily summary", "", "| Date | Orders | Units | Revenue |", "|---|---:|---:|---:|"]
    for day in sorted(daily):
        info = daily[day]
        lines.append(f"| {day.isoformat()} | {info['orders']} | {info['units']} | ${info['revenue']:,.2f} |")
    total_revenue = sum((info["revenue"] for info in daily.values()), Decimal("0.00"))
    lines += ["", f"Total revenue: **${total_revenue:,.2f}**", "", "## Revenue by day and category", "", "| Date | Category | Revenue |", "|---|---|---:|"]
    for (day, category), amount in sorted(pivot.items()):
        lines.append(f"| {day.isoformat()} | {category} | ${amount:,.2f} |")
    (output_dir / "report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    args = parser.parse_args(argv)
    args.output.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", handlers=[logging.StreamHandler(), logging.FileHandler(args.output / "run.log", encoding="utf-8")])
    try:
        result = process(args.input, args.output, args.config)
    except (OSError, ValueError, yaml.YAMLError) as exc:
        LOG.error("Run failed: %s", exc)
        return 1
    print(f"Processed {result['files']} files: {len(result['accepted'])} accepted, {len(result['rejects'])} rejected")
    return 0


if __name__ == "__main__":
    sys.exit(main())
