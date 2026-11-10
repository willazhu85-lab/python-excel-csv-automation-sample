import csv
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

from sample.run import clean_row, load_config, map_headers, parse_date, parse_money, parse_quantity, process

ROOT = Path(__file__).resolve().parents[1]
CONFIG = load_config(ROOT / "config" / "columns.yml")

@pytest.mark.parametrize(("value", "expected"), [
    ("2025-01-02", date(2025, 1, 2)), ("1/2/2025", date(2025, 1, 2)),
    ("Jan 03, 2025", date(2025, 1, 3)), ("05-Jan-2025", date(2025, 1, 5)),
])
def test_parse_supported_dates(value, expected):
    assert parse_date(value) == expected

def test_invalid_date():
    with pytest.raises(ValueError, match="invalid or missing date"):
        parse_date("not a date")

@pytest.mark.parametrize(("value", "expected"), [("$1,234.50", Decimal("1234.50")), ("(£12.00)", Decimal("-12.00")), ("€0.5", Decimal("0.50"))])
def test_parse_currency_and_grouping(value, expected):
    assert parse_money(value) == expected

@pytest.mark.parametrize("value", [None, "", "oops", "NaN"])
def test_invalid_or_missing_money(value):
    with pytest.raises(ValueError):
        parse_money(value)

@pytest.mark.parametrize(("value", "expected"), [("3", 3), (3, 3)])
def test_parse_quantity(value, expected):
    assert parse_quantity(value) == expected

@pytest.mark.parametrize("value", [0, -2, "1.5", "x", None])
def test_reject_invalid_quantity(value):
    with pytest.raises(ValueError):
        parse_quantity(value)

def test_header_mapping_ignores_case_and_whitespace():
    result = map_headers([" DATE ", " ITEM ", "Qty"], CONFIG)
    assert result == {"date": 0, "product": 1, "quantity": 2}

def test_missing_required_field_is_rejected():
    mapping = map_headers(["Date", "Product", "Category", "Qty", "Unit Price", "Amount"], CONFIG)
    mapping["_headers"] = ["Date", "Product", "Category", "Qty", "Unit Price", "Amount"]
    row, reason = clean_row({"Date": "2025-01-01", "Product": "", "Category": "X", "Qty": "1", "Unit Price": "2", "Amount": "2"}, mapping, CONFIG)
    assert row is None and reason == "missing required field: product"

def test_amount_must_match_quantity_times_price():
    headers = ["Date", "Product", "Category", "Qty", "Unit Price", "Amount"]
    mapping = map_headers(headers, CONFIG); mapping["_headers"] = headers
    row, reason = clean_row(dict(zip(headers, ["2025-01-01", "A", "X", 2, "$5", "$8"])), mapping, CONFIG)
    assert row is None and "amount mismatch" in reason

def test_end_to_end_manual_totals_and_all_bad_rows_rejected(tmp_path):
    output = tmp_path / "out"
    result = process(ROOT / "tests" / "fixtures", output)
    assert len(result["accepted"]) == 6
    assert len(result["rejects"]) == 5
    assert sum((row["amount"] for row in result["accepted"]), Decimal("0")) == Decimal("87.00")
    totals = {}
    for row in result["accepted"]:
        key = row["date"].isoformat()
        totals[key] = totals.get(key, Decimal("0")) + row["amount"]
    assert totals == {"2025-01-01": Decimal("20.00"), "2025-01-02": Decimal("27.00"), "2025-01-03": Decimal("12.00"), "2025-01-04": Decimal("20.00"), "2025-01-05": Decimal("8.00")}
    reasons = [item["reason"] for item in result["rejects"]]
    assert "duplicate row" in reasons
    assert any(reason.startswith("missing required field") for reason in reasons)
    assert any("invalid or missing date" in reason for reason in reasons)
    assert any("amount mismatch" in reason for reason in reasons)
    assert any("positive whole number" in reason for reason in reasons)
    with (output / "rejects.csv").open(encoding="utf-8") as stream:
        assert len(list(csv.DictReader(stream))) == 5
    workbook = load_workbook(output / "report.xlsx", data_only=True)
    assert workbook["Daily Summary"].max_row == 6
    assert workbook["Daily Summary"]["D2"].value == 20
    assert workbook["Daily by Category"].max_row == 7
    assert "87.00" in (output / "report.md").read_text()
