"""Create deterministic, fictional sales files for the customer demo."""
from __future__ import annotations

import csv
import random
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "data" / "input"
SEED = 20261110
CATEGORIES = ["Books", "Electronics", "Grocery", "Home", "Office", "Outdoor"]
# Different store weights create distinct operating scales.
STORES = [("Store North", 0.40), ("Store South", 0.29), ("Store East", 0.20), ("Store West", 0.11)]
SCHEMAS = [
    ["Sale Date", "Item", "Product Category", "Qty", "Price Each", "Total", "Store"],
    ["DATE", "Product Name", "Category", "Units", "Unit Price", "Amount", "Store"],
    ["transaction date", "product", "category", "quantity", "unit price", "sales amount", "store"],
    ["Date", "Item", "Category", "Quantity", "Price", "Revenue", "Store"],
    ["SALE DATE", "PRODUCT", "CATEGORY", "QTY", "UNIT PRICE", "TOTAL", "STORE"],
]


def date_text(day: date, style: int) -> str:
    return (day.isoformat(), day.strftime("%m/%d/%Y"), day.strftime("%d-%b-%Y"),
            day.strftime("%b %d, %Y"), day.strftime("%Y/%m/%d"))[style % 5]


def money_text(cents: int, style: int) -> str:
    amount = cents / 100
    if style % 4 == 0:
        return f"${amount:,.2f}"
    if style % 4 == 1:
        return f"{amount:,.2f}"
    if style % 4 == 2:
        return f"£{amount:.2f}"
    return f"{amount:.2f}"


def make_order(rng: random.Random, day: date, serial: int, promo: bool = False) -> dict[str, object]:
    store = rng.choices(STORES, weights=[weight for _, weight in STORES], k=1)[0][0]
    category = rng.choice(CATEGORIES)
    quantity = rng.choices(range(1, 9), weights=[28, 24, 17, 12, 8, 5, 3, 3], k=1)[0]
    # Prices and basket sizes differ by store and category; promotion adds volume and discount.
    store_factor = {"Store North": 1.12, "Store South": 1.0, "Store East": 0.86, "Store West": 0.72}[store]
    category_factor = {"Books": 0.75, "Electronics": 2.8, "Grocery": 0.55,
                       "Home": 1.5, "Office": 0.9, "Outdoor": 1.8}[category]
    unit_cents = max(199, int(rng.lognormvariate(7.0, 0.62) * store_factor * category_factor))
    if promo and category in {"Electronics", "Home", "Outdoor"}:
        unit_cents = round(unit_cents * 0.85)
    return {
        "date": date_text(day, serial + rng.randrange(5)),
        "product": f"{category} item {rng.randint(1, 90):02d}",
        "category": category,
        "quantity": quantity,
        "unit_price": money_text(unit_cents, serial + rng.randrange(4)),
        "amount": money_text(unit_cents * quantity, serial * 3 + rng.randrange(4)),
        "store": store,
    }


def write_file(file_index: int, rows: list[dict[str, object]]) -> None:
    headers = SCHEMAS[file_index]
    OUTPUT.mkdir(parents=True, exist_ok=True)
    suffix = ".csv" if file_index in {0, 2, 4} else ".xlsx"
    path = OUTPUT / f"demo_{file_index + 1}{suffix}"
    if suffix == ".csv":
        with path.open("w", newline="", encoding="utf-8") as stream:
            writer = csv.writer(stream)
            writer.writerow(headers)
            for row in rows:
                writer.writerow([row[key] for key in ("date", "product", "category", "quantity", "unit_price", "amount", "store")])
        return

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sales"
    sheet.append(headers)
    for row in rows:
        sheet.append([row[key] for key in ("date", "product", "category", "quantity", "unit_price", "amount", "store")])
    if file_index == 3:
        unrelated = workbook.create_sheet("Instructions")
        unrelated.append(["This sheet is intentionally unrelated to sales data."])
    workbook.save(path)


def main() -> None:
    rng = random.Random(SEED)
    start = date(2025, 2, 1)
    accepted: list[dict[str, object]] = []
    serial = 0
    for offset in range(30):
        day = start + timedelta(days=offset)
        weekend = day.weekday() >= 5
        # A mid-month sale adds demand on Feb 15; weekends are busier than weekdays.
        if offset == 14:
            order_count = rng.randint(66, 70)
        elif weekend:
            order_count = rng.randint(60, 70)
        else:
            order_count = rng.randint(40, 50)
        for _ in range(order_count):
            serial += 1
            accepted.append(make_order(rng, day, serial, promo=(offset == 14)))

    # 60 rejects among roughly 1,570 source rows (about 3.8%): balanced issue types.
    rejects: list[dict[str, object]] = []
    for index in range(12):
        rejects.append({**make_order(rng, start + timedelta(days=index), serial + index), "date": ""})
        rejects.append({**make_order(rng, start + timedelta(days=index + 5), serial + 20 + index), "date": "not-a-date"})
        rejects.append({**make_order(rng, start + timedelta(days=index + 10), serial + 40 + index), "quantity": -1})
        mismatch = make_order(rng, start + timedelta(days=index + 15), serial + 60 + index)
        # Change the extended amount by one cent while keeping a valid currency format.
        wrong_cents = int(Decimal(str(mismatch["amount"]).replace("$", "").replace("£", "").replace(",", "")) * 100) + 1
        mismatch["amount"] = money_text(wrong_cents, index)
        rejects.append(mismatch)
        rejects.append(dict(accepted[rng.randrange(len(accepted))]))
    all_rows = accepted + rejects
    rng.shuffle(all_rows)

    for old_file in OUTPUT.iterdir() if OUTPUT.exists() else ():
        if old_file.is_file() and old_file.suffix.lower() in {".csv", ".xlsx"}:
            old_file.unlink()
    partitions: list[list[dict[str, object]]] = [[] for _ in range(5)]
    for index, row in enumerate(all_rows):
        partitions[index % len(partitions)].append(row)
    for index, rows in enumerate(partitions):
        write_file(index, rows)
    print(f"Generated 5 fictional files with {len(all_rows)} data rows ({len(accepted)} valid, {len(rejects)} intentionally invalid) in {OUTPUT}")


if __name__ == "__main__":
    main()
