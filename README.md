Sample project for demonstration only.

# Daily Sales Report Automation

This sample combines fictional CSV and Excel sales exports, maps inconsistent column names, normalizes and validates each row, records rejected rows with a reason, and creates an Excel and Markdown daily report.

The reproducible demo inputs contain 1,637 rows across five CSV/XLSX files, 30 days, six categories, and four fictional stores. Regenerate them with `python scripts/make_demo_data.py` (fixed seed: `20261110`).

## Architecture

```text
CSV / XLSX files
       |
       v
Header mapping (config/columns.yml)
       |
       v
Normalize + validate ---> rejects.csv (source file, row, reason)
       |
       v
Deduplicate + aggregate
       |
       +----> report.xlsx (daily summary + daily/category revenue)
       +----> report.md
       +----> run.log
```

## Sample report preview

The fixed-seed demo has **1,577 accepted orders** across 30 days, four fictional stores, and six categories. It creates busier weekends, different store volumes, varied transaction values, and a mid-month promotion on February 15. Each day has 40–70 accepted orders. The input contains 1,637 rows; 60 are rejected (3.7%).

The generated daily summary is shown below. The workbook also contains a daily/category breakdown; see [`output/report.md`](output/report.md) and [`output/report.xlsx`](output/report.xlsx).

## Daily summary

| Date | Orders | Units | Revenue |
|---|---:|---:|---:|
| 2025-02-01 | 64 | 175 | $2,790.50 |
| 2025-02-02 | 70 | 202 | $3,871.68 |
| 2025-02-03 | 46 | 142 | $3,145.26 |
| 2025-02-04 | 47 | 140 | $1,837.46 |
| 2025-02-05 | 50 | 160 | $2,724.47 |
| 2025-02-06 | 43 | 131 | $2,727.33 |
| 2025-02-07 | 46 | 155 | $2,658.60 |
| 2025-02-08 | 65 | 178 | $3,383.20 |
| 2025-02-09 | 65 | 203 | $3,712.84 |
| 2025-02-10 | 40 | 120 | $2,096.18 |
| 2025-02-11 | 41 | 123 | $2,675.96 |
| 2025-02-12 | 45 | 126 | $2,911.47 |
| 2025-02-13 | 47 | 148 | $2,284.79 |
| 2025-02-14 | 40 | 107 | $2,594.10 |
| 2025-02-15 | 70 | 188 | $3,459.09 |
| 2025-02-16 | 67 | 187 | $3,551.11 |
| 2025-02-17 | 49 | 118 | $2,463.71 |
| 2025-02-18 | 45 | 154 | $3,057.65 |
| 2025-02-19 | 48 | 154 | $2,798.88 |
| 2025-02-20 | 41 | 134 | $2,089.32 |
| 2025-02-21 | 48 | 153 | $2,206.91 |
| 2025-02-22 | 65 | 188 | $4,098.21 |
| 2025-02-23 | 66 | 179 | $3,652.81 |
| 2025-02-24 | 41 | 122 | $2,576.35 |
| 2025-02-25 | 50 | 156 | $2,995.03 |
| 2025-02-26 | 48 | 126 | $1,962.60 |
| 2025-02-27 | 46 | 139 | $3,033.00 |
| 2025-02-28 | 45 | 130 | $2,281.54 |
| 2025-03-01 | 69 | 211 | $3,877.52 |
| 2025-03-02 | 70 | 199 | $3,038.72 |

Total revenue: **$86,556.29**

## Revenue by day and category

| Date | Category | Revenue |
|---|---|---:|
| 2025-02-01 | Books | $407.84 |
| 2025-02-01 | Electronics | $417.31 |
| 2025-02-01 | Grocery | $187.26 |
| 2025-02-01 | Home | $372.67 |
| 2025-02-01 | Office | $299.49 |
| 2025-02-01 | Outdoor | $1,105.93 |
| 2025-02-02 | Books | $389.09 |
| 2025-02-02 | Electronics | $1,910.97 |
| 2025-02-02 | Grocery | $225.92 |
| 2025-02-02 | Home | $789.11 |
| 2025-02-02 | Office | $156.66 |
| 2025-02-02 | Outdoor | $399.93 |
| 2025-02-03 | Books | $328.96 |
| 2025-02-03 | Electronics | $1,191.13 |
| 2025-02-03 | Grocery | $138.69 |
| 2025-02-03 | Home | $802.43 |
| 2025-02-03 | Office | $66.50 |
| 2025-02-03 | Outdoor | $617.55 |
| 2025-02-04 | Books | $186.86 |
| 2025-02-04 | Electronics | $824.32 |
| 2025-02-04 | Grocery | $174.41 |
| 2025-02-04 | Home | $296.06 |
| 2025-02-04 | Office | $131.01 |
| 2025-02-04 | Outdoor | $224.80 |
| 2025-02-05 | Books | $183.14 |
| 2025-02-05 | Electronics | $888.95 |
| 2025-02-05 | Grocery | $117.45 |
| 2025-02-05 | Home | $566.12 |
| 2025-02-05 | Office | $180.53 |
| 2025-02-05 | Outdoor | $788.28 |
| 2025-02-06 | Books | $243.38 |
| 2025-02-06 | Electronics | $721.96 |
| 2025-02-06 | Grocery | $118.74 |
| 2025-02-06 | Home | $683.11 |
| 2025-02-06 | Office | $109.99 |
| 2025-02-06 | Outdoor | $850.15 |
| 2025-02-07 | Books | $189.89 |
| 2025-02-07 | Electronics | $846.53 |
| 2025-02-07 | Grocery | $158.87 |
| 2025-02-07 | Home | $421.66 |
| 2025-02-07 | Office | $400.09 |
| 2025-02-07 | Outdoor | $641.56 |
| 2025-02-08 | Books | $247.64 |
| 2025-02-08 | Electronics | $734.86 |
| 2025-02-08 | Grocery | $330.52 |
| 2025-02-08 | Home | $786.15 |
| 2025-02-08 | Office | $243.87 |
| 2025-02-08 | Outdoor | $1,040.16 |
| 2025-02-09 | Books | $377.57 |
| 2025-02-09 | Electronics | $1,252.61 |
| 2025-02-09 | Grocery | $192.92 |
| 2025-02-09 | Home | $445.24 |
| 2025-02-09 | Office | $374.21 |
| 2025-02-09 | Outdoor | $1,070.29 |
| 2025-02-10 | Books | $235.80 |
| 2025-02-10 | Electronics | $755.41 |
| 2025-02-10 | Grocery | $129.85 |
| 2025-02-10 | Home | $247.64 |
| 2025-02-10 | Office | $334.83 |
| 2025-02-10 | Outdoor | $392.65 |
| 2025-02-11 | Books | $95.22 |
| 2025-02-11 | Electronics | $1,241.48 |
| 2025-02-11 | Grocery | $96.32 |
| 2025-02-11 | Home | $619.51 |
| 2025-02-11 | Office | $255.45 |
| 2025-02-11 | Outdoor | $367.98 |
| 2025-02-12 | Books | $142.94 |
| 2025-02-12 | Electronics | $1,903.62 |
| 2025-02-12 | Grocery | $90.15 |
| 2025-02-12 | Home | $435.51 |
| 2025-02-12 | Office | $263.24 |
| 2025-02-12 | Outdoor | $76.01 |
| 2025-02-13 | Books | $337.49 |
| 2025-02-13 | Electronics | $502.50 |
| 2025-02-13 | Grocery | $141.33 |
| 2025-02-13 | Home | $668.33 |
| 2025-02-13 | Office | $300.05 |
| 2025-02-13 | Outdoor | $335.09 |
| 2025-02-14 | Books | $131.22 |
| 2025-02-14 | Electronics | $1,182.75 |
| 2025-02-14 | Grocery | $73.41 |
| 2025-02-14 | Home | $266.62 |
| 2025-02-14 | Office | $253.10 |
| 2025-02-14 | Outdoor | $687.00 |
| 2025-02-15 | Books | $325.56 |
| 2025-02-15 | Electronics | $590.75 |
| 2025-02-15 | Grocery | $184.16 |
| 2025-02-15 | Home | $416.14 |
| 2025-02-15 | Office | $678.89 |
| 2025-02-15 | Outdoor | $1,263.59 |
| 2025-02-16 | Books | $309.79 |
| 2025-02-16 | Electronics | $1,061.25 |
| 2025-02-16 | Grocery | $107.60 |
| 2025-02-16 | Home | $548.68 |
| 2025-02-16 | Office | $260.01 |
| 2025-02-16 | Outdoor | $1,263.78 |
| 2025-02-17 | Books | $292.19 |
| 2025-02-17 | Electronics | $649.27 |
| 2025-02-17 | Grocery | $109.86 |
| 2025-02-17 | Home | $503.70 |
| 2025-02-17 | Office | $550.50 |
| 2025-02-17 | Outdoor | $358.19 |
| 2025-02-18 | Books | $135.21 |
| 2025-02-18 | Electronics | $1,188.67 |
| 2025-02-18 | Grocery | $161.63 |
| 2025-02-18 | Home | $719.70 |
| 2025-02-18 | Office | $233.44 |
| 2025-02-18 | Outdoor | $619.00 |
| 2025-02-19 | Books | $566.77 |
| 2025-02-19 | Electronics | $298.45 |
| 2025-02-19 | Grocery | $87.69 |
| 2025-02-19 | Home | $426.66 |
| 2025-02-19 | Office | $359.42 |
| 2025-02-19 | Outdoor | $1,059.89 |
| 2025-02-20 | Books | $207.40 |
| 2025-02-20 | Electronics | $662.49 |
| 2025-02-20 | Grocery | $97.30 |
| 2025-02-20 | Home | $108.33 |
| 2025-02-20 | Office | $335.42 |
| 2025-02-20 | Outdoor | $678.38 |
| 2025-02-21 | Books | $244.80 |
| 2025-02-21 | Electronics | $804.70 |
| 2025-02-21 | Grocery | $70.60 |
| 2025-02-21 | Home | $89.35 |
| 2025-02-21 | Office | $270.23 |
| 2025-02-21 | Outdoor | $727.23 |
| 2025-02-22 | Books | $31.89 |
| 2025-02-22 | Electronics | $1,668.73 |
| 2025-02-22 | Grocery | $220.08 |
| 2025-02-22 | Home | $1,129.79 |
| 2025-02-22 | Office | $196.52 |
| 2025-02-22 | Outdoor | $851.20 |
| 2025-02-23 | Books | $319.82 |
| 2025-02-23 | Electronics | $873.32 |
| 2025-02-23 | Grocery | $452.89 |
| 2025-02-23 | Home | $521.27 |
| 2025-02-23 | Office | $231.91 |
| 2025-02-23 | Outdoor | $1,253.60 |
| 2025-02-24 | Books | $145.53 |
| 2025-02-24 | Electronics | $1,353.79 |
| 2025-02-24 | Grocery | $120.40 |
| 2025-02-24 | Home | $255.67 |
| 2025-02-24 | Office | $150.19 |
| 2025-02-24 | Outdoor | $550.77 |
| 2025-02-25 | Books | $102.76 |
| 2025-02-25 | Electronics | $665.05 |
| 2025-02-25 | Grocery | $337.52 |
| 2025-02-25 | Home | $564.31 |
| 2025-02-25 | Office | $393.40 |
| 2025-02-25 | Outdoor | $931.99 |
| 2025-02-26 | Books | $149.55 |
| 2025-02-26 | Electronics | $721.52 |
| 2025-02-26 | Grocery | $120.78 |
| 2025-02-26 | Home | $339.34 |
| 2025-02-26 | Office | $197.08 |
| 2025-02-26 | Outdoor | $434.33 |
| 2025-02-27 | Books | $272.75 |
| 2025-02-27 | Electronics | $1,010.40 |
| 2025-02-27 | Grocery | $215.21 |
| 2025-02-27 | Home | $82.39 |
| 2025-02-27 | Office | $89.38 |
| 2025-02-27 | Outdoor | $1,362.87 |
| 2025-02-28 | Books | $263.94 |
| 2025-02-28 | Electronics | $329.54 |
| 2025-02-28 | Grocery | $140.67 |
| 2025-02-28 | Home | $528.61 |
| 2025-02-28 | Office | $657.40 |
| 2025-02-28 | Outdoor | $361.38 |
| 2025-03-01 | Books | $124.48 |
| 2025-03-01 | Electronics | $1,225.86 |
| 2025-03-01 | Grocery | $264.01 |
| 2025-03-01 | Home | $612.28 |
| 2025-03-01 | Office | $816.55 |
| 2025-03-01 | Outdoor | $834.34 |
| 2025-03-02 | Books | $103.02 |
| 2025-03-02 | Electronics | $1,040.62 |
| 2025-03-02 | Grocery | $292.28 |
| 2025-03-02 | Home | $497.47 |
| 2025-03-02 | Office | $376.90 |
| 2025-03-02 | Outdoor | $728.43 |

## Rejections by reason

| Reason | Rows |
|---|---:|
| Missing required date | 12 |
| Invalid date | 12 |
| Non-positive or non-whole quantity | 12 |
| Amount does not match quantity × unit price | 12 |
| Duplicate row | 12 |


## Run

Requires Python 3. Install the pinned dependencies and run the sample:

```sh
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python -m sample.run --input data/input --output output
```

The optional `--config path/to/columns.yml` argument selects another YAML column mapping. Dates are parsed using a small explicit list of common formats. The first worksheet in each workbook is treated as the data sheet; other worksheets are ignored.

## How to adapt to your files

Edit `config/columns.yml`. Under `columns`, each key is a canonical field used by the program (`date`, `product`, `category`, `quantity`, `unit_price`, or `amount`); add the exact header names from your exports to that field's alias list. Header matching ignores capitalization and extra whitespace. For example, if your amount column is named `Line Total`, add it to `amount`:

```yaml
columns:
  amount: [amount, total, sales amount, revenue, line total]
```

Keep each source header mapped to only one canonical field. The `required` list controls which canonical columns must be present. The program currently reads the first worksheet in an XLSX file; change the input workbook layout or extend `read_records` if the sales data is on another sheet. Keep dates within the supported formats listed in `sample/run.py` and verify ambiguous numeric dates against the source's locale.

### Schedule a daily run

On macOS or Linux, create a cron entry with absolute paths, for example:

```cron
0 7 * * * cd /absolute/path/to/project && /absolute/path/to/project/.venv/bin/python -m sample.run --input /absolute/path/to/project/data/input --output /absolute/path/to/project/output >> /absolute/path/to/project/output/cron.log 2>&1
```

On Windows, create a **Basic Task** in Task Scheduler, set a daily trigger, and use the project's `.venv\Scripts\python.exe` as the program. Set `-m sample.run --input C:\path\to\project\data\input --output C:\path\to\project\output` as arguments and the project directory as **Start in**. Ensure upstream exports arrive before the scheduled time.

## Test

```sh
.venv/bin/pytest -q
```

Tests cover header matching, supported and invalid dates, currency and grouping separators, missing or malformed amounts, valid and invalid quantities, missing required fields, amount consistency, duplicate rejection, report contents, and fixed end-to-end totals.

## Limitations

- This is a local demonstration, not a hosted service or production scheduler.
- Input supports CSV and XLSX only. The first worksheet is used; formulas are read from cached values and may be empty if the workbook has not been recalculated.
- A row is considered a duplicate when all normalized sales fields match. Distinct legitimate transactions with identical values cannot be distinguished without a transaction ID.
- Ambiguous numeric dates use month/day/year before day/month/year. Configure source exports consistently to avoid ambiguity.
- Amount must equal quantity multiplied by unit price exactly to the cent; discounts, tax, shipping, refunds, and exchange rates are not modeled.
- Files are processed from one local directory. The sample does not manage secure storage, access control, backups, or concurrent runs.
