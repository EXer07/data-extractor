# Power BI, API, and Excel Data Extraction Handoff

Use this file to continue the project on another computer. The objective is to
extract public agricultural datasets into verifiable Excel files without
inventing, estimating, or silently dropping records.

## Non-negotiable data rules

1. Use only records returned by the specified Power BI report, API, website, or
   source workbook.
2. Do not create zero-filled rows for provinces or periods absent from the
   source.
3. Do not infer undocumented units. Record the unit as `ไม่ระบุใน API` when the
   source does not define it.
4. Keep source files unchanged. Save transformations as new files.
5. Preserve original fields in raw-data sheets. Put descriptions or units in a
   separate metadata sheet.
6. Validate row counts, columns, duplicate rows, date ranges, and file integrity
   before delivery.
7. Never commit API keys, secrets, `.env`, local configuration, or generated
   datasets.

## Setup on another computer

The upstream scraper requires Python 3.11.

```bash
git clone https://github.com/holstt/powerbi-table-scraper.git
cd powerbi-table-scraper
python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Copy the project scripts from this repository after cloning, or clone this
working repository if it has been pushed.

Create a local `config.yml` from `config.example.yml`. The file is ignored by
Git.

## Power BI extraction

### Report 1: provincial rice prices

URL:

```text
https://app.powerbi.com/view?r=eyJrIjoiODgwMzM4NWUtMDcyOC00ZWNhLWE5YWMtOWNmNjAxOTBlNmJmIiwidCI6IjFiMTZiOWU4LWNiZjgtNGRjNi1hODY3LTJlNDIzNDJiM2Y4NiIsImMiOjEwfQ%3D%3D
```

Known semantic tables:

- `Rice (2)`: rice area and production data
- `Sheet1`: rice market-event schedule
- `price start 2023`: raw price observations for 2023–2026, despite the table
  name

Run:

```bash
python extract_all.py
```

Default output:

```text
powerbi_all_data.xlsx
```

Expected counts at the time of the last extraction:

| Table | Rows |
|---|---:|
| Rice (2) | 23,555 |
| Sheet1 | 10 |
| price start 2023 | 116,604 |

The Power BI public query API has a 30,000-row response limit. `extract_all.py`
splits large price queries by Buddhist year and month, then combines the
partitions. Do not accept exactly 30,000 rows as complete without checking.

### Report 2: satellite rice-supply forecast

URL:

```text
https://app.powerbi.com/view?r=eyJrIjoiYTgxMmRkNzYtOGMyNC00ZjM3LWFjZjctZDU5MWM1OGExZjNiIiwidCI6ImQyM2EzNDJjLTBlYWYtNDg2ZS1hMjgxLWJhZWQ4ZWY3NTFjMyIsImMiOjEwfQ%3D%3D
```

Known semantic tables and last observed counts:

| Table | Rows |
|---|---:|
| supply_latest_GISTDA | 25,464 |
| location | 855 |
| time | 23 |
| ricemill_capacity | 845 |
| ricetype | 7 |
| param_selected_data_source | 3 |

Example programmatic call:

```python
from pathlib import Path
from extract_all import extract_all

counts = extract_all(
    POWER_BI_URL,
    Path("powerbi_2_all_data.xlsx").resolve(),
)
print(counts)
```

Use `rename_powerbi_2_excel.py` to create the reader-facing workbook:

```text
คาดการณ์อุปทานข้าวจากภาพถ่ายดาวเทียม.xlsx
```

Reader-facing sheet names:

- `สรุปชุดข้อมูล`
- `ผลผลิตข้าวรายพื้นที่`
- `ข้อมูลจังหวัดและอำเภอ`
- `งวดเวลาและฤดูเพาะปลูก`
- `กำลังการผลิตโรงสี`
- `ประเภทข้าว`
- `แหล่งข้อมูล`

## Reader-facing Excel outputs for Report 1

Run:

```bash
python create_three_excel_files.py
```

It creates three files in `excel_split_corrected/`:

- `rice_2.xlsx`: sheet `เนื้อที่และผลผลิตข้าว`
- `sheet1.xlsx`: sheet `กำหนดการตลาดนัดข้าว`
- `by_province.xlsx`: one price sheet per province, named
  `ราคาข้าว_<จังหวัด>`

`by_province.xlsx` must contain only records from `price start 2023`. Do not mix
production or market-schedule rows into this workbook. Each province sheet keeps
the original `Year` and `YearBE` columns.

The older `split_excel_by_province.py` creates separate files per province and
does not match the final requirement. Prefer `create_three_excel_files.py`.

## RAG preparation

Run:

```bash
python build_rag_dataset.py
```

The script creates `rag_data/` containing:

- cleaned raw Parquet tables
- aggregate Parquet tables
- `database/rice_rag.duckdb` for numerical questions
- JSONL documents for vector retrieval
- schema, metric definitions, data-quality report, routing configuration, SQL
  examples, and a sample chatbot system prompt

Use hybrid retrieval:

- SQL/DuckDB for totals, averages, filtering, rankings, comparisons, and trends
- vector search for definitions, descriptions, and narrative context

Never calculate final numerical answers from vector-search snippets.

## DIT rice-mill API

The key-management screen displays `/ricetrade/mill`, but the working endpoint
is:

```text
GET https://api.dit.go.th/api/ricetrade/mill
```

Authentication:

```text
X-API-Key: <key>
```

Required query parameters:

- `startdate=YYYY-MM-DD`
- `enddate=YYYY-MM-DD`
- `limit`: 1–5,000
- `offset`: pagination offset

Store the key outside source control:

```bash
export DIT_API_KEY="..."
```

Never write the key into `SKILLS.md`, source code, logs, Git history, or the
generated workbook.

OpenAPI specification:

```text
https://api.dit.go.th/api/openapi.json
```

Last verified API scope:

- no records before 2017
- data from April 2017 through June 2026
- 7,215 rows
- 65 provinces
- 111 months per returned province
- no duplicate rows

The API returned no records for these 12 provinces:

- กระบี่
- ชุมพร
- นราธิวาส
- บึงกาฬ
- ปัตตานี
- พังงา
- ภูเก็ต
- ยะลา
- ระนอง
- สตูล
- สมุทรสงคราม
- สุราษฎร์ธานี

Do not add these provinces as zero-valued rows.

API fields with explicit units:

| Field | Unit |
|---|---|
| กำลังการผลิต_ตันต่อวัน | ตันต่อวัน |
| ข้าวเปลือก_ตัน | ตัน |
| ข้าวสาร_ตัน | ตัน |
| ปลายข้าว_ตัน | ตัน |

The API schema does not explicitly declare units for the small, medium, and
large mill-count fields. Label their unit as undocumented rather than assuming.

Latest generated workbook:

```text
ข้อมูลโรงสีข้าวทั้งหมด_2017-2026.xlsx
```

It contains:

- `ข้อมูลจาก API`: the unchanged API records
- `คำอธิบายคอลัมน์`: field meanings, documented units, and unit caveats

## Filtering a supplied Excel workbook

For `gistda_rice_20260430.xlsx`, the province field is `p_name`. The exact value
for Bangkok is:

```text
กรุงเทพมหานคร
```

The last filtered output contained 505 rows and retained the six original
columns. Always filter with an exact match and compare the result back to the
source rows.

## Power BI scraper compatibility fixes

The upstream project required these local fixes:

1. Remove the unused `click.Option` import from `src/config.py`.
2. Guard Windows-only `subprocess.CREATE_NO_WINDOW` in
   `src/scraper/driver.py`.
3. Do not switch into Power BI sandbox iframes when `.tableEx` already exists in
   the main document.

Without the iframe guard, the scraper enters the first visual sandbox and cannot
find the report table.

## Validation checklist

Before delivering any Excel file:

1. Confirm every API page or Power BI partition was fetched.
2. Reconcile source totals with exported row counts.
3. Confirm column names and order.
4. Check duplicate full rows and duplicate IDs.
5. Check minimum and maximum dates.
6. Confirm missing provinces are genuinely absent from the source.
7. Read the completed workbook back and compare it with the source data.
8. Run:

   ```bash
   unzip -t output.xlsx
   ```

9. Report the source URL, date range, row count, column count, and any source
   limitations to the user.

## Generated files and Git

The repository `.gitignore` excludes:

- `.venv/`, Python caches, and IDE files
- local config and environment files
- Excel, CSV, Parquet, ZIP, DuckDB, and SQLite outputs
- `rag_data/`, `excel_by_province/`, and `excel_split_corrected/`

Commit reusable scripts and documentation, not generated datasets or secrets.

## Unfinished website request

`https://tradereport.moc.go.th/en` is a portal containing many independent
trade datasets rather than one dataset. Before extraction, require the user to
specify:

- export, import, re-export, balance, dashboard, or an API dataset
- country/product/HS-code dimension
- monthly or yearly granularity
- date range
- desired currency or quantity fields

Do not claim that the entire portal has been extracted without an explicit
scope.
