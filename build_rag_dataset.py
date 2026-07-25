"""Prepare the extracted Power BI workbook for a hybrid SQL + vector RAG system."""

from __future__ import annotations

import hashlib
import json
import logging
import shutil
from pathlib import Path
from typing import Any, Iterable

import duckdb
import pandas as pd

logger = logging.getLogger(__name__)

SOURCE_WORKBOOK = Path("powerbi_all_data.xlsx")
OUTPUT_DIR = Path("rag_data")

PRICE_COLUMNS = {
    "price start 2023.total": "source_total",
    "price start 2023.startdate": "source_start_date",
    "price start 2023.enddate": "source_end_date",
    "price start 2023.data.ProvinceID": "province_id",
    "price start 2023.data.ProvinceName": "province",
    "price start 2023.data.SurveyDate": "survey_timestamp_ms",
    "price start 2023.data.ProjectTypeName": "project_type",
    "price start 2023.data.PaddyName": "paddy_name",
    "price start 2023.data.MinPrice": "min_price",
    "price start 2023.data.MaxPrice": "max_price",
    "price start 2023.AvgPrice": "avg_price",
    "price start 2023.type": "rice_category",
    "price start 2023.Month": "month",
    "price start 2023.Year": "gregorian_year",
    "price start 2023.YearBE": "buddhist_year",
}

PRODUCTION_COLUMNS = {
    "Rice (2).ประเทศ/ภาค/จังหวัด": "province",
    "Rice (2).ชนิด": "rice_type",
    "Rice (2).เนื้อที่": "planted_area",
    "Rice (2).เนื้อที่เก็บเกี่ยว": "harvested_area",
    "Rice (2).ผลผลิต": "production",
    "Rice (2).ปี": "buddhist_year",
    "Rice (2).ภาค": "region",
    "Rice (2).ข้าวนาปี/นาปรัง": "crop_season",
    "Rice (2).อำเภอ": "district",
    "Rice (2).ตำบล": "subdistrict",
    "Rice (2).owner": "data_owner",
}

MARKET_COLUMNS = {
    "Sheet1.ที่": "schedule_id",
    "Sheet1.สถานที่": "location",
    "Sheet1.อำเภอ": "district",
    "Sheet1.จังหวัด": "province",
    "Sheet1.ช่วงเวลา": "period_text",
    "Sheet1.ประเภทโครงการ": "project_type",
    "Sheet1.รายละเอียด": "details",
}


def _stable_ids(frame: pd.DataFrame, prefix: str) -> pd.Series:
    def digest(row: pd.Series) -> str:
        raw = "|".join("" if pd.isna(value) else str(value) for value in row)
        return f"{prefix}_{hashlib.sha256(raw.encode('utf-8')).hexdigest()[:20]}"

    return frame.astype(object).apply(digest, axis=1)


def _write_json(path: Path, value: Any) -> None:
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )


def _write_jsonl(path: Path, documents: Iterable[dict[str, Any]]) -> int:
    count = 0
    with path.open("w", encoding="utf-8") as output:
        for document in documents:
            output.write(json.dumps(document, ensure_ascii=False, default=str))
            output.write("\n")
            count += 1
    return count


def _clean_text_columns(frame: pd.DataFrame) -> pd.DataFrame:
    for column in frame.select_dtypes(include="object").columns:
        frame[column] = frame[column].map(
            lambda value: value.strip() if isinstance(value, str) else value
        )
    return frame


def load_and_clean() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    prices = pd.read_excel(SOURCE_WORKBOOK, sheet_name="price start 2023").rename(
        columns=PRICE_COLUMNS
    )
    production = pd.read_excel(SOURCE_WORKBOOK, sheet_name="Rice (2)").rename(
        columns=PRODUCTION_COLUMNS
    )
    markets = pd.read_excel(SOURCE_WORKBOOK, sheet_name="Sheet1").rename(
        columns=MARKET_COLUMNS
    )

    prices = _clean_text_columns(prices)
    production = _clean_text_columns(production)
    markets = _clean_text_columns(markets)

    prices["survey_date"] = pd.to_datetime(
        prices["survey_timestamp_ms"], unit="ms", utc=True
    ).dt.tz_localize(None)
    prices["survey_date_text"] = prices["survey_date"].dt.strftime("%Y-%m-%d")
    prices["record_id"] = _stable_ids(
        prices.drop(columns=["record_id"], errors="ignore"), "price"
    )

    production["record_id"] = _stable_ids(
        production.drop(columns=["record_id"], errors="ignore"), "production"
    )
    markets["record_id"] = markets["schedule_id"].map(
        lambda value: f"market_{int(value):04d}"
    )

    price_order = [
        "record_id",
        "survey_date",
        "survey_date_text",
        "province_id",
        "province",
        "project_type",
        "paddy_name",
        "rice_category",
        "min_price",
        "max_price",
        "avg_price",
        "month",
        "gregorian_year",
        "buddhist_year",
        "source_total",
        "source_start_date",
        "source_end_date",
        "survey_timestamp_ms",
    ]
    production_order = [
        "record_id",
        "province",
        "region",
        "district",
        "subdistrict",
        "rice_type",
        "crop_season",
        "buddhist_year",
        "planted_area",
        "harvested_area",
        "production",
        "data_owner",
    ]
    market_order = [
        "record_id",
        "schedule_id",
        "project_type",
        "province",
        "district",
        "location",
        "period_text",
        "details",
    ]
    return prices[price_order], production[production_order], markets[market_order]


def build_aggregates(
    prices: pd.DataFrame, production: pd.DataFrame
) -> dict[str, pd.DataFrame]:
    price_metrics = {
        "record_count": ("record_id", "count"),
        "avg_price": ("avg_price", "mean"),
        "lowest_price": ("min_price", "min"),
        "highest_price": ("max_price", "max"),
        "first_survey_date": ("survey_date", "min"),
        "last_survey_date": ("survey_date", "max"),
    }
    price_by_province_month = (
        prices.groupby(
            [
                "buddhist_year",
                "month",
                "province",
                "rice_category",
                "paddy_name",
            ],
            dropna=False,
        )
        .agg(**price_metrics)
        .reset_index()
    )
    price_by_province_year = (
        prices.groupby(["buddhist_year", "province", "rice_category"], dropna=False)
        .agg(**price_metrics)
        .reset_index()
    )
    price_by_rice_type_year = (
        prices.groupby(["buddhist_year", "rice_category", "paddy_name"], dropna=False)
        .agg(**price_metrics)
        .reset_index()
    )

    production_by_region_year = (
        production.groupby(
            ["buddhist_year", "region", "rice_type", "crop_season"],
            dropna=False,
        )
        .agg(
            record_count=("record_id", "count"),
            planted_area=("planted_area", "sum"),
            harvested_area=("harvested_area", "sum"),
            production=("production", "sum"),
            province_count=("province", "nunique"),
        )
        .reset_index()
    )
    production_by_province_year = (
        production.groupby(
            ["buddhist_year", "province", "rice_type", "crop_season"],
            dropna=False,
        )
        .agg(
            record_count=("record_id", "count"),
            planted_area=("planted_area", "sum"),
            harvested_area=("harvested_area", "sum"),
            production=("production", "sum"),
        )
        .reset_index()
    )

    return {
        "price_by_province_month": price_by_province_month,
        "price_by_province_year": price_by_province_year,
        "price_by_rice_type_year": price_by_rice_type_year,
        "production_by_region_year": production_by_region_year,
        "production_by_province_year": production_by_province_year,
    }


def build_documents(
    aggregates: dict[str, pd.DataFrame], markets: pd.DataFrame
) -> dict[str, list[dict[str, Any]]]:
    price_documents = []
    for row in aggregates["price_by_province_year"].itertuples(index=False):
        doc_id = hashlib.sha256(
            f"{row.buddhist_year}|{row.province}|{row.rice_category}".encode()
        ).hexdigest()[:20]
        text = (
            f"สรุปราคาข้าว จังหวัด{row.province} ปี พ.ศ. {row.buddhist_year} "
            f"หมวด {row.rice_category}: ราคาเฉลี่ย {row.avg_price:,.2f} บาท "
            f"ราคาต่ำสุด {row.lowest_price:,.0f} บาท "
            f"ราคาสูงสุด {row.highest_price:,.0f} บาท "
            f"จาก {row.record_count:,} รายการสำรวจ "
            f"ช่วงวันที่ {row.first_survey_date:%Y-%m-%d} ถึง "
            f"{row.last_survey_date:%Y-%m-%d}."
        )
        price_documents.append(
            {
                "id": f"price_summary_{doc_id}",
                "document_type": "annual_province_price_summary",
                "text": text,
                "metadata": {
                    "source_table": "price_by_province_year",
                    "province": row.province,
                    "buddhist_year": int(row.buddhist_year),
                    "rice_category": row.rice_category,
                    "record_count": int(row.record_count),
                },
            }
        )

    production_documents = []
    for row in aggregates["production_by_region_year"].itertuples(index=False):
        doc_id = hashlib.sha256(
            f"{row.buddhist_year}|{row.region}|{row.rice_type}|{row.crop_season}".encode()
        ).hexdigest()[:20]
        text = (
            f"สรุปข้อมูลการผลิตข้าว {row.region} ปี พ.ศ. {row.buddhist_year} "
            f"ชนิด {row.rice_type} ฤดู {row.crop_season}: "
            f"เนื้อที่ {row.planted_area:,.0f}, "
            f"เนื้อที่เก็บเกี่ยว {row.harvested_area:,.0f}, "
            f"ผลผลิต {row.production:,.0f}, ครอบคลุม {row.province_count} จังหวัด. "
            "หน่วยของเนื้อที่และผลผลิตไม่ได้ระบุชัดเจนในแหล่งข้อมูลต้นทาง."
        )
        production_documents.append(
            {
                "id": f"production_summary_{doc_id}",
                "document_type": "annual_region_production_summary",
                "text": text,
                "metadata": {
                    "source_table": "production_by_region_year",
                    "region": row.region,
                    "buddhist_year": int(row.buddhist_year),
                    "rice_type": row.rice_type,
                    "crop_season": row.crop_season,
                },
            }
        )

    market_documents = []
    for row in markets.itertuples(index=False):
        market_documents.append(
            {
                "id": row.record_id,
                "document_type": "market_schedule",
                "text": (
                    f"กำหนดการ{row.project_type} จังหวัด{row.province} "
                    f"อำเภอ{row.district} ณ {row.location} ช่วงเวลา {row.period_text}. "
                    f"รายละเอียด: {row.details}."
                ),
                "metadata": {
                    "source_table": "market_schedule",
                    "schedule_id": int(row.schedule_id),
                    "province": row.province,
                    "district": row.district,
                    "project_type": row.project_type,
                    "period_text": row.period_text,
                },
            }
        )

    return {
        "annual_price_summaries": price_documents,
        "annual_production_summaries": production_documents,
        "market_schedule": market_documents,
    }


def build_metadata(
    prices: pd.DataFrame,
    production: pd.DataFrame,
    markets: pd.DataFrame,
    aggregates: dict[str, pd.DataFrame],
    document_counts: dict[str, int],
) -> tuple[dict[str, Any], dict[str, Any], list[dict[str, Any]]]:
    schema = {
        "dataset_name": "TPSO Rice Price and Production",
        "source": "Public Power BI report",
        "tables": {
            "rice_prices": {
                "description": "Raw provincial rice price observations.",
                "primary_key": "record_id (stable content hash; source has no explicit row key)",
                "row_count": len(prices),
                "columns": [
                    {"name": name, "dtype": str(dtype)}
                    for name, dtype in prices.dtypes.items()
                ],
            },
            "rice_production": {
                "description": "Rice area and production observations by location and crop type.",
                "primary_key": "record_id (stable content hash; source has no explicit row key)",
                "row_count": len(production),
                "columns": [
                    {"name": name, "dtype": str(dtype)}
                    for name, dtype in production.dtypes.items()
                ],
            },
            "market_schedule": {
                "description": "Rice market and absorption-event schedule.",
                "primary_key": "record_id",
                "row_count": len(markets),
                "columns": [
                    {"name": name, "dtype": str(dtype)}
                    for name, dtype in markets.dtypes.items()
                ],
            },
        },
        "aggregate_tables": {
            name: {"row_count": len(frame), "columns": list(frame.columns)}
            for name, frame in aggregates.items()
        },
        "vector_documents": document_counts,
    }
    metrics = {
        "avg_price": {
            "definition": "Arithmetic mean of raw avg_price observations in the selected group.",
            "sql": "AVG(avg_price)",
            "unit": "Baht as displayed by the source; price denomination is not explicitly documented.",
        },
        "lowest_price": {
            "definition": "Minimum raw min_price in the selected group.",
            "sql": "MIN(min_price)",
            "unit": "Baht as displayed by the source.",
        },
        "highest_price": {
            "definition": "Maximum raw max_price in the selected group.",
            "sql": "MAX(max_price)",
            "unit": "Baht as displayed by the source.",
        },
        "planted_area": {
            "definition": "Sum of the source planted-area field.",
            "sql": "SUM(planted_area)",
            "unit": "Not explicitly documented in the public source; verify before presenting a unit.",
        },
        "harvested_area": {
            "definition": "Sum of the source harvested-area field.",
            "sql": "SUM(harvested_area)",
            "unit": "Not explicitly documented in the public source; verify before presenting a unit.",
        },
        "production": {
            "definition": "Sum of the source production field.",
            "sql": "SUM(production)",
            "unit": "Not explicitly documented in the public source; verify before presenting a unit.",
        },
    }
    dictionary_documents = [
        {
            "id": "data_dictionary_overview",
            "document_type": "data_dictionary",
            "text": (
                "ชุดข้อมูลประกอบด้วย rice_prices สำหรับข้อมูลราคา, rice_production "
                "สำหรับพื้นที่และผลผลิต และ market_schedule สำหรับกำหนดการตลาด. "
                "คำถามที่ต้องคำนวณตัวเลขควรใช้ SQL กับ DuckDB; vector search "
                "เหมาะกับการค้นหาบริบทและเอกสารสรุป."
            ),
            "metadata": {"source_table": "metadata", "language": "th"},
        },
        {
            "id": "calendar_definition",
            "document_type": "data_dictionary",
            "text": (
                "buddhist_year คือปีพุทธศักราช และ gregorian_year คือปีคริสต์ศักราช. "
                "โดยทั่วไป พ.ศ. เท่ากับ ค.ศ. บวก 543 ในข้อมูลราคา."
            ),
            "metadata": {"source_table": "metadata", "language": "th"},
        },
        {
            "id": "price_metric_definition",
            "document_type": "metric_definition",
            "text": (
                "avg_price คือราคาเฉลี่ยที่มากับข้อมูลต้นทาง; lowest_price ใช้ "
                "MIN(min_price), highest_price ใช้ MAX(max_price). แหล่งข้อมูลแสดง "
                "หน่วยเป็นบาท แต่ไม่ได้ระบุ denomination อย่างชัดเจน."
            ),
            "metadata": {"source_table": "metadata", "language": "th"},
        },
    ]
    return schema, metrics, dictionary_documents


def build_quality_report(
    prices: pd.DataFrame, production: pd.DataFrame, markets: pd.DataFrame
) -> dict[str, Any]:
    return {
        "source_workbook": str(SOURCE_WORKBOOK.resolve()),
        "tables": {
            "rice_prices": {
                "rows": len(prices),
                "duplicate_full_rows_excluding_record_id": int(
                    prices.drop(columns=["record_id"]).duplicated().sum()
                ),
                "null_counts": prices.isna().sum().astype(int).to_dict(),
                "min_price_greater_than_max_price": int(
                    (prices["min_price"] > prices["max_price"]).sum()
                ),
                "avg_price_outside_min_max": int(
                    (
                        (prices["avg_price"] < prices["min_price"])
                        | (prices["avg_price"] > prices["max_price"])
                    ).sum()
                ),
                "year_calendar_mismatch": int(
                    (prices["buddhist_year"] != prices["gregorian_year"] + 543).sum()
                ),
                "date_range": {
                    "min": prices["survey_date"].min(),
                    "max": prices["survey_date"].max(),
                },
            },
            "rice_production": {
                "rows": len(production),
                "duplicate_full_rows_excluding_record_id": int(
                    production.drop(columns=["record_id"]).duplicated().sum()
                ),
                "null_counts": production.isna().sum().astype(int).to_dict(),
                "zero_production_rows": int((production["production"] == 0).sum()),
                "harvested_area_greater_than_planted_area": int(
                    (production["harvested_area"] > production["planted_area"]).sum()
                ),
            },
            "market_schedule": {
                "rows": len(markets),
                "duplicate_schedule_ids": int(
                    markets["schedule_id"].duplicated().sum()
                ),
                "null_counts": markets.isna().sum().astype(int).to_dict(),
                "period_parse_status": (
                    "period_text is preserved as Thai free text; no exact normalized "
                    "date is generated because several values are ranges or pending dates"
                ),
            },
        },
    }


def write_duckdb(
    path: Path,
    prices: pd.DataFrame,
    production: pd.DataFrame,
    markets: pd.DataFrame,
    aggregates: dict[str, pd.DataFrame],
) -> None:
    connection = duckdb.connect(str(path))
    try:
        all_tables = {
            "rice_prices": prices,
            "rice_production": production,
            "market_schedule": markets,
            **aggregates,
        }
        for name, frame in all_tables.items():
            connection.register("_frame", frame)
            connection.execute(
                f'CREATE OR REPLACE TABLE "{name}" AS SELECT * FROM _frame'
            )
            connection.unregister("_frame")
        connection.execute(
            """
            CREATE OR REPLACE VIEW latest_price_observations AS
            SELECT *
            FROM rice_prices
            WHERE survey_date = (SELECT MAX(survey_date) FROM rice_prices)
            """
        )
    finally:
        connection.close()


def write_readme(
    output_dir: Path,
    row_counts: dict[str, int],
    document_counts: dict[str, int],
) -> None:
    content = f"""# TPSO Rice RAG Dataset

This package is designed for a hybrid SQL + vector RAG chatbot.

## Contents

- `database/rice_rag.duckdb`: authoritative structured data for numeric questions.
- `raw/*.parquet`: cleaned source tables.
- `aggregates/*.parquet`: precomputed groupings for common questions.
- `documents/*.jsonl`: Thai retrieval documents ready for embedding.
- `metadata/schema.json`: table and column catalogue.
- `metadata/metric_definitions.json`: calculation rules and unit caveats.
- `metadata/data_quality_report.json`: validation results.
- `examples/sql_examples.sql`: sample SQL for tool-calling agents.

## Source row counts

- rice_prices: {row_counts["rice_prices"]:,}
- rice_production: {row_counts["rice_production"]:,}
- market_schedule: {row_counts["market_schedule"]:,}

## Vector document counts

{chr(10).join(f"- {name}: {count:,}" for name, count in document_counts.items())}

## Recommended routing

1. Use DuckDB for totals, averages, rankings, filters, comparisons, and trends.
2. Use vector search for definitions, market-event descriptions, and narrative summaries.
3. Return the SQL filters, source table, row count, and data date with numerical answers.
4. Do not state production/area units more specifically than the metadata permits.
"""
    (output_dir / "README.md").write_text(content, encoding="utf-8")


def write_sql_examples(path: Path) -> None:
    path.write_text(
        """-- Average price by province for a Buddhist year
SELECT province, ROUND(AVG(avg_price), 2) AS avg_price, COUNT(*) AS observations
FROM rice_prices
WHERE buddhist_year = 2568
GROUP BY province
ORDER BY avg_price DESC;

-- Monthly trend for one province and rice category
SELECT buddhist_year, month, ROUND(AVG(avg_price), 2) AS avg_price
FROM rice_prices
WHERE province = 'บุรีรัมย์' AND rice_category = 'ข้าวหอม'
GROUP BY buddhist_year, month
ORDER BY buddhist_year, month;

-- Production summary by region and year
SELECT buddhist_year, region, SUM(production) AS production
FROM rice_production
GROUP BY buddhist_year, region
ORDER BY buddhist_year, production DESC;

-- Market schedule by province
SELECT province, district, location, period_text, project_type, details
FROM market_schedule
WHERE province = 'บุรีรัมย์';
""",
        encoding="utf-8",
    )


def write_chatbot_assets(output_dir: Path) -> None:
    rag_config = {
        "strategy": "hybrid_sql_vector",
        "structured_store": {
            "engine": "duckdb",
            "path": "database/rice_rag.duckdb",
            "use_for": [
                "counts",
                "averages",
                "rankings",
                "comparisons",
                "filters",
                "time trends",
            ],
        },
        "vector_store": {
            "source": "documents/all_documents.jsonl",
            "id_field": "id",
            "text_field": "text",
            "metadata_field": "metadata",
            "recommended_metadata_filters": [
                "document_type",
                "province",
                "region",
                "buddhist_year",
                "rice_category",
                "rice_type",
                "project_type",
            ],
        },
        "routing": {
            "sql_signals_th": [
                "เท่าไร",
                "เฉลี่ย",
                "สูงสุด",
                "ต่ำสุด",
                "อันดับ",
                "เปรียบเทียบ",
                "แนวโน้ม",
                "กี่รายการ",
                "รวม",
            ],
            "vector_signals_th": [
                "คืออะไร",
                "หมายถึง",
                "อธิบาย",
                "กำหนดการ",
                "รายละเอียด",
                "บริบท",
            ],
        },
        "answer_contract": {
            "required_for_numeric_answers": [
                "source_table",
                "filters",
                "calculation",
                "rows_used",
                "data_date_or_period",
            ],
            "unit_rule": (
                "Do not infer a more specific area, production, or price "
                "denomination than metric_definitions.json documents."
            ),
        },
    }
    _write_json(output_dir / "metadata" / "rag_config.json", rag_config)
    (output_dir / "examples" / "chatbot_system_prompt.md").write_text(
        """# Example system prompt

คุณเป็นผู้ช่วยวิเคราะห์ข้อมูลราคาข้าว ผลผลิตข้าว และกำหนดการตลาดข้าว

กฎการตอบ:

1. ถ้าคำถามต้องคำนวณตัวเลข ให้สร้าง SQL แบบ read-only สำหรับ DuckDB ก่อนตอบ
2. ใช้ vector retrieval สำหรับคำอธิบาย นิยาม และกำหนดการ
3. ห้ามคำนวณยอดรวมจากข้อความที่ค้นด้วย vector search
4. แยกปี พ.ศ. (`buddhist_year`) และ ค.ศ. (`gregorian_year`) ให้ถูกต้อง
5. เมื่อรายงานตัวเลข ให้ระบุตาราง เงื่อนไข จำนวนแถว และช่วงวันที่ที่ใช้
6. หาก metadata ไม่ได้ยืนยันหน่วย ห้ามเดาหน่วยที่ละเอียดกว่าแหล่งข้อมูล
7. หากไม่มีข้อมูลตามเงื่อนไข ให้บอกว่าไม่พบข้อมูล ห้ามสร้างค่าประมาณเอง
8. SQL ต้องเป็น SELECT หรือ WITH...SELECT เท่านั้น และต้องกำหนด LIMIT
   เมื่อแสดงรายการรายละเอียด
""",
        encoding="utf-8",
    )


def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="[%(asctime)s] [%(levelname)s] %(message)s",
    )
    if not SOURCE_WORKBOOK.exists():
        raise FileNotFoundError(f"Source workbook not found: {SOURCE_WORKBOOK}")

    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    for subdirectory in [
        "raw",
        "aggregates",
        "documents",
        "metadata",
        "database",
        "examples",
    ]:
        (OUTPUT_DIR / subdirectory).mkdir(parents=True, exist_ok=True)

    logger.info("Loading and cleaning source workbook")
    prices, production, markets = load_and_clean()
    aggregates = build_aggregates(prices, production)

    raw_tables = {
        "rice_prices": prices,
        "rice_production": production,
        "market_schedule": markets,
    }
    for name, frame in raw_tables.items():
        frame.to_parquet(OUTPUT_DIR / "raw" / f"{name}.parquet", index=False)
    for name, frame in aggregates.items():
        frame.to_parquet(OUTPUT_DIR / "aggregates" / f"{name}.parquet", index=False)

    documents = build_documents(aggregates, markets)
    document_counts = {
        name: _write_jsonl(OUTPUT_DIR / "documents" / f"{name}.jsonl", values)
        for name, values in documents.items()
    }

    schema, metrics, dictionary_documents = build_metadata(
        prices, production, markets, aggregates, document_counts
    )
    document_counts["data_dictionary"] = _write_jsonl(
        OUTPUT_DIR / "documents" / "data_dictionary.jsonl",
        dictionary_documents,
    )
    all_documents = (
        document
        for group in [*documents.values(), dictionary_documents]
        for document in group
    )
    document_counts["all_documents"] = _write_jsonl(
        OUTPUT_DIR / "documents" / "all_documents.jsonl",
        all_documents,
    )

    schema["vector_documents"] = document_counts
    _write_json(OUTPUT_DIR / "metadata" / "schema.json", schema)
    _write_json(OUTPUT_DIR / "metadata" / "metric_definitions.json", metrics)
    _write_json(
        OUTPUT_DIR / "metadata" / "data_quality_report.json",
        build_quality_report(prices, production, markets),
    )
    write_duckdb(
        OUTPUT_DIR / "database" / "rice_rag.duckdb",
        prices,
        production,
        markets,
        aggregates,
    )
    write_sql_examples(OUTPUT_DIR / "examples" / "sql_examples.sql")
    write_chatbot_assets(OUTPUT_DIR)
    write_readme(
        OUTPUT_DIR,
        {name: len(frame) for name, frame in raw_tables.items()},
        document_counts,
    )
    logger.info("RAG package saved to %s", OUTPUT_DIR.resolve())


if __name__ == "__main__":
    main()
