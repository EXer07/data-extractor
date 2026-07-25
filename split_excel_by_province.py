"""Split powerbi_all_data.xlsx into one workbook per province and year."""

from __future__ import annotations

import re
import shutil
from pathlib import Path

import pandas as pd

from build_rag_dataset import load_and_clean

SOURCE_PATH = Path("powerbi_all_data.xlsx")
OUTPUT_DIR = Path("excel_by_province")


def safe_filename(value: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]', "_", value).strip().rstrip(".")
    return name or "ไม่ระบุจังหวัด"


def write_sheet(writer, sheet_name: str, frame: pd.DataFrame) -> None:
    frame.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(frame), len(frame.columns) - 1)
    worksheet.set_row(0, 25)

    header_format = writer.book.add_format(
        {
            "bold": True,
            "font_color": "white",
            "bg_color": "#1F4E78",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        }
    )
    integer_format = writer.book.add_format({"num_format": "#,##0"})
    date_format = writer.book.add_format({"num_format": "yyyy-mm-dd"})

    for column_index, column_name in enumerate(frame.columns):
        worksheet.write(0, column_index, column_name, header_format)
        lengths = frame[column_name].astype(str).str.len()
        width = min(
            max(len(str(column_name)) + 2, int(lengths.quantile(0.95)) + 2),
            38,
        )
        cell_format = None
        if pd.api.types.is_integer_dtype(frame[column_name]):
            cell_format = integer_format
        elif pd.api.types.is_datetime64_any_dtype(frame[column_name]):
            cell_format = date_format
        worksheet.set_column(column_index, column_index, width, cell_format)


def main() -> None:
    if not SOURCE_PATH.exists():
        raise FileNotFoundError(SOURCE_PATH)

    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(parents=True)

    prices, production, markets = load_and_clean()
    markets = markets.copy()
    markets["buddhist_year"] = 2569
    province_aliases = {"อยุธยา": "พระนครศรีอยุธยา"}
    for frame in (prices, production, markets):
        frame["province"] = frame["province"].replace(province_aliases)

    price_columns = [
        "survey_date",
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
    ]
    production_columns = [
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
    market_columns = [
        "buddhist_year",
        "schedule_id",
        "project_type",
        "province",
        "district",
        "location",
        "period_text",
        "details",
    ]

    provinces = sorted(
        set(prices["province"].dropna())
        | set(production["province"].dropna())
        | set(markets["province"].dropna())
    )

    manifest_rows = []
    for province in provinces:
        province_prices = prices.loc[prices["province"] == province, price_columns]
        province_production = production.loc[
            production["province"] == province, production_columns
        ]
        province_markets = markets.loc[markets["province"] == province, market_columns]

        summary_parts = []
        for data_type, frame in [
            ("price", province_prices),
            ("production", province_production),
            ("market_schedule", province_markets),
        ]:
            if frame.empty:
                continue
            summary = (
                frame.groupby("buddhist_year", as_index=False)
                .size()
                .rename(columns={"size": "row_count"})
            )
            summary.insert(0, "data_type", data_type)
            summary_parts.append(summary)
        summary_frame = pd.concat(summary_parts, ignore_index=True)
        summary_frame.insert(0, "province", province)

        output_path = OUTPUT_DIR / f"{safe_filename(province)}.xlsx"
        with pd.ExcelWriter(
            output_path,
            engine="xlsxwriter",
            datetime_format="yyyy-mm-dd",
            date_format="yyyy-mm-dd",
        ) as writer:
            write_sheet(writer, "Summary", summary_frame)

            for year in sorted(province_prices["buddhist_year"].unique()):
                year_frame = province_prices.loc[
                    province_prices["buddhist_year"] == year
                ]
                write_sheet(writer, f"Prices_{int(year)}", year_frame)

            for year in sorted(province_production["buddhist_year"].unique()):
                year_frame = province_production.loc[
                    province_production["buddhist_year"] == year
                ]
                write_sheet(writer, f"Production_{int(year)}", year_frame)

            for year in sorted(province_markets["buddhist_year"].unique()):
                year_frame = province_markets.loc[
                    province_markets["buddhist_year"] == year
                ]
                write_sheet(writer, f"Markets_{int(year)}", year_frame)

        manifest_rows.append(
            {
                "province": province,
                "file": output_path.name,
                "price_rows": len(province_prices),
                "production_rows": len(province_production),
                "market_rows": len(province_markets),
                "price_years": ", ".join(
                    str(int(year))
                    for year in sorted(province_prices["buddhist_year"].unique())
                ),
                "production_years": ", ".join(
                    str(int(year))
                    for year in sorted(province_production["buddhist_year"].unique())
                ),
            }
        )

    manifest = pd.DataFrame(manifest_rows)
    manifest_path = OUTPUT_DIR / "00_manifest.xlsx"
    with pd.ExcelWriter(manifest_path, engine="xlsxwriter") as writer:
        write_sheet(writer, "Provinces", manifest)

    print(f"Created {len(provinces)} province workbooks in {OUTPUT_DIR.resolve()}")


if __name__ == "__main__":
    main()
