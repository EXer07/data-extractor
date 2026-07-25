"""Export the Power BI dataset as a clean, reader-friendly Excel workbook."""

from pathlib import Path

import pandas as pd

from build_rag_dataset import load_and_clean

OUTPUT_PATH = Path("rice_dataset_clean.xlsx")


def format_data_sheet(writer, sheet_name: str, frame: pd.DataFrame) -> None:
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(frame), len(frame.columns) - 1)
    worksheet.set_row(0, 26)

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
    for column_index, column_name in enumerate(frame.columns):
        worksheet.write(0, column_index, column_name, header_format)
        sample_lengths = frame[column_name].astype(str).str.len()
        width = min(
            max(len(column_name) + 2, int(sample_lengths.quantile(0.95)) + 2), 36
        )
        worksheet.set_column(column_index, column_index, width)


def main() -> None:
    prices, production, markets = load_and_clean()

    price_columns = [
        "survey_date",
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
        "schedule_id",
        "project_type",
        "province",
        "district",
        "location",
        "period_text",
        "details",
    ]

    prices = prices[price_columns]
    production = production[production_columns]
    markets = markets[market_columns]
    yearly_summary = (
        prices.groupby(["gregorian_year", "buddhist_year"], as_index=False)
        .agg(
            row_count=("province", "size"),
            first_survey_date=("survey_date", "min"),
            last_survey_date=("survey_date", "max"),
            province_count=("province", "nunique"),
            average_price=("avg_price", "mean"),
        )
        .sort_values("gregorian_year")
    )
    yearly_summary["average_price"] = yearly_summary["average_price"].round(2)

    with pd.ExcelWriter(
        OUTPUT_PATH,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd",
        date_format="yyyy-mm-dd",
    ) as writer:
        yearly_summary.to_excel(writer, sheet_name="Summary", index=False)
        prices.to_excel(writer, sheet_name="Rice_Prices", index=False)
        production.to_excel(writer, sheet_name="Rice_Production", index=False)
        markets.to_excel(writer, sheet_name="Market_Schedule", index=False)

        for sheet_name, frame in {
            "Summary": yearly_summary,
            "Rice_Prices": prices,
            "Rice_Production": production,
            "Market_Schedule": markets,
        }.items():
            format_data_sheet(writer, sheet_name, frame)

        summary = writer.sheets["Summary"]
        summary.set_column("A:B", 18)
        summary.set_column("C:F", 20)
        summary.set_column("E:E", 18, writer.book.add_format({"num_format": "#,##0"}))
        summary.set_column(
            "F:F", 18, writer.book.add_format({"num_format": "#,##0.00"})
        )

        prices_sheet = writer.sheets["Rice_Prices"]
        prices_sheet.set_column(
            "A:A", 14, writer.book.add_format({"num_format": "yyyy-mm-dd"})
        )
        prices_sheet.set_column(
            "G:I", 14, writer.book.add_format({"num_format": "#,##0"})
        )

        production_sheet = writer.sheets["Rice_Production"]
        production_sheet.set_column(
            "H:J", 18, writer.book.add_format({"num_format": "#,##0"})
        )

    print(OUTPUT_PATH.resolve())


if __name__ == "__main__":
    main()
