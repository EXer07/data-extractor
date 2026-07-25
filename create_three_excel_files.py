"""Create the three Excel deliverables requested from powerbi_all_data.xlsx."""

from __future__ import annotations

import re
import shutil
from pathlib import Path

import pandas as pd

SOURCE_PATH = Path("powerbi_all_data.xlsx")
OUTPUT_DIR = Path("excel_split_corrected")


def safe_sheet_name(value: str, used: set[str]) -> str:
    base = re.sub(r"[\[\]:*?/\\]", "_", str(value)).strip()[:31] or "ไม่ระบุจังหวัด"
    candidate = base
    suffix = 2
    while candidate in used:
        tail = f"_{suffix}"
        candidate = f"{base[: 31 - len(tail)]}{tail}"
        suffix += 1
    used.add(candidate)
    return candidate


def write_formatted_sheet(
    writer: pd.ExcelWriter, sheet_name: str, frame: pd.DataFrame
) -> None:
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
    for column_index, column_name in enumerate(frame.columns):
        worksheet.write(0, column_index, column_name, header_format)
        lengths = frame[column_name].astype(str).str.len()
        width = min(
            max(len(str(column_name)) + 2, int(lengths.quantile(0.95)) + 2),
            38,
        )
        cell_format = (
            integer_format
            if pd.api.types.is_integer_dtype(frame[column_name])
            else None
        )
        worksheet.set_column(column_index, column_index, width, cell_format)


def write_single_sheet_file(
    output_path: Path, sheet_name: str, frame: pd.DataFrame
) -> None:
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        write_formatted_sheet(writer, sheet_name, frame)


def main() -> None:
    if not SOURCE_PATH.exists():
        raise FileNotFoundError(SOURCE_PATH)
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(parents=True)

    rice_2 = pd.read_excel(SOURCE_PATH, sheet_name="Rice (2)")
    sheet_1 = pd.read_excel(SOURCE_PATH, sheet_name="Sheet1")
    prices = pd.read_excel(SOURCE_PATH, sheet_name="price start 2023")

    write_single_sheet_file(
        OUTPUT_DIR / "rice_2.xlsx",
        "เนื้อที่และผลผลิตข้าว",
        rice_2,
    )
    write_single_sheet_file(
        OUTPUT_DIR / "sheet1.xlsx",
        "กำหนดการตลาดนัดข้าว",
        sheet_1,
    )

    province_column = "price start 2023.data.ProvinceName"
    provinces = sorted(
        prices[province_column].dropna().astype(str).str.strip().unique()
    )
    used_sheet_names: set[str] = set()
    with pd.ExcelWriter(OUTPUT_DIR / "by_province.xlsx", engine="xlsxwriter") as writer:
        for province in provinces:
            province_frame = prices.loc[
                prices[province_column].astype(str).str.strip() == province
            ].copy()
            write_formatted_sheet(
                writer,
                safe_sheet_name(f"ราคาข้าว_{province}", used_sheet_names),
                province_frame,
            )

    print(f"rice_2 rows: {len(rice_2)}")
    print(f"sheet1 rows: {len(sheet_1)}")
    print(f"by_province sheets: {len(provinces)}, rows: {len(prices)}")
    print(OUTPUT_DIR.resolve())


if __name__ == "__main__":
    main()
