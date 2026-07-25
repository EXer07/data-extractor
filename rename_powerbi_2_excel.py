"""Create a reader-friendly Excel copy of the second Power BI dataset."""

from pathlib import Path

import pandas as pd

SOURCE_PATH = Path("powerbi_2_all_data.xlsx")
OUTPUT_PATH = Path("คาดการณ์อุปทานข้าวจากภาพถ่ายดาวเทียม.xlsx")

SHEET_NAMES = {
    "Summary": "สรุปชุดข้อมูล",
    "supply_latest_GISTDA": "ผลผลิตข้าวรายพื้นที่",
    "location": "ข้อมูลจังหวัดและอำเภอ",
    "time": "งวดเวลาและฤดูเพาะปลูก",
    "ricemill_capacity": "กำลังการผลิตโรงสี",
    "ricetype": "ประเภทข้าว",
    "param_selected_data_source": "แหล่งข้อมูล",
}

COLUMN_NAMES = {
    "Month Name Abbr": "ชื่อเดือนย่อ",
    "Year BE Abbr": "ปี พ.ศ. ย่อ",
    "Half Month Name": "รหัสครึ่งเดือน",
    "period": "รหัสงวด",
    "subdistrict_id": "รหัสพื้นที่",
    "Month Number": "เลขเดือน",
    "Half Month Number": "ลำดับครึ่งเดือน",
    "Number of Weeks Forward": "จำนวนสัปดาห์ล่วงหน้า",
    "season": "ฤดูเพาะปลูก",
}


def clean_columns(frame: pd.DataFrame, source_sheet: str) -> pd.DataFrame:
    if source_sheet == "Summary":
        frame = frame.copy()
        frame["table"] = frame["table"].replace(SHEET_NAMES)
        return frame.rename(
            columns={
                "table": "ชุดข้อมูล",
                "rows": "จำนวนแถว",
                "columns": "จำนวนคอลัมน์",
            }
        )

    prefix = f"{source_sheet}."
    renamed = {}
    for column in frame.columns:
        name = str(column)
        if name.startswith(prefix):
            name = name[len(prefix) :]
        renamed[column] = name

    frame = frame.rename(columns=renamed).rename(columns=COLUMN_NAMES)
    if source_sheet == "param_selected_data_source":
        frame = frame.rename(columns={"data_source": "แหล่งข้อมูล"})
    return frame


def write_formatted_sheet(writer, sheet_name: str, frame: pd.DataFrame) -> None:
    frame.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(frame), len(frame.columns) - 1)
    worksheet.set_row(0, 26)

    header_format = writer.book.add_format(
        {
            "bold": True,
            "font_color": "white",
            "bg_color": "#2F5597",
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
            40,
        )
        cell_format = (
            integer_format
            if pd.api.types.is_integer_dtype(frame[column_name])
            else None
        )
        worksheet.set_column(column_index, column_index, width, cell_format)


def main() -> None:
    if not SOURCE_PATH.exists():
        raise FileNotFoundError(SOURCE_PATH.resolve())

    workbook = pd.ExcelFile(SOURCE_PATH)
    with pd.ExcelWriter(OUTPUT_PATH, engine="xlsxwriter") as writer:
        for source_sheet in workbook.sheet_names:
            frame = pd.read_excel(SOURCE_PATH, sheet_name=source_sheet)
            frame = clean_columns(frame, source_sheet)
            write_formatted_sheet(writer, SHEET_NAMES[source_sheet], frame)

    print(OUTPUT_PATH.resolve())


if __name__ == "__main__":
    main()
