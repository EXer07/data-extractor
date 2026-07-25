"""Extract every exposed raw table from a public Power BI report."""

from __future__ import annotations

import gzip
import json
import logging
import re
import urllib.request
import uuid
from pathlib import Path
from time import sleep
from typing import Any

import pandas as pd
import yaml
from selenium import webdriver

logger = logging.getLogger(__name__)

LOAD_WAIT_SECONDS = 15
MAX_ROWS_PER_TABLE = 30_000


def _post_json(url: str, resource_key: str, payload: dict[str, Any]) -> dict[str, Any]:
    headers = {
        "ActivityId": str(uuid.uuid4()),
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": "https://app.powerbi.com",
        "Referer": "https://app.powerbi.com/",
        "RequestId": str(uuid.uuid4()),
        "X-PowerBI-ResourceKey": resource_key,
    }
    request = urllib.request.Request(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=headers,
    )
    with urllib.request.urlopen(request, timeout=120) as response:
        body = response.read()
        if response.headers.get("Content-Encoding") == "gzip":
            body = gzip.decompress(body)
        return json.loads(body)


def _capture_report_context(url: str) -> dict[str, Any]:
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.set_capability("goog:loggingPrefs", {"performance": "ALL"})
    driver = webdriver.Chrome(options=options)
    try:
        driver.get(url)
        sleep(LOAD_WAIT_SECONDS)
        requests = []
        for entry in driver.get_log("performance"):
            message = json.loads(entry["message"])["message"]
            if message["method"] == "Network.requestWillBeSent":
                requests.append(message["params"]["request"])
    finally:
        driver.quit()

    schema_request = next(
        request
        for request in requests
        if request["method"] == "POST" and "conceptualschema" in request["url"].lower()
    )
    query_request = next(
        request
        for request in requests
        if request["method"] == "POST" and "querydata" in request["url"].lower()
    )
    example_payload = json.loads(query_request["postData"])
    app_context = example_payload["queries"][0]["ApplicationContext"]

    return {
        "resource_key": query_request["headers"]["X-PowerBI-ResourceKey"],
        "schema_url": schema_request["url"],
        "schema_payload": json.loads(schema_request["postData"]),
        "query_url": query_request["url"],
        "model_id": example_payload["modelId"],
        "application_context": app_context,
    }


def _column_select(source: str, entity: str, property_name: str) -> dict[str, Any]:
    return {
        "Column": {
            "Expression": {"SourceRef": {"Source": source}},
            "Property": property_name,
        },
        "Name": f"{entity}.{property_name}",
        "NativeReferenceName": property_name,
    }


def _build_query(
    model_id: int,
    application_context: dict[str, Any],
    entity: str,
    columns: list[str],
    filters: list[tuple[str, Any]] | None = None,
) -> dict[str, Any]:
    source = "t"
    projections = list(range(len(columns)))
    semantic_query: dict[str, Any] = {
        "Version": 2,
        "From": [{"Name": source, "Entity": entity, "Type": 0}],
        "Select": [
            _column_select(source, entity, property_name) for property_name in columns
        ],
    }
    if filters:
        semantic_query["Where"] = []
        for filter_property, filter_value in filters:
            literal = (
                f"{filter_value}L"
                if isinstance(filter_value, int)
                else f"'{str(filter_value).replace(chr(39), chr(39) * 2)}'"
            )
            semantic_query["Where"].append(
                {
                    "Condition": {
                        "In": {
                            "Expressions": [
                                {
                                    "Column": {
                                        "Expression": {"SourceRef": {"Source": source}},
                                        "Property": filter_property,
                                    }
                                }
                            ],
                            "Values": [[{"Literal": {"Value": literal}}]],
                        }
                    }
                }
            )

    command = {
        "SemanticQueryDataShapeCommand": {
            "Query": semantic_query,
            "Binding": {
                "Primary": {"Groupings": [{"Projections": projections}]},
                "DataReduction": {
                    "DataVolume": 6,
                    "Primary": {"Window": {"Count": MAX_ROWS_PER_TABLE}},
                },
                "Version": 1,
            },
            "ExecutionMetricsKind": 1,
        }
    }
    return {
        "version": "1.0.0",
        "queries": [
            {
                "Query": {"Commands": [command]},
                "QueryId": "",
                "ApplicationContext": application_context,
            }
        ],
        "cancelQueries": [],
        "modelId": model_id,
    }


def _decode_value(
    value: Any, dictionary_name: str | None, dictionaries: dict[str, Any]
) -> Any:
    if dictionary_name and isinstance(value, int):
        values = dictionaries.get(dictionary_name)
        if isinstance(values, list) and 0 <= value < len(values):
            return values[value]
    return value


def _decode_rows(response: dict[str, Any], fallback_columns: list[str]) -> pd.DataFrame:
    result = response["results"][0]["result"]["data"]
    descriptor = result.get("descriptor", {}).get("Select", [])
    column_names = [
        column.get(
            "Name",
            fallback_columns[index]
            if index < len(fallback_columns)
            else f"column_{index + 1}",
        )
        for index, column in enumerate(descriptor)
    ]
    if not column_names:
        column_names = fallback_columns

    datasets = result["dsr"]["DS"]
    decoded_rows: list[list[Any]] = []
    for dataset in datasets:
        dictionaries = dataset.get("ValueDicts", {})
        for phase in dataset.get("PH", []):
            for rows in phase.values():
                if not isinstance(rows, list) or not rows:
                    continue
                schema = rows[0].get("S", [])
                dictionary_names = [item.get("DN") for item in schema]
                previous = [None] * len(column_names)

                for encoded in rows:
                    if schema and any(item.get("N") in encoded for item in schema):
                        current = [
                            encoded.get(
                                item.get("N"),
                                previous[index] if index < len(previous) else None,
                            )
                            for index, item in enumerate(schema)
                        ]
                        decoded_rows.append(current)
                        previous = current
                        continue

                    values = encoded.get("C", [])
                    repeat_mask = encoded.get("R", 0)
                    null_mask = encoded.get("Ø", 0)
                    value_index = 0
                    current: list[Any] = []
                    for column_index in range(len(column_names)):
                        bit = 1 << column_index
                        if repeat_mask & bit:
                            value = previous[column_index]
                        elif null_mask & bit:
                            value = None
                        else:
                            value = (
                                values[value_index]
                                if value_index < len(values)
                                else None
                            )
                            value_index += 1
                            dictionary_name = (
                                dictionary_names[column_index]
                                if column_index < len(dictionary_names)
                                else None
                            )
                            value = _decode_value(value, dictionary_name, dictionaries)
                        current.append(value)
                    decoded_rows.append(current)
                    previous = current

    return pd.DataFrame(decoded_rows, columns=column_names)


def _query_table(
    context: dict[str, Any],
    entity: str,
    columns: list[str],
    filters: list[tuple[str, Any]] | None = None,
) -> pd.DataFrame:
    payload = _build_query(
        context["model_id"],
        context["application_context"],
        entity,
        columns,
        filters,
    )
    response = _post_json(
        context["query_url"],
        context["resource_key"],
        payload,
    )
    return _decode_rows(response, columns)


def _safe_sheet_name(name: str, used: set[str]) -> str:
    base = re.sub(r"[\[\]:*?/\\]", "_", name).strip()[:31] or "Table"
    candidate = base
    suffix = 2
    while candidate in used:
        tail = f"_{suffix}"
        candidate = f"{base[: 31 - len(tail)]}{tail}"
        suffix += 1
    used.add(candidate)
    return candidate


def extract_all(url: str, output_path: Path) -> dict[str, int]:
    context = _capture_report_context(url)
    schema_response = _post_json(
        context["schema_url"],
        context["resource_key"],
        context["schema_payload"],
    )
    schema = next(
        item["schema"]
        for item in schema_response["schemas"]
        if item["modelId"] == context["model_id"]
    )

    tables: dict[str, pd.DataFrame] = {}
    for entity in schema["Entities"]:
        columns = [
            prop["Name"] for prop in entity.get("Properties", []) if "Column" in prop
        ]
        if not columns:
            continue
        logger.info("Extracting %s (%d columns)", entity["Name"], len(columns))
        dataframe = _query_table(context, entity["Name"], columns)

        # Public Power BI queries return at most 30,000 rows. Split large fact
        # tables by year and combine the partitions to avoid silent truncation.
        if len(dataframe) >= MAX_ROWS_PER_TABLE and "YearBE" in columns:
            year_frame = _query_table(context, entity["Name"], ["YearBE"])
            year_column = year_frame.columns[0]
            partitions = []
            for year in year_frame[year_column].dropna().unique():
                logger.info("Extracting %s year %s", entity["Name"], year)
                partition = _query_table(
                    context,
                    entity["Name"],
                    columns,
                    [("YearBE", int(year))],
                )
                if len(partition) >= MAX_ROWS_PER_TABLE:
                    if "Month" not in columns:
                        raise RuntimeError(
                            f"{entity['Name']} year {year} still reached the "
                            f"{MAX_ROWS_PER_TABLE:,}-row Power BI limit"
                        )
                    month_frame = _query_table(
                        context,
                        entity["Name"],
                        ["Month"],
                        [("YearBE", int(year))],
                    )
                    month_column = month_frame.columns[0]
                    for month in month_frame[month_column].dropna().unique():
                        logger.info(
                            "Extracting %s year %s month %s",
                            entity["Name"],
                            year,
                            month,
                        )
                        month_partition = _query_table(
                            context,
                            entity["Name"],
                            columns,
                            [("YearBE", int(year)), ("Month", int(month))],
                        )
                        if len(month_partition) >= MAX_ROWS_PER_TABLE:
                            raise RuntimeError(
                                f"{entity['Name']} year {year} month {month} "
                                f"still reached the {MAX_ROWS_PER_TABLE:,}-row "
                                "Power BI limit"
                            )
                        partitions.append(month_partition)
                else:
                    partitions.append(partition)
            dataframe = pd.concat(partitions, ignore_index=True)

        tables[entity["Name"]] = dataframe

    used_sheet_names: set[str] = set()
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        summary = pd.DataFrame(
            [
                {
                    "table": table_name,
                    "rows": len(dataframe),
                    "columns": len(dataframe.columns),
                }
                for table_name, dataframe in tables.items()
            ]
        )
        summary.to_excel(writer, sheet_name="Summary", index=False)
        for table_name, dataframe in tables.items():
            sheet_name = _safe_sheet_name(table_name, used_sheet_names)
            dataframe.to_excel(writer, sheet_name=sheet_name, index=False)

    return {table_name: len(dataframe) for table_name, dataframe in tables.items()}


def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="[%(asctime)s] [%(levelname)s] %(message)s",
    )
    with Path("config.yml").open() as config_file:
        config = yaml.safe_load(config_file)
    url = config["console"]["url"]
    output_path = Path("powerbi_all_data.xlsx").resolve()
    counts = extract_all(url, output_path)
    logger.info("Saved %s", output_path)
    for table_name, row_count in counts.items():
        logger.info("%s: %d rows", table_name, row_count)


if __name__ == "__main__":
    main()
