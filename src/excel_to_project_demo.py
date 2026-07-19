"""End-to-end public demo: read an XLSX task plan and emit project JSON.

The XLSX reader intentionally uses Python's standard library only. It supports
the small, explicit table contract documented in docs/05-ai-specification.md;
it is not intended to replace a production-grade spreadsheet parser.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from document_to_project_demo import transform


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"x": MAIN_NS, "r": DOC_REL_NS}

HEADER_ALIASES = {
    "module": {"模块", "阶段", "一级分类"},
    "card": {"卡片", "工作包", "里程碑", "分组"},
    "task": {"任务", "事项", "工作内容", "任务名称"},
    "owner": {"负责人", "执行人", "owner"},
    "start_date": {"开始日期", "计划开始", "开始时间"},
    "end_date": {"截止日期", "计划完成", "结束日期", "结束时间"},
    "notes": {"备注", "说明"},
}
DATE_FIELDS = {"start_date", "end_date"}


def column_index(cell_ref: str) -> int:
    """Convert an Excel cell reference such as C7 to a zero-based column."""
    match = re.match(r"([A-Z]+)", cell_ref.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    result = 0
    for char in match.group(1):
        result = result * 26 + ord(char) - ord("A") + 1
    return result - 1


def excel_date_to_iso(value: Any) -> str:
    """Convert Excel's 1900 date serial to YYYY-MM-DD."""
    if value in (None, ""):
        return ""
    if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        return value
    try:
        serial = float(value)
    except (TypeError, ValueError):
        return str(value).strip()
    # Excel's 1900 date system includes the historical leap-year bug.
    converted = datetime(1899, 12, 30) + timedelta(days=serial)
    return converted.date().isoformat()


def _shared_strings(archive: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for item in root.findall("x:si", NS):
        strings.append("".join(node.text or "" for node in item.iter(f"{{{MAIN_NS}}}t")))
    return strings


def _sheet_paths(archive: ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    result: list[tuple[str, str]] = []
    for sheet in workbook.findall("x:sheets/x:sheet", NS):
        relationship_id = sheet.attrib[f"{{{DOC_REL_NS}}}id"]
        target = targets[relationship_id].replace("\\", "/").lstrip("/")
        if not target.startswith("xl/"):
            target = f"xl/{target}"
        result.append((sheet.attrib["name"], target))
    return result


def _cell_value(cell: ET.Element, shared_strings: list[str]) -> Any:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(f"{{{MAIN_NS}}}t"))
    value_node = cell.find("x:v", NS)
    if value_node is None or value_node.text is None:
        return ""
    raw = value_node.text
    if cell_type == "s":
        return shared_strings[int(raw)]
    if cell_type == "b":
        return raw == "1"
    if cell_type in {"str", "e"}:
        return raw
    try:
        numeric = float(raw)
        return int(numeric) if numeric.is_integer() else numeric
    except ValueError:
        return raw


def read_sheet_rows(
    archive: ZipFile, sheet_path: str, shared_strings: list[str]
) -> list[tuple[int, list[Any]]]:
    """Read non-empty rows while retaining their original Excel row numbers."""
    root = ET.fromstring(archive.read(sheet_path))
    rows: list[tuple[int, list[Any]]] = []
    for row in root.findall(".//x:sheetData/x:row", NS):
        indexed: dict[int, Any] = {}
        for cell in row.findall("x:c", NS):
            indexed[column_index(cell.attrib["r"])] = _cell_value(cell, shared_strings)
        if not indexed:
            continue
        width = max(indexed) + 1
        values = [indexed.get(index, "") for index in range(width)]
        rows.append((int(row.attrib["r"]), values))
    return rows


def _header_mapping(values: list[Any]) -> dict[int, str]:
    normalized = [str(value).strip().lower() for value in values]
    mapping: dict[int, str] = {}
    for index, value in enumerate(normalized):
        for field, aliases in HEADER_ALIASES.items():
            if value in aliases:
                mapping[index] = field
                break
    return mapping


def read_xlsx_payload(path: Path, sheet_name: str | None = None) -> dict[str, Any]:
    """Read the first matching task table from an XLSX workbook."""
    with ZipFile(path) as archive:
        shared_strings = _shared_strings(archive)
        sheets = _sheet_paths(archive)
        if sheet_name:
            sheets = [item for item in sheets if item[0] == sheet_name]
            if not sheets:
                raise ValueError(f"Worksheet not found: {sheet_name}")

        for current_name, sheet_path in sheets:
            rows = read_sheet_rows(archive, sheet_path, shared_strings)
            project_name = path.stem
            for row_index, (excel_row, values) in enumerate(rows):
                mapping = _header_mapping(values)
                if "task" not in mapping.values():
                    if row_index == 0 and values and str(values[0]).strip():
                        project_name = str(values[0]).strip()
                    continue

                records: list[dict[str, Any]] = []
                previous_source_row = excel_row
                for source_row, raw_values in rows[row_index + 1 :]:
                    # A blank line ends the contiguous task table. This avoids
                    # treating footer instructions as task data.
                    if records and source_row > previous_source_row + 1:
                        break
                    previous_source_row = source_row
                    record: dict[str, Any] = {
                        "source_sheet": current_name,
                        "source_row": source_row,
                    }
                    for column, field in mapping.items():
                        value = raw_values[column] if column < len(raw_values) else ""
                        record[field] = (
                            excel_date_to_iso(value) if field in DATE_FIELDS else value
                        )
                    if any(str(record.get(field, "")).strip() for field in mapping.values()):
                        records.append(record)

                return {"project_name": project_name, "rows": records}

    raise ValueError(
        "No task table found. A header row containing '任务' or a supported alias is required."
    )


def convert_xlsx(path: Path, sheet_name: str | None = None) -> dict[str, Any]:
    """Read an XLSX workbook and apply the deterministic project transform."""
    return transform(read_xlsx_payload(path, sheet_name))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert an XLSX task plan to nested project JSON."
    )
    parser.add_argument("input", type=Path, help="Path to an .xlsx file")
    parser.add_argument("--sheet", help="Optional worksheet name")
    parser.add_argument(
        "--output",
        type=Path,
        help="Optional output path. Prints JSON to stdout when omitted.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.input.suffix.lower() != ".xlsx":
        raise SystemExit("Only .xlsx input is supported by this public demo.")
    try:
        result = convert_xlsx(args.input, args.sheet)
    except (KeyError, ValueError, ET.ParseError, OSError) as error:
        raise SystemExit(f"Could not convert workbook: {error}") from error
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered + "\n", encoding="utf-8")
        print(f"Wrote {args.output}")
    else:
        sys.stdout.write(rendered + "\n")


if __name__ == "__main__":
    main()
