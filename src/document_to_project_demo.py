"""Deterministic, sanitized demo for mapping extracted table rows to project JSON.

This module starts after file parsing. It does not contain production APIs,
customer schemas, model calls, or authentication logic.
"""

from __future__ import annotations

import argparse
import json
from collections import OrderedDict
from datetime import date
from pathlib import Path
from typing import Any


DEFAULT_MODULE = "待分类"
DEFAULT_CARD = "任务清单"


def clean_text(value: Any) -> str:
    """Return a trimmed string; None and non-string scalar values are supported."""
    if value is None:
        return ""
    return str(value).strip()


def is_iso_date(value: str) -> bool:
    if not value:
        return True
    try:
        date.fromisoformat(value)
        return True
    except ValueError:
        return False


def transform(payload: dict[str, Any]) -> dict[str, Any]:
    """Transform normalized table rows into a stable nested project structure."""
    project_name = clean_text(payload.get("project_name"))
    rows = payload.get("rows")
    if not project_name:
        raise ValueError("project_name is required")
    if not isinstance(rows, list):
        raise ValueError("rows must be a list")

    warnings: list[dict[str, Any]] = []
    modules: OrderedDict[str, OrderedDict[str, list[dict[str, Any]]]] = (
        OrderedDict()
    )
    valid_row_count = 0
    default_module_count = 0
    default_card_count = 0

    for row_number, raw_row in enumerate(rows, start=1):
        if not isinstance(raw_row, dict):
            warnings.append(
                {
                    "code": "INVALID_ROW",
                    "row": row_number,
                    "message": "Row is not an object and was skipped.",
                }
            )
            continue

        title = clean_text(raw_row.get("task"))
        if not title:
            warnings.append(
                {
                    "code": "MISSING_TASK_TITLE",
                    "row": row_number,
                    "message": "Task title is empty and the row was skipped.",
                }
            )
            continue

        module_name = clean_text(raw_row.get("module")) or DEFAULT_MODULE
        card_name = clean_text(raw_row.get("card")) or DEFAULT_CARD
        if module_name == DEFAULT_MODULE:
            default_module_count += 1
            warnings.append(
                {
                    "code": "DEFAULT_MODULE_USED",
                    "row": row_number,
                    "message": f"Missing module; grouped under '{DEFAULT_MODULE}'.",
                }
            )
        if card_name == DEFAULT_CARD:
            default_card_count += 1
            warnings.append(
                {
                    "code": "DEFAULT_CARD_USED",
                    "row": row_number,
                    "message": f"Missing card; grouped under '{DEFAULT_CARD}'.",
                }
            )

        start_date = clean_text(raw_row.get("start_date"))
        end_date = clean_text(raw_row.get("end_date"))
        for field_name, field_value in (
            ("start_date", start_date),
            ("end_date", end_date),
        ):
            if not is_iso_date(field_value):
                warnings.append(
                    {
                        "code": "INVALID_DATE",
                        "row": row_number,
                        "field": field_name,
                        "value": field_value,
                        "message": "Date must use YYYY-MM-DD format.",
                    }
                )

        if is_iso_date(start_date) and is_iso_date(end_date):
            if start_date and end_date and start_date > end_date:
                warnings.append(
                    {
                        "code": "DATE_RANGE_CONFLICT",
                        "row": row_number,
                        "message": "start_date is later than end_date.",
                    }
                )

        task = {
            "id": f"task-{row_number:04d}",
            "title": title,
            "owner": clean_text(raw_row.get("owner")) or None,
            "start_date": start_date or None,
            "end_date": end_date or None,
            "notes": clean_text(raw_row.get("notes")) or None,
            "source_ref": {
                "sheet": clean_text(raw_row.get("source_sheet")) or None,
                "row": raw_row.get("source_row", row_number),
            },
        }
        modules.setdefault(module_name, OrderedDict()).setdefault(
            card_name, []
        ).append(task)
        valid_row_count += 1

    nested_modules: list[dict[str, Any]] = []
    for module_index, (module_name, cards) in enumerate(modules.items(), start=1):
        nested_cards: list[dict[str, Any]] = []
        for card_index, (card_name, tasks) in enumerate(cards.items(), start=1):
            nested_cards.append(
                {
                    "id": f"module-{module_index:02d}-card-{card_index:02d}",
                    "name": card_name,
                    "tasks": tasks,
                }
            )
        nested_modules.append(
            {
                "id": f"module-{module_index:02d}",
                "name": module_name,
                "cards": nested_cards,
            }
        )

    return {
        "schema_version": "demo-1.0",
        "project": {
            "name": project_name,
            "modules": nested_modules,
        },
        "quality": {
            "source_rows": len(rows),
            "mapped_tasks": valid_row_count,
            "coverage_rate": round(valid_row_count / len(rows), 4) if rows else 0,
            "default_module_rows": default_module_count,
            "default_card_rows": default_card_count,
            "warning_count": len(warnings),
        },
        "warnings": warnings,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Map extracted table rows to a nested project JSON."
    )
    parser.add_argument("input", type=Path, help="Path to normalized input JSON")
    parser.add_argument(
        "--output",
        type=Path,
        help="Optional output path. Prints to stdout when omitted.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    with args.input.open("r", encoding="utf-8") as source:
        payload = json.load(source)
    result = transform(payload)
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered + "\n", encoding="utf-8")
        print(f"Wrote {args.output}")
    else:
        print(rendered)


if __name__ == "__main__":
    main()
