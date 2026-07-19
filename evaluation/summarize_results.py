"""Recalculate the public evaluation snapshot from anonymized row-level data."""

from __future__ import annotations

import argparse
import csv
import json
from collections import Counter, defaultdict
from pathlib import Path
from statistics import mean, median
from typing import Any


def load_results(path: Path) -> list[dict[str, Any]]:
    with path.open("r", encoding="utf-8", newline="") as source:
        rows = list(csv.DictReader(source))
    if not rows:
        raise ValueError("Evaluation CSV is empty.")
    for row in rows:
        row["ai_score"] = float(row["ai_score"])
        row["latency_seconds"] = float(row["latency_seconds"])
        row["human_score"] = float(row["human_score"])
    return rows


def percentile(values: list[float], fraction: float) -> float:
    ordered = sorted(values)
    index = round((len(ordered) - 1) * fraction)
    return ordered[index]


def summarize(rows: list[dict[str, Any]]) -> dict[str, Any]:
    tiers: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        tiers[row["tier"]].append(row)

    human_counts = Counter(row["human_status"] for row in rows)
    ai_counts = Counter(row["ai_status"] for row in rows)
    return {
        "sample_count": len(rows),
        "ai_score": {
            "mean": round(mean(row["ai_score"] for row in rows), 1),
            "status_distribution": dict(sorted(ai_counts.items())),
        },
        "human_review": {
            "mean": round(mean(row["human_score"] for row in rows), 1),
            "status_distribution": dict(sorted(human_counts.items())),
        },
        "latency_seconds": {
            "mean": round(mean(row["latency_seconds"] for row in rows), 1),
            "median": round(median(row["latency_seconds"] for row in rows), 1),
            "p95_nearest_rank": round(
                percentile([row["latency_seconds"] for row in rows], 0.95), 1
            ),
            "max": round(max(row["latency_seconds"] for row in rows), 1),
        },
        "score_gap": {
            "mean_ai_minus_human": round(
                mean(row["ai_score"] - row["human_score"] for row in rows), 1
            ),
            "ai_pass_but_human_fail": sum(
                row["ai_status"] == "pass" and row["human_status"] == "fail"
                for row in rows
            ),
        },
        "by_tier": {
            tier: {
                "count": len(items),
                "ai_mean": round(mean(item["ai_score"] for item in items), 1),
                "human_mean": round(mean(item["human_score"] for item in items), 1),
            }
            for tier, items in sorted(tiers.items())
        },
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        default=Path(__file__).with_name("anonymized_results.csv"),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).with_name("summary.json"),
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    summary = summarize(load_results(args.input))
    args.output.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
