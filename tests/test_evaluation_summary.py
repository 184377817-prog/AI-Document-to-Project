import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "evaluation"))

from summarize_results import load_results, summarize  # noqa: E402


class EvaluationSummaryTests(unittest.TestCase):
    def test_public_snapshot_is_reproducible(self):
        rows = load_results(ROOT / "evaluation" / "anonymized_results.csv")
        summary = summarize(rows)

        self.assertEqual(20, summary["sample_count"])
        self.assertEqual(89.5, summary["ai_score"]["mean"])
        self.assertEqual(83.1, summary["human_review"]["mean"])
        self.assertEqual(
            {"conditional": 2, "fail": 5, "pass": 13},
            summary["human_review"]["status_distribution"],
        )
        self.assertEqual(40.1, summary["latency_seconds"]["mean"])
        self.assertEqual(4, summary["score_gap"]["ai_pass_but_human_fail"])


if __name__ == "__main__":
    unittest.main()
