import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from excel_to_project_demo import (  # noqa: E402
    column_index,
    convert_xlsx,
    excel_date_to_iso,
    read_xlsx_payload,
)


class ExcelToProjectTests(unittest.TestCase):
    workbook = ROOT / "demo" / "sample_project_plan.xlsx"

    def test_excel_utilities(self):
        self.assertEqual(0, column_index("A1"))
        self.assertEqual(26, column_index("AA9"))
        self.assertEqual("2026-07-01", excel_date_to_iso(46204))

    def test_reads_real_xlsx_rows_and_source_references(self):
        payload = read_xlsx_payload(self.workbook)

        self.assertEqual("市场活动排期项目", payload["project_name"])
        self.assertEqual(8, len(payload["rows"]))
        self.assertEqual("项目计划", payload["rows"][0]["source_sheet"])
        self.assertEqual(5, payload["rows"][0]["source_row"])
        self.assertEqual("2026-07-01", payload["rows"][0]["start_date"])

    def test_end_to_end_conversion(self):
        result = convert_xlsx(self.workbook)

        self.assertEqual(8, result["quality"]["source_rows"])
        self.assertEqual(8, result["quality"]["mapped_tasks"])
        self.assertEqual(1.0, result["quality"]["coverage_rate"])
        self.assertEqual(1, result["quality"]["default_card_rows"])
        self.assertEqual(
            ["筹备", "设计", "开发", "上线"],
            [module["name"] for module in result["project"]["modules"]],
        )
        self.assertEqual(
            "任务清单",
            result["project"]["modules"][3]["cards"][1]["name"],
        )


if __name__ == "__main__":
    unittest.main()
