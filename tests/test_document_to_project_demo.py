import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from document_to_project_demo import transform  # noqa: E402


class TransformTests(unittest.TestCase):
    def test_groups_rows_and_preserves_source_order(self):
        result = transform(
            {
                "project_name": "测试项目",
                "rows": [
                    {"module": "阶段 A", "card": "分组 1", "task": "任务 1"},
                    {"module": "阶段 A", "card": "分组 1", "task": "任务 2"},
                    {"module": "阶段 B", "card": "分组 2", "task": "任务 3"},
                ],
            }
        )

        modules = result["project"]["modules"]
        self.assertEqual(["阶段 A", "阶段 B"], [item["name"] for item in modules])
        self.assertEqual(
            ["任务 1", "任务 2"],
            [item["title"] for item in modules[0]["cards"][0]["tasks"]],
        )
        self.assertEqual(1.0, result["quality"]["coverage_rate"])

    def test_uses_default_groups_and_emits_warnings(self):
        result = transform(
            {
                "project_name": "测试项目",
                "rows": [{"task": "未分组任务"}],
            }
        )

        module = result["project"]["modules"][0]
        self.assertEqual("待分类", module["name"])
        self.assertEqual("任务清单", module["cards"][0]["name"])
        warning_codes = {item["code"] for item in result["warnings"]}
        self.assertIn("DEFAULT_MODULE_USED", warning_codes)
        self.assertIn("DEFAULT_CARD_USED", warning_codes)

    def test_skips_missing_titles_and_flags_date_conflicts(self):
        result = transform(
            {
                "project_name": "测试项目",
                "rows": [
                    {"module": "阶段 A"},
                    {
                        "module": "阶段 A",
                        "card": "分组 1",
                        "task": "日期冲突任务",
                        "start_date": "2026-07-03",
                        "end_date": "2026-07-01",
                    },
                ],
            }
        )

        self.assertEqual(1, result["quality"]["mapped_tasks"])
        self.assertEqual(0.5, result["quality"]["coverage_rate"])
        warning_codes = {item["code"] for item in result["warnings"]}
        self.assertIn("MISSING_TASK_TITLE", warning_codes)
        self.assertIn("DATE_RANGE_CONFLICT", warning_codes)

    def test_requires_project_name_and_list_rows(self):
        with self.assertRaises(ValueError):
            transform({"rows": []})
        with self.assertRaises(ValueError):
            transform({"project_name": "测试项目", "rows": {}})


if __name__ == "__main__":
    unittest.main()
