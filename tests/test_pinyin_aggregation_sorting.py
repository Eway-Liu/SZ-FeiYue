import importlib.util
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import unittest

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]
BUILD_SCRIPT = ROOT / "scripts" / "build_cases.py"

spec = importlib.util.spec_from_file_location("build_cases", BUILD_SCRIPT)
assert spec is not None and spec.loader is not None
build_cases = importlib.util.module_from_spec(spec)
spec.loader.exec_module(build_cases)


class PinyinAggregationSortingTest(unittest.TestCase):
    def test_sort_key_uses_toneless_pinyin_casefold_and_name_tiebreaker(self) -> None:
        sort_key = getattr(build_cases, "pinyin_sort_key", None)
        self.assertIsNotNone(sort_key)

        self.assertLess(sort_key("北京航空航天大学"), sort_key("上海交通大学"))
        self.assertEqual(sort_key("AI工程")[0], sort_key("ai工程")[0])
        self.assertEqual(sort_key("李")[0], sort_key("里")[0])
        self.assertNotEqual(sort_key("李"), sort_key("里"))
        self.assertLess(
            sort_key("香港科技大学（广州）"),
            sort_key("香港科技大学（深圳）"),
        )

    def test_generated_university_and_major_headings_follow_pinyin(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            scripts_dir = temp_root / "scripts"
            docs_dir = temp_root / "docs"
            scripts_dir.mkdir()
            (docs_dir / "seniors").mkdir(parents=True)
            shutil.copy2(BUILD_SCRIPT, scripts_dir / "build_cases.py")
            (docs_dir / "index.md").write_text(
                "<!-- LAST_UPDATED_START -->\nold\n<!-- LAST_UPDATED_END -->\n",
                encoding="utf-8",
            )

            workbook = Workbook()
            sheet = workbook.active
            sheet.append([
                "昵称", "高考年份", "选科", "高考分数", "高考排名",
                "录取院校", "录取专业", "院校评价", "专业评价",
                "给学弟学妹的建议", "提交时间",
            ])
            rows = [
                ("甲", "上海交通大学", "软件工程"),
                ("乙", "北京航空航天大学", "人工智能"),
                ("丙", "电子科技大学", "电子信息科学与技术"),
                ("丁", "复旦大学", "汉语言文学"),
                ("戊", "哈尔滨工业大学（深圳）", "计算机科学与技术"),
            ]
            for index, (nickname, university, major) in enumerate(rows, start=1):
                sheet.append([
                    nickname, "2026", "物理", str(660 - index), str(100 + index),
                    university, major, "院校评价", "专业评价", "建议",
                    f"2026-08-02 12:0{index}:00",
                ])
            workbook.save(docs_dir / "submissions.xlsx")

            result = subprocess.run(
                [sys.executable, "scripts/build_cases.py"],
                cwd=temp_root,
                text=True,
                capture_output=True,
                check=False,
            )
            self.assertEqual(result.returncode, 0, result.stderr)

            university_page = (docs_dir / "cases" / "by-university.md").read_text(
                encoding="utf-8"
            )
            major_page = (docs_dir / "cases" / "by-major.md").read_text(
                encoding="utf-8"
            )
            heading_pattern = re.compile(r"^## (.+?)（\d+）$", re.MULTILINE)

            self.assertEqual(
                heading_pattern.findall(university_page),
                [
                    "北京航空航天大学",
                    "电子科技大学",
                    "复旦大学",
                    "哈尔滨工业大学（深圳）",
                    "上海交通大学",
                ],
            )
            self.assertEqual(
                heading_pattern.findall(major_page),
                [
                    "电子信息科学与技术",
                    "汉语言文学",
                    "计算机科学与技术",
                    "人工智能",
                    "软件工程",
                ],
            )


if __name__ == "__main__":
    unittest.main()
