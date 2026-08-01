import re
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
BUILD_SCRIPT = ROOT / "scripts" / "build_cases.py"


class CaseAggregationLinksTest(unittest.TestCase):
    def test_identity_links_target_exact_case_and_exclude_review_text(self) -> None:
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
            sheet.append([
                "Link Student", "2026", "物理", "650", "1000",
                "Link University", "Software Engineering",
                "University review", "Major review", "Advice body",
                "2026-08-01 12:00:00",
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

            case_index = (docs_dir / "cases" / "index.md").read_text(encoding="utf-8")
            stem_match = re.search(r"\]\((case-[0-9a-f]{10})/\)", case_index)
            self.assertIsNotNone(stem_match)
            stem = stem_match.group(1)

            university_page = (docs_dir / "cases" / "by-university.md").read_text(encoding="utf-8")
            major_page = (docs_dir / "cases" / "by-major.md").read_text(encoding="utf-8")
            experience_page = (docs_dir / "experience.md").read_text(encoding="utf-8")

            self.assertIn(
                f"- [**Link Student | Software Engineering**]({stem}/)：University review",
                university_page,
            )
            self.assertIn(
                f"- [**Link Student | Link University**]({stem}/)：Major review",
                major_page,
            )
            self.assertIn(
                f"- [**Link Student | 650 | Link University | Software Engineering**]"
                f"(cases/{stem}/)：Advice body",
                experience_page,
            )


if __name__ == "__main__":
    unittest.main()
