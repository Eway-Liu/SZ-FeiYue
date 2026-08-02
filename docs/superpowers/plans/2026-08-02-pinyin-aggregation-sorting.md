# Pinyin Aggregation Sorting Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the generated “按院校” and “按专业” headings sort by each complete name's toneless Chinese pinyin instead of Unicode code points.

**Architecture:** Keep the existing page-generation pipeline and add one internal `pinyin_sort_key(name)` helper in `scripts/build_cases.py`. The two aggregation loops will share this helper; dependency declarations and focused integration tests ensure local, CI, and generated output stay consistent.

**Tech Stack:** Python 3.11+, `unittest`, `openpyxl`, `pypinyin`, MkDocs Material, GitHub Actions

## Global Constraints

- Use the complete name's toneless pinyin as the primary key and compare it case-insensitively.
- Use the original name after case-folding as the deterministic secondary key when pinyin is identical.
- Preserve non-Chinese characters through transliteration so they participate in comparison.
- Change only university and major heading order; preserve item order within every heading and all other page orderings.
- Do not add initial-letter grouping, an index, or frontend interactions.

---

### Task 1: Declare the pypinyin build dependency

**Files:**
- Modify: `tests/test_pages_workflow.py`
- Modify: `.github/workflows/pages.yml`
- Modify: `README.md`

**Interfaces:**
- Consumes: the existing GitHub Actions `Install dependencies` step and README local-install command.
- Produces: identical `pypinyin` availability for local builds and CI builds.

- [ ] **Step 1: Write the failing dependency-declaration test**

Add the README path beside the existing workflow constant:

```python
README_PATH = ROOT / "README.md"
```

Read it in `PagesWorkflowTest.setUpClass`:

```python
cls.readme_source = README_PATH.read_text(encoding="utf-8")
```

Add this test method:

```python
def test_pinyin_dependency_is_installed_in_ci_and_documented_locally(self) -> None:
    install_step = next(
        step
        for step in self.jobs["build"]["steps"]
        if step.get("name") == "Install dependencies"
    )

    self.assertIn("pypinyin", install_step["run"])
    self.assertIn(
        "pip install mkdocs-material pyyaml openpyxl pypinyin",
        self.readme_source,
    )
```

- [ ] **Step 2: Run the workflow tests and verify RED**

Run:

```powershell
python -m unittest discover -s tests -p "test_pages_workflow.py" -v
```

Expected: FAIL in `test_pinyin_dependency_is_installed_in_ci_and_documented_locally` because neither installation command contains `pypinyin`.

- [ ] **Step 3: Add the minimal dependency declarations**

Change the workflow install command to:

```yaml
pip install mkdocs-material pyyaml openpyxl pypinyin
```

Change the README prerequisite command to:

```bash
pip install mkdocs-material pyyaml openpyxl pypinyin
```

- [ ] **Step 4: Run the workflow tests and verify GREEN**

Run:

```powershell
python -m unittest discover -s tests -p "test_pages_workflow.py" -v
```

Expected: all workflow tests PASS.

- [ ] **Step 5: Install the declared dependency for the remaining local test cycle**

Run:

```powershell
python -m pip install pypinyin
```

Expected: installation succeeds and `python -c "from pypinyin import lazy_pinyin; print(lazy_pinyin('拼音'))"` prints `['pin', 'yin']`.

- [ ] **Step 6: Commit the dependency declaration**

```powershell
git add tests/test_pages_workflow.py .github/workflows/pages.yml README.md
git commit -m "build: add pypinyin dependency"
```

### Task 2: Sort aggregation headings with one shared pinyin key

**Files:**
- Create: `tests/test_pinyin_aggregation_sorting.py`
- Modify: `scripts/build_cases.py`
- Regenerate: `docs/cases/by-university.md`
- Regenerate: `docs/cases/by-major.md`

**Interfaces:**
- Consumes: `pypinyin.Style`, `pypinyin.lazy_pinyin`, and the existing `uni_map` and `major_map` string keys.
- Produces: `pinyin_sort_key(name: str) -> tuple[str, str]`, used only as the `key` for the two aggregation-heading `sorted()` calls.

- [ ] **Step 1: Write focused failing unit and build-integration tests**

Create `tests/test_pinyin_aggregation_sorting.py` with the following test fixture and assertions:

```python
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
```

- [ ] **Step 2: Run the new tests and verify RED**

Run:

```powershell
python tests/test_pinyin_aggregation_sorting.py -v
```

Expected: two failures. The unit test reports that `pinyin_sort_key` is absent, and the integration test reports Unicode heading order instead of the expected pinyin order.

- [ ] **Step 3: Add the minimal shared sort key and use it in both loops**

Add the import near `openpyxl`:

```python
from pypinyin import Style, lazy_pinyin
```

Add this helper near the other normalization helpers:

```python
def pinyin_sort_key(name: str) -> tuple[str, str]:
    original_name = str(name)
    pinyin_name = "".join(
        lazy_pinyin(original_name, style=Style.NORMAL)
    ).casefold()
    return pinyin_name, original_name.casefold()
```

Replace only the university heading loop:

```python
for uni in sorted(uni_map.keys(), key=pinyin_sort_key):
```

Replace only the major heading loop:

```python
for maj in sorted(major_map.keys(), key=pinyin_sort_key):
```

- [ ] **Step 4: Run the focused tests and verify GREEN**

Run:

```powershell
python tests/test_pinyin_aggregation_sorting.py -v
```

Expected: 2 tests PASS.

- [ ] **Step 5: Run the existing regression suite**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: every test PASS with no errors or failures.

- [ ] **Step 6: Regenerate the committed aggregation pages**

Run:

```powershell
python scripts/build_cases.py
```

Expected: the command reports the generated case and aggregation pages. Inspect `git diff -- docs/cases/by-university.md docs/cases/by-major.md` and confirm that only heading blocks move; their counts and contents remain unchanged.

- [ ] **Step 7: Verify the generated headings and strict site build**

Run:

```powershell
rg -n "^## " docs/cases/by-university.md docs/cases/by-major.md
mkdocs build --strict
git diff --check
```

Expected: university headings begin in pinyin order with `北京航空航天大学`, `电子科技大学`, `复旦大学`, and `哈尔滨工业大学（深圳）`; major headings are in complete-name pinyin order; MkDocs exits 0; `git diff --check` prints nothing.

- [ ] **Step 8: Commit the tested behavior and generated pages**

```powershell
git add tests/test_pinyin_aggregation_sorting.py scripts/build_cases.py docs/cases/by-university.md docs/cases/by-major.md
git commit -m "feat: sort aggregation headings by pinyin"
```
