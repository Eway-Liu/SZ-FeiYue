# Case Aggregation Detail Links Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the bold identity text on the university, major, and experience aggregation pages link to the exact case detail page.

**Architecture:** Keep routing in the existing Markdown generator. Carry each case's stable `stem` through the three aggregation loops and emit page-relative links around only the bold identity text; no JavaScript or route changes are needed.

**Tech Stack:** Python 3.11+, `unittest`, `openpyxl`, MkDocs Material, Markdown.

## Global Constraints

- Only the bold identity text is clickable; review and advice bodies remain plain text.
- `Anonymous` and `Null` identities link to their own cases exactly like named identities.
- University and major pages use `case-<hash>/`; the root-level experience page uses `cases/case-<hash>/`.
- Preserve stable case URLs, ordering, filtering, null handling, layout, and deployment configuration.
- Add no JavaScript and no new runtime dependency.

---

## File Structure

- Create `tests/test_case_aggregation_links.py`: generator regression test with a temporary one-row workbook and content tree.
- Modify `scripts/build_cases.py`: carry `stem` through aggregation tuples and render links.
- Regenerate `docs/cases/by-university.md`, `docs/cases/by-major.md`, `docs/experience.md`, and the generator-managed timestamp in `docs/index.md`.

### Task 1: Add the generator regression test

**Files:**
- Create: `tests/test_case_aggregation_links.py`
- Reference: `scripts/build_cases.py:442-471`
- Reference: `scripts/build_cases.py:605-664`

**Interfaces:**
- Consumes: `python scripts/build_cases.py`, one `.xlsx` under `docs/`, homepage update markers, and `docs/seniors/`.
- Produces: `CaseAggregationLinksTest.test_identity_links_target_exact_case_and_exclude_review_text()` covering all three pages.

- [ ] **Step 1: Write the failing integration test**

Create `tests/test_case_aggregation_links.py`:

```python
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
```

- [ ] **Step 2: Verify the RED state**

Run `python -m unittest tests.test_case_aggregation_links -v`.

Expected: the subprocess succeeds, then the first `assertIn` fails because aggregation identities are not links. Fixture or import errors are not the intended failure.

- [ ] **Step 3: Commit the regression test**

```powershell
git add tests/test_case_aggregation_links.py
git commit -m "test: cover links from case aggregation pages"
```

### Task 2: Generate links around each identity

**Files:**
- Modify: `scripts/build_cases.py:442-471`
- Modify: `scripts/build_cases.py:605-664`
- Test: `tests/test_case_aggregation_links.py`

**Interfaces:**
- Consumes: case dictionaries containing `title`, `stem`, `nickname`, `university`, `major`, and review/advice strings.
- Produces: `[**<identity>**](<page-relative-case-route>)：<plain-body>`.

- [ ] **Step 1: Carry `stem` through the experience aggregation**

```python
shown.append((c["title"], adv, c["stem"]))

for title, adv, stem in shown:
    lines.append(f"- [**{title}**](cases/{case_link(stem)})：{adv}")
```

- [ ] **Step 2: Carry `stem` through the university aggregation**

```python
uni_map: dict[str, list[tuple[str, str, str]]] = defaultdict(list)
uni_map[uni].append((f"{nick} | {maj}", review, c["stem"]))

shown: list[tuple[str, str, str]] = []
for prefix, review, stem in items:
    txt = show_or_skip_null(review)
    if txt is not None:
        shown.append((prefix, txt, stem))

for prefix, txt, stem in shown:
    lines.append(f"- [**{prefix}**]({case_link(stem)})：{txt}")
```

- [ ] **Step 3: Carry `stem` through the major aggregation**

```python
major_map: dict[str, list[tuple[str, str, str]]] = defaultdict(list)
major_map[maj].append((f"{nick} | {uni}", review, c["stem"]))

shown: list[tuple[str, str, str]] = []
for prefix, review, stem in items:
    txt = show_or_skip_null(review)
    if txt is not None:
        shown.append((prefix, txt, stem))

for prefix, txt, stem in shown:
    lines.append(f"- [**{prefix}**]({case_link(stem)})：{txt}")
```

- [ ] **Step 4: Verify GREEN and run the complete suite**

```powershell
python -m unittest tests.test_case_aggregation_links -v
python -m unittest discover -s tests -v
```

Expected: the focused test and all existing tests pass. The exact assertions prove that each body begins after the link's closing parenthesis.

- [ ] **Step 5: Commit the generator implementation**

```powershell
git add scripts/build_cases.py
git commit -m "feat: link case identities from aggregation pages"
```

### Task 3: Regenerate and verify the checked-in site

**Files:**
- Modify: `docs/cases/by-university.md`
- Modify: `docs/cases/by-major.md`
- Modify: `docs/experience.md`
- Modify: `docs/index.md`
- Verify: `site/cases/by-university/index.html`
- Verify: `site/cases/by-major/index.html`
- Verify: `site/experience/index.html`

**Interfaces:**
- Consumes: the repository workbook and updated generator.
- Produces: checked-in linked Markdown and a strict MkDocs build under `site/`.

- [ ] **Step 1: Regenerate derived Markdown**

Run `python scripts/build_cases.py`.

Expected: the generator reports imported and built case counts and exits successfully.

- [ ] **Step 2: Review generated scope**

```powershell
git diff --check
git diff --stat
git diff -- docs/cases/by-university.md docs/cases/by-major.md docs/experience.md docs/index.md
```

Expected: three aggregation pages gain only the planned links; `docs/index.md` changes only in its generated timestamp. No detail page, raw submission, layout, or deployment change.

- [ ] **Step 3: Build the site strictly**

Run `python -m mkdocs build --strict`.

Expected: exit code `0` with no strict-mode warning or broken-link error.

- [ ] **Step 4: Inspect built HTML link boundaries**

Run:

```powershell
rg -n "case-[0-9a-f]{10}/" site/cases/by-university/index.html site/cases/by-major/index.html site/experience/index.html
```

Expected: identities are inside `<a>` elements targeting case routes; review/advice text follows `</a>`.

- [ ] **Step 5: Validate all three routes in a browser**

Start `python -m mkdocs serve --dev-addr 127.0.0.1:8000`. Open `/cases/by-university/`, `/cases/by-major/`, and `/experience/`; on each page click a named identity plus an `Anonymous` or `Null` identity where present.

Expected: each click opens the matching case detail title, body text is not linked, adjacent navigation works, and the browser console has no new error.

- [ ] **Step 6: Re-run final verification**

```powershell
python -m unittest discover -s tests -v
python -m mkdocs build --strict
git diff --check
```

Expected: all tests pass, strict build succeeds, and diff check emits no error.

- [ ] **Step 7: Commit generated content**

```powershell
git add docs/cases/by-university.md docs/cases/by-major.md docs/experience.md docs/index.md
git commit -m "docs: regenerate case aggregation links"
```
