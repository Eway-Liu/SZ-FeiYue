# Case Overview Exam Year Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a case-format explanation and the Gaokao year to every link on the “案例总览” page without changing any other case display.

**Architecture:** Keep the existing shared case title unchanged. Carry normalized `exam_year` and `gaokao_score` values in the transient case dictionaries, then compose a dedicated overview label only while generating `docs/cases/index.md`.

**Tech Stack:** Python 3, `unittest`, `openpyxl`, Markdown, MkDocs Material

## Global Constraints

- Overview label format must be exactly `昵称 | 高考年份 | 高考分数 | 院校 | 专业`.
- Format copy in the generated Markdown must be exactly: `案例格式：` followed by the inline-code text `昵称 | 高考年份 | 高考分数 | 院校 | 专业` and the Chinese full stop `。`.
- High school exam year must appear only on “案例总览”; detail, university, major, and experience displays remain unchanged.
- Preserve existing case URLs and sorting behavior.
- Do not add CSS, JavaScript, routes, dependencies, or deployment changes.
- Do not publish or deploy the site.

---

### Task 1: Specify the overview-only display behavior

**Files:**
- Modify: `tests/test_case_aggregation_links.py:69-101`

**Interfaces:**
- Consumes: generated Markdown written by `scripts/build_cases.py` into `docs/cases/index.md`, `docs/cases/by-university.md`, `docs/cases/by-major.md`, and `docs/experience.md`.
- Produces: a regression test that requires the overview format explanation and overview-only year label while preserving the existing labels on the other aggregation pages.

- [ ] **Step 1: Update the generation test with the new expectations**

Replace the existing lookup of the old overview title inside the `for` loop with the following expectations, while retaining the existing link-target and review-text assertions:

```python
self.assertIn(
    "案例格式：`昵称 | 高考年份 | 高考分数 | 院校 | 专业`。",
    case_index,
)

for nickname, score, _, university, major, university_review, major_review, advice, _ in cases:
    title = f"{nickname} | {score} | {university} | {major}"
    overview_title = f"{nickname} | 2026 | {score} | {university} | {major}"
    self.assertIn(overview_title, case_stems)
    stem = case_stems[overview_title]
    self.assertNotIn(f"**{nickname} | 2026 |", university_page)
    self.assertNotIn(f"**{nickname} | 2026 |", major_page)
    self.assertNotIn(f"**{nickname} | 2026 |", experience_page)
```

Keep using `title` in the existing `experience_page` assertion so that the test explicitly proves the shared title remains unchanged.

- [ ] **Step 2: Run the focused test and verify RED**

Run:

```powershell
python -m unittest tests/test_case_aggregation_links.py -v
```

Expected: FAIL because `docs/cases/index.md` generated in the temporary fixture does not contain the format explanation or `2026` in its case labels.

---

### Task 2: Generate overview labels with the exam year

**Files:**
- Modify: `scripts/build_cases.py:562-604`
- Test: `tests/test_case_aggregation_links.py`

**Interfaces:**
- Consumes: `display(meta.get(...))`, `display_nickname(meta.get("nickname"))`, the existing stable `stem`, and the unchanged shared `title`.
- Produces: transient case dictionary keys `exam_year: str` and `gaokao_score: str`, plus overview-only link labels assembled in the required field order.

- [ ] **Step 1: Carry normalized year and score values in both case-building branches**

Add these keys to both the validation-error case dictionary and the normal case dictionary:

```python
"exam_year": display(meta_err.get("exam_year")),
"gaokao_score": display(meta_err.get("gaokao_score")),
```

Use `meta` instead of `meta_err` in the normal branch:

```python
"exam_year": display(meta.get("exam_year")),
"gaokao_score": display(meta.get("gaokao_score")),
```

- [ ] **Step 2: Add the format explanation and overview-only title assembly**

Change only the “案例总览” generation block:

```python
lines.append(f"当前收录：**{len(cases_sorted)}** 条。点击标题进入详情页。排序不分先后。")
lines.append("")
lines.append("案例格式：`昵称 | 高考年份 | 高考分数 | 院校 | 专业`。")
lines.append("")
for c in cases_sorted:
    overview_title = " | ".join(
        [
            c["nickname"],
            c["exam_year"],
            c["gaokao_score"],
            c["university"],
            c["major"],
        ]
    )
    lines.append(f"- [{overview_title}]({case_link(c['stem'])})")
```

Do not modify `title_of()`, `render_case_page()`, the aggregation maps, `case_link()`, or the sort key.

- [ ] **Step 3: Run the focused test and verify GREEN**

Run:

```powershell
python -m unittest tests/test_case_aggregation_links.py -v
```

Expected: PASS; the overview labels include `2026`, and the university, major, and experience labels do not.

- [ ] **Step 4: Inspect the scoped diff**

Run:

```powershell
git -c safe.directory=D:/SZ-FeiYue diff --check
git -c safe.directory=D:/SZ-FeiYue diff -- scripts/build_cases.py tests/test_case_aggregation_links.py
```

Expected: no whitespace errors and no changes outside the overview data plumbing and test assertions.

- [ ] **Step 5: Commit the tested implementation**

```powershell
git -c safe.directory=D:/SZ-FeiYue add scripts/build_cases.py tests/test_case_aggregation_links.py
git -c safe.directory=D:/SZ-FeiYue commit -m "feat: show exam year in case overview"
```

---

### Task 3: Regenerate content and verify the rendered site

**Files:**
- Modify (generated): `docs/cases/index.md`
- Verify unchanged unless the source spreadsheet legitimately regenerates identical content: `docs/cases/*.md`, `docs/cases_raw/*.md`, `docs/cases/by-university.md`, `docs/cases/by-major.md`, `docs/experience.md`, `docs/index.md`, `docs/seniors/index.md`

**Interfaces:**
- Consumes: the repository’s single `docs/*.xlsx` source file and the updated generator.
- Produces: committed generated overview Markdown and fresh test/build/browser evidence.

- [ ] **Step 1: Regenerate site content**

Run:

```powershell
python scripts/build_cases.py
```

Expected: command exits `0`; `docs/cases/index.md` gains the format line and exam year in each case label.

- [ ] **Step 2: Confirm generated scope**

Run:

```powershell
git -c safe.directory=D:/SZ-FeiYue status --short
git -c safe.directory=D:/SZ-FeiYue diff --check
git -c safe.directory=D:/SZ-FeiYue diff -- docs/cases/index.md docs/cases/by-university.md docs/cases/by-major.md docs/experience.md
```

Expected: the intended visible change is confined to `docs/cases/index.md`; other aggregation pages have no year-format change.

- [ ] **Step 3: Run the complete automated test suite**

Run:

```powershell
python -m unittest discover -s tests -v
```

Expected: all tests PASS with zero failures and zero errors.

- [ ] **Step 4: Build the static site strictly**

Run:

```powershell
python -m mkdocs build --strict
```

Expected: exit code `0` with no strict-mode build error.

- [ ] **Step 5: Validate the rendered overview and adjacent pages in a browser**

Start the local server in a hidden process:

```powershell
Start-Process python -ArgumentList '-m','mkdocs','serve','--dev-addr','127.0.0.1:8000' -WindowStyle Hidden -PassThru
```

Use browser automation to open `http://127.0.0.1:8000/cases/` and verify:

- the “案例格式” line appears between the collection summary and list;
- the first case label has five fields in the required order;
- clicking the first link opens its existing case detail page;
- `/cases/by-university/`, `/cases/by-major/`, and `/experience/` retain their prior labels without the extra year;
- the browser console contains no new error.

- [ ] **Step 6: Commit the regenerated overview**

```powershell
git -c safe.directory=D:/SZ-FeiYue add docs/cases/index.md
git -c safe.directory=D:/SZ-FeiYue commit -m "docs: regenerate case overview with exam year"
```

- [ ] **Step 7: Run final repository checks**

Run:

```powershell
git -c safe.directory=D:/SZ-FeiYue status --short
git -c safe.directory=D:/SZ-FeiYue log -3 --oneline
```

Expected: no uncommitted task changes remain; the plan, implementation, and regenerated content commits are visible.
