# Aggregation Format Copy Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the format explanations on “按院校”“按专业”“查看经验” use the same separate `案例格式：…` structure as “所有案例”.

**Architecture:** Keep all aggregation data, links, titles, and sorting unchanged. Update only the three fixed Markdown-copy blocks in `scripts/build_cases.py`, cover the generated behavior with the existing end-to-end generator test, then regenerate the three target pages.

**Tech Stack:** Python 3, `unittest`, `openpyxl`, Markdown, MkDocs Material

## Global Constraints

- “按院校” explanation must be exactly `展示**院校评价**的聚合结果。院校按拼音排序。`, followed by a blank line, the `案例格式：` prefix, inline-code text `昵称 | 专业：评价`, and `。`.
- “按专业” explanation must be exactly `展示**专业评价**的聚合结果。专业按拼音排序。`, followed by a blank line, the `案例格式：` prefix, inline-code text `昵称 | 院校：评价`, and `。`.
- “查看经验” explanation must be exactly `本页汇总所有已收录案例的 **给学弟学妹的建议**。排序不分先后。`, followed by a blank line, the `案例格式：` prefix, inline-code text `昵称 | 高考分数 | 院校 | 专业：评价`, and `。`.
- Do not change list content, link targets, review/advice text, sort behavior, case URLs, or empty states.
- Do not modify “所有案例”, case detail pages, CSS, JavaScript, dependencies, routes, deployment, or publishing.

---

### Task 1: Unify generated aggregation format explanations

**Files:**
- Modify: `tests/test_case_aggregation_links.py:68-86`
- Modify: `scripts/build_cases.py:457-458,629-664`
- Modify (generated): `docs/cases/by-university.md`
- Modify (generated): `docs/cases/by-major.md`
- Modify (generated): `docs/experience.md`

**Interfaces:**
- Consumes: the existing spreadsheet fixture and real `scripts/build_cases.py` subprocess used by `CaseAggregationLinksTest`.
- Produces: three generated Markdown pages whose purpose/sort copy is separated from an exact `案例格式：…` paragraph, with all existing entries and links preserved.

- [ ] **Step 1: Write the failing generated-output assertions**

After reading `university_page`, `major_page`, and `experience_page`, add literal behavior assertions:

```python
self.assertIn(
    "展示**院校评价**的聚合结果。院校按拼音排序。\n\n"
    "案例格式：`昵称 | 专业：评价`。",
    university_page,
)
self.assertIn(
    "展示**专业评价**的聚合结果。专业按拼音排序。\n\n"
    "案例格式：`昵称 | 院校：评价`。",
    major_page,
)
self.assertIn(
    "本页汇总所有已收录案例的 **给学弟学妹的建议**。排序不分先后。\n\n"
    "案例格式：`昵称 | 高考分数 | 院校 | 专业：评价`。",
    experience_page,
)
```

These assertions exercise the real generator against a controlled workbook. The production regression they catch is recombining the purpose/sort text with the format text or emitting the wrong format fields.

- [ ] **Step 2: Run the focused test and verify RED**

Run:

```powershell
D:\ANACONDA\python.exe -B -m unittest tests/test_case_aggregation_links.py -v
```

Expected: FAIL at the first new assertion because the current generated page still contains `格式为` in a single paragraph.

- [ ] **Step 3: Implement the three minimal generator-copy changes**

Replace the “查看经验” copy block with:

```python
lines.append("本页汇总所有已收录案例的 **给学弟学妹的建议**。排序不分先后。")
lines.append("")
lines.append("案例格式：`昵称 | 高考分数 | 院校 | 专业：评价`。")
lines.append("")
```

Replace the “按院校” copy block with:

```python
lines.append("展示**院校评价**的聚合结果。院校按拼音排序。")
lines.append("")
lines.append("案例格式：`昵称 | 专业：评价`。")
lines.append("")
```

Replace the “按专业” copy block with:

```python
lines.append("展示**专业评价**的聚合结果。专业按拼音排序。")
lines.append("")
lines.append("案例格式：`昵称 | 院校：评价`。")
lines.append("")
```

Do not modify any loop, tuple, sort key, link helper, review/advice value, or empty-state branch.

- [ ] **Step 4: Run the focused test and verify GREEN**

Run:

```powershell
D:\ANACONDA\python.exe -B -m unittest tests/test_case_aggregation_links.py -v
```

Expected: PASS, including the existing assertions for exact links, review text, overview years, and non-overview year exclusion.

- [ ] **Step 5: Regenerate repository content**

Run with a Python environment that provides `openpyxl`, `PyYAML`, and IANA timezone data:

```powershell
python scripts/build_cases.py
```

Expected: exit code `0`; the three target generated pages receive the separated format paragraphs. If `docs/index.md` is refreshed as an unrelated timestamp side effect, restore only line 7 to `最后更新时间：2026/08/01  17:47:51` with `apply_patch`.

- [ ] **Step 6: Inspect generated scope and content**

Run:

```powershell
git -c safe.directory=D:/SZ-FeiYue status --short
git -c safe.directory=D:/SZ-FeiYue diff --check
git -c safe.directory=D:/SZ-FeiYue diff -- scripts/build_cases.py tests/test_case_aggregation_links.py docs/cases/by-university.md docs/cases/by-major.md docs/experience.md docs/index.md
```

Expected: task changes are limited to the generator, test, and the three target generated pages; `docs/index.md` has no remaining diff. Existing list items and their link targets are unchanged.

- [ ] **Step 7: Run complete automated verification**

Run:

```powershell
D:\ANACONDA\python.exe -B -m unittest discover -s tests -v
python -m mkdocs build --strict
```

Expected: all 11 tests PASS with zero failures/errors, and MkDocs strict build exits `0` with no new task-introduced warning.

- [ ] **Step 8: Verify the three rendered pages in a trusted browser**

Serve the site locally and use the installed trusted Browser client. Open:

```text
http://127.0.0.1:8000/sz-feiyue/cases/by-university/
http://127.0.0.1:8000/sz-feiyue/cases/by-major/
http://127.0.0.1:8000/sz-feiyue/experience/
```

For each route, verify the purpose/sort sentence and separate `案例格式` paragraph render in order, existing entries remain visible and clickable, and browser console error/warning logs are empty.

- [ ] **Step 9: Commit the verified change**

```powershell
git -c safe.directory=D:/SZ-FeiYue add scripts/build_cases.py tests/test_case_aggregation_links.py docs/cases/by-university.md docs/cases/by-major.md docs/experience.md
git -c safe.directory=D:/SZ-FeiYue commit -m "fix: unify aggregation format copy"
git -c safe.directory=D:/SZ-FeiYue status --short
```

Expected: commit succeeds and no uncommitted task changes remain.
