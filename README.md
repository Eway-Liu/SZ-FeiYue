# SZ-FeiYue (SZ FeiYue Handbook)

A community-maintained handbook for Shenzhen Middle School (深圳中学) students, built as a MkDocs site and published on GitHub Pages.

Website: https://Eway-Liu.github.io/SZ-FeiYue/

## What this project is

**SZ-FeiYue** is a structured, searchable collection of:

- Real admission cases (score/rank, admitted university/major, and personal reviews)
- Aggregated views by **university** and by **major**
- A curated “experience” page that aggregates advice across all cases
- Long-form alumni articles under “Seniors’ Posts” (学长学姐说)

> Note: The site content is primarily in Chinese.

## Repository layout

- `mkdocs.yml`: MkDocs configuration (navigation, theme, etc.)
- `docs/`: Site content
	- `index.md`: Home page (includes a “Last updated” marker block)
	- `submit.md`: Submission instructions (questionnaire + long-form article email)
	- `cases_raw/`: Source case front-matter files (auto-generated)
	- `cases/`: Rendered case pages + index pages (auto-generated)
	- `seniors/`: Long-form posts and an auto-generated index
- `scripts/build_cases.py`: Builds/updates generated pages from the submission spreadsheet
- `.github/workflows/pages.yml`: GitHub Pages build/deploy workflow

## How the content is generated

This repo uses an **offline spreadsheet** (exported from the submission form) as the single source of truth for admission cases.

The build pipeline is:

1. Put exactly **one** `.xlsx` file under `docs/` (the exported submissions spreadsheet)
2. Run `python scripts/build_cases.py`
3. The script will:
	 - Convert the spreadsheet into `docs/cases_raw/*.md` (front matter only)
	 - Generate case pages under `docs/cases/` (one page per case)
	 - Generate aggregation pages:
		 - `docs/cases/index.md`
		 - `docs/cases/by-university.md`
		 - `docs/cases/by-major.md`
		 - `docs/experience.md`
	 - Auto-generate `docs/seniors/index.md`
	 - Update the “Last updated” line in `docs/index.md` (between marker comments)

GitHub Actions runs the same sequence on every push to `main` and deploys the built site to GitHub Pages.

## Local preview

### Prerequisites

- Python 3.11+ recommended

Install dependencies:

```bash
python -m pip install --upgrade pip
pip install mkdocs-material pyyaml openpyxl
```

### Build generated pages

```bash
python scripts/build_cases.py
```

Important constraints:

- `scripts/build_cases.py` expects **exactly one** `.xlsx` file inside `docs/`.
- If the questionnaire header names change, the script may fail with a “missing required fields” error; update the header matching rules in `scripts/build_cases.py`.

### Serve the site

```bash
mkdocs serve
```

Then open the local URL shown in the terminal.

### Build static site (strict)

```bash
mkdocs build --strict
```

## Contributing

Contributions are welcome.

### Add or edit long-form alumni posts

- Add Markdown files to `docs/seniors/` (images are currently not supported in the submission workflow).
- Re-run `python scripts/build_cases.py` to regenerate `docs/seniors/index.md`.

### Update admission cases

Maintainers typically:

1. Export the latest submissions spreadsheet from the form
2. Replace the `.xlsx` under `docs/` (keep only one)
3. Run `python scripts/build_cases.py`
4. Preview with `mkdocs serve`
5. Commit and push to `main`

## Contact

For long-form article submissions, see `docs/submit.md` .

## Credits

Made by Eway Liu and contributors.