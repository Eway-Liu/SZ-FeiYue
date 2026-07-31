# CloudBase GitHub Actions Deployment Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Publish every successful `main` build to GitHub Pages and CloudBase independently, using the same complete MkDocs artifact so all existing case pages are overwritten with their newest versions.

**Architecture:** The existing build job remains the single producer of `site/`. It uploads both the GitHub Pages artifact and a short-lived generic artifact; sibling deployment jobs consume those artifacts independently. The CloudBase job authenticates with GitHub repository secrets, uploads the complete artifact to `/sz-feiyue/`, then verifies the remote file listing and public URL.

**Tech Stack:** GitHub Actions, Python 3.11, MkDocs Material, CloudBase CLI v3 (`@cloudbase/cli@latest`), Python `unittest` and PyYAML.

## Global Constraints

- Target branch: `main`; retain `workflow_dispatch` support.
- Target CloudBase environment: `sms-teacher-ranking-d4bd8db87b87`.
- Target CloudBase path: `/sz-feiyue/`.
- Target URL: `https://sms-teacher-ranking-d4bd8db87b87-1331414357.tcloudbaseapp.com/sz-feiyue/`.
- Upload the complete `site/` directory on every deployment; same-path case files must be overwritten.
- Do not clear `/sz-feiyue/` before upload and do not touch `cloud-admin/`, `__auth/`, or unrelated paths.
- Never publish `.xlsx` files.
- Read credentials only from `TENCENTCLOUD_SECRET_ID` and `TENCENTCLOUD_SECRET_KEY` GitHub repository secrets.
- GitHub Pages and CloudBase deployment failures must remain isolated from each other.

---

## File map

- Create `tests/test_pages_workflow.py`: structural regression tests for the dual-deployment workflow, target environment/path, secret references, privacy guard, and independent job dependencies.
- Modify `.github/workflows/pages.yml`: produce one site artifact and deploy it independently to GitHub Pages and CloudBase.
- Modify `README.md`: document automatic deployment behavior and the two required GitHub repository secrets.

### Task 1: Add independent CloudBase deployment to the Pages workflow

**Files:**
- Create: `tests/test_pages_workflow.py`
- Modify: `.github/workflows/pages.yml`
- Modify: `README.md`
- Test: `tests/test_pages_workflow.py`

**Interfaces:**
- Consumes: MkDocs output directory `site/`; GitHub secrets `TENCENTCLOUD_SECRET_ID` and `TENCENTCLOUD_SECRET_KEY`.
- Produces: artifact `cloudbase-site`; GitHub Actions job `deploy-cloudbase`; complete upload to `/sz-feiyue/` in environment `sms-teacher-ranking-d4bd8db87b87`.

- [ ] **Step 1: Write the workflow contract test**

Create `tests/test_pages_workflow.py`:

```python
from pathlib import Path
import unittest

import yaml


ROOT = Path(__file__).resolve().parents[1]
WORKFLOW_PATH = ROOT / ".github" / "workflows" / "pages.yml"


class PagesWorkflowTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.source = WORKFLOW_PATH.read_text(encoding="utf-8")
        cls.workflow = yaml.safe_load(cls.source)
        cls.jobs = cls.workflow["jobs"]

    def test_deployments_are_independent_siblings(self) -> None:
        self.assertEqual(self.jobs["deploy"]["needs"], "build")
        self.assertEqual(self.jobs["deploy-cloudbase"]["needs"], "build")

    def test_build_publishes_both_artifact_types(self) -> None:
        steps = self.jobs["build"]["steps"]
        uses = {step.get("uses") for step in steps}
        self.assertIn("actions/upload-pages-artifact@v3", uses)
        self.assertIn("actions/upload-artifact@v4", uses)
        cloudbase_artifact = next(
            step for step in steps if step.get("uses") == "actions/upload-artifact@v4"
        )
        self.assertEqual(cloudbase_artifact["with"]["name"], "cloudbase-site")
        self.assertEqual(cloudbase_artifact["with"]["path"], "site")

    def test_cloudbase_target_and_credentials_are_explicit(self) -> None:
        job = self.jobs["deploy-cloudbase"]
        self.assertEqual(job["env"]["CLOUDBASE_ENV_ID"], "sms-teacher-ranking-d4bd8db87b87")
        self.assertEqual(job["env"]["CLOUDBASE_PATH"], "/sz-feiyue/")
        self.assertIn("secrets.TENCENTCLOUD_SECRET_ID", self.source)
        self.assertIn("secrets.TENCENTCLOUD_SECRET_KEY", self.source)
        self.assertIn('tcb hosting deploy ./site "$CLOUDBASE_PATH"', self.source)

    def test_workflow_guards_private_sources_and_verifies_deployment(self) -> None:
        self.assertIn("-name '*.xlsx'", self.source)
        self.assertIn("tcb hosting list", self.source)
        self.assertIn("curl --fail", self.source)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the new test and verify RED**

Run:

```powershell
C:\Users\Eway\.conda\envs\SZ-feiyue\python.exe -m unittest tests/test_pages_workflow.py -v
```

Expected: FAIL because `deploy-cloudbase`, `actions/upload-artifact@v4`, CloudBase secret references, and the `.xlsx` output guard do not yet exist in the workflow.

- [ ] **Step 3: Implement the dual-deployment workflow**

Update `.github/workflows/pages.yml` with these structural changes:

1. Rename the workflow to `Deploy to GitHub Pages and CloudBase`.
2. Rename the concurrency group to `site-deploy`.
3. After `mkdocs build --strict`, add:

```yaml
      - name: Verify private sources are excluded
        run: |
          if find site -type f -name '*.xlsx' | grep -q .; then
            echo "Build output contains private .xlsx files" >&2
            exit 1
          fi

      - name: Upload CloudBase artifact
        uses: actions/upload-artifact@v4
        with:
          name: cloudbase-site
          path: site
          if-no-files-found: error
          retention-days: 1
```

4. Keep the existing `deploy` job for GitHub Pages.
5. Add the sibling CloudBase job:

```yaml
  deploy-cloudbase:
    name: Deploy to CloudBase
    runs-on: ubuntu-latest
    needs: build
    timeout-minutes: 10
    env:
      CLOUDBASE_ENV_ID: sms-teacher-ranking-d4bd8db87b87
      CLOUDBASE_PATH: /sz-feiyue/
      CLOUDBASE_URL: https://sms-teacher-ranking-d4bd8db87b87-1331414357.tcloudbaseapp.com/sz-feiyue/
    steps:
      - name: Download site artifact
        uses: actions/download-artifact@v4
        with:
          name: cloudbase-site
          path: site

      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version: "20"

      - name: Install CloudBase CLI
        run: |
          npm install --global @cloudbase/cli@latest
          tcb --version

      - name: Authenticate CloudBase CLI
        env:
          TENCENTCLOUD_SECRET_ID: ${{ secrets.TENCENTCLOUD_SECRET_ID }}
          TENCENTCLOUD_SECRET_KEY: ${{ secrets.TENCENTCLOUD_SECRET_KEY }}
        run: |
          test -n "$TENCENTCLOUD_SECRET_ID"
          test -n "$TENCENTCLOUD_SECRET_KEY"
          tcb login --apiKeyId "$TENCENTCLOUD_SECRET_ID" --apiKey "$TENCENTCLOUD_SECRET_KEY"

      - name: Upload complete site to CloudBase
        run: tcb hosting deploy ./site "$CLOUDBASE_PATH" --env-id "$CLOUDBASE_ENV_ID" --yes

      - name: Verify CloudBase files
        run: tcb hosting list --env-id "$CLOUDBASE_ENV_ID" --json

      - name: Verify CloudBase URL
        run: |
          curl --fail --show-error --silent \
            --retry 5 --retry-delay 5 \
            --output /dev/null \
            "${CLOUDBASE_URL}?build=${GITHUB_SHA}"
```

The CloudBase CLI login and hosting commands follow the current official CLI guidance: `tcb login --apiKeyId ... --apiKey ...` and `tcb hosting deploy <localPath> <cloudPath> --env-id <envId> --yes`.

- [ ] **Step 4: Document the repository secrets and deployment behavior**

Append to the existing `CloudBase deployment` section in `README.md`:

```markdown
### Automatic deployment

Pushes to `main` and manual runs of `.github/workflows/pages.yml` build the site once and deploy the same complete artifact independently to GitHub Pages and CloudBase. Uploading the complete artifact overwrites every existing case with the same path, so edits to case content are reflected after the next successful run.

Configure these GitHub repository secrets under **Settings → Secrets and variables → Actions**:

- `TENCENTCLOUD_SECRET_ID`
- `TENCENTCLOUD_SECRET_KEY`

Use credentials from a Tencent Cloud CAM user limited to the permissions needed for CloudBase static hosting in environment `sms-teacher-ranking-d4bd8db87b87`.
```

- [ ] **Step 5: Run the contract test and verify GREEN**

Run:

```powershell
C:\Users\Eway\.conda\envs\SZ-feiyue\python.exe -m unittest tests/test_pages_workflow.py -v
```

Expected: four tests PASS with no failures or errors.

- [ ] **Step 6: Run a fresh production build and privacy check**

Build in a temporary copy so generated case files do not alter the user's working tree:

```powershell
$scratch = Join-Path ([System.IO.Path]::GetTempPath()) ('sz-feiyue-actions-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch | Out-Null
Copy-Item -LiteralPath docs -Destination (Join-Path $scratch 'docs') -Recurse
Copy-Item -LiteralPath scripts -Destination (Join-Path $scratch 'scripts') -Recurse
Copy-Item -LiteralPath mkdocs.yml -Destination (Join-Path $scratch 'mkdocs.yml')
C:\Users\Eway\.conda\envs\SZ-feiyue\python.exe (Join-Path $scratch 'scripts\build_cases.py')
C:\Users\Eway\.conda\envs\SZ-feiyue\Scripts\mkdocs.exe build --strict --config-file (Join-Path $scratch 'mkdocs.yml') --site-dir (Join-Path $scratch 'site')
if (Get-ChildItem -Recurse -File (Join-Path $scratch 'site') -Filter *.xlsx) { throw 'Build contains .xlsx files' }
```

Expected: commands exit 0; MkDocs reports a successful strict build; the privacy check produces no error.

- [ ] **Step 7: Validate YAML, diff hygiene, and exact scope**

Run:

```powershell
C:\Users\Eway\.conda\envs\SZ-feiyue\python.exe -c "import pathlib,yaml; yaml.safe_load(pathlib.Path('.github/workflows/pages.yml').read_text(encoding='utf-8')); print('YAML_OK')"
git diff --check
git diff -- .github/workflows/pages.yml tests/test_pages_workflow.py README.md
```

Expected: `YAML_OK`; `git diff --check` exits 0; the diff contains only the workflow, its test, and the documented secret/setup changes for this task.

- [ ] **Step 8: Commit the implementation**

```powershell
git add -- .github/workflows/pages.yml tests/test_pages_workflow.py README.md docs/superpowers/plans/2026-07-31-cloudbase-github-actions.md
git commit -m "ci: deploy site to CloudBase"
```

Expected: a commit containing only the automatic CloudBase deployment implementation, test, README update, and this implementation plan. Existing user-owned case and script changes remain unstaged.

## Post-implementation user action

Before the first automated CloudBase run, add `TENCENTCLOUD_SECRET_ID` and `TENCENTCLOUD_SECRET_KEY` to the GitHub repository's Actions secrets. The workflow cannot authenticate until both secrets exist. The next push to `main` or manual workflow dispatch will exercise the real CloudBase deployment; local validation cannot substitute for that remote run.
