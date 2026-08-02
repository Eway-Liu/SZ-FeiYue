import json
from html.parser import HTMLParser
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest

import yaml


ROOT = Path(__file__).resolve().parents[1]
WORKFLOW_PATH = ROOT / ".github" / "workflows" / "pages.yml"
ROOT_ENTRY_PATH = ROOT / "cloudbase-root" / "index.html"
README_PATH = ROOT / "README.md"


class RedirectDocumentParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.refresh_targets: list[str] = []
        self.links: list[str] = []

    def handle_starttag(
        self, tag: str, attrs: list[tuple[str, str | None]]
    ) -> None:
        attributes = dict(attrs)
        if tag == "meta" and attributes.get("http-equiv", "").lower() == "refresh":
            content = attributes.get("content")
            if content is not None:
                self.refresh_targets.append(content)
        if tag == "a":
            href = attributes.get("href")
            if href is not None:
                self.links.append(href)


class PagesWorkflowTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.source = WORKFLOW_PATH.read_text(encoding="utf-8")
        cls.readme_source = README_PATH.read_text(encoding="utf-8")
        cls.workflow = yaml.safe_load(cls.source)
        cls.jobs = cls.workflow["jobs"]

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

    def test_cloudbase_cli_is_pinned(self) -> None:
        self.assertIn("npm install --global @cloudbase/cli@3.7.0", self.source)
        self.assertNotIn("@cloudbase/cli@latest", self.source)

    def test_cloudbase_upload_uses_safe_single_concurrency(self) -> None:
        steps = self.jobs["deploy-cloudbase"]["steps"]
        upload_step = next(
            step for step in steps if step.get("name") == "Upload complete site to CloudBase"
        )
        self.assertRegex(
            upload_step["run"], r"(?:^|\s)--concurrency\s+1(?:\s|$)"
        )

    def test_cloudbase_root_entry_redirects_to_full_site(self) -> None:
        parser = RedirectDocumentParser()
        parser.feed(ROOT_ENTRY_PATH.read_text(encoding="utf-8"))

        self.assertIn("0; url=/sz-feiyue/", parser.refresh_targets)
        self.assertIn("/sz-feiyue/", parser.links)

    def test_workflow_uploads_and_verifies_cloudbase_root_entry(self) -> None:
        steps = self.jobs["deploy-cloudbase"]["steps"]
        upload_step = next(
            step for step in steps if step.get("name") == "Upload CloudBase root entry"
        )

        self.assertIn("./cloudbase-root/index.html /index.html", upload_step["run"])
        self.assertIn('"/index.html"', self.source)

    def test_workflow_guards_private_sources_case_insensitively(self) -> None:
        self.assertIn("-iname '*.xlsx'", self.source)

    def test_workflow_verifies_deployed_files_and_content(self) -> None:
        self.assertIn("tcb hosting list", self.source)
        self.assertIn("JSON.parse", self.source)
        self.assertIn('`${process.env.CLOUDBASE_PATH}index.html`', self.source)
        self.assertIn('`${process.env.CLOUDBASE_PATH}404.html`', self.source)
        self.assertIn("sha256sum site/index.html", self.source)
        self.assertIn("sha256sum", self.source)
        self.assertIn("--retry 5", self.source)
        self.assertIn("--max-time 20", self.source)
        self.assertIn('test "$local_sha256" = "$remote_sha256"', self.source)

    def test_cloudbase_file_verification_script_executes(self) -> None:
        steps = self.jobs["deploy-cloudbase"]["steps"]
        verification_step = next(
            step for step in steps if step.get("name") == "Verify CloudBase files"
        )
        run = verification_step["run"]
        script = run.split("node - <<'NODE'\n", 1)[1].rsplit("\nNODE", 1)[0]
        node = shutil.which("node")
        self.assertIsNotNone(node, "Node.js is required to validate the workflow script")

        listing = {
            "files": [
                {
                    "url": "https://example.test/sz-feiyue/index.html",
                    "key": "sz-feiyue/404.html",
                },
                {"key": "index.html"},
            ]
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            Path(temp_dir, "hosting-files.json").write_text(
                json.dumps(listing), encoding="utf-8"
            )
            env = os.environ.copy()
            env["CLOUDBASE_PATH"] = "/sz-feiyue/"
            result = subprocess.run(
                [node, "-"],
                input=script,
                text=True,
                cwd=temp_dir,
                env=env,
                capture_output=True,
                check=False,
            )

        self.assertEqual(result.returncode, 0, result.stderr)


if __name__ == "__main__":
    unittest.main()
