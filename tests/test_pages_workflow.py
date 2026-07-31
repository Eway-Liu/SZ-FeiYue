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

    def test_cloudbase_cli_is_pinned(self) -> None:
        self.assertIn("npm install --global @cloudbase/cli@3.7.0", self.source)
        self.assertNotIn("@cloudbase/cli@latest", self.source)

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


if __name__ == "__main__":
    unittest.main()
