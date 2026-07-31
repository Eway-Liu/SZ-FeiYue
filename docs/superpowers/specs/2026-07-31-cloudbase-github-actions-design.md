# CloudBase GitHub Actions Deployment Design

## Goal

Every push to `main` and every manual workflow run builds the MkDocs site once, then independently publishes the same build output to GitHub Pages and CloudBase static hosting.

## Deployment targets

- GitHub Pages: keep the existing Pages deployment.
- CloudBase environment: `sms-teacher-ranking-d4bd8db87b87`.
- CloudBase path: `/sz-feiyue/`.
- CloudBase URL: `https://sms-teacher-ranking-d4bd8db87b87-1331414357.tcloudbaseapp.com/sz-feiyue/`.

## Workflow architecture

The workflow has three jobs:

1. `build` checks out the repository, installs Python dependencies, regenerates case pages, runs `mkdocs build --strict`, and uploads the resulting `site/` directory as both the GitHub Pages artifact and a reusable workflow artifact.
2. `deploy-pages` depends on `build` and deploys the Pages artifact with the existing official GitHub Pages action.
3. `deploy-cloudbase` depends on `build`, downloads the reusable artifact, installs the current CloudBase CLI, authenticates with repository secrets, uploads the complete artifact to `/sz-feiyue/`, and verifies the hosted file list.

The two deployment jobs are siblings. A CloudBase deployment failure does not prevent the GitHub Pages job from running, and a GitHub Pages failure does not prevent the CloudBase job from running.

## Case update behavior

Every CloudBase deployment uploads the complete generated `site/` directory. Existing files with the same paths are overwritten, so edits to any existing case are reflected on the next successful deployment. The workflow does not delete the whole `/sz-feiyue/` directory before upload, avoiding an unnecessary outage and preserving unrelated CloudBase paths.

Removing a file from Git does not automatically remove an old file with a different path from CloudBase. Exact deletion synchronization is outside the current requirement and can be added later if needed.

## Authentication and secrets

The workflow references these GitHub Actions repository secrets:

- `TENCENTCLOUD_SECRET_ID`
- `TENCENTCLOUD_SECRET_KEY`

The credentials must belong to a Tencent Cloud CAM user with only the permissions required to deploy CloudBase static hosting in the target environment. Secret values are never stored in the repository or printed intentionally.

The CloudBase environment ID and deployment path are non-secret workflow environment variables.

## Safety and privacy

- `mkdocs.yml` excludes `*.xlsx`, so the source submission spreadsheet is not published.
- The workflow deploys only the generated `site/` artifact.
- It does not delete or change `cloud-admin/`, `__auth/`, or other CloudBase paths.
- The existing CloudBase default domain remains the deployment target; custom-domain configuration is not part of this change.

## Verification

Before deployment, `mkdocs build --strict` must succeed and the build artifact must contain no `.xlsx` files.

After deployment, the workflow lists files under `/sz-feiyue/` and performs an HTTP request to the CloudBase site URL. A failed upload or verification request fails only the CloudBase deployment job.

Local validation will check workflow YAML syntax, expected job dependencies, secret references, deployment path, and a fresh strict MkDocs build.
