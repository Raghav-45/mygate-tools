# mygate-tools

Three Chrome (Manifest V3) extensions for the MyGate admin dashboard:

| Tool | Purpose | Folder |
|------|---------|--------|
| **MyGate Dump Tool** | Autonomous multi-year ticket dump exporter (bypasses MyGate's 1-year export restriction) | `mygate-dump-tool` |
| **MyGate Report Tool** | Generate pending tickets report from MyGate Dashboard GraphQL API | `mygate-report-tool` |
| **MyGate Summary Tool** | Complaint summary sheet automation | `mygate-summary-tool` |

## Trigger a release (GitHub)

The GitHub Actions workflow `Manual Release` packages the 3 folders into `.zip` assets
and attaches them to a GitHub Release. To trigger it:

1. Go to the repo on GitHub → **Actions** tab.
2. Select **Manual Release** on the left → **Run workflow**.
3. Fill in:
   - **Version / tag** — e.g. `1.0.1` (prefixed with `v` automatically).
   - **Release notes** — what changed.
   - **Mark as pre-release** — toggle if needed.
4. Click **Run workflow**.

Once it finishes, open the generated Release (tagged `v<version>`) and download the `.zip` assets:
`mygate-dump-tool.zip`, `mygate-report-tool.zip`, `mygate-summary-tool.zip`.

## Build locally (preview before releasing)

```sh
./build_release.sh
```

This creates the same three `.zip` files in `./dist` so you can verify them locally.

## Install an extension in Chrome

The `.zip` files are unpacked extensions, installed via **Load unpacked**:

1. Download and **unzip** the desired `.zip` into its own folder (e.g. `mygate-dump-tool/`).
2. Open Chrome → `chrome://extensions`.
3. Turn on **Developer mode** (top-right corner).
4. Click **Load unpacked** and select the unzipped folder.
5. Pin the extension to the toolbar and sign in to MyGate before using it.

Repeat for each tool you want installed.
