# mygate-tools

Five tools for the MyGate admin dashboard — all Chrome (Manifest V3) extensions:

| Tool | Purpose | Folder |
|------|---------|--------|
| **MyGate Handoff Tool** | Two independent steps: generate and download selected original workbooks, then prepare the Outlook draft. Summary is off by default. No saved-file or file-import mode. | `mygate-handoff-tool` |
| **MyGate Dump Tool** | Autonomous multi-year ticket dump exporter (bypasses MyGate's 1-year export restriction) | `mygate-dump-tool` |
| **MyGate Report Tool** | Generate pending tickets report from MyGate Dashboard GraphQL API | `mygate-report-tool` |
| **MyGate Summary Tool** | Complaint summary sheet automation | `mygate-summary-tool` |
| **MyGate Analysis Tool** | One-click converter: turns a master helpdesk dump into the standardized "Mygate Complaints Dashboard as on DD.MM.YY" workbook — Interactive & Overview KPI sheets, 10 live charts, Data / Analysis sheets and a hidden Calc engine; byte-identical to the manual dashboard it replaces | `mygate-analysis-tool` |

## MyGate Handoff Tool (recommended)

1. Sign in to MyGate, choose **Report From** and **Report Till**, then click
   **Generate & Download**. Reports run in this order: **Pending Tickets Report →
   Master Dump → Analysis Dashboard → Summary Sheet**. Each has an on/off toggle;
   Summary is off by default, so the default batch downloads three files. Analysis
   uses the newly generated dump: turning Dump off also turns Analysis off;
   turning Analysis on enables Dump. Summary covers the month of Report Till,
   using the standalone Summary tool's existing rules. The API delay defaults to
   **1.0 second**, adjustable from 1.0–5.0 seconds in **0.1-second** increments.
2. **Email is disabled in v1.0.0.** The Outlook card has a blocking overlay and
   disabled controls. Attach the downloaded reports manually for now. A later
   extension release can enable it by changing `release-features.js`; saved settings
   cannot enable it, and the background rejects draft requests in this build.

Step 1 never opens Outlook. Report selections are locked until settings load and
saved immediately when toggled. A run uses its explicit current selection, so a
report-only run never falls back to saved/default tools. The temporary batch
for step 2 is replaced on each generation and cleared when Chrome starts again;
there is no saved-file picker, import, or reuse mode. Downloaded files remain on disk.
Email settings persist. The existing standalone tools are unchanged.

Build with `./build_release.sh`, then unzip `dist/mygate-handoff-tool.zip` and load
the folder at `chrome://extensions` using **Load unpacked**. You can also load
`mygate-handoff-tool/` directly from this repository.

The original generator bodies are embedded without changes in isolated factories.
`node scripts/sync_handoff_sources.mjs` refreshes those copies and the analysis assets
from the standalone tools; the release build runs this automatically.

Run the offline handoff integration checks with:

```sh
node mygate-handoff-tool/tests/run_tests.mjs
```

These build actual workbooks using synthetic API responses and check dump-to-analysis
handoff, independent steps, email retries, failed downloads, partial dumps and aborts.
Live MyGate and Outlook verification still requires signed-in browser accounts.

For personal distribution without Chrome Web Store registration, use the existing
GitHub **Manual Release** workflow to publish the extension ZIP. Users replace the
contents of the same unpacked extension folder and click **Reload** at
`chrome://extensions`. Pushing code or publishing a GitHub release alone does not
update their installed extension. The [Windows updater script](scripts/Update-MyGateExtensions.ps1)
replaces files for extension folders already installed from `Downloads`; Chrome
still needs to reload each extension. Self-hosted CRX automatic updates on
Windows/macOS require managed enterprise policies; Linux also supports self-hosted
extension updates.

To run the updater from an elevated Command Prompt after the script is available
in the public repository:

```cmd
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$p=Join-Path $env:TEMP 'Update-MyGateExtensions.ps1'; Invoke-WebRequest 'https://raw.githubusercontent.com/Raghav-45/mygate-tools/main/scripts/Update-MyGateExtensions.ps1' -OutFile $p; & $p"
```

It updates only matching MyGate extension folders already present directly in the
current user's `Downloads` folder. It does not install extensions into Chrome;
after it finishes, open `chrome://extensions` and click **Reload** for each updated
extension. Close Chrome before running if extension files are in use.

## Trigger a release (GitHub)

The GitHub Actions workflow `Manual Release` packages the 5 extension folders into `.zip` assets
and attaches them to a GitHub Release. To trigger it:

1. Go to the repo on GitHub → **Actions** tab.
2. Select **Manual Release** on the left → **Run workflow**.
3. Fill in:
   - **Version / tag** — use `0.0.3` for this release (prefixed with `v` automatically).
   - **Release notes** — what changed.
   - **Mark as pre-release** — toggle if needed.
4. Click **Run workflow**.

Once it finishes, open the generated Release (tagged `v<version>`) and download the `.zip` assets:
`mygate-dump-tool.zip`, `mygate-report-tool.zip`, `mygate-summary-tool.zip`,
`mygate-analysis-tool.zip`, `mygate-handoff-tool.zip`.

## Build locally (preview before releasing)

```sh
./build_release.sh
```

This creates the same five `.zip` files in `./dist` so you can verify them locally.

## Install an extension in Chrome

The `.zip` files are unpacked extensions, installed via **Load unpacked**:

1. Download and **unzip** the desired `.zip` into its own folder (e.g. `mygate-dump-tool/`).
2. Open Chrome → `chrome://extensions`.
3. Turn on **Developer mode** (top-right corner).
4. Click **Load unpacked** and select the unzipped folder.
5. Pin the extension to the toolbar and sign in to MyGate before using it.

Repeat for each tool you want installed.

## Run the Analysis Tool engine tests

The workbook engine ships with a self-contained Node regression suite (63 checks —
transformation rules plus a full mini end-to-end build):

```sh
node mygate-analysis-tool/tests/run_tests.mjs
```

You can also generate a workbook from any dump straight from the CLI:

```sh
node mygate-analysis-tool/tests/generate.mjs dump.xlsx out.xlsx --as-on 22.09.26
```
