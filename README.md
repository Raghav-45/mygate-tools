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
   using the standalone Summary tool's existing rules. Each generator uses its
   original delay: **Pending Report 1 second**, **Master Dump 2 seconds**, and
   **Summary Sheet 1.5 seconds**. Analysis runs locally. There is no shared speed
   slider, and previously saved shared-delay settings are ignored.
2. Open the previous report email in **Outlook → Sent Items**. From that Outlook
   tab, open the extension and click **Prepare Reply Draft**. The tool checks the
   open report's subject and uses Outlook's **Reply all** so the daily report stays
   in the same conversation. It keeps the native reply subject, existing recipients,
   signature and quoted history. To/Cc are inherited from Reply all and can be
   edited directly in Outlook; the extension never enters recipients again. It
   inserts the new Pending Tickets table when available, and attaches the newly
   generated Master Dump and Analysis. Review the draft and send it manually.
   If the open email is wrong, another draft is open, or Reply all is unavailable,
   step 2 stops. Retries reuse only the reply belonging to the current batch,
   without regenerating reports or duplicating the body and attachments. Each
   file is submitted once per reply. If an upload cannot be verified, retries
   check the existing attachments rather than uploading it again; attach any
   missing files manually from Downloads. Step 1 downloads remain available if
   Outlook fails.

Step 1 never opens Outlook. Report selections are locked until settings load and
saved immediately when toggled. A run uses its explicit current selection, so a
report-only run never falls back to saved/default tools. The temporary batch
for step 2 is replaced on each generation and cleared when Chrome starts again;
there is no saved-file picker, import, or reuse mode. Downloaded files remain on disk.
Email settings persist. Master Dump waits for a new cloud export entry and checks
ticket dates, statuses, and duplicate IDs. It also verifies the row count against a
separate read-only MyGate count for that range. It stops on an invalid export or timeout
without requesting another export or saving a partial dump. Pending Report and
Master Dump counts are independent; differences between them do not block the
handoff. See [dump timeout behavior](mygate-dump-tool/KNOWN_ISSUES.md).

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
handoff, independent report counts, stale or reused cloud files, duplicate tickets,
email retries, failed downloads, partial dumps and aborts.
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

Open **PowerShell or Command Prompt as administrator**, then use the block for
that terminal. Windows Terminal often opens PowerShell. Run the first line to
download the updater; if it finishes without an error, run the second line.

**PowerShell / Windows Terminal:**

```powershell
Invoke-WebRequest -UseBasicParsing -ErrorAction Stop -Uri 'https://raw.githubusercontent.com/Raghav-45/mygate-tools/main/scripts/Update-MyGateExtensions.ps1' -OutFile "$env:TEMP\Update-MyGateExtensions.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\Update-MyGateExtensions.ps1"
```

**Command Prompt (CMD):**

```cmd
powershell.exe -NoProfile -Command "Invoke-WebRequest -UseBasicParsing -ErrorAction Stop -Uri 'https://raw.githubusercontent.com/Raghav-45/mygate-tools/main/scripts/Update-MyGateExtensions.ps1' -OutFile (Join-Path $env:TEMP 'Update-MyGateExtensions.ps1')"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%TEMP%\Update-MyGateExtensions.ps1"
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
   - **Version / tag** — choose a new, unused version, such as `0.0.6` (prefixed with `v` automatically).
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
