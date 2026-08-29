# React Port — Notes & Decisions

Status note for the human reviewer. This document records (a) the inventory of the three
original extensions that were read before porting, and (b) every non-obvious decision made
during the port. Source of truth is always the original folders at the repo root; this file
flags anything that differs or that I was not able to verify without a live MyGate account.

---

## 1. Inventory — what the three originals do

All three are Chrome Manifest V3 extensions for the MyGate residential admin dashboard
(`dashboard.mygate.com`). They share a GraphQL API at `https://api.dashboard.mygate.com/graphql/`
with `content-type: application/json`, `origin: https://dashboard.mygate.com`,
`referer: https://dashboard.mygate.com/` request headers.

### Structural facts (verified with `diff`/`shasum`)

- `exceljs.min.js` is byte-identical in all three folders (same SHA-256). This is the ~930KB
  vendored copy that the port replaces with a single `exceljs` npm dependency.
- `popup.css` is byte-identical in all three folders (624 lines) — a shared "design system"
  that the port re-implements once as shared Tailwind-based components + a shared stylesheet.
- `popup.html` differs only in title/subtitle, the form controls (dates vs. month), the
  footer link (`https://dashboard.mygate.com/home/society/generatedReports` for dump-tool vs.
  `https://dashboard.mygate.com/` for the other two), and the about-modal is identical.
  The logo SVG and the author/about modal are identical across all three.
- Each tool has its **own** `SAMPLE_FALLBACK_TOKEN` (dump and report only; summary has none),
  its own token-discovery heuristics, its own storage keys, and its own message protocol.

### 1.1 `mygate-dump-tool` (v1.0.0)

**Popup** (440px wide, Inter font): date inputs `fromDate`/`toDate` (defaults `2024-01-01` →
today), a hidden "Polling Speed" drawer with a 1.0–5.0s step-0.5 slider (default 2.0), a
"Generate Master Dump" button, an abort button, progress bar, KPI cards (Yearly Chunks,
Total Merged Rows, Status), a chunk status table, a "downloaded automatically" banner, and the
shared about modal. Sends to background:
- `START_DUMP_EXPORT` `{ fromDate, toDate, requestDelayMs }`
- `ABORT_DUMP_EXPORT`

Listens for: `DUMP_PROGRESS_UPDATE`, `DUMP_FINISHED`, `DUMP_ABORTED`.

**Background** (`background.js`, 563 lines):
- `discoverActiveAuthToken()` — executes a script in open `*.mygate.com` tabs; scans
  `localStorage` for a key containing `"token"` whose value is `>20` chars and contains no `{`.
  Falls back to `SAMPLE_FALLBACK_TOKEN` if nothing is found (silently).
- `sliceIntoYears(from, to)` — slices a date range into ≤365-day chunks, oldest-first, then
  **reverses** so chunks are processed newest→oldest (the resulting file is ordered newest first).
- Per chunk (processed strictly sequentially):
  1. POST `getAdminSrList` with `isDownload: true`, `downloadFilters` carrying the *spaces*
     `'New ', 'Reopened ', 'In Progress ', 'Job Done ', 'On Hold '` statuses and `From`/`To`
     dates formatted `d:m:yyyy` (no leading zeros), plus epoch-condition filters
     (`from_date ≥ midnight-of-from`, `to_date ≤ midnight-of-to + 86399`). Wait for the request
     itself to succeed (auth error → status `Failed (Login Required)`).
  2. Poll `getDownloadReportList` up to **45** times, one sleep of `requestDelayMs` per cycle,
     looking for `report_name === 'Helpdesk Report'`, `status === 'Success'`, a `report_link`
     not already used, and `download_filters.From/Date From` + `To/Date To` matching the chunk.
     No match after 45 → `Timed out`.
  3. `fetch(report_link)` → `arrayBuffer` → `ExcelJS.xlsx.load` → worksheet[0]. Header row = 3,
     data from row 4. Column map is re-derived from the header text (case-insensitive match on
     e.g. `id`/`ticket id`/`i.d`, `created date`/`date`, `category`, `sub category`, `flat`/
     `house`, `subject`/`description`, `status`), with a hard-coded default map
     `{id:1, createdDate:2, category:4, subCategory:5, flat:7, subject:9, status:10}`.
     Cell values that are `Date` instances are normalized to UTC noon
     `new Date(Date.UTC(y, m, d, 12, 0, 0))` while merging. Rows are collected only when
     `rVals[colMap.id]` is defined and `!== ''`.
  4. Merge into a master in-memory list; the master xlsx is written at the end.
- Master workbook shape: row1 merged A1:H1 title **`DLF Independent Floors`** (blue fill
  `FF4D93D9`, Aptos 12, center), row2 merged A2:H2 subtitle
  `Help Desk Report: From <d-m-yyyy> To <d-m-yyyy>` (same fill/font), row3 headers
  `Sr No. / I.D / Created Date / Category / Sub Category / Flat / Subject / Status`. Every cell
  gets `thin` black borders. Data rows: Sr No 1-indexed; alignment centered except Category
  (col 4) and Subject (col 7) which are left-aligned; Subject wraps. Column widths 10/12/22/28/
  20/16/45/15. Row heights 24/20/24.
- Filename: `MyGate_Master_Helpdesk_Dump_<fromDate>_to_<toDate>.xlsx` (raw YYYY-MM-DD).
- Storage key `dumpScanState` (`isScanning`, `pct`, `statusText`, `chunks`, `totalRows`,
  `isDone`; `isAborted` on abort). Progress percentages: chunk base up to 85 (i/n), +5 after
  request sent, +15 while downloading, +25 once merged, 96 while compiling the workbook, 100
  done.
- `KNOWN_ISSUES.md` (not ported, but read) documents a CloudFront timeout workaround.

### 1.2 `mygate-report-tool` (v2.2.0)

**Purpose:** pending-tickets counts per category from the GraphQL API (no export files come
from MyGate — this tool builds its own xlsx from counts).

**Popup:** date inputs (same defaults), a **category pill selector** (fetched from background
via `GET_CATEGORIES_LIST`), "Toggle All", a Scan Speed drawer (0.2–3.0s, default 1.0), KPIs
(Total / Resolved / Open), and a results table with category rows + a highlighted yellow
(`#FEF08A`) totals row. Storage keys: `requestDelay` (**seconds**, stored bare),
`selectedCatIds` (array of category ids, persisted on every pill toggle), `ticketsScanState`.

**Background** (`background.js`, 400 lines):
- Fixed `CATEGORIES` list: 7 `[name, id]` pairs (Accounts Billing, Construction Or Project
  Related, Design Related Issue, Estate Infra Outer Area from the plot, FM Common Area Related
  Issue, IT WIFI Network, Products Appliances) — page-scoped ids, treated as opaque.
- `discoverActiveAuthToken()` — richer than dump-tool: scans `localStorage` keys matching
  `token / auth_token / authorization / access_token / jwt / mygate_token / user`, accepts a
  value `>20` chars not starting with `{`, and also unwraps JSON objects looking for
  `token/accessToken/jwt/authorization`. Then falls back to browser cookies via
  `chrome.cookies.getAll({ domain: 'mygate.com' })` matching `token / auth_token /
  authorization / access_token / jwt / session_token` with value `>20`. Logs to console on
  discovery. Fallback `SAMPLE_FALLBACK_TOKEN` if nothing found.
- For each selected category, three count queries (`getAdminSrList` with `requiredFields:[id]`,
  `pagination:{count:1,page:1}`, `date_filter equal created_date`, `category equal <id>`,
  `from_date gte`, `to_date lte (+86399)`, `mygate_status`):
  - TOTAL: `["open","hold","re_opened","job_done","in_progress","closed"]` → `operation: "in"`
  - RESOLVED: `["closed"]` → `operation: "equal"`
  - OPEN: `["open","hold","re_opened","job_done","in_progress"]` → `operation: "in"`
  Note the `operation` is chosen by `statuses.length === 1 ? 'equal' : 'in'`.
- Request uses `credentials: 'include'` and attaches `authorization` only when the discovered/
  fallback token is non-empty. Errors: `HTTP 401/403` → specific message; `data.errors` →
  first error message; missing numeric `totalCount` → auth-verification message.
- Sleeps `requestDelayMs` after **each** count query (default 1000ms).
- Storage `ticketsScanState`: `{ isScanning, pct, statusText, results, summary, isDone }`.
  Per-category pct = `round((i+1)/n*100)`.
- Messages: `SCAN_PROGRESS_UPDATE` (with `stepIndex`, `totalSteps`, `statusText`),
  `CATEGORY_COMPLETED`, `SCAN_FINISHED` (`results`, `summary`, `reportMeta:{fromDate,toDate}`),
  `SCAN_ERROR`, `SCAN_ABORTED`.
- Excel: sheet `Pending Tickets`, row1 merged A1:D1 title
  `Pending Mygate Tickets - From <d-m-yyyy> To <d-m-yyyy>` (blue fill, Aptos 12, center),
  row2 headers `Category / Total / Resolved / Open` (blue fill, left align), data rows +
  `Total` row (Aptos 12, col 1 left, others right). Thin black borders are applied to **every**
  cell of the used range `A1:D<last>` in a final loop (so title + header rows get borders
  too, even though they weren't explicitly bordered when created). Column widths 55/9/12/8.
- Filename: `Pending_Mygate_Tickets_Report_<d-m-yyyy of toDate>.xlsx` (un-padded `d` and `m`).

### 1.3 `mygate-summary-tool` (v1.0.0)

**Purpose:** per-day complaint summary for a selected month (prev-open / received / closed /
pending) by calling the GraphQL API once per metric per day.

**Popup:** a `<input type="month">` picker (defaults to current month), Scan Speed drawer
(0.5–3.0s step 0.5, default 1.5), KPIs (Total Received, Total Closed, Final Pending), and a
table of the current day's rows (Date / Prev Open / Received / Closed — the Pending column is
only in the xlsx, **not** the table). Storage keys: `requestDelay` (seconds), `summaryScanState`.

**Background** (`background.js`, 329 lines):
- `getAuthToken()` — tab localStorage scan (key = any, value `>40` chars, no `{`, no space),
  then cookie scan via `chrome.cookies.getAll({domain:'mygate.com'})` where the cookie name
  contains `token` or `auth` and value `>30`. **No fallback token**: null → throws
  `No MyGate token found! Please open dashboard.mygate.com and log in.` (Note: because the
  cookie check uses name *contains* `"auth"`, cookies like `_gid`/`_ga` *do not* match `auth`,
  but a cookie such as `gatewayauth` would.)
- Dates are `DD-MM-YYYY` strings. `epoch(dStr)` parses `dd-mm-yyyy`. For each day:
  - `prevOpen` = count(`01-01-2024` → that day, statuses `open/hold/re_opened/in_progress/job_done`)
  - `received` = count(day → day, all six statuses incl. `closed`)
  - `closed` = count(day → day, `["closed"]`)
  - `pending = prevOpen + received - closed` → row `{ date: DD-MMM-YYYY, prevOpen, received,
    closed, pending }`, pushed then the list re-sorted chronologically.
  - Rows are iterated **newest-first internally?** No — days iterate forward, but each row is
    inserted and the array re-sorted by date each iteration. There is a per-day abort check.
  - The `to_date` filter is `day + 86399` (end of day). The **hard-coded** `01-01-2024` base
    date for `prevOpen` is ported **as-is** (see §5 flag).
- Storage `summaryScanState`: `{ isScanning, year, month, pct, stepText, rows, isDone }`.
  Messages: `SUMMARY_PROGRESS` (`stepText`, `pct`, optional `row`), `SUMMARY_DONE`,
  `SUMMARY_ABORTED`, `SUMMARY_ERROR`.
- Excel: sheet `Sheet1`, `ws.views = [{showGridLines:true}]`, col widths 27.14/14.14/13.86/
  11.57/12.00. Row1 merged A1:E1 **`Complaint Summary Sheet`** (Calibri 14 bold, center, white
  fill, thin black borders, row height 18). Row2 headers
  `Date / Previous day Open Complaints / Today Received Complaints / Today Closed Complaints /
  Pending` (set via `row.values`), height 45, wrap text, white fill, borders. Data rows sorted
  chronologically: date cell = UTC-noon `Date` with `numFmt 'dd-mm-yyyy'`; Pending cell =
  formula `(B{r}+C{r})-D{r}` with `result: r.pending`. All cells Calibri 11, white fill,
  center aligned, thin black borders.
- Filename: `Complaint Summary Sheet - <MMM>-<yyyy>.xlsx`.
- `CHECK_AUTH_STATUS` handler exists in background but **no popup code sends it** → dead code
  in the original; ported for parity in the message handler switch but unused (flagged in §5).

---

## 2. Genuinely shared vs. tool-specific (the real diffs)

**Genuinely shared (goes into `packages/shared`):**
- GraphQL endpoint constant + the fetch-boilerplate client (`POST`, JSON, origin/referer
  headers, credentials, error normalization).
- The three count-query builders are *nearly* identical between report-tool and summary-tool
  (`date_filter`, `from/to` epoch filters, `mygate_status in|equal`) — shared as one builder
  taking statuses + a filter date pair. (dump-tool's `getAdminSrList` request is deliberately
  different: `isDownload`, full required fields, `downloadFilters`.)
- The "write workbook → base64 data URL → `chrome.downloads.download`" helper (identical
  pattern in all three, incl. `saveAs:false` and lastError handling).
- The shared xlsx styling constants (`FF4D93D9` fill, Aptos fonts, thin borders) for the dump
  & report workbooks (summary uses its own Calibri/white-white style — kept local to summary).
- Date helpers common to ≥2 tools: `getMidnightEpoch`, `formatHeaderDate` (dump) ≈
  `formatTitleDate` (report) — same `d-m-yyyy` un-padded output; `formatFilenameDate` (report)
  ≈ report's own filename formatting; `arrayBufferToBase64`; `sleep`.
- The shared UI "design system" (all identical `popup.css`): brand header + logo, about modal,
  settings drawer + speed slider, progress bar, KPI cards, alerts, primary/abort buttons,
  auto-download banner, table styling, body sizing 440×560–640.
- The shared popup.html boilerplate (fonts, meta, body wrap, top accent bar).

**Tool-specific (kept out of shared):**
- dump-tool: `sliceIntoYears`, the `isDownload` request payload, the `getDownloadReportList`
  polling loop + report matcher, the chunk-parse / column-map logic, the master-workbook
  layout, its `dumpScanState` protocol, its password-less `localStorage` token hunt.
- report-tool: the `CATEGORIES` table, the pill selection UX, its dual tab+cookie discovery
  with JSON-unwrap, its `ticketsScanState` protocol and per-category scan loop.
- summary-tool: the day-by-day month loop, the `01-01-2024` base-date logic, the
  formula-based workbook, its `summaryScanState` protocol, the `CHECK_AUTH_STATUS` handler.
- Tokens: each tool keeps its own fallback token constant (report/summary). No token is shared.

Token discovery is *similar but not identical* in all three (key lists, length thresholds, tab
URL patterns, cookie vs localStorage precedence). I factored a small, parametrizable
`discoverAuthToken` into shared while preserving each tool's exact thresholds and order — the
per-tool call sites pass their tool-specific options verbatim from the originals.

---

## 3. Port architecture decisions

- **Monorepo:** npm workspaces under `react-port/` with `packages/shared` + `apps/dump-tool`,
  `apps/report-tool`, `apps/summary-tool`. Shared source is consumed directly as TS (no build
  step) via Vite; `tsc` checks the whole program via per-app tsconfig `paths`.
- **Stack:** React 18.3 + TypeScript 5 + Vite 6 + `@crxjs/vite-plugin` (MV3 build, bundles the
  service worker as a module worker). Tailwind CSS 3.4 replaces the hand-written popup.css —
  the shared design-system CSS is re-created as Tailwind utilities in shared components plus a
  small base layer (`body` sizing, fonts, scrollbar). ESLint 9 flat config + Prettier.
  Vitest 3 + jsdom + Testing Library.
- **exceljs:** single `exceljs@4.4.0` dependency at the shared package; the three ~930KB vendored
  copies are not used at all. The service worker imports it through the shared `excel` helpers.
- **Chrome API surface:** thin shared helpers (`storage`, `downloads`, messaging) so the browser
  globals are isolated; app code never touches `chrome` directly except through these. This is
  what makes Vitest mocking tractable. A tiny `chrome` stub is provided in the test setup.
- **Worker structure:** each app splits `background/index.ts` (thin chrome wiring) from pure
  business modules (`slice.ts`, `payloads.ts`, `parseChunk.ts`, `workbook/*.ts`) that are unit
  tested directly. Popup logic is split into small components + custom hooks.
- **Behavior preservation:** all status strings, messages, pct math, storage shapes, filenames,
  excel layouts, sort orders, and mygate-specific quirks are ported verbatim; see §5.

---

## 4. Fixtures for self-testing (no live MyGate)

Because no live calls may be made, all foreground dependencies are mocked:
- `getDownloadReportList` responses are fabricated from the field names read in
  `background.js` (`report_name`, `status`, `report_link`, `download_filters`).
- Chunk xlsx files are *produced by exceljs itself* in tests (header row 3 + data rows) so the
  dump-tool parser is exercised on a real, valid XLSX buffer.
- KPI / summary / table updates are verified by rendering the React popups under jsdom with a
  mocked `chrome.runtime` bus that emits the same messages the worker would emit.

Note: `requestDelay` is stored inconsistently across the originals (dump-tool stores
milliseconds — a real bug, see §5; report- and summary-tool store seconds). The forms are
ported faithfully, bug included.

---

## 5. Bugs & quirks ported 1:1 (flagged, not fixed)

1. **dump-tool `requestDelay` units mismatch (bug).** Popup stores `requestDelay: val*1000`
   (ms) on slider input, but on restore assigns that ms value back to the seconds-based range
   input, and prints `${stored.requestDelay.toFixed(1)}s`. So after a reload the slider will
   show the ms value clamped to max 5.0 and the label shows e.g. `2000.0s`. The actual poll
   delay used on Generate is always correct (`sliderSeconds * 1000`). **Ported exactly.**
2. **Report-tool `SCAN_PROGRESS_UPDATE` "Downloading Excel..." has no `stepIndex`, so the
   popup computes `NaN%`** for that one transient message. Ported exactly.
3. **dump-tool poll cap is 45**, while `KNOWN_ISSUES.md` documents ~60. Ported the code (45);
   the doc was not ported.
4. **summary-tool hard-codes `01-01-2024` as the "previous day open" base date.** Any data
   before Jan 2024 is invisible to the prev-open count. Ported as-is.
5. **summary-tool `CHECK_AUTH_STATUS` handler is dead code** in the original popup; the
   handler is ported for behavioral parity but never invoked.
6. **dump-tool auth failure path** sets `chunks` rows to `Auth Error` / `Failed (Login Required)`
   and returns without setting `isDone`; on popup reopen this shows as interrupted state.
   Ported as-is.
7. **report-tool** only attaches `authorization` when a token string is non-empty — but because
   the SAMPLE fallback is anonymous, the header is effectively always attached. Ported as-is.
8. dump-tool tab-scan accepts the **first** `localStorage` key containing `token` with no
   JSON-parse; report-tool's is stricter. Kept per-tool (not unified).
9. **report-tool mid-category abort is silent.** If an abort arrives between the sleep and the
   count requests (the `break`s after each sleep), the loop exits and `if (!activeScanAbort)`
   means the finished block is skipped with no state flush and no `SCAN_ABORTED` — the popup
   keeps showing "Stopping..." until it is closed. Ported verbatim (tested).

## 6. Things I deliberately did NOT do

- No `git push`, no force-push/rebase/history rewrite; everything is on the local
  `react-port` branch for human review.
- No live MyGate traffic of any kind (rule §3.3). All shapes inferred from source.
- Did not "improve" category lists, statuses, Excel layouts, or messages; did not re-order
  dumped chunks; did not fix any of the §5 quirks.
- Did not keep any vendored `exceljs.min.js` under `react-port/`.

## 7. How decisions were recorded as I went

- Decisions log appended to this file (or placed in README/MANUAL-TEST-CHECKLIST where they
  belong) as commits were made, so `git log` on `react-port` tells the full story.

## 8. Porting log

- **Milestone 3 — dump-tool ported + tested.** Everything in §1.1 is now implemented in
  `apps/dump-tool` (`src/dump/{dumpState,slice,payloads,parseChunkXlsx,masterWorkbook,
  runDumpExport}.ts`, `src/background/index.ts`, `src/popup/App.tsx`) with 25 Vitest tests
  covering slicing, the chunk parser, the master workbook, the full request→poll→merge→download
  run, and the popup (message broadcast, ms-delay persistence, storage replay, quirk labels).
  Ported-decision notes:
  - `discoverAuthToken` is reused from shared with dump's exact localStorage options.
  - `postGraphQL` defaults to `credentials: 'include'`, but the original dump request sends no
    cookies, so dump's deps pass `credentials: 'same-origin'`.
  - `runDumpExport` takes an injectable `DumpDeps` (token, post, fetch, workbook load/download,
    storage, messaging, sleep, log, abortState, poll cap) so the whole run is unit-tested with
    a fake GraphQL server that fabricates `getDownloadReportList` responses and real XLSX
    buffers; both the auth-failure path (`isScanning` stays true) and the abort paths are
    covered. `abortState` is a shared mutable object between popup message wiring and the run.
  - Popup quirks ported 1:1: `requestDelay` stored in ms and fed back into the seconds slider
    on reload (`2000.0s` label, clamped slider) — §5.1; from>to validation message preserved.
  - Shared fic: `packages/shared/src/ui/index.tsx` imports `../styles/base.css` (not
    `./styles/base.css`). Icons (GearIcon, StopIcon, ZapIcon, …) exported from the shared
    barrel for the settings/abort/about affordances.
- **Build detail:** Tailwind content globs must be absolute (`path.join(__dirname, …)` in
  `tailwind.config.js`). Each `vite build` runs with CWD inside an app workspace; root-relative
  globs silently matched nothing, so no utilities were generated. Config is ESM (loaded fine via
  PostCSS/Tailwind) and prettier-formatted.
- **Milestone 4 — report-tool ported + tested.** Everything in §1.2 is now implemented in
  `apps/report-tool` (`src/report/{reportState,categories,countQueries,runReportScan,
  reportWorkbook}.ts`, `src/background/index.ts`, `src/popup/App.tsx`) with 32 Vitest tests
  (count-query payload+error mapping, workbook layout, the full per-category
  total→resolved→open scan incl. all abort/error paths, and the popup). Ported-decision notes:
  - `countQueries.ts` reproduces the exact conditions list (`date_filter` equal `created_date`,
    `category` in `[<id>]`, `from_date`/`to_date` equal epoch strings, `mygate_status`
    in/equal) and error strings from §1.2, layered over shared `postGraphQL`
    (`credentials: 'include'`; `authorization` attached only when the token is non-empty).
    `fetchCategoryCount` is `fetchImpl`-injectable for tests.
  - `runReportScan.ts` takes an injectable `ReportScanDeps` (discoverToken + report's exact
    localStorage/cookie options + console-log on discovery, fallback token, countRequest,
    sleep, storage, messaging, downloadWorkbook, log, abortState). One count query per status,
    a `sleep(requestDelayMs)` — default 1000 — after **each** of the 3 steps per category;
    per-category pct `round((i+1)/n*100)`; epochs `getMidnightEpoch(from)` and
    `getMidnightEpoch(to)+86399`. Storage key `ticketsScanState`.
  - Quirks ported 1:1 and locked by tests: the "Downloading Excel..." `SCAN_PROGRESS_UPDATE`
    carries no stepIndex/totalSteps → popup renders `NaN%` (§5.2); a mid-category abort does
    `break`, skips the finished block and returns without flushing state or sending
    `SCAN_ABORTED` — the popup stays on "Stopping..." (see §5.9); empty category selection
    (only reachable via stale storage) falls through to an empty workbook download.
  - Popup behavior: category pills fetched via `GET_CATEGORIES_LIST`, default all-selected,
    selection persisted under `selectedCatIds`; `requestDelay` stored as **seconds**
    (`Number(delaySeconds)` after slider input, matching §1.2); scan-state replay on reopen
    (in-flight → progress; done → banner); BOTH original validation messages preserved.
  - Workbook: headers are written explicitly into row 2 (trailing `sheet.columns`-style headers
    aren't used — exceljs auto-writes header text into row 1, which would clash with the merged
    title). Column widths set via `getColumn(i).width`.
  - Shared fic: `AutoDownloadBanner` gained an optional `message` prop (default = dump's
    existing text) so report can pass its own "Excel File Downloaded Automatically!". The test
    chrome stub also grew a `setResponse(type, value)` hook (used to seed
    `GET_CATEGORIES_LIST`).