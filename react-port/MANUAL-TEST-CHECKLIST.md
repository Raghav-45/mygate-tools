# Manual Test Checklist

Each tool is an MV3 extension. Build with `npm run build --workspace @mygate/<tool>`,
load `apps/<tool>/dist` via **Load unpacked**, pin it, and sign in at
`dashboard.mygate.com`.

Some behaviors below are **ported quirks** (§5 in `NOTES.md`) — if you see them,
that is the expected behavior, not a bug in the port.

## 1. MyGate Dump Tool (v1.0.0)

Sign in to `dashboard.mygate.com` and open its Reports page first (the
background injects into that tab to read the auth token).

- [ ] Slider behaves like the original: delay stored as **ms**; after reload the
      label can show e.g. `2000.0s` and the slider clamps to 5.0 (quirk §5.1),
      but the actual poll interval on Generate is always seconds-derived.
- [ ] Generate runs the multi-year scan: status advances per poll, scanned
      chunk counts update, and all report links are fetched and merged.
- [ ] Auto-scroll table fills with month rows; KPI counts update live;
      "Polling Limit Reached (45)" appears after ~45 polls and the scan stops
      (quirk §5.3).
- [ ] A completed run auto-downloads
      `MyGate_Master_Helpdesk_Dump_<from>_to_<to>.xlsx` with the blue/Aptos
      design, merged title, data rows and the green total row; banner + success
      alert show.
- [ ] Abort button stops the scan and shows the abort state.
- [ ] Close the popup mid-scan, reopen: progress and rows are replayed from
      `chunksScanState` storage.
- [ ] Auth-failure scenario (no token): rows show `Auth Error` / `Failed (Login
      Required)` and reopening shows the interrupted state (quirk §5.6).

## 2. MyGate Report Tool — Pending Tickets Report

- [ ] Both date fields prefill (From `2024-01-01`, Till today); category pills
      load via `GET_CATEGORIES_LIST` and default to all selected; `Toggle All`
      works; selection persists in `selectedCatIds`.
- [ ] `Report From date cannot be after Till date.` shows for an inverted range;
      `Please select at least one Ticket Category.` shows when none are ticked.
- [ ] Scan per category: Total → Resolved → Open count queries with the chosen
      delay (drawer 0.2–3.0 s, step 0.2, stored in seconds).
- [ ] Rows stream into the table as each category completes; Total/Resolved/
      Open KPIs update; the yellow total row appears.
- [ ] During the transient `Downloading Excel...` update the popup briefly
      shows **NaN%** (quirk §5.2 — correct).
- [ ] Completed run downloads `Pending_Mygate_Tickets_Report_<d-m-yyyy>.xlsx`
      (blue title/headers, category rows, yellow totals) and shows the banner.
- [ ] Abort mid-category: button goes to "Stopping..." and may stay there until
      the popup is closed (quirk §5.9 — correct).
- [ ] Close/reopen mid-scan: progress + results replay from `ticketsScanState`.

## 3. MyGate Summary Tool — Complaint Summary Sheet

- [ ] `Report Month` defaults to the current month (`YYYY-MM`); clearing it and
      pressing Generate shows `Please select a report month.`.
- [ ] Drawer slider is **0.5–3.0 s (step 0.5)**, stored as seconds; label shows
      `1.5s` by default.
- [ ] Generate shows the results section immediately with zeroed KPIs, then
      "Connecting MyGate API..." → per-day `Processing 01-01-2024...` →
      `Completed ...` rows streaming into the Date/Prev Open/Received/Closed
      table.
- [ ] KPIs: **Total Received**, **Total Closed**, **Final Pending** (last row's
      pending after the chronological sort).
- [ ] Completed run downloads `Complaint Summary Sheet - <MONTH>-<year>.xlsx`
      (Calibri/white style, merged title, wrapped headers, Pending is a
      formula `(B+C)-D`) and shows the banner.
- [ ] `Stop Live Scan` → `Live scan stopped by user.` alert when the abort lands
      on a day boundary; aborting between the daily count queries can leave the
      popup on "Stopping..." until reopened (quirk §5.10 — correct).
- [ ] `prevOpen` always counts from `01-01-2024` regardless of the selected
      month (quirk §5.4 — correct).
- [ ] Close/reopen mid-scan or after a finish: rows + KPIs + banner replay from
      `summaryScanState`.

## Cross-tool

- [ ] All three extensions can be installed side by side (distinct names/ids);
      the shared package is bundled into each, so no external runtime deps.
- [ ] No network requests go to anything except `https://api.mygate.com/`
      (GraphQL) while scanning; the dashboard tab is only probed for tokens.