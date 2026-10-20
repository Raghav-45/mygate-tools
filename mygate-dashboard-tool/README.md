# MyGate Dashboard Tool

A Python CLI that turns a **MyGate Master Helpdesk Dump** (produced by the
[MyGate Dump Tool](../mygate-dump-tool/)) into the polished
**"Mygate Complaints Dashboard as on DD.MM.YY"** workbook — a drop-in
replacement for the dashboards previously assembled by hand in Excel.

It is part of the same data pipeline:

```
mygate-dump-tool  ──produces──▶  MyGate_Master_Helpdesk_Dump_<from>_to_<to>.xlsx
                                        │
mygate-dashboard-tool  ──reads──▶       ▼
                 Mygate Complaints Dashboard as on DD.MM.YY.xlsx
```

Pair each dashboard with the dump that was taken on the same day: the dump is
the point-in-time snapshot (pending tickets only), and the dashboard is its
analysis.

## Usage

```sh
# as-on date is derived automatically from the dump file name (..._to_2026-10-01)
python3 mygate-dashboard-tool/make_dashboard.py path/to/MyGate_Master_Helpdesk_Dump_2024-01-01_to_2026-10-01.xlsx
# → Mygate Complaints Dashboard as on 01.10.26.xlsx

# or override it explicitly and pick an output path
python3 mygate-dashboard-tool/make_dashboard.py dump.xlsx \
    --as-on 22.09.26 --out "Mygate Complaints Dashboard as on 22.09.26.xlsx"
```

Requirements: Python 3.9+ and `openpyxl` (`pip install openpyxl`). Nothing else
— no Excel, no macros, no browser.

Both `DD.MM.YY` and `YYYY-MM-DD` are accepted with `--as-on`.

## What the generated workbook contains

Exactly the 5-sheet layout of the hand-made dashboard:

| Sheet | Role |
|-------|------|
| **Interactive** | Filterable dashboard — 4 dropdowns (Category / Sub Category / Phase / Status) drive a live `COUNTIFS` engine; 5 charts (monthly trend line, category bar, sub-category bar, status doughnut, phase bar) |
| **Overview (static)** | Snapshot of the whole period: totals, In Progress / Reopened / On Hold counts, static charts |
| **Data** | The parsed dump rows, sorted by created date then I.D., re-numbered `Sr No.` 1..N, with derived real-date (`Date`) and `Phase` columns; frozen header, filterable |
| **Analysis** | Count tables (category, sub-category, status, phase, monthly trend) for the static charts |
| **Calc** (hidden) | The live `COUNTIFS` engine: filter criteria picked up from the Interactive dropdowns plus per-chart count formulas |

Reopen the file in Excel and everything self-calculates
(`fullCalcOnLoad` is set). Add a new dropdown selection or edit a data row and
the Interactive charts/KPIs follow — the same behavior the hand-made file had.

All dropdown enumerations (categories, sub-categories, phases, statuses) are
derived from the data in the dump — no hardcoded lists.

## Verified equivalence (golden tests)

`mygate-dashboard-tool/tests/` runs 13 tests. The two **golden tests** load the
real hand-made reference files found in `/Users/aditya/Downloads/mygate-new-tool
data dumps` (override the folder with the `MYGATE_REF_DIR` env var):

- `test_golden_nikhil_data_roundtrip` — Nikhil's 852 Data rows from
  *"Mygate Complaints Dashboard as on 22.09.26.xlsm"* are converted back into a
  dump and fed through the tool: the result matches his workbook **1:1** (same
  Data sheet cell-for-cell, same KPI formulas, same charts/dropdowns/defined
  names, same Overview text, matching counts).
- `test_golden_real_dump_integration` — running on the actual paired dump
  (`MyGate_Master_Helpdesk_Dump_2024-01-01_to_2026-09-22.xlsx`) agrees with
  Nikhil's snapshot on every one of the 819 shared rows. The only differences
  are tickets whose status/category changed, or which were created/closed,
  between the morning snapshot Nikhil worked from and the later dump download
  (13 changed cells on 10 shared rows) — legitimate snapshot drift, not
  transformation errors, and the test reports them explicitly.

Run everything with:

```sh
python3 -m pytest mygate-dashboard-tool/tests/ -v
```

## Deviations from the hand-made file (all deliberate fixes)

1. **Monthly trend spans every month in the data, zero-filled.** The hand-made
   tables stopped at Jul-2026 and skipped gaps; the generated file continues
   through the latest data month (e.g. Aug/Sep-2026) so short months are not
   lost and no bars are missing.
2. **Real statuses.** The original dropdown/doughnut used `"HOLD"` and
   `"Repeated"`, which exist nowhere in the data (that slice counted 0 while 2
   tickets were `On Hold`). The tool enumerates the statuses actually present:
   `New / In Progress / Reopened / On Hold`.
3. **`On Hold` on the Overview sheet counts correctly** (no more `COUNTIF(…,"HOLD")`
   = 0).
4. **Output is `.xlsx`, not `.xlsm`.** No VBA is used: the dropdowns and live
   `COUNTIFS` engine are pure Excel features.
5. **Fixable quirk of the original kept for compatibility:** the Interactive
   engine reads data ranges through row `$10000` (the original stopped at
   `$9752`), so the file tolerates append-fixes in Excel; formulas differ
   accordingly (tests normalize this). Phase borders are applied properly in
   the Data sheet, which the hand-made file skipped.

## Why a Python CLI and not a Chrome extension

The other tools here are Chrome extensions, but the dashboard needs **real
charts, data validation dropdowns and a hidden calculation sheet** — ExcelJS
(used by the extensions) cannot create charts, while `openpyxl` writes exactly
what the original hand-made dashboards were themselves produced with (their
internal metadata says openpyxl). A native Python build therefore reproduces
the reference workbook most faithfully.

## Status / limitations

- Pending tickets only — the dump already makes that guarantee.
- `J/m` "DLP Status" on the Analysis sheet can't be derived from the dump (MyGate
  provides no closure/purchase dates) — same as the original, which left that
  section broken; here it shows a single "Not Specified" count.
