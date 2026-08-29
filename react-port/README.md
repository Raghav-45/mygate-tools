# MyGate Tools — React/TypeScript Port

Behavior-identical re-implementation of the three MyGate Chrome extensions in
React + TypeScript + Vite, with a shared package, a single test setup, and unit
tests for every worker, workbook and popup behavior.

The vanilla JavaScript originals live in the sibling `mygate-*-tool/` folders
(unchanged); this port reproduces their behavior 1:1, including bugs and quirks
(flagged in `NOTES.md` §5, not fixed).

| Tool | App | Tests | Status |
|------|-----|-------|--------|
| MyGate Dump Tool | `apps/dump-tool` | 25 | Ported (M3) |
| MyGate Report Tool | `apps/report-tool` | 32 | Ported (M4) |
| MyGate Summary Tool | `apps/summary-tool` | 32 | Ported (M5) |

## Layout

```
react-port/
  packages/shared/     shared TS source (auth discovery, GraphQL client, date helpers,
                       excel download/styles, chrome messaging/storage, UI components)
  apps/dump-tool/      0.0 roadmap: multi-year dump exporter
  apps/report-tool/    pending-tickets report from the GraphQL API
  apps/summary-tool/   complaint summary sheet (monthly)
  NOTES.md             inventory of the originals + port decisions + quirks log
```

## Prerequisites

- Node 20+ and npm (workspaces).
- The lockfile is checked in; `npm install` wires the monorepo.

## Build

```sh
npm install
npm run build          # builds all three extensions into <app>/dist
```

Each app uses `@crxjs/vite-plugin`, so `dist/` contains a loadable MV3
extension (`manifest.json` + a bundled service worker + popup).

## Install in Chrome (Load unpacked)

1. `npm run build --workspace @mygate/dump-tool` (or `report-tool` / `summary-tool`).
2. Chrome → `chrome://extensions` → enable **Developer mode**.
3. **Load unpacked** → select `apps/<tool>/dist`.
4. Pin the extension; sign in at `dashboard.mygate.com` before running a scan.

## Checks

```sh
npm run typecheck      # tsc across all workspaces
npm run lint           # eslint across the repo
npm run test           # vitest: 89 tests (25 dump + 32 report + 32 summary)
npm run build          # three vite builds
npx prettier --check . # formatting gate
```

The full gate must stay green before any milestone commit.

## Manual QA

See `MANUAL-TEST-CHECKLIST.md` — including the ported quirks you can reproduce
in Chrome (`NaN%` label, ms-vs-seconds slider, silent mid-scan aborts, etc.).

## Decisions & quirks

`NOTES.md` §1 inventories each original extension byte-for-byte; §3 records the
architecture decisions (dependency injection, shared `chrome` stub, explicit
vitest imports, UTC test timezone); §5 lists every bug ported verbatim.