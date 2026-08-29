import type { Workbook } from 'exceljs'
import { discoverAuthToken, downloadWorkbook, getMidnightEpoch, sleep } from '@mygate/shared'
import { CATEGORY_NAME } from './categories'
import { fetchCategoryCount, STATUS_OPEN, STATUS_RESOLVED, STATUS_TOTAL } from './countQueries'
import { buildPendingTicketsWorkbook, reportFilename } from './reportWorkbook'
import type { CategoryData, ReportScanParams, ReportSummary } from './reportState'

export const SAMPLE_FALLBACK_TOKEN =
  'zr2Er9wrdfhTyiY01Lnvr03cje9oeeH7wR6XvFmeSR87okw1qW4QyAuRkoSaOkff'

export interface ReportScanDeps {
  /** Returns a discovered token or `null`; the fallback token is used when null. */
  discoverToken: () => Promise<string | null>
  fallbackToken: string
  /** Run one `getAdminSrList` count query and return totalCount. */
  countRequest: (
    categoryId: number,
    fromEpoch: number,
    toEpoch: number,
    statuses: string[],
    token: string,
  ) => Promise<number>
  sleep: (ms: number) => Promise<void>
  storageSet: (items: Record<string, unknown>) => Promise<void>
  sendMessage: (message: unknown) => void
  downloadWorkbook: (workbook: Workbook, filename: string) => Promise<number>
  logError: (message: string, ...args: unknown[]) => void
  abortState: { aborted: boolean }
}

export function createReportDeps(overrides: Partial<ReportScanDeps> = {}): ReportScanDeps {
  return {
    discoverToken: () =>
      discoverAuthToken({
        tabQueryUrl: '*://*.mygate.com/*',
        localStorage: {
          keyContains: [
            'token',
            'auth_token',
            'authorization',
            'access_token',
            'jwt',
            'mygate_token',
            'user',
          ],
          minLength: 20,
          reject: [],
          rejectPrefix: ['{'],
          unwrap: true,
          unwrapFields: ['token', 'accessToken', 'jwt', 'authorization'],
        },
        cookie: {
          names: ['token', 'auth_token', 'authorization', 'access_token', 'jwt', 'session_token'],
          minLength: 20,
        },
        log: (msg) => console.log(msg),
      }),
    fallbackToken: SAMPLE_FALLBACK_TOKEN,
    countRequest: (categoryId, fromEpoch, toEpoch, statuses, token) =>
      fetchCategoryCount(categoryId, fromEpoch, toEpoch, statuses, token),
    sleep,
    storageSet: (items) => chrome.storage.local.set(items),
    sendMessage: (message) => {
      void chrome.runtime.sendMessage(message).catch(() => {})
    },
    downloadWorkbook,
    logError: (message, ...args) => console.error(message, ...args),
    abortState: { aborted: false },
    ...overrides,
  }
}

const INITIAL_SUMMARY: ReportSummary = { total: 0, resolved: 0, open: 0 }

function selectedCategories(ids: number[]): Array<{ name: string; id: number }> {
  return ids
    .map((id) => {
      const name = CATEGORY_NAME.get(id)
      return name !== undefined ? { name, id } : null
    })
    .filter((c): c is { name: string; id: number } => c !== null && c !== undefined)
}

async function handleAbort(deps: ReportScanDeps): Promise<void> {
  await deps.storageSet({ ticketsScanState: { isScanning: false, isAborted: true } })
  deps.sendMessage({ type: 'SCAN_ABORTED', results: [] })
}

/**
 * Full report scan, step-for-step matching the original `runFullReportScan`:
 * per selected category it fetches total -> resolved -> open counts (one
 * `sleep(requestDelayMs)` between steps), then compiles the Pending Tickets
 * workbook and auto-downloads it.
 *
 * Ported quirks (kept verbatim, see NOTES.md §8):
 *  - the "Downloading Excel..." progress message carries no stepIndex/totalSteps,
 *    so the popup renders `NaN%`;
 *  - when an abort lands mid-category the `break` skips the finished block but the
 *    scan state is not flushed and no SCAN_ABORTED is sent -> the popup stays on
 *    "Stopping..." until it is closed.
 */
export async function runReportScan(params: ReportScanParams, deps: ReportScanDeps): Promise<void> {
  const { fromDate, toDate, requestDelayMs = 1000, selectedCategoryIds = [] } = params

  const tokenToUse = (await deps.discoverToken()) || deps.fallbackToken

  const cats = selectedCategories(selectedCategoryIds)
  if (cats.length === 0) return

  const fromEpoch = getMidnightEpoch(fromDate)
  const toEpoch = getMidnightEpoch(toDate) + 86399

  const results: CategoryData[] = []
  const summary: ReportSummary = { ...INITIAL_SUMMARY }
  const totalSteps = cats.length

  await deps.storageSet({
    ticketsScanState: {
      isScanning: true,
      pct: 0,
      statusText: 'Connecting MyGate API...',
      results: [],
      summary: INITIAL_SUMMARY,
      isDone: false,
    },
  })

  for (let i = 0; i < cats.length; i++) {
    if (deps.abortState.aborted) {
      await handleAbort(deps)
      return
    }

    const cat = cats[i]
    const step = i + 1

    // --- total ---------------------------------------------------------------
    deps.sendMessage({
      type: 'SCAN_PROGRESS_UPDATE',
      stepIndex: step,
      totalSteps,
      currentCategory: cat.name,
      statusText: `Fetching total tickets for ${cat.name}...`,
    })
    await deps.sleep(requestDelayMs)
    if (deps.abortState.aborted) break

    let total: number
    try {
      total = await deps.countRequest(cat.id, fromEpoch, toEpoch, STATUS_TOTAL, tokenToUse)
    } catch (e) {
      deps.logError('Error fetching category total:', e)
      const errorMessage = e instanceof Error ? e.message : String(e)
      await deps.storageSet({
        ticketsScanState: { isScanning: false, error: errorMessage },
      })
      deps.sendMessage({ type: 'SCAN_ERROR', category: cat.name, errorMessage })
      return
    }

    // --- resolved -------------------------------------------------------------
    deps.sendMessage({
      type: 'SCAN_PROGRESS_UPDATE',
      stepIndex: step,
      totalSteps,
      currentCategory: cat.name,
      statusText: `Fetching resolved tickets for ${cat.name}...`,
    })
    await deps.sleep(requestDelayMs)
    if (deps.abortState.aborted) break

    let resolved: number
    try {
      resolved = await deps.countRequest(cat.id, fromEpoch, toEpoch, STATUS_RESOLVED, tokenToUse)
    } catch (e) {
      deps.logError('Error fetching category resolved:', e)
      const errorMessage = e instanceof Error ? e.message : String(e)
      await deps.storageSet({
        ticketsScanState: { isScanning: false, error: errorMessage },
      })
      deps.sendMessage({ type: 'SCAN_ERROR', category: cat.name, errorMessage })
      return
    }

    // --- open ------------------------------------------------------------------
    deps.sendMessage({
      type: 'SCAN_PROGRESS_UPDATE',
      stepIndex: step,
      totalSteps,
      currentCategory: cat.name,
      statusText: `Fetching open tickets for ${cat.name}...`,
    })
    await deps.sleep(requestDelayMs)
    if (deps.abortState.aborted) break

    let open: number
    try {
      open = await deps.countRequest(cat.id, fromEpoch, toEpoch, STATUS_OPEN, tokenToUse)
    } catch (e) {
      deps.logError('Error fetching category open:', e)
      const errorMessage = e instanceof Error ? e.message : String(e)
      await deps.storageSet({
        ticketsScanState: { isScanning: false, error: errorMessage },
      })
      deps.sendMessage({ type: 'SCAN_ERROR', category: cat.name, errorMessage })
      return
    }

    summary.total += total
    summary.resolved += resolved
    summary.open += open
    results.push({ name: cat.name, id: cat.id, total, resolved, open })

    const pct = Math.round((step / totalSteps) * 100)
    await deps.storageSet({
      ticketsScanState: {
        isScanning: true,
        pct,
        statusText: `Completed ${cat.name}`,
        results,
        summary,
        isDone: false,
      },
    })
    deps.sendMessage({ type: 'CATEGORY_COMPLETED', data: results[results.length - 1], summary })
  }

  if (!deps.abortState.aborted) {
    await deps.storageSet({
      ticketsScanState: {
        isScanning: true,
        pct: 100,
        statusText: 'Downloading Excel...',
        results,
        summary,
        isDone: false,
      },
    })
    deps.sendMessage({ type: 'SCAN_PROGRESS_UPDATE', statusText: 'Downloading Excel...' })

    const workbook = buildPendingTicketsWorkbook(results, summary, fromDate, toDate)
    await deps.downloadWorkbook(workbook, reportFilename(toDate))

    await deps.storageSet({
      ticketsScanState: {
        isScanning: false,
        pct: 100,
        statusText: 'Report Generated!',
        results,
        summary,
        isDone: true,
      },
    })
    deps.sendMessage({
      type: 'SCAN_FINISHED',
      results,
      summary,
      reportMeta: { fromDate, toDate },
    })
  }
}
