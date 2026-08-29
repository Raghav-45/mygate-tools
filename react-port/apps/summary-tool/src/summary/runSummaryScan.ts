import { Workbook } from 'exceljs'
import {
  discoverAuthToken,
  downloadWorkbook,
  formatDDMMYYYY,
  formatDDMMMYYYY,
  parseDateToTime,
  sleep,
} from '@mygate/shared'
import { fetchDayCount, STATUS_ALL, STATUS_CLOSED, STATUS_PREV_OPEN } from './countQueries'
import type { SummaryRow } from './summaryState'
import { buildSummaryWorkbook, summaryFilename } from './summaryWorkbook'

const DAY_MS = 1000 * 60 * 60 * 24

/** Hard-coded base date used for the daily `prevOpen` roll-forward count (quirk, see NOTES §5). */
export const BASE_PREV_DATE = '01-01-2024'

export interface SummaryScanDeps {
  /** Raw token discovery, resolves `null` when no token exists (never throws). */
  discoverToken: () => Promise<string | null>
  /** Token for the scan; throws the original "No MyGate token found!" error when missing. */
  getToken: () => Promise<string>
  count: (fromDate: string, toDate: string, statuses: string[], token: string) => Promise<number>
  sleep: (ms: number) => Promise<void>
  storageSet: (items: Record<string, unknown>) => Promise<void>
  sendMessage: (message: unknown) => void
  downloadWorkbook: (workbook: Workbook, filename: string) => Promise<number>
  abortState: { aborted: boolean }
}

export function createSummaryDeps(overrides: Partial<SummaryScanDeps> = {}): SummaryScanDeps {
  const discoverToken = (): Promise<string | null> =>
    discoverAuthToken({
      tabQueryUrl: '*://dashboard.mygate.com/*',
      localStorage: {
        minLength: 40,
        reject: ['{', ' '],
        rejectPrefix: [],
        unwrap: false,
        unwrapFields: [],
      },
      cookie: { contains: ['token', 'auth'], minLength: 30 },
    })

  return {
    discoverToken,
    getToken: async () => {
      const token = await discoverToken()
      if (!token) {
        throw new Error('No MyGate token found! Please open dashboard.mygate.com and log in.')
      }
      return token
    },
    count: fetchDayCount,
    sleep,
    storageSet: (items) => chrome.storage.local.set(items),
    sendMessage: (message) => {
      void chrome.runtime.sendMessage(message).catch(() => {})
    },
    downloadWorkbook,
    abortState: { aborted: false },
    ...overrides,
  }
}

/**
 * Ports the original `runSummaryScan` loop for one month:
 *  - for each day: `prevOpen` counts BASE_PREV_DATE..day, `received` = day..day
 *    (all statuses), `closed` = day..day (closed only); `pending = prevOpen + received - closed`
 *  - two SUMMARY_PROGRESS messages per day (Processing + Completed, the latter carrying the row)
 *  - after each count, `sleep(delayMs)` then a silent abort check (return without touching
 *    storage / sending SUMMARY_ABORTED — quirk); only the top-of-loop abort flushes
 *    `{ isScanning: false, isAborted: true }` and sends SUMMARY_ABORTED
 *  - errors inside the loop throw; the caller's catch sends SUMMARY_ERROR
 */
export async function runSummaryScan(
  year: number,
  month: number,
  delaySec: number,
  deps: SummaryScanDeps,
): Promise<void> {
  const token = await deps.getToken()
  const delayMs = Math.round(delaySec * 1000)

  const startDate = new Date(year, month - 1, 1)
  const lastDay = new Date(year, month, 0).getDate()
  const endDate = new Date(year, month - 1, lastDay)
  const now = new Date()
  if (endDate > now) {
    endDate.setTime(now.getTime())
  }
  endDate.setHours(23, 59, 59, 999)

  const totalDays = Math.floor((endDate.getTime() - startDate.getTime()) / DAY_MS) + 1
  const current = new Date(startDate)
  let dayIdx = 0
  const rows: SummaryRow[] = []

  await deps.storageSet({
    summaryScanState: {
      isScanning: true,
      year,
      month,
      pct: 0,
      stepText: 'Connecting MyGate API...',
      rows: [],
      isDone: false,
    },
  })

  while (current <= endDate) {
    if (deps.abortState.aborted) {
      await deps.storageSet({ summaryScanState: { isScanning: false, isAborted: true } })
      deps.sendMessage({ type: 'SUMMARY_ABORTED' })
      return
    }

    const dStr = formatDDMMYYYY(current)
    const pct = Math.round((dayIdx / totalDays) * 100)

    await deps.storageSet({
      summaryScanState: {
        isScanning: true,
        year,
        month,
        pct,
        stepText: `Processing ${dStr}...`,
        rows,
        isDone: false,
      },
    })
    deps.sendMessage({ type: 'SUMMARY_PROGRESS', stepText: `Processing ${dStr}...`, pct })

    const prevOpen = await deps.count(BASE_PREV_DATE, dStr, STATUS_PREV_OPEN, token)
    if (deps.abortState.aborted) return
    await deps.sleep(delayMs)

    const received = await deps.count(dStr, dStr, STATUS_ALL, token)
    if (deps.abortState.aborted) return
    await deps.sleep(delayMs)

    const closed = await deps.count(dStr, dStr, STATUS_CLOSED, token)
    if (deps.abortState.aborted) return
    await deps.sleep(delayMs)

    const pending = prevOpen + received - closed
    rows.push({ date: formatDDMMMYYYY(current), prevOpen, received, closed, pending })
    rows.sort((a, b) => parseDateToTime(a.date) - parseDateToTime(b.date))

    dayIdx++
    const newPct = Math.round((dayIdx / totalDays) * 100)

    await deps.storageSet({
      summaryScanState: {
        isScanning: true,
        year,
        month,
        pct: newPct,
        stepText: `Completed ${dStr}`,
        rows,
        isDone: false,
      },
    })
    deps.sendMessage({
      type: 'SUMMARY_PROGRESS',
      stepText: `Completed ${dStr}`,
      pct: newPct,
      row: rows[rows.length - 1],
    })

    current.setDate(current.getDate() + 1)
  }

  await deps.storageSet({
    summaryScanState: {
      isScanning: true,
      year,
      month,
      pct: 100,
      stepText: 'Downloading Excel...',
      rows,
      isDone: false,
    },
  })
  deps.sendMessage({ type: 'SUMMARY_PROGRESS', stepText: 'Downloading Excel...', pct: 100 })

  await deps.downloadWorkbook(buildSummaryWorkbook(rows), summaryFilename(year, month))

  await deps.storageSet({
    summaryScanState: {
      isScanning: false,
      year,
      month,
      pct: 100,
      stepText: 'Report Generated!',
      rows,
      isDone: true,
    },
  })
  deps.sendMessage({ type: 'SUMMARY_DONE' })
}
