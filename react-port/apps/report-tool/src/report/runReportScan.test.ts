import { getMidnightEpoch } from '@mygate/shared'
import { describe, expect, it } from 'vitest'
import { STATUS_OPEN, STATUS_RESOLVED, STATUS_TOTAL } from './countQueries'
import { runReportScan } from './runReportScan'
import type { ReportScanDeps } from './runReportScan'
import type { ReportScanParams } from './reportState'

function createHarness(overrides: Partial<ReportScanDeps> = {}) {
  const messages: unknown[] = []
  const storageSets: Record<string, unknown>[] = []
  const downloads: string[] = []
  const countCalls: Array<{ categoryId: number; statuses: string[]; token: string }> = []
  const sleepMs: number[] = []
  const logErrors: unknown[] = []
  const countResults = new Map<string, number>()

  const record = (categoryId: number, statuses: string[], token: string) => {
    countCalls.push({ categoryId, statuses, token })
    const key = `${categoryId}:${statuses.join('|')}`
    const value = countResults.get(key)
    if (value === undefined) throw new Error(`no count result registered for ${key}`)
    return value
  }

  const deps: ReportScanDeps = {
    discoverToken: async () => null,
    fallbackToken: 'FAKE_FALLBACK_TOKEN',
    countRequest: async (categoryId, _fromEpoch, _toEpoch, statuses, token) =>
      record(categoryId, statuses, token),
    sleep: async (ms) => {
      sleepMs.push(ms)
    },
    storageSet: async (items) => {
      storageSets.push(items)
    },
    sendMessage: (message) => {
      messages.push(message)
    },
    downloadWorkbook: async (_wb, filename) => {
      downloads.push(filename)
      return 123
    },
    logError: (message, ...args) => {
      logErrors.push([message, ...args])
    },
    abortState: { aborted: false },
    ...overrides,
  }

  return { deps, messages, storageSets, downloads, countCalls, sleepMs, logErrors, countResults }
}

const PARAMS: ReportScanParams = {
  fromDate: '2024-01-01',
  toDate: '2024-01-31',
  requestDelayMs: 1000,
  selectedCategoryIds: [252434, 277747],
}

function seedCategory(h: ReturnType<typeof createHarness>, id: number, counts: number[]) {
  h.countResults.set(`${id}:${STATUS_TOTAL.join('|')}`, counts[0])
  h.countResults.set(`${id}:${STATUS_RESOLVED.join('|')}`, counts[1])
  h.countResults.set(`${id}:${STATUS_OPEN.join('|')}`, counts[2])
}

function progressUpdates(messages: unknown[]) {
  return messages.filter(
    (m): m is { type: string; stepIndex?: number; totalSteps?: number; statusText: string } =>
      (m as { type?: string }).type === 'SCAN_PROGRESS_UPDATE',
  )
}

describe('runReportScan', () => {
  it('fetches total -> resolved -> open per category, finishes and downloads the report', async () => {
    const h = createHarness()
    seedCategory(h, 252434, [10, 4, 6])
    seedCategory(h, 277747, [5, 5, 0])

    await runReportScan(PARAMS, h.deps)

    expect(h.downloads).toEqual(['Pending_Mygate_Tickets_Report_31-1-2024.xlsx'])

    const last = h.storageSets[h.storageSets.length - 1]
    const state = last.ticketsScanState as {
      isScanning: boolean
      isDone: boolean
      pct: number
      statusText: string
      results: unknown[]
      summary: { total: number; resolved: number; open: number }
    }
    expect(state.isScanning).toBe(false)
    expect(state.isDone).toBe(true)
    expect(state.pct).toBe(100)
    expect(state.statusText).toBe('Report Generated!')
    expect(state.results).toHaveLength(2)
    expect(state.summary).toEqual({ total: 15, resolved: 9, open: 6 })

    // One count query per status (total, resolved, open) per category.
    expect(h.countCalls.map((c) => c.statuses)).toEqual([
      STATUS_TOTAL,
      STATUS_RESOLVED,
      STATUS_OPEN,
      STATUS_TOTAL,
      STATUS_RESOLVED,
      STATUS_OPEN,
    ])
    expect(h.countCalls.every((c) => c.token === 'FAKE_FALLBACK_TOKEN')).toBe(true)

    // Epochs match the original: midnight of from, midnight of to + 86399.
    const used = h.countCalls.map((c) => c.statuses.join('|'))
    expect(used.filter((s) => s === STATUS_TOTAL.join('|'))).toHaveLength(2)

    // SCAN_FINISHED carries results, summary and reportMeta.
    const finished = h.messages[h.messages.length - 1] as {
      type: string
      reportMeta: { fromDate: string; toDate: string }
    }
    expect(finished.type).toBe('SCAN_FINISHED')
    expect(finished.reportMeta).toEqual({ fromDate: '2024-01-01', toDate: '2024-01-31' })

    // Progress updates carry stepIndex/totalSteps for every step except the
    // final "Downloading Excel..." one (leaving the popup's pct as NaN).
    const updates = progressUpdates(h.messages)
    const withSteps = updates.filter((u) => u.stepIndex !== undefined)
    expect(withSteps).toHaveLength(6)
    expect(withSteps[0]).toMatchObject({ stepIndex: 1, totalSteps: 2 })
    expect(withSteps[5]).toMatchObject({ stepIndex: 2, totalSteps: 2 })

    const downloading = updates.find((u) => u.statusText === 'Downloading Excel...')
    expect(downloading).toBeDefined()
    expect(downloading).not.toHaveProperty('stepIndex')
    expect(downloading).not.toHaveProperty('totalSteps')

    // 3 sleeps per category at the requested delay.
    expect(h.sleepMs).toEqual(Array(6).fill(1000))
  })

  it('prepends the discovered token when one is found', async () => {
    const h = createHarness({ discoverToken: async () => 'MY_DISCOVERED_TOKEN' })
    seedCategory(h, 252434, [1, 1, 0])
    await runReportScan({ ...PARAMS, selectedCategoryIds: [252434] }, h.deps)
    expect(h.countCalls.every((c) => c.token === 'MY_DISCOVERED_TOKEN')).toBe(true)
  })

  it('returns early without side effects when no category survives selection', async () => {
    const h = createHarness()
    await runReportScan({ ...PARAMS, selectedCategoryIds: [999999] }, h.deps)
    expect(h.downloads).toEqual([])
    expect(h.messages).toEqual([])
    expect(h.storageSets).toEqual([])
  })

  it('uses the fallback token and reports category errors via SCAN_ERROR', async () => {
    const h = createHarness()
    seedCategory(h, 277747, [5, 5, 0])

    await runReportScan({ ...PARAMS, selectedCategoryIds: [252434, 277747] }, h.deps)

    // First category's "total" query throws (no result registered).
    const errorMsg = h.messages[h.messages.length - 1] as {
      type: string
      errorMessage: string
      category: string
    }
    expect(errorMsg.type).toBe('SCAN_ERROR')
    expect(errorMsg.category).toBe('Accounts Billing')
    expect(errorMsg.errorMessage).toMatch(/no count result registered for 252434:/)

    const last = h.storageSets[h.storageSets.length - 1]
    expect(last.ticketsScanState).toEqual({ isScanning: false, error: errorMsg.errorMessage })
    expect(h.downloads).toEqual([])
    expect(h.logErrors.length).toBeGreaterThan(0)
  })

  it('handles abort at loop top with SCAN_ABORTED and aborted scan state', async () => {
    const h = createHarness({ abortState: { aborted: true } })
    seedCategory(h, 252434, [1, 1, 0])

    await runReportScan(PARAMS, h.deps)

    const last = h.storageSets[h.storageSets.length - 1]
    expect(last.ticketsScanState).toEqual({ isScanning: false, isAborted: true })
    const aborted = h.messages[h.messages.length - 1] as { type: string }
    expect(aborted.type).toBe('SCAN_ABORTED')
    expect(h.downloads).toEqual([])
  })

  it('mid-category abort breaks out silently (ported bug: no state flush, no SCAN_ABORTED)', async () => {
    const h = createHarness()
    seedCategory(h, 252434, [10, 4, 6])
    seedCategory(h, 277747, [5, 5, 0])

    // First sleep of the first category flips the abort flag -> the break path.
    let firstSleep = true
    h.deps.sleep = async () => {
      if (firstSleep) {
        firstSleep = false
        h.deps.abortState.aborted = true
      }
    }

    await runReportScan(PARAMS, h.deps)

    expect(h.downloads).toEqual([])
    expect(h.messages.some((m) => (m as { type?: string }).type === 'SCAN_FINISHED')).toBe(false)
    expect(h.messages.some((m) => (m as { type?: string }).type === 'SCAN_ABORTED')).toBe(false)
    // Scan state is left untouched from the last completed step.
    const last = h.storageSets[h.storageSets.length - 1]
    const state = last.ticketsScanState as { isScanning?: boolean; isAborted?: boolean }
    expect(state.isScanning).not.toBe(false)
  })

  it('uses midnight-of-from and midnight-of-to + 86399 epochs', async () => {
    const h = createHarness()
    seedCategory(h, 252434, [1, 1, 0])
    const seen: Array<number | undefined> = []
    let pushed = 0
    h.deps.countRequest = async (categoryId, fromEpoch, toEpoch, statuses, _token) => {
      seen.push(fromEpoch)
      seen.push(toEpoch)
      pushed++
      return h.countResults.get(`${categoryId}:${statuses.join('|')}`) ?? 0
    }

    await runReportScan({ ...PARAMS, selectedCategoryIds: [252434] }, h.deps)
    expect(pushed).toBe(3)

    for (let i = 0; i < seen.length; i += 2) {
      expect(seen[i]).toBe(getMidnightEpoch('2024-01-01'))
      expect(seen[i + 1]).toBe(getMidnightEpoch('2024-01-31') + 86399)
    }
  })
})
