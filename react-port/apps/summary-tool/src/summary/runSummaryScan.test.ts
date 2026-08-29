import { formatDDMMYYYY, formatDDMMMYYYY } from '@mygate/shared'
import { describe, expect, it } from 'vitest'
import { STATUS_ALL, STATUS_CLOSED, STATUS_PREV_OPEN } from './countQueries'
import { BASE_PREV_DATE, createSummaryDeps, runSummaryScan } from './runSummaryScan'
import type { SummaryScanDeps } from './runSummaryScan'
import type { SummaryScanState } from './summaryState'

const YEAR = 2024
const MONTH = 1

function createHarness(overrides: Partial<SummaryScanDeps> = {}) {
  const messages: unknown[] = []
  const storageSets: Record<string, unknown>[] = []
  const downloads: string[] = []
  const countCalls: Array<{ from: string; to: string; statuses: string[]; token: string }> = []
  const sleepMs: number[] = []
  const countResults = new Map<string, number>()

  const seed = (from: string, to: string, statuses: string[], value: number) => {
    countResults.set(`${from}|${to}|${statuses.join('|')}`, value)
  }

  const deps: SummaryScanDeps = {
    discoverToken: async () => 'TEST_TOKEN',
    getToken: async () => 'TEST_TOKEN',
    count: async (from, to, statuses, token) => {
      countCalls.push({ from, to, statuses, token })
      const key = `${from}|${to}|${statuses.join('|')}`
      const value = countResults.get(key)
      if (value === undefined) throw new Error(`no count result registered for ${key}`)
      return value
    },
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
    abortState: { aborted: false },
    ...overrides,
  }

  return { deps, messages, storageSets, downloads, countCalls, sleepMs, countResults, seed }
}

function dayStrings(): string[] {
  return Array.from({ length: 31 }, (_, i) => {
    const d = new Date(YEAR, MONTH - 1, i + 1)
    return formatDDMMYYYY(d)
  })
}

function seedMonth(
  h: ReturnType<typeof createHarness>,
  prevOpen: number,
  received: number,
  closed: number,
) {
  for (const dd of dayStrings()) {
    h.seed(BASE_PREV_DATE, dd, STATUS_PREV_OPEN, prevOpen)
    h.seed(dd, dd, STATUS_ALL, received)
    h.seed(dd, dd, STATUS_CLOSED, closed)
  }
}

function lastState(h: ReturnType<typeof createHarness>): SummaryScanState {
  const last = h.storageSets[h.storageSets.length - 1]
  return last.summaryScanState as SummaryScanState
}

describe('runSummaryScan', () => {
  it('runs prevOpen -> received -> closed per day and finishes with a download', async () => {
    const h = createHarness()
    seedMonth(h, 100, 5, 3)

    await runSummaryScan(YEAR, MONTH, 1.5, h.deps)

    expect(h.downloads).toEqual(['Complaint Summary Sheet - Jan-2024.xlsx'])
    expect(h.countCalls).toHaveLength(31 * 3)

    // prevOpen always counts from the hard-coded base date; received/closed are single-day.
    for (const dd of dayStrings()) {
      expect(h.countCalls.some((c) => c.from === BASE_PREV_DATE && c.to === dd)).toBe(true)
      expect(h.countCalls.some((c) => c.from === dd && c.to === dd)).toBe(true)
    }

    // Every query carries the token; three sleeps per day at the configured delay.
    expect(h.countCalls.every((c) => c.token === 'TEST_TOKEN')).toBe(true)
    expect(h.sleepMs).toHaveLength(31 * 3)
    expect(h.sleepMs.every((ms) => ms === 1500)).toBe(true)

    // Progress: Processing + Completed message per day, then Downloading, then DONE.
    const progress = h.messages.filter(
      (m): m is { type: string; stepText: string; pct: number } =>
        (m as { type?: string }).type === 'SUMMARY_PROGRESS',
    )
    expect(progress).toHaveLength(31 * 2 + 1)
    expect(progress[0]).toMatchObject({ stepText: 'Processing 01-01-2024...', pct: 0 })
    expect(progress[1]).toMatchObject({ stepText: 'Completed 01-01-2024', pct: 3 })
    expect(progress[progress.length - 1]).toMatchObject({
      stepText: 'Downloading Excel...',
      pct: 100,
    })

    // First message of each day flow is the initial Connecting state.
    expect(h.storageSets[0].summaryScanState).toMatchObject({
      isScanning: true,
      pct: 0,
      stepText: 'Connecting MyGate API...',
      isDone: false,
    })

    // Completed rows carry the resolved date, pending math and chronological sort.
    const completedRows = messagesWithRows(h)
    expect(completedRows).toHaveLength(31)
    expect(completedRows[0].row).toMatchObject({
      date: '01-Jan-2024',
      prevOpen: 100,
      received: 5,
      closed: 3,
      pending: 102,
    })
    const dates = completedRows.map((m) => m.row.date)
    expect(dates).toEqual(
      dayStrings().map((dd) =>
        formatDDMMMYYYY(new Date(YEAR, MONTH - 1, Number(dd.split('-')[0]))),
      ),
    )

    // Final state: done, rows persisted.
    const state = lastState(h)
    expect(state).toMatchObject({
      isScanning: false,
      pct: 100,
      stepText: 'Report Generated!',
      isDone: true,
    })
    expect(state.rows ?? []).toHaveLength(31)
    expect(h.messages[h.messages.length - 1]).toEqual({ type: 'SUMMARY_DONE' })
  })

  it('aborts at the top of the next day loop with SUMMARY_ABORTED and no download', async () => {
    const h = createHarness()
    seedMonth(h, 100, 5, 3)
    h.deps.abortState.aborted = true

    await runSummaryScan(YEAR, MONTH, 1.0, h.deps)

    const state = lastState(h)
    expect(state).toMatchObject({ isScanning: false, isAborted: true })
    expect(h.messages).toContainEqual({ type: 'SUMMARY_ABORTED' })
    expect(h.downloads).toEqual([])
  })

  it('aborting between daily counts returns silently without SUMMARY_ABORTED (ported quirk)', async () => {
    const h = createHarness()
    seedMonth(h, 100, 5, 3)
    h.deps.sleep = async () => {
      h.deps.abortState.aborted = true
    }

    await runSummaryScan(YEAR, MONTH, 1.0, h.deps)

    // No abort flush, no done -> popup stays stuck on the last Processing state.
    expect(h.messages.some((m) => (m as { type?: string }).type === 'SUMMARY_ABORTED')).toBe(false)
    expect(h.messages.some((m) => (m as { type?: string }).type === 'SUMMARY_DONE')).toBe(false)
    expect(h.downloads).toEqual([])
    const state = lastState(h)
    expect(state.isScanning).toBe(true)
    expect(state.stepText).toBe('Processing 01-01-2024...')
  })

  it('rejects when no token is found (no fallback in the summary tool)', async () => {
    const h = createHarness({
      getToken: async () => {
        throw new Error('No MyGate token found! Please open dashboard.mygate.com and log in.')
      },
    })

    await expect(runSummaryScan(YEAR, MONTH, 1.0, h.deps)).rejects.toThrow(
      'No MyGate token found! Please open dashboard.mygate.com and log in.',
    )
    expect(h.downloads).toEqual([])
    expect(h.storageSets).toEqual([])
  })

  it('lets per-day count errors propagate so the caller can send SUMMARY_ERROR', async () => {
    const h = createHarness()
    seedMonth(h, 100, 5, 3)
    h.deps.count = async () => {
      throw new Error('HTTP 500')
    }

    await expect(runSummaryScan(YEAR, MONTH, 1.0, h.deps)).rejects.toThrow('HTTP 500')
    expect(h.downloads).toEqual([])
    expect(h.messages.some((m) => (m as { type?: string }).type === 'SUMMARY_DONE')).toBe(false)
  })

  it('default createSummaryDeps.getToken throws the original message when nothing is discoverable', async () => {
    const deps = createSummaryDeps()
    await expect(deps.getToken()).rejects.toThrow(
      'No MyGate token found! Please open dashboard.mygate.com and log in.',
    )
  })
})

function messagesWithRows(h: ReturnType<typeof createHarness>) {
  return h.messages.filter(
    (m): m is { type: string; stepText: string; pct: number; row: { date: string } } =>
      (m as { type?: string }).type === 'SUMMARY_PROGRESS' &&
      (m as { row?: unknown }).row !== undefined,
  )
}
