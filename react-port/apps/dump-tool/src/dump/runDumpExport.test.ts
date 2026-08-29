import { Workbook } from 'exceljs'
import { describe, expect, it } from 'vitest'
import { runMasterDumpExport } from './runDumpExport'
import type { DumpDeps, DumpExportParams } from './runDumpExport'
import type { ChunkResult } from './dumpState'

async function makeChunkBuffer(rowCount: number): Promise<Uint8Array> {
  const wb = new Workbook()
  const ws = wb.addWorksheet('chunk')
  const pos = [1, 2, 4, 5, 7, 9, 10]
  const names = ['I.D', 'Created Date', 'Category', 'Sub Category', 'Flat', 'Subject', 'Status']
  pos.forEach((col, i) => {
    ws.getRow(3).getCell(col).value = names[i]
  })
  for (let r = 0; r < rowCount; r++) {
    const rowNum = 4 + r
    ws.getRow(rowNum).getCell(1).value = 1000 + r
    ws.getRow(rowNum).getCell(2).value = new Date(2024, 0, 15)
    ws.getRow(rowNum).getCell(4).value = 'Accounts'
    ws.getRow(rowNum).getCell(5).value = 'Billing'
    ws.getRow(rowNum).getCell(7).value = 'A-101'
    ws.getRow(rowNum).getCell(9).value = `Subject ${r}`
    ws.getRow(rowNum).getCell(10).value = 'closed'
  }
  return wb.xlsx.writeBuffer() as unknown as Uint8Array
}

/** GraphQL fake server with the same contract the worker relies on. */
type ServerFn = (
  operationName: string,
  variables: Record<string, unknown> | undefined,
  authorization: string | undefined,
) => Promise<unknown>

function createHarness(server: ServerFn, rowsPerChunk = 2) {
  const messages: unknown[] = []
  const storageSets: Record<string, unknown>[] = []
  const downloads: string[] = []
  const sleepMs: number[] = []
  const postApiCalls: { op: string; auth?: string }[] = []
  const logErrors: unknown[] = []

  const deps: DumpDeps = {
    discoverToken: async () => null,
    fallbackToken: 'FAKE_FALLBACK_TOKEN',
    postApi: async (payload, opts) => {
      const op = (payload as { operationName: string }).operationName
      postApiCalls.push({ op, auth: opts?.authorization })
      const vars = (payload as { variables?: Record<string, unknown> }).variables
      return server(op, vars, opts?.authorization)
    },
    fetchFile: async () => makeChunkBuffer(rowsPerChunk),
    loadChunkWorkbook: async (buffer) => {
      const wb = new Workbook()
      await wb.xlsx.load(buffer)
      return wb
    },
    downloadWorkbook: async (_wb, filename) => {
      downloads.push(filename)
      return 123
    },
    storageSet: async (items) => {
      storageSets.push(items)
    },
    sendMessage: (message) => {
      messages.push(message)
    },
    sleep: async (ms) => {
      sleepMs.push(ms)
    },
    logError: (message, ...args) => {
      logErrors.push([message, ...args])
    },
    abortState: { aborted: false },
    pollMaxAttempts: 45,
  }

  return { deps, messages, storageSets, downloads, sleepMs, postApiCalls, logErrors }
}

const PARAMS: DumpExportParams = {
  fromDate: '2024-01-01',
  toDate: '2025-12-31',
  requestDelayMs: 2000,
}

function lastState(h: ReturnType<typeof createHarness>) {
  return h.storageSets[h.storageSets.length - 1].dumpScanState as unknown as {
    isScanning: boolean
    isDone: boolean
    isAborted?: boolean
    pct: number
    statusText: string
    totalRows: number
    chunks: ChunkResult[]
  }
}

describe('runMasterDumpExport', () => {
  it('slices, requests, polls, merges, builds the master workbook and downloads it', async () => {
    // Fake server: within a chunk, the report is ready on the 3rd poll.
    let active: { From: string; To: string } | null = null
    let pollsSinceRequest = 0

    const h = createHarness(async (op, vars) => {
      if (op === 'getAdminSrList') {
        const df = (
          vars as {
            requestData?: { downloadFilters?: { From: string; To: string } }
          }
        ).requestData?.downloadFilters
        active = df ? { From: df.From, To: df.To } : null
        pollsSinceRequest = 0
        return { data: { getAdminSrList: {} } }
      }
      pollsSinceRequest += 1
      if (pollsSinceRequest < 3 || !active) {
        return { data: { getDownloadReportList: { dataResponse: { data: [] } } } }
      }
      return {
        data: {
          getDownloadReportList: {
            dataResponse: {
              data: [
                {
                  report_name: 'Helpdesk Report',
                  status: 'Success',
                  // Unique per chunk so the worker's usedReportLinks dedupe
                  // does not reject the second chunk's report.
                  report_link: `https://cloudfront.example/${active.From.replaceAll(':', '')}.xlsx`,
                  download_filters: { From: active.From, To: active.To },
                },
              ],
            },
          },
        },
      }
    })

    await runMasterDumpExport(PARAMS, h.deps)

    expect(h.downloads).toEqual(['MyGate_Master_Helpdesk_Dump_2024-01-01_to_2025-12-31.xlsx'])

    const state = lastState(h)
    expect(state.isDone).toBe(true)
    expect(state.pct).toBe(100)
    expect(state.statusText).toBe('Master Dump Downloaded!')
    expect(state.totalRows).toBe(4)
    expect(state.chunks.map((c) => c.status)).toEqual(['Merged ✅', 'Merged ✅'])

    const finished = h.messages[h.messages.length - 1] as {
      type: string
      totalRows: number
      chunks: ChunkResult[]
    }
    expect(finished.type).toBe('DUMP_FINISHED')
    expect(finished.totalRows).toBe(4)

    expect(h.postApiCalls.filter((c) => c.op === 'getAdminSrList').length).toBe(2)
    expect(h.postApiCalls.filter((c) => c.op === 'getDownloadReportList').length).toBe(6)
    h.postApiCalls.forEach((c) => expect(c.auth).toBe('FAKE_FALLBACK_TOKEN'))
    h.sleepMs.forEach((ms) => expect(ms).toBe(2000))
  })

  it('times out a chunk after the poll cap and completes the master dump', async () => {
    let polls = 0
    const h = createHarness(async (op) => {
      if (op === 'getAdminSrList') return { data: { getAdminSrList: {} } }
      polls += 1
      return { data: { getDownloadReportList: { dataResponse: { data: [] } } } }
    })

    await runMasterDumpExport(
      { fromDate: '2024-01-01', toDate: '2024-06-30', requestDelayMs: 2000 },
      h.deps,
    )

    expect(polls).toBe(45)
    const state = lastState(h)
    expect(state.chunks[0].status).toBe('Timed out')
    expect(state.chunks[0].rowsFound).toBe('0')
    expect(state.isDone).toBe(true)
  })

  it('marks the chunk as auth-failed and stops (ported quirk: isScanning stays true)', async () => {
    const h = createHarness(async (op) => {
      if (op === 'getAdminSrList') throw new Error('HTTP 401')
      return { data: {} }
    })

    await runMasterDumpExport(PARAMS, h.deps)

    const state = lastState(h)
    expect(state.chunks[0].status).toBe('Failed (Login Required)')
    expect(state.chunks[0].rowsFound).toBe('Auth Error')
    expect(state.statusText).toBe('Error: Please open dashboard.mygate.com and log in!')
    expect(state.isScanning).toBe(true) // ported quirk
    expect(state.isDone).toBe(false)
    expect(h.messages.some((m) => (m as { type: string }).type === 'DUMP_FINISHED')).toBe(false)
  })

  it('aborts cleanly when the abort flag is set before starting', async () => {
    const h = createHarness(async (op) => {
      if (op === 'getAdminSrList') return { data: { getAdminSrList: {} } }
      return { data: { getDownloadReportList: { dataResponse: { data: [] } } } }
    })
    h.deps.abortState.aborted = true

    await runMasterDumpExport(PARAMS, h.deps)

    expect(lastState(h)).toEqual({ isScanning: false, isAborted: true })
    expect((h.messages[h.messages.length - 1] as { type: string }).type).toBe('DUMP_ABORTED')
    expect(h.logErrors).toEqual([])
  })

  it('aborts mid-run before processing the second chunk', async () => {
    let exports = 0
    const h = createHarness(async (op) => {
      if (op === 'getAdminSrList') {
        exports += 1
        if (exports >= 2) h.deps.abortState.aborted = true
        return { data: { getAdminSrList: {} } }
      }
      return { data: { getDownloadReportList: { dataResponse: { data: [] } } } }
    })

    await runMasterDumpExport(PARAMS, h.deps)

    expect(lastState(h).isAborted).toBe(true)
    expect((h.messages[h.messages.length - 1] as { type: string }).type).toBe('DUMP_ABORTED')
  })
})
