import { Workbook } from 'exceljs'
import {
  discoverAuthToken,
  downloadWorkbook,
  formatFilterDate,
  postGraphQL,
  sleep,
} from '@mygate/shared'
import type { GraphQLRequest, GraphQLRequestOptions } from '@mygate/shared'
import type { ChunkResult } from './dumpState'
import { buildMasterWorkbook } from './masterWorkbook'
import { parseChunkXlsx } from './parseChunkXlsx'
import type { ChunkParseContext, MergedRow } from './parseChunkXlsx'
import { buildExportRequestPayload, buildPollStatusPayload } from './payloads'
import type { DownloadReport } from './payloads'
import { sliceIntoYears } from './slice'

export const SAMPLE_FALLBACK_TOKEN =
  'zbdzHQCrz1uOVa3Z9QrFabEIA600Udb6lZPrM2SFIkC597iOyCDKllxwR9ZD7Jqa'

/** Match the original hard cap (known to disagree with KNOWN_ISSUES.md's ~60). */
export const MAX_POLL_ATTEMPTS = 45

export interface DumpExportParams {
  fromDate: string
  toDate: string
  requestDelayMs: number
}

export interface DumpDeps {
  /** Returns a discovered token or `null`; the fallback token is used when null. */
  discoverToken: () => Promise<string | null>
  fallbackToken: string
  /** GraphQL POST; non-2xx responses reject (caller distinguishes the rest). */
  postApi: (payload: GraphQLRequest, opts?: GraphQLRequestOptions) => Promise<unknown>
  /** Download a finished cloud report as bytes. */
  fetchFile: (url: string) => Promise<ArrayBuffer | Uint8Array>
  /** Parse a chunk report buffer into an exceljs workbook. */
  loadChunkWorkbook: (buffer: ArrayBuffer | Uint8Array) => Promise<Workbook>
  downloadWorkbook: (workbook: Workbook, filename: string) => Promise<number>
  storageSet: (items: Record<string, unknown>) => Promise<void>
  sendMessage: (message: unknown) => void
  sleep: (ms: number) => Promise<void>
  logError: (message: string, ...args: unknown[]) => void
  abortState: { aborted: boolean }
  pollMaxAttempts: number
}

interface GetAdminSrListBody {
  errors?: unknown
  data?: { getAdminSrList?: unknown }
}

interface PollBody {
  data?: {
    getDownloadReportList?: {
      dataResponse?: { data?: DownloadReport[] }
    }
  }
}

export function createDumpDeps(overrides: Partial<DumpDeps> = {}): DumpDeps {
  return {
    discoverToken: () =>
      discoverAuthToken({
        tabQueryUrl: '*://*.mygate.com/*',
        localStorage: {
          keyContains: ['token'],
          minLength: 20,
          reject: ['{'],
          rejectPrefix: [],
          unwrap: false,
          unwrapFields: [],
        },
      }),
    fallbackToken: SAMPLE_FALLBACK_TOKEN,
    // The original dump request sends NO cookies (default fetch credentials);
    // it authenticates purely through the authorization header.
    postApi: (payload, opts) => postGraphQL(payload, { ...opts, credentials: 'same-origin' }),
    fetchFile: async (url) => {
      const res = await fetch(url)
      return res.arrayBuffer()
    },
    loadChunkWorkbook: async (buffer) => {
      const wb = new Workbook()
      await wb.xlsx.load(buffer)
      return wb
    },
    downloadWorkbook,
    storageSet: (items) => chrome.storage.local.set(items),
    sendMessage: (message) => {
      void chrome.runtime.sendMessage(message).catch(() => {})
    },
    sleep,
    logError: (message, ...args) => console.error(message, ...args),
    abortState: { aborted: false },
    pollMaxAttempts: MAX_POLL_ATTEMPTS,
    ...overrides,
  }
}

async function updateState(
  deps: DumpDeps,
  pct: number,
  statusText: string,
  chunks: ChunkResult[],
  totalRows: number,
): Promise<void> {
  await deps.storageSet({
    dumpScanState: {
      isScanning: true,
      pct,
      statusText,
      chunks,
      totalRows,
      isDone: false,
    },
  })
  deps.sendMessage({ type: 'DUMP_PROGRESS_UPDATE', pct, statusText, chunks, totalRows })
}

async function handleAbort(deps: DumpDeps): Promise<void> {
  await deps.storageSet({
    dumpScanState: { isScanning: false, isAborted: true },
  })
  deps.sendMessage({ type: 'DUMP_ABORTED' })
}

export async function runMasterDumpExport(params: DumpExportParams, deps: DumpDeps): Promise<void> {
  const { fromDate, toDate, requestDelayMs = 2000 } = params

  const tokenToUse = (await deps.discoverToken()) || deps.fallbackToken

  const chunks = sliceIntoYears(fromDate, toDate)
  const chunkResults: ChunkResult[] = []
  let totalRowsMerged = 0

  await deps.storageSet({
    dumpScanState: {
      isScanning: true,
      pct: 0,
      statusText: `Sliced into ${chunks.length} yearly chunks. Initiating...`,
      chunks: [],
      totalRows: 0,
      isDone: false,
    },
  })

  const ctx: ChunkParseContext = { headerRowValues: null }
  const allMergedRows: MergedRow[] = []
  const usedReportLinks = new Set<string>()

  // Process chunks strictly sequentially to prevent MyGate cloud job collisions.
  for (let i = 0; i < chunks.length; i++) {
    if (deps.abortState.aborted) {
      await handleAbort(deps)
      return
    }

    const chunk = chunks[i]
    const filterFrom = formatFilterDate(chunk.fromDate)
    const filterTo = formatFilterDate(chunk.toDate)
    const cr: ChunkResult = {
      rangeStr: `${chunk.fromDate} to ${chunk.toDate}`,
      filterFrom,
      filterTo,
      rowsFound: 'Requesting...',
      status: 'Initiating...',
      downloadUrl: null,
    }
    chunkResults.push(cr)

    const basePct = Math.round((i / chunks.length) * 85)
    await updateState(deps, basePct, `Processing ${cr.rangeStr}...`, chunkResults, totalRowsMerged)

    try {
      const res = await deps.postApi(buildExportRequestPayload(chunk), {
        authorization: tokenToUse,
      })
      const data = res as GetAdminSrListBody
      if (data.errors || !data?.data?.getAdminSrList) {
        throw new Error('Session expired! Please open dashboard.mygate.com and log in.')
      }
      cr.status = 'Generating in Cloud...'
      cr.rowsFound = 'Waiting for Cloud...'
      await updateState(
        deps,
        basePct + 5,
        `Generating report for ${cr.rangeStr}...`,
        chunkResults,
        totalRowsMerged,
      )
    } catch (e) {
      // Original behavior: the auth-failure path never sets isDone, so the popup
      // stays in the "scanning" state on the next open (flagged in NOTES.md).
      deps.logError('Error requesting export:', e)
      cr.rowsFound = 'Auth Error'
      cr.status = 'Failed (Login Required)'
      await updateState(
        deps,
        0,
        'Error: Please open dashboard.mygate.com and log in!',
        chunkResults,
        totalRowsMerged,
      )
      return
    }

    // Poll until this chunk report is ready.
    let chunkReady = false
    let pollAttempts = 0
    const pollPayload = buildPollStatusPayload()

    while (!chunkReady && pollAttempts < deps.pollMaxAttempts) {
      if (deps.abortState.aborted) {
        await handleAbort(deps)
        return
      }
      pollAttempts++
      await deps.sleep(requestDelayMs)

      try {
        const pollBody = (await deps.postApi(pollPayload, {
          authorization: tokenToUse,
        })) as PollBody
        const reports = pollBody?.data?.getDownloadReportList?.dataResponse?.data || []

        const match = reports.find((r) => {
          if (
            r.report_name !== 'Helpdesk Report' ||
            r.status !== 'Success' ||
            !r.report_link ||
            usedReportLinks.has(r.report_link)
          )
            return false
          const df = r.download_filters || {}
          const matchFrom = df['From'] === filterFrom || df['Date From'] === filterFrom
          const matchTo = df['To'] === filterTo || df['Date To'] === filterTo
          return matchFrom && matchTo
        })

        if (match && match.report_link) {
          usedReportLinks.add(match.report_link)
          cr.downloadUrl = match.report_link
          chunkReady = true
        }
      } catch (e) {
        deps.logError('Polling error:', e)
      }
    }

    if (!cr.downloadUrl) {
      cr.status = 'Timed out'
      cr.rowsFound = '0'
      continue
    }

    cr.status = 'Downloading & Parsing...'
    await updateState(
      deps,
      basePct + 15,
      `Merging ${cr.rangeStr}...`,
      chunkResults,
      totalRowsMerged,
    )

    try {
      const arrayBuf = await deps.fetchFile(cr.downloadUrl)
      const chunkWb = await deps.loadChunkWorkbook(arrayBuf)
      const { rows, rowsInChunk } = parseChunkXlsx(chunkWb, ctx)
      allMergedRows.push(...rows)

      cr.rowsFound = rowsInChunk.toLocaleString()
      cr.status = 'Merged ✅'
      totalRowsMerged += rowsInChunk
      await updateState(
        deps,
        basePct + 25,
        `Merged ${rowsInChunk} rows from ${cr.rangeStr}`,
        chunkResults,
        totalRowsMerged,
      )
    } catch (e) {
      deps.logError('Error downloading chunk:', e)
      cr.status = 'Merge Error'
    }
  }

  // Build standardized master spreadsheet and auto-download.
  await updateState(
    deps,
    96,
    'Compiling standardized master spreadsheet...',
    chunkResults,
    totalRowsMerged,
  )

  const masterWb = buildMasterWorkbook(allMergedRows, fromDate, toDate)
  const filename = `MyGate_Master_Helpdesk_Dump_${fromDate}_to_${toDate}.xlsx`

  await deps.downloadWorkbook(masterWb, filename)

  await deps.storageSet({
    dumpScanState: {
      isScanning: false,
      pct: 100,
      statusText: 'Master Dump Downloaded!',
      chunks: chunkResults,
      totalRows: totalRowsMerged,
      isDone: true,
    },
  })

  deps.sendMessage({ type: 'DUMP_FINISHED', chunks: chunkResults, totalRows: totalRowsMerged })
}
