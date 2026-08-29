/** Shared message + storage contracts between the dump-tool worker and popup. */

export interface ChunkResult {
  rangeStr: string
  filterFrom: string
  filterTo: string
  rowsFound: string
  status: string
  downloadUrl: string | null
}

export interface DumpScanState {
  isScanning: boolean
  pct: number
  statusText: string
  chunks: ChunkResult[]
  totalRows: number
  isDone: boolean
  isAborted?: boolean
}

export interface DumpProgressUpdate {
  type: 'DUMP_PROGRESS_UPDATE'
  pct: number
  statusText: string
  chunks: ChunkResult[]
  totalRows: number
}

export interface DumpFinished {
  type: 'DUMP_FINISHED'
  chunks: ChunkResult[]
  totalRows: number
}

export interface DumpAborted {
  type: 'DUMP_ABORTED'
}

export type DumpWorkerMessage = DumpProgressUpdate | DumpFinished | DumpAborted

export interface StartDumpExport {
  type: 'START_DUMP_EXPORT'
  params: { fromDate: string; toDate: string; requestDelayMs: number }
}

export interface AbortDumpExport {
  type: 'ABORT_DUMP_EXPORT'
}
