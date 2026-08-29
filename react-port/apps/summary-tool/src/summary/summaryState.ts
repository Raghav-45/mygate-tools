export interface SummaryRow {
  date: string
  prevOpen: number
  received: number
  closed: number
  pending: number
}

export interface SummaryScanState {
  isScanning: boolean
  isAborted?: boolean
  error?: string
  year?: number
  month?: number
  pct?: number
  stepText?: string
  rows?: SummaryRow[]
  isDone?: boolean
}

export type PopupToWorkerMessage =
  | { type: 'START_SUMMARY_SCAN'; year: number; month: number; requestDelay: number }
  | { type: 'ABORT_SUMMARY_SCAN' }
  | { type: 'CHECK_AUTH_STATUS' }

export type StartSummaryScan = Extract<PopupToWorkerMessage, { type: 'START_SUMMARY_SCAN' }>
export type AbortSummaryScan = Extract<PopupToWorkerMessage, { type: 'ABORT_SUMMARY_SCAN' }>
export type CheckAuthStatus = Extract<PopupToWorkerMessage, { type: 'CHECK_AUTH_STATUS' }>

export type SummaryWorkerMessage =
  | { type: 'SUMMARY_PROGRESS'; stepText?: string; pct?: number; row?: SummaryRow }
  | { type: 'SUMMARY_DONE' }
  | { type: 'SUMMARY_ABORTED' }
  | { type: 'SUMMARY_ERROR'; error: string }
