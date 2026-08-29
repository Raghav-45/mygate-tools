export interface CategoryData {
  name: string
  id: number
  total: number
  resolved: number
  open: number
}

export interface ReportSummary {
  total: number
  resolved: number
  open: number
}

export interface ReportScanParams {
  fromDate: string
  toDate: string
  requestDelayMs: number
  selectedCategoryIds: number[]
}

export interface ReportScanState {
  isScanning: boolean
  isAborted?: boolean
  error?: string
  pct?: number
  statusText?: string
  results?: CategoryData[]
  summary?: ReportSummary
  isDone?: boolean
}

export type PopupToWorkerMessage =
  | { type: 'START_REPORT_SCAN'; params: ReportScanParams }
  | { type: 'ABORT_REPORT_SCAN' }
  | { type: 'GET_CATEGORIES_LIST' }

export type StartReportScan = Extract<PopupToWorkerMessage, { type: 'START_REPORT_SCAN' }>
export type AbortReportScan = Extract<PopupToWorkerMessage, { type: 'ABORT_REPORT_SCAN' }>
export type GetCategoriesList = Extract<PopupToWorkerMessage, { type: 'GET_CATEGORIES_LIST' }>

export type ReportWorkerMessage =
  | {
      type: 'SCAN_PROGRESS_UPDATE'
      stepIndex?: number
      totalSteps?: number
      currentCategory?: string
      statusText: string
    }
  | { type: 'CATEGORY_COMPLETED'; data: CategoryData; summary: ReportSummary }
  | {
      type: 'SCAN_FINISHED'
      results: CategoryData[]
      summary: ReportSummary
      reportMeta: { fromDate: string; toDate: string }
    }
  | { type: 'SCAN_ABORTED'; results: CategoryData[] }
  | { type: 'SCAN_ERROR'; category: string; errorMessage: string }
