import { CATEGORIES } from '../report/categories'
import type {
  AbortReportScan,
  GetCategoriesList,
  PopupToWorkerMessage,
} from '../report/reportState'
import { createReportDeps, runReportScan } from '../report/runReportScan'

const abortState = { aborted: false }
const deps = createReportDeps({ abortState })

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  const msg = message as PopupToWorkerMessage

  switch (msg.type) {
    case 'START_REPORT_SCAN':
      abortState.aborted = false
      void runReportScan(msg.params, deps)
      break
    case 'ABORT_REPORT_SCAN':
      abortState.aborted = true
      void (message as AbortReportScan)
      break
    case 'GET_CATEGORIES_LIST':
      void (message as GetCategoriesList)
      sendResponse({ categories: CATEGORIES })
      break
  }

  // Keep the port open while the async worker run is in flight.
  return true
})
