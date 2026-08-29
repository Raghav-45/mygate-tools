import type { PopupToWorkerMessage } from '../summary/summaryState'
import { createSummaryDeps, runSummaryScan } from '../summary/runSummaryScan'

const abortState = { aborted: false }
const deps = createSummaryDeps({ abortState })

chrome.runtime.onMessage.addListener((rawMessage, _sender, sendResponse) => {
  const message = rawMessage as PopupToWorkerMessage

  switch (message.type) {
    case 'CHECK_AUTH_STATUS':
      void deps
        .discoverToken()
        .then((token) => sendResponse({ hasToken: !!token }))
        .catch(() => sendResponse({ hasToken: false }))
      break

    case 'ABORT_SUMMARY_SCAN':
      abortState.aborted = true
      sendResponse({ aborted: true })
      break

    case 'START_SUMMARY_SCAN':
      abortState.aborted = false
      void runSummaryScan(message.year, message.month, message.requestDelay || 1.5, deps).catch(
        (err: unknown) => {
          const errorMessage = err instanceof Error ? err.message : String(err)
          void deps.storageSet({ summaryScanState: { isScanning: false, error: errorMessage } })
          deps.sendMessage({ type: 'SUMMARY_ERROR', error: errorMessage })
        },
      )
      sendResponse({ started: true })
      break
  }

  return true
})
