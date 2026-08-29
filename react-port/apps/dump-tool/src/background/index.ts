import type { AbortDumpExport, StartDumpExport } from '../dump/dumpState'
import { createDumpDeps, runMasterDumpExport } from '../dump/runDumpExport'

const abortState = { aborted: false }
const deps = createDumpDeps({ abortState })

chrome.runtime.onMessage.addListener((message) => {
  const msg = message as { type?: string }

  if (msg.type === 'START_DUMP_EXPORT') {
    abortState.aborted = false
    const start = message as StartDumpExport
    void runMasterDumpExport(start.params, deps)
  } else if (msg.type === 'ABORT_DUMP_EXPORT') {
    abortState.aborted = true
    void (message as AbortDumpExport)
  }

  // Keep the port open while the async worker run is in flight.
  return true
})
