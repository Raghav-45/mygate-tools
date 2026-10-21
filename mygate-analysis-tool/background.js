/**
 * MyGate Analysis Tool — Background Service Worker
 *
 * Converts a MyGate Master Helpdesk Dump (.xlsx) into the standardized
 * "Mygate Complaints Dashboard as on <DD.MM.YY>.xlsx" workbook and
 * auto-downloads it.
 *
 * The conversion engine (engine.js + constants.js + chart_templates/) is a
 * faithful port of the frozen reference generator: the produced workbook is
 * part-for-part byte-identical to that reference build (both stamp their
 * own build time into docProps/core.xml and the zip container).
 */

importScripts('constants.js', 'engine.js')

function base64ToArrayBuffer(base64) {
  const binary = atob(base64)
  const len = binary.length
  const bytes = new Uint8Array(len)
  for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

function arrayBufferToBase64(buffer) {
  let binary = ''
  const bytes = new Uint8Array(buffer)
  const len = bytes.byteLength
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i])
  }
  return btoa(binary)
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === 'BUILD_DASHBOARD') {
    runAnalysisBuild(message.params)
  }
  return true
})

// Fast single-pass build (seconds). Progress messages are kept so the popup
// can show the same live UX as the other MyGate tools.
async function runAnalysisBuild(params) {
  const { dumpBase64, asOn, dumpFileName } = params

  await updateState({
    isBuilding: true,
    pct: 5,
    statusText: 'Reading the dump file...',
    isDone: false,
  })

  try {
    const dumpBytes = base64ToArrayBuffer(dumpBase64)

    await updateState({
      isBuilding: true,
      pct: 35,
      statusText: 'Parsing dump & aggregating categories...',
      isDone: false,
    })

    const result = await self.AnalysisEngine.generate(dumpBytes, {
      dumpFileName,
      asOn: asOn || null,
    })

    await updateState({
      isBuilding: true,
      pct: 85,
      statusText: 'Packing the analysis workbook...',
      isDone: false,
    })

    const base64 = arrayBufferToBase64(result.bytes.buffer || result.bytes)
    const dataUrl =
      'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' +
      base64

    await new Promise((resolve, reject) => {
      chrome.downloads.download(
        {
          url: dataUrl,
          filename: result.info.filename,
          saveAs: false,
        },
        (downloadId) => {
          if (chrome.runtime.lastError || !downloadId) {
            reject(
              new Error(
                chrome.runtime.lastError?.message ||
                  'Download blocked by browser permission',
              ),
            )
          } else {
            resolve(downloadId)
          }
        },
      )
    })

    await chrome.storage.local.set({
      dashboardScanState: {
        isBuilding: false,
        pct: 100,
        statusText: 'Analysis Downloaded!',
        info: result.info,
        isDone: true,
      },
    })
    chrome.runtime
      .sendMessage({ type: 'DASHBOARD_DONE', info: result.info })
      .catch(() => {})
  } catch (err) {
    console.error('Analysis build failed:', err)
    await chrome.storage.local.set({
      dashboardScanState: {
        isBuilding: false,
        pct: 0,
        statusText:
          'Error: ' + (err && err.message ? err.message : String(err)),
        isDone: false,
        isFailed: true,
      },
    })
    chrome.runtime
      .sendMessage({
        type: 'DASHBOARD_FAILED',
        error: err && err.message ? err.message : String(err),
      })
      .catch(() => {})
  }
}

function updateState(state) {
  return chrome.storage.local.set({ dashboardScanState: state })
}
