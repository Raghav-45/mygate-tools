import type { Workbook } from 'exceljs'

/** ArrayBuffer/typed-array -> base64 (replicates the original helper). */
export function arrayBufferToBase64(buffer: ArrayBuffer | Uint8Array): string {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer)
  let binary = ''
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i])
  }
  return btoa(binary)
}

export function xlsxDataUrl(buf: ArrayBuffer | Uint8Array): string {
  return `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${arrayBufferToBase64(buf)}`
}

export interface DownloadOptions {
  url: string
  filename: string
  saveAs: boolean
}

export type DownloadFunction = (
  options: DownloadOptions,
  callback: (downloadId?: number) => void,
) => void

function defaultDownloader(options: DownloadOptions, callback: (downloadId?: number) => void) {
  chrome.downloads.download(options, callback)
}

/**
 * Serialize a workbook to xlsx and hand it to `chrome.downloads.download`
 * (`saveAs: false`, matching the original extensions). Resolves with the
 * download id; rejects when the browser blocks the download.
 */
export async function downloadWorkbook(
  workbook: Workbook,
  filename: string,
  downloader: DownloadFunction = defaultDownloader,
): Promise<number> {
  const buf = await workbook.xlsx.writeBuffer()
  const dataUrl = xlsxDataUrl(new Uint8Array(buf))

  return new Promise((resolve, reject) => {
    downloader({ url: dataUrl, filename, saveAs: false }, (downloadId) => {
      if (chrome.runtime.lastError || !downloadId) {
        reject(
          new Error(chrome.runtime.lastError?.message || 'Download blocked by browser permission'),
        )
      } else {
        resolve(downloadId)
      }
    })
  })
}
