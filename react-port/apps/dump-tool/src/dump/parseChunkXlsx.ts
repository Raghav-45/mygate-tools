import type { Workbook } from 'exceljs'

export interface MergedRow {
  id: unknown
  createdDate: unknown
  category: unknown
  subCategory: unknown
  flat: unknown
  subject: unknown
  status: unknown
}

export interface ColMap {
  id: number
  createdDate: number
  category: number
  subCategory: number
  flat: number
  subject: number
  status: number
}

export const DEFAULT_COL_MAP: ColMap = {
  id: 1,
  createdDate: 2,
  category: 4,
  subCategory: 5,
  flat: 7,
  subject: 9,
  status: 10,
}

/**
 * Holds the header row captured from the first parseable chunk. The original
 * worker records it exactly once (`if (!headerRowValues && rowCount >= 3)`) and
 * reuses it for every subsequent chunk — reproduced here as a mutable context.
 */
export interface ChunkParseContext {
  headerRowValues: unknown[] | null
}

export interface ChunkParseResult {
  rows: MergedRow[]
  rowsInChunk: number
  colMap: ColMap
}

/** Parse one chunk worksheet (header on row 3, data from row 4) into MergedRows. */
export function parseChunkXlsx(wb: Workbook, ctx: ChunkParseContext): ChunkParseResult {
  const ws = wb.worksheets[0]

  if (!ctx.headerRowValues && ws.rowCount >= 3) {
    ctx.headerRowValues = (ws.getRow(3).values as unknown[]) ?? null
  }

  const colMap: ColMap = { ...DEFAULT_COL_MAP }
  if (ctx.headerRowValues) {
    ctx.headerRowValues.forEach((val, idx) => {
      if (!val) return
      const s = String(val).toLowerCase().trim()
      if (s === 'id' || s === 'ticket id' || s === 'i.d') colMap.id = idx
      else if (s === 'created date' || s === 'date') colMap.createdDate = idx
      else if (s === 'category') colMap.category = idx
      else if (s === 'sub category' || s === 'subcategory') colMap.subCategory = idx
      else if (s === 'flat' || s === 'house') colMap.flat = idx
      else if (s === 'subject' || s === 'description') colMap.subject = idx
      else if (s === 'status' || s === 'mygate status') colMap.status = idx
    })
  }

  const rows: MergedRow[] = []
  let rowsInChunk = 0
  for (let r = 4; r <= ws.rowCount; r++) {
    const rVals = ws.getRow(r).values
    if (!rVals || !Array.isArray(rVals)) continue
    const cells = rVals as unknown[]
    if (cells.length <= 1) continue
    const idV = cells[colMap.id]
    if (idV === undefined || idV === '') continue

    const getVal = (idx: number) => {
      const v = cells[idx]
      if (v && v instanceof Date) {
        return new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), 12, 0, 0))
      }
      return v !== undefined && v !== null ? v : ''
    }
    rows.push({
      id: getVal(colMap.id),
      createdDate: getVal(colMap.createdDate),
      category: getVal(colMap.category),
      subCategory: getVal(colMap.subCategory),
      flat: getVal(colMap.flat),
      subject: getVal(colMap.subject),
      status: getVal(colMap.status),
    })
    rowsInChunk++
  }

  return { rows, rowsInChunk, colMap }
}
