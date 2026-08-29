import { Workbook } from 'exceljs'
import type { Alignment } from 'exceljs'
import {
  cellBorder,
  MONTH_NAMES,
  parseDateToTime,
  parseDateToUTCNoon,
  whiteFill,
} from '@mygate/shared'
import type { SummaryRow } from './summaryState'

export function summaryFilename(year: number, month: number): string {
  return `Complaint Summary Sheet - ${MONTH_NAMES[month - 1]}-${year}.xlsx`
}

const HEADERS = [
  'Date',
  'Previous day Open Complaints',
  'Today Received Complaints',
  'Today Closed Complaints',
  'Pending',
]

const COL_WIDTHS = [27.14, 14.14, 13.86, 11.57, 12.0]

/**
 * The original summary-tool uses vertical `center` (Excel accepts it; the
 * exceljs type only allows `middle`), so the alignment is cast.
 */
function center(extra: Partial<Alignment> = {}): Partial<Alignment> {
  return { horizontal: 'center', vertical: 'center', ...extra } as Partial<Alignment>
}

/**
 * Builds the "Complaint Summary Sheet" workbook with the summary-tool's own
 * Calibri/white style (deliberately different from the dump/report blue design):
 *  - sheet "Sheet1", gridlines shown, widths 27.14/14.14/13.86/11.57/12.00
 *  - A1:E1 merged bold title (row height 18)
 *  - row 2 headers (height 45, wrap text)
 *  - data rows: UTC-noon Date cell with `dd-mm-yyyy` numFmt; the Pending cell is
 *    a formula `(B+C)-D` with its resolved value
 */
export function buildSummaryWorkbook(rows: SummaryRow[]): Workbook {
  const wb = new Workbook()
  const ws = wb.addWorksheet('Sheet1')
  ws.views = [{ showGridLines: true }]

  ws.columns = COL_WIDTHS.map((width) => ({ width }))

  ws.mergeCells('A1:E1')
  const r1 = ws.getRow(1)
  r1.height = 18
  const tCell = r1.getCell(1)
  tCell.value = 'Complaint Summary Sheet'
  tCell.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FF000000' } }
  tCell.alignment = center()
  for (let c = 1; c <= 5; c++) {
    const cell = r1.getCell(c)
    cell.border = cellBorder
    cell.fill = whiteFill
  }

  const r2 = ws.getRow(2)
  r2.values = HEADERS
  r2.height = 45
  for (let c = 1; c <= 5; c++) {
    const cell = r2.getCell(c)
    cell.font = { name: 'Calibri', size: 11, color: { argb: 'FF000000' } }
    cell.fill = whiteFill
    cell.alignment = center({ wrapText: true })
    cell.border = cellBorder
  }

  const sorted = [...rows].sort((a, b) => parseDateToTime(a.date) - parseDateToTime(b.date))
  sorted.forEach((r, idx) => {
    const rowNum = idx + 3
    const row = ws.getRow(rowNum)

    row.getCell(1).value = parseDateToUTCNoon(r.date)
    row.getCell(2).value = r.prevOpen
    row.getCell(3).value = r.received
    row.getCell(4).value = r.closed
    row.getCell(5).value = { formula: `(B${rowNum}+C${rowNum})-D${rowNum}`, result: r.pending }

    for (let c = 1; c <= 5; c++) {
      const cell = row.getCell(c)
      cell.font = { name: 'Calibri', size: 11, color: { argb: 'FF000000' } }
      cell.fill = whiteFill
      cell.alignment = center()
      cell.border = cellBorder
      if (c === 1) cell.numFmt = 'dd-mm-yyyy'
    }
  })

  return wb
}
