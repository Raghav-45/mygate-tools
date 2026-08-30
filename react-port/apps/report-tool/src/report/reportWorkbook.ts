import { Workbook } from 'exceljs'
import {
  aptosTitleFont,
  blueFill,
  cellBorder,
  formatFilenameDate,
  formatHeaderDate,
} from '@mygate/shared'
import type { CategoryData, ReportSummary } from './reportState'

export function reportFilename(toDate: string): string {
  return `Pending_Mygate_Tickets_Report_${formatFilenameDate(toDate)}.xlsx`
}

/**
 * Builds the "Pending Tickets" report workbook. Layout matches the original
 * report-tool exactly:
 *  - sheet "Pending Tickets"
 *  - A1:D1 merged blue title `Pending Mygate Tickets - From <d-m-yyyy> To <d-m-yyyy>`
 *  - row 2 blue headers Category/Total/Resolved/Open
 *  - data rows from row 3, then a yellow bold TOTALS row
 *  - thin borders applied to every cell A1..D<last> in a final loop
 *  - column widths 55 / 9 / 12 / 8
 */
export function buildPendingTicketsWorkbook(
  results: CategoryData[],
  summary: ReportSummary,
  fromDate: string,
  toDate: string,
): Workbook {
  const workbook = new Workbook()
  const sheet = workbook.addWorksheet('Pending Tickets')

  // Column widths 55 / 9 / 12 / 8 (original layout).
  sheet.getColumn(1).width = 55
  sheet.getColumn(2).width = 9
  sheet.getColumn(3).width = 12
  sheet.getColumn(4).width = 8

  const titleCell = sheet.getCell('A1')
  titleCell.value = `Pending Mygate Tickets - From ${formatHeaderDate(fromDate)} To ${formatHeaderDate(toDate)}`
  titleCell.fill = blueFill
  titleCell.font = aptosTitleFont
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' }
  sheet.mergeCells('A1:D1')
  sheet.getRow(1).height = 20

  const headerNames = ['Category', 'Total', 'Resolved', 'Open']
  const headerRow = sheet.getRow(2)
  headerRow.height = 20
  headerNames.forEach((name, idx) => {
    const cell = headerRow.getCell(idx + 1)
    cell.value = name
    cell.fill = blueFill
    cell.font = aptosTitleFont
    cell.border = cellBorder
    cell.alignment = { vertical: 'middle', horizontal: 'left' }
  })

  for (const data of results) {
    sheet.addRow([data.name, data.total, data.resolved, data.open])
  }

  sheet.addRow(['TOTALS', summary.total, summary.resolved, summary.open])

  const lastRow = sheet.rowCount
  for (let r = 3; r <= lastRow; r++) {
    const row = sheet.getRow(r)
    row.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' }
    row.getCell(2).alignment = { vertical: 'middle', horizontal: 'right' }
    row.getCell(3).alignment = { vertical: 'middle', horizontal: 'right' }
    row.getCell(4).alignment = { vertical: 'middle', horizontal: 'right' }
  }

  for (let r = 1; r <= lastRow; r++) {
    for (let c = 1; c <= 4; c++) {
      sheet.getCell(r, c).border = cellBorder
    }
  }

  return workbook
}
