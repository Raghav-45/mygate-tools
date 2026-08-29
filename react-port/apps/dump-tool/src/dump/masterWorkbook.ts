import { Workbook } from 'exceljs'
import type { CellValue } from 'exceljs'
import { aptosTitleFont, blueFill, cellBorder, formatHeaderDate } from '@mygate/shared'
import type { MergedRow } from './parseChunkXlsx'

const COL_HEADERS = [
  'Sr No.',
  'I.D',
  'Created Date',
  'Category',
  'Sub Category',
  'Flat',
  'Subject',
  'Status',
]

const COL_WIDTHS = [10, 12, 22, 28, 20, 16, 45, 15]

export function buildMasterWorkbook(rows: MergedRow[], fromDate: string, toDate: string): Workbook {
  const masterWb = new Workbook()
  const masterWs = masterWb.addWorksheet('Master Dump')

  // Row 1: Title
  masterWs.mergeCells('A1:H1')
  for (let c = 1; c <= 8; c++) {
    const cell = masterWs.getCell(1, c)
    cell.fill = blueFill
    cell.border = cellBorder
  }
  const titleCell = masterWs.getCell('A1')
  titleCell.value = 'DLF Independent Floors'
  titleCell.font = aptosTitleFont
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
  masterWs.getRow(1).height = 24

  // Row 2: Subtitle
  masterWs.mergeCells('A2:H2')
  for (let c = 1; c <= 8; c++) {
    const cell = masterWs.getCell(2, c)
    cell.fill = blueFill
    cell.border = cellBorder
  }
  const subCell = masterWs.getCell('A2')
  subCell.value = `Help Desk Report: From ${formatHeaderDate(fromDate)} To ${formatHeaderDate(toDate)}`
  subCell.font = aptosTitleFont
  subCell.alignment = { horizontal: 'center', vertical: 'middle' }
  masterWs.getRow(2).height = 20

  // Row 3: Headers
  const hRow = masterWs.getRow(3)
  hRow.height = 24
  COL_HEADERS.forEach((h, idx) => {
    const cell = hRow.getCell(idx + 1)
    cell.value = h
    cell.fill = blueFill
    cell.font = aptosTitleFont
    cell.border = cellBorder
    cell.alignment = { horizontal: 'center', vertical: 'middle' }
  })

  // Rows 4+: Data
  let currRowIdx = 4
  rows.forEach((item, idx) => {
    const row = masterWs.getRow(currRowIdx)
    const vals = [
      idx + 1, // Sr No. (1-indexed all the way)
      item.id,
      item.createdDate,
      item.category,
      item.subCategory,
      item.flat,
      item.subject,
      item.status,
    ]
    vals.forEach((val, cIdx) => {
      const cell = row.getCell(cIdx + 1)
      cell.value = val as CellValue
      cell.font = aptosTitleFont
      cell.border = cellBorder
      // Center Sr No., I.D, Created Date, Sub Category, Flat, Status; left align Category and Subject
      const isLeft = cIdx === 3 || cIdx === 6
      cell.alignment = {
        horizontal: isLeft ? 'left' : 'center',
        vertical: 'middle',
        wrapText: cIdx === 6,
      }
    })
    currRowIdx++
  })

  COL_WIDTHS.forEach((w, i) => {
    masterWs.getColumn(i + 1).width = w
  })

  return masterWb
}
