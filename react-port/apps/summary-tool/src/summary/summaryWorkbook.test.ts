import { describe, expect, it } from 'vitest'
import { parseDateToUTCNoon } from '@mygate/shared'
import type { SummaryRow } from './summaryState'
import { buildSummaryWorkbook, summaryFilename } from './summaryWorkbook'
import { STATUS_ALL } from './countQueries'
import { BASE_PREV_DATE } from './runSummaryScan'

const ROWS: SummaryRow[] = [
  { date: '15-Jan-2024', prevOpen: 10, received: 5, closed: 3, pending: 12 },
  { date: '01-Jan-2024', prevOpen: 8, received: 5, closed: 4, pending: 9 },
  { date: '31-Jan-2024', prevOpen: 12, received: 5, closed: 3, pending: 14 },
]

describe('summaryFilename', () => {
  it('uses the original `Complaint Summary Sheet - <MONTH>-<year>.xlsx` shape', () => {
    expect(summaryFilename(2024, 1)).toBe('Complaint Summary Sheet - Jan-2024.xlsx')
    expect(summaryFilename(2024, 2)).toBe('Complaint Summary Sheet - Feb-2024.xlsx')
  })
})

describe('buildSummaryWorkbook', () => {
  const wb = buildSummaryWorkbook(ROWS)
  const ws = wb.getWorksheet('Sheet1')!

  it('uses `Sheet1` with gridlines shown', () => {
    expect(ws).toBeDefined()
    expect(ws.views[0].showGridLines).toBe(true)
  })

  it('applies the summary-tool column widths 27.14/14.14/13.86/11.57/12.00', () => {
    expect(ws.getColumn(1).width).toBe(27.14)
    expect(ws.getColumn(2).width).toBe(14.14)
    expect(ws.getColumn(3).width).toBe(13.86)
    expect(ws.getColumn(4).width).toBe(11.57)
    expect(ws.getColumn(5).width).toBe(12.0)
  })

  it('merges A1:E1 into the bold Calibri title and sets row height 18', () => {
    expect(ws.getCell('A1').value).toBe('Complaint Summary Sheet')
    expect(ws.getCell('A1').font).toMatchObject({
      name: 'Calibri',
      size: 14,
      bold: true,
      color: { argb: 'FF000000' },
    })
    expect(ws.getCell('A1').alignment.horizontal).toBe('center')
    expect(ws.model.merges.map(String)).toEqual(['A1:E1'])
    expect(ws.getRow(1).height).toBe(18)
  })

  it('writes the wrapped white header row with the exact column texts', () => {
    const row = ws.getRow(2)
    expect(row.height).toBe(45)
    expect(row.getCell(1).value).toBe('Date')
    expect(row.getCell(2).value).toBe('Previous day Open Complaints')
    expect(row.getCell(3).value).toBe('Today Received Complaints')
    expect(row.getCell(4).value).toBe('Today Closed Complaints')
    expect(row.getCell(5).value).toBe('Pending')
    for (let c = 1; c <= 5; c++) {
      const cell = row.getCell(c)
      expect(cell.font).toMatchObject({ name: 'Calibri', size: 11, color: { argb: 'FF000000' } })
      expect(cell.alignment.wrapText).toBe(true)
      expect(cell.alignment.horizontal).toBe('center')
    }
  })

  it('sorts rows by date and writes a formula Pending cell with its resolved value', () => {
    const r1 = ws.getRow(3)
    expect(r1.getCell(1).value).toEqual(parseDateToUTCNoon('01-Jan-2024'))
    expect(r1.getCell(1).numFmt).toBe('dd-mm-yyyy')
    expect(r1.getCell(2).value).toBe(8)
    expect(r1.getCell(3).value).toBe(5)
    expect(r1.getCell(4).value).toBe(4)
    expect(r1.getCell(5).value).toEqual({ formula: '(B3+C3)-D3', result: 9 })

    const r2 = ws.getRow(4)
    expect(r2.getCell(5).value).toEqual({ formula: '(B4+C4)-D4', result: 12 })

    const r3 = ws.getRow(5)
    expect(r3.getCell(5).value).toEqual({ formula: '(B5+C5)-D5', result: 14 })
  })
})

describe('summary scan status constants', () => {
  it('exposes the base prev date constant used by every prevOpen daily count', () => {
    expect(BASE_PREV_DATE).toBe('01-01-2024')
    expect(STATUS_ALL).toHaveLength(6)
  })
})
