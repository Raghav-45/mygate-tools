import { describe, expect, it } from 'vitest'
import { buildPendingTicketsWorkbook, reportFilename } from './reportWorkbook'

const RESULTS = [
  { name: 'Accounts Billing', id: 252434, total: 10, resolved: 4, open: 6 },
  { name: 'IT WIFI Network', id: 277816, total: 5, resolved: 5, open: 0 },
]

const SUMMARY = { total: 15, resolved: 9, open: 6 }

describe('reportFilename', () => {
  it('uses the unpadded d-m-yyyy filename format', () => {
    expect(reportFilename('2024-01-31')).toBe('Pending_Mygate_Tickets_Report_31-1-2024.xlsx')
    expect(reportFilename('2024-12-05')).toBe('Pending_Mygate_Tickets_Report_5-12-2024.xlsx')
  })
})

describe('buildPendingTicketsWorkbook', () => {
  it('lays out the sheet exactly like the original report-tool', () => {
    const wb = buildPendingTicketsWorkbook(RESULTS, SUMMARY, '2024-01-01', '2024-01-31')
    const ws = wb.worksheets[0]

    expect(ws.name).toBe('Pending Tickets')
    expect(ws.model.merges).toEqual(['A1:D1'])

    const title = ws.getCell('A1')
    expect(title.value).toBe('Pending Mygate Tickets - From 1-1-2024 To 31-1-2024')
    expect((title.fill as { fgColor: { argb: string } }).fgColor.argb).toBe('FF4D93D9')
    expect(title.alignment).toMatchObject({ horizontal: 'center', vertical: 'middle' })

    // Header row.
    expect(ws.getCell('A2').value).toBe('Category')
    expect(ws.getCell('B2').value).toBe('Total')
    expect(ws.getCell('C2').value).toBe('Resolved')
    expect(ws.getCell('D2').value).toBe('Open')
    expect((ws.getCell('A2').fill as { fgColor: { argb: string } }).fgColor.argb).toBe('FF4D93D9')

    // Data rows.
    expect(ws.getCell('A3').value).toBe('Accounts Billing')
    expect(ws.getCell('B3').value).toBe(10)
    expect(ws.getCell('C3').value).toBe(4)
    expect(ws.getCell('D3').value).toBe(6)
    expect(ws.getCell('A4').value).toBe('IT WIFI Network')
    expect(ws.getCell('D4').value).toBe(0)

    // TOTALS row (no bold/yellow stip).
    expect(ws.getCell('A5').value).toBe('TOTALS')
    expect(ws.getCell('B5').value).toBe(15)
    expect(ws.getCell('C5').value).toBe(9)
    expect(ws.getCell('D5').value).toBe(6)
    expect(ws.getRow(5).font?.bold).toBeFalsy()

    // Column widths 55 / 9 / 12 / 8.
    expect(ws.getColumn(1).width).toBe(55)
    expect(ws.getColumn(2).width).toBe(9)
    expect(ws.getColumn(3).width).toBe(12)
    expect(ws.getColumn(4).width).toBe(8)

    // Borders on every cell in A1:D<last>.
    for (let r = 1; r <= 5; r++) {
      for (let c = 1; c <= 4; c++) {
        expect(ws.getCell(r, c).border?.left).toBeDefined()
        expect(ws.getCell(r, c).border?.right).toBeDefined()
      }
    }
  })

  it('number columns are right aligned', () => {
    const wb = buildPendingTicketsWorkbook(RESULTS, SUMMARY, '2024-01-01', '2024-01-31')
    const ws = wb.worksheets[0]
    expect(ws.getCell('B3').alignment).toMatchObject({ horizontal: 'right' })
    expect(ws.getCell('A3').alignment).toMatchObject({ horizontal: 'left' })
  })
})
