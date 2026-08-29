import { describe, expect, it } from 'vitest'
import { buildMasterWorkbook } from './masterWorkbook'
import type { MergedRow } from './parseChunkXlsx'

const rows: MergedRow[] = [
  {
    id: 1001,
    createdDate: new Date(Date.UTC(2024, 0, 15, 12, 0, 0)),
    category: 'Accounts',
    subCategory: 'Billing',
    flat: 'A-101',
    subject: 'Water leak',
    status: 'closed',
  },
  {
    id: 1002,
    createdDate: '',
    category: 'IT',
    subCategory: '',
    flat: 'B-22',
    subject: 'Wifi issue in flat',
    status: 'open',
  },
]

describe('buildMasterWorkbook', () => {
  it('lays out the standardized master sheet', () => {
    const wb = buildMasterWorkbook(rows, '2024-01-01', '2025-12-31')
    const ws = wb.worksheets[0]

    expect(ws.getCell('A1').value).toBe('DLF Independent Floors')
    expect(ws.getCell('A2').value).toBe('Help Desk Report: From 1-1-2024 To 31-12-2025')

    const headers = [
      'Sr No.',
      'I.D',
      'Created Date',
      'Category',
      'Sub Category',
      'Flat',
      'Subject',
      'Status',
    ]
    headers.forEach((h, i) => {
      expect(ws.getRow(3).getCell(i + 1).value).toBe(h)
    })

    // Row 4 = first data row; Sr No. is 1-indexed.
    expect(ws.getRow(4).getCell(1).value).toBe(1)
    expect(ws.getRow(4).getCell(2).value).toBe(1001)
    expect(ws.getRow(5).getCell(1).value).toBe(2)
    expect(ws.getRow(5).getCell(5).value).toBe('')

    expect(ws.getRow(1).height).toBe(24)
    expect(ws.getRow(2).height).toBe(20)
    expect(ws.getRow(3).height).toBe(24)

    const widths = [10, 12, 22, 28, 20, 16, 45, 15]
    widths.forEach((w, i) => {
      expect(ws.getColumn(i + 1).width).toBe(w)
    })
  })

  it('merges the title rows across A:H and fills them', () => {
    const wb = buildMasterWorkbook(rows, '2024-01-01', '2025-12-31')
    const ws = wb.worksheets[0]
    expect((ws.model.merges as string[]).sort()).toEqual(['A1:H1', 'A2:H2'])

    const a1 = ws.getCell('A1')
    expect(a1.isMerged).toBe(true)
    expect(ws.getCell('H1').isMerged).toBe(true)
    expect((a1.fill as { fgColor: { argb: string } }).fgColor.argb).toBe('FF4D93D9')
    expect((a1.font as { name: string }).name).toBe('Aptos')
  })
})
