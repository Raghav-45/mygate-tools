import { Workbook } from 'exceljs'
import type { CellValue } from 'exceljs'
import { describe, expect, it } from 'vitest'
import { DEFAULT_COL_MAP, parseChunkXlsx } from './parseChunkXlsx'
import type { ChunkParseContext } from './parseChunkXlsx'

type HeaderSpec = [number, string][] // [exceljs column, header text]

/** Build a chunk worksheet with a fixed header row (row 3) and data rows (4+). */
function buildChunkWorkbook(headers: HeaderSpec, dataRows: unknown[][]): Workbook {
  const wb = new Workbook()
  const ws = wb.addWorksheet('chunk')
  headers.forEach(([col, val]) => {
    ws.getRow(3).getCell(col).value = val
  })
  dataRows.forEach((row, r) => {
    const rowNum = 4 + r
    row.forEach((val, c) => {
      if (val === undefined) return
      ws.getRow(rowNum).getCell(c + 1).value = val as CellValue
    })
  })
  return wb
}

function ctx(): ChunkParseContext {
  return { headerRowValues: null }
}

const DEFAULT_HEADERS: HeaderSpec = [
  [1, 'I.D'],
  [2, 'Created Date'],
  [4, 'Category'],
  [5, 'Sub Category'],
  [7, 'Flat'],
  [9, 'Subject'],
  [10, 'Status'],
]

describe('parseChunkXlsx', () => {
  it('uses the default column map when there is no header row', () => {
    // No headers: the fallback mapping (1,2,4,5,7,9,10) must pick these out.
    const wb = buildChunkWorkbook(
      [],
      [
        [
          1001,
          new Date(2024, 0, 15, 0, 0, 0),
          undefined,
          'Accounts',
          'Billing',
          undefined,
          'A-101',
          undefined,
          'Subject text here',
          'closed',
        ],
      ],
    )
    const { colMap, rows, rowsInChunk } = parseChunkXlsx(wb, ctx())
    expect(colMap).toEqual(DEFAULT_COL_MAP)
    expect(rowsInChunk).toBe(1)
    expect(rows[0].id).toBe(1001)
    expect((rows[0].createdDate as Date).getTime()).toBe(
      new Date(Date.UTC(2024, 0, 15, 12, 0, 0)).getTime(),
    )
    expect(rows[0].subject).toBe('Subject text here')
    expect(rows[0].status).toBe('closed')
  })

  it('repairs the column map from the header row (row 3) when values differ', () => {
    const wb = buildChunkWorkbook(
      [
        [1, 'Ticket ID'],
        [2, 'Date'],
        [3, 'xyz'],
        [4, 'Subcategory'],
        [5, 'House'],
        [6, 'Description'],
        [7, 'MyGate Status'],
      ],
      [
        [
          'ABC123',
          new Date(2024, 2, 3, 0, 0, 0),
          'XYZ-Col-3',
          'SubCatVal',
          'B-22',
          'Long subject',
          'open',
        ],
      ],
    )
    const { colMap, rows } = parseChunkXlsx(wb, ctx())
    // No 'Category' header so the default index (4) is retained (faithful to
    // the original loop — category and subcategory both resolve to column 4).
    expect(colMap).toEqual({
      id: 1,
      createdDate: 2,
      category: 4,
      subCategory: 4,
      flat: 5,
      subject: 6,
      status: 7,
    })
    expect(rows[0].id).toBe('ABC123')
    expect(rows[0].createdDate).toBeInstanceOf(Date)
    expect(rows[0].category).toBe('SubCatVal')
    expect(rows[0].flat).toBe('B-22')
    expect(rows[0].subject).toBe('Long subject')
    expect(rows[0].status).toBe('open')
  })

  it('skips rows with an empty id and keeps partial rows', () => {
    const wb = buildChunkWorkbook(DEFAULT_HEADERS, [
      [
        1001,
        new Date(2024, 0, 1),
        undefined,
        'A',
        'B',
        undefined,
        'C',
        undefined,
        'Subject',
        'open',
      ],
      ['', new Date(2024, 0, 2), undefined, 'A', 'B', undefined, 'C', undefined, 'Subject', 'open'],
    ])
    const { rows, rowsInChunk } = parseChunkXlsx(wb, ctx())
    expect(rowsInChunk).toBe(1)
    expect(rows[0].id).toBe(1001)
  })

  it('captures the header from the first parseable chunk and reuses it', () => {
    const context = ctx()
    const wb = buildChunkWorkbook(DEFAULT_HEADERS, [
      [
        1001,
        new Date(2024, 0, 15),
        undefined,
        'A',
        'B',
        undefined,
        'C',
        undefined,
        'Subject',
        'closed',
      ],
    ])
    parseChunkXlsx(wb, context)
    expect(context.headerRowValues).not.toBeNull()

    const wb2 = buildChunkWorkbook([[1, 'I.D']], [[1002]])
    const { rows } = parseChunkXlsx(wb2, context)
    // The second chunk's own (shorter) header must not override the captured one.
    expect(rows[0].id).toBe(1002)
  })

  it('handles an empty worksheet without crashing', () => {
    const wb = new Workbook()
    wb.addWorksheet('empty')
    const { rows, rowsInChunk } = parseChunkXlsx(wb, ctx())
    expect(rowsInChunk).toBe(0)
    expect(rows).toEqual([])
  })
})
