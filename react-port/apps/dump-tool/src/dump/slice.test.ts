import { describe, expect, it } from 'vitest'
import { sliceIntoYears } from './slice'

describe('sliceIntoYears', () => {
  it('splits a multi-year range into ≤365-day chunks and returns them newest-first', () => {
    const chunks = sliceIntoYears('2024-01-01', '2026-12-31')
    expect(chunks).toEqual([
      { fromDate: '2026-01-01', toDate: '2026-12-31' },
      { fromDate: '2025-01-01', toDate: '2025-12-31' },
      { fromDate: '2024-01-01', toDate: '2024-12-31' },
    ])
  })

  it('keeps a single full year as one chunk', () => {
    expect(sliceIntoYears('2024-01-01', '2024-12-31')).toEqual([
      { fromDate: '2024-01-01', toDate: '2024-12-31' },
    ])
  })

  it('clamps the tail chunk to the requested end date', () => {
    const chunks = sliceIntoYears('2024-06-01', '2025-03-15')
    expect(chunks).toEqual([{ fromDate: '2024-06-01', toDate: '2025-03-15' }])
  })

  it('spans an exact year boundary into two chunks', () => {
    const chunks = sliceIntoYears('2024-01-15', '2026-01-14')
    expect(chunks).toEqual([
      { fromDate: '2025-01-15', toDate: '2026-01-14' },
      { fromDate: '2024-01-15', toDate: '2025-01-14' },
    ])
  })

  it('slices across two chunk boundaries', () => {
    const chunks = sliceIntoYears('2023-06-01', '2025-06-30')
    expect(chunks).toEqual([
      { fromDate: '2025-06-01', toDate: '2025-06-30' },
      { fromDate: '2024-06-01', toDate: '2025-05-31' },
      { fromDate: '2023-06-01', toDate: '2024-05-31' },
    ])
  })

  it('returns an empty list when the end is before the start', () => {
    expect(sliceIntoYears('2024-06-01', '2024-05-31')).toEqual([])
  })
})
