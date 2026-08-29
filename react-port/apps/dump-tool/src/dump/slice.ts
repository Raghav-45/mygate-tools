export interface YearChunk {
  fromDate: string
  toDate: string
}

/**
 * Slices a multi-year date range into ≤365-day chunks (1 year minus 1 day each).
 * Returns them newest-first, exactly like the original `sliceIntoYears`.
 */
export function sliceIntoYears(fromDateStr: string, toDateStr: string): YearChunk[] {
  const chunks: YearChunk[] = []
  const currStart = new Date(fromDateStr)
  const finalEnd = new Date(toDateStr)

  let start = currStart
  while (start <= finalEnd) {
    const currEnd = new Date(start)
    currEnd.setFullYear(currEnd.getFullYear() + 1)
    currEnd.setDate(currEnd.getDate() - 1) // 1 year minus 1 day

    if (currEnd > finalEnd) {
      currEnd.setTime(finalEnd.getTime())
    }

    const sStr = start.toISOString().split('T')[0]
    const eStr = currEnd.toISOString().split('T')[0]
    chunks.push({ fromDate: sStr, toDate: eStr })

    start = new Date(currEnd)
    start.setDate(start.getDate() + 1)
  }
  return chunks.reverse()
}
