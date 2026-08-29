/**
 * Date helpers shared across the three tools.
 *
 * Every helper accepts the `YYYY-MM-DD` string form that the popup date inputs
 * produce, and reproduces (character for character) the output formats that the
 * original extensions relied on for MyGate filter values, Excel cells and
 * filenames. See `NOTES.md` §1 for the per-tool usage.
 */

export const MONTH_NAMES = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
]

/** YYYY-MM-DD (or `null`) to local-midnight epoch seconds. */
export function getMidnightEpoch(dateStr: string | null | undefined): number {
  if (!dateStr) return Math.floor(Date.now() / 1000)
  const [year, month, day] = dateStr.split('-').map(Number)
  const date = new Date(year, month - 1, day, 0, 0, 0)
  return Math.floor(date.getTime() / 1000)
}

/** YYYY-MM-DD -> `d:m:yyyy` (no leading zeros) — MyGate `downloadFilters` format. */
export function formatFilterDate(dateStr: string): string {
  const [y, m, d] = dateStr.split('-')
  return `${Number(d)}:${Number(m)}:${Number(y)}`
}

/** YYYY-MM-DD -> `d-m-yyyy` (no leading zeros) — Excel title/subtitle format. */
export function formatHeaderDate(dateStr: string): string {
  const [y, m, d] = dateStr.split('-')
  return `${Number(d)}-${Number(m)}-${Number(y)}`
}

/** YYYY-MM-DD -> `d-m-yyyy` (no leading zeros) — Excel filename date format. */
export function formatFilenameDate(dateStr: string): string {
  const [y, m, d] = dateStr.split('-')
  return `${Number(d)}-${Number(m)}-${Number(y)}`
}

/** Date -> `DD-MM-YYYY` (zero-padded). */
export function formatDDMMYYYY(date: Date): string {
  const dd = String(date.getDate()).padStart(2, '0')
  const mm = String(date.getMonth() + 1).padStart(2, '0')
  return `${dd}-${mm}-${date.getFullYear()}`
}

/** Date -> `DD-MMM-YYYY` (e.g. `09-Jan-2024`). */
export function formatDDMMMYYYY(date: Date): string {
  const dd = String(date.getDate()).padStart(2, '0')
  return `${dd}-${MONTH_NAMES[date.getMonth()]}-${date.getFullYear()}`
}

/** Parse `DD-MMM-YYYY` to a local-midnight timestamp (used for sorting). */
export function parseDateToTime(dStr: string): number {
  const [dd, mmm, yyyy] = dStr.split('-')
  const mIdx = MONTH_NAMES.indexOf(mmm)
  return new Date(Number(yyyy), mIdx, Number(dd)).getTime()
}

/** Parse `DD-MMM-YYYY` to a UTC-noon Date (used for Excel date cells). */
export function parseDateToUTCNoon(dStr: string): Date {
  const [dd, mmm, yyyy] = dStr.split('-')
  const mIdx = MONTH_NAMES.indexOf(mmm)
  return new Date(Date.UTC(Number(yyyy), mIdx, Number(dd), 12, 0, 0))
}
