import { GraphqlHttpError, postGraphQL } from '@mygate/shared'

export const STATUS_PREV_OPEN = ['open', 'hold', 're_opened', 'in_progress', 'job_done']
export const STATUS_ALL = [...STATUS_PREV_OPEN, 'closed']
export const STATUS_CLOSED = ['closed']

/** dd-mm-yyyy -> local-midnight epoch *seconds* (matches the original `epoch()`). */
export function epochDDMMYYYY(dateStr: string): number {
  const [dd, mm, yyyy] = dateStr.split('-').map(Number)
  return Math.floor(new Date(yyyy, mm - 1, dd, 0, 0, 0).getTime() / 1000)
}

export interface SummaryCountPayload {
  operationName: string
  variables: {
    requestData: {
      requiredFields: string[]
      pagination: { count: number; page: number }
      sorting: unknown[]
      conditions: Array<{ name: string; values: Array<string | number>; operation: string }>
    }
  }
  query: string
}

/**
 * Count-request builder for the summary tool. The original sends a different
 * condition shape than report-tool: keys `name`/`values` (arrays), operations
 * `equal`/`gte`/`lte`/`in`, and epoch *seconds*. Kept local to the summary app.
 */
export function buildSummaryCountRequest(
  fromDate: string,
  toDate: string,
  statuses: string[],
): SummaryCountPayload {
  return {
    operationName: 'getAdminSrList',
    variables: {
      requestData: {
        requiredFields: ['id'],
        pagination: { count: 1, page: 1 },
        sorting: [],
        conditions: [
          { name: 'date_filter', operation: 'equal', values: ['created_date'] },
          { name: 'from_date', values: [epochDDMMYYYY(fromDate)], operation: 'gte' },
          { name: 'to_date', values: [epochDDMMYYYY(toDate) + 86399], operation: 'lte' },
          { name: 'mygate_status', values: statuses, operation: 'in' },
        ],
      },
    },
    query: `query getAdminSrList($requestData: DataListInput) {
      getAdminSrList(requestData: $requestData) {
        dataResponse { totalCount }
      }
    }`,
  }
}

interface SummaryCountBody {
  errors?: unknown
  data?: { getAdminSrList?: { dataResponse?: { totalCount?: unknown } } }
}

/**
 * Run one count query and translate failures exactly like the original
 * `getCount`: any non-2xx status -> `HTTP <status>`; GraphQL errors or a missing
 * `getAdminSrList` -> the session-expired message; otherwise the count, `0`
 * when `totalCount` is absent. The `authorization` header is always attached
 * (the token is guaranteed non-null by the time the scan starts), and no cookies
 * are sent (original default fetch credentials).
 */
export async function fetchDayCount(
  fromDate: string,
  toDate: string,
  statuses: string[],
  token: string,
  fetchImpl: typeof fetch = fetch,
): Promise<number> {
  let body: SummaryCountBody
  try {
    body = await postGraphQL<SummaryCountBody>(
      buildSummaryCountRequest(fromDate, toDate, statuses),
      { authorization: token, credentials: 'same-origin' },
      fetchImpl,
    )
  } catch (err) {
    if (err instanceof GraphqlHttpError) {
      throw new Error(`HTTP ${err.status}`)
    }
    throw err
  }

  if (body.errors || (body.data && !body.data.getAdminSrList)) {
    throw new Error('Session expired! Please refresh dashboard.mygate.com and log in.')
  }

  const totalCount = body?.data?.getAdminSrList?.dataResponse?.totalCount
  return typeof totalCount === 'number' ? totalCount : 0
}
