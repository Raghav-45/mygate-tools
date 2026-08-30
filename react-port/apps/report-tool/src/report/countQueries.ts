import { GraphqlHttpError, postGraphQL, type GraphQLRequest } from '@mygate/shared'

export const STATUS_TOTAL = ['open', 'hold', 're_opened', 'job_done', 'in_progress', 'closed']
export const STATUS_OPEN = STATUS_TOTAL.slice(0, 5)
export const STATUS_RESOLVED = ['closed']

// Matches the original report-tool `COUNT_QUERY` (plus the hard-coded
// Authorization header the original dropped only when the token was empty).
export const COUNT_QUERY = `
  query getAdminSrList($requestData: DataListInput) {
    getAdminSrList(requestData: $requestData) {
      dataResponse {
        totalCount
      }
    }
  }
`.trim()

export interface CountQueryPayload {
  operationName: string
  variables: {
    requestData: {
      requiredFields: string[]
      pagination: { count: number; page: number }
      sorting: unknown[]
      conditions: Array<{ name: string; operation: string; values: unknown[] }>
    }
  }
  query: string
}

export function buildCountQueryPayload(
  categoryId: number,
  fromEpoch: number,
  toEpoch: number,
  statuses: string[],
): CountQueryPayload {
  return {
    operationName: 'getAdminSrList',
    variables: {
      requestData: {
        requiredFields: ['id'],
        pagination: { count: 1, page: 1 },
        sorting: [],
        conditions: [
          { name: 'date_filter', operation: 'equal', values: ['created_date'] },
          { name: 'category', operation: 'equal', values: [categoryId] },
          { name: 'from_date', operation: 'gte', values: [fromEpoch] },
          { name: 'to_date', operation: 'lte', values: [toEpoch] },
          {
            name: 'mygate_status',
            operation: statuses.length === 1 ? 'equal' : 'in',
            values: statuses,
          },
        ],
      },
    },
    query: COUNT_QUERY,
  }
}

export function buildCategoryCountRequest(
  categoryId: number,
  fromEpoch: number,
  toEpoch: number,
  statuses: string[],
): GraphQLRequest {
  const payload = buildCountQueryPayload(categoryId, fromEpoch, toEpoch, statuses)
  return {
    operationName: payload.operationName,
    variables: payload.variables,
    query: payload.query,
  }
}

interface CountResponseShape {
  errors?: Array<{ message?: string }>
  data?: { getAdminSrList?: { dataResponse?: { totalCount?: unknown } } }
}

/**
 * Run one `getAdminSrList` count query and return the numeric totalCount.
 * Mirrors the original report-tool translation of failures:
 *  - 401/403 => auth prompt message
 *  - other bad status => `HTTP Error ${status}: ${statusText}`
 *  - GraphQL errors => `MyGate API Error: ${first message}`
 *  - missing totalCount => auth-verification message
 */
export async function fetchCategoryCount(
  categoryId: number,
  fromEpoch: number,
  toEpoch: number,
  statuses: string[],
  token: string | null,
  fetchImpl: typeof fetch = fetch,
): Promise<number> {
  const request = buildCategoryCountRequest(categoryId, fromEpoch, toEpoch, statuses)
  const opts = token && token.trim() !== '' ? { authorization: token } : {}

  let body: CountResponseShape
  try {
    body = await postGraphQL<CountResponseShape>(request, opts, fetchImpl)
  } catch (err) {
    if (err instanceof GraphqlHttpError) {
      if (err.status === 401 || err.status === 403) {
        throw new Error(
          'HTTP 401/403 Unauthorized: Please log into dashboard.mygate.com in Google Chrome.',
        )
      }
      throw new Error(`HTTP Error ${err.status}: ${err.statusText}`)
    }
    throw err
  }

  if (body?.errors && body.errors.length > 0) {
    throw new Error(`MyGate API Error: ${body.errors[0]?.message ?? 'Unknown'}`)
  }

  const totalCount = body?.data?.getAdminSrList?.dataResponse?.totalCount
  if (typeof totalCount !== 'number') {
    throw new Error(
      'Authentication Verification Failed: MyGate GraphQL returned empty data. Please ensure you are logged into dashboard.mygate.com in Google Chrome.',
    )
  }
  return totalCount
}
