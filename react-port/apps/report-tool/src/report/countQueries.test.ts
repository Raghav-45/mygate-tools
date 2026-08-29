import { describe, expect, it, vi } from 'vitest'
import {
  buildCountQueryPayload,
  fetchCategoryCount,
  STATUS_OPEN,
  STATUS_RESOLVED,
  STATUS_TOTAL,
} from './countQueries'

const FROM_EPOCH = 1704067200000
const TO_EPOCH = 1706659286399

function jsonResponse(body: unknown, status = 200, statusText = 'OK') {
  return Promise.resolve(
    new Response(JSON.stringify(body), { status, statusText }),
  ) as Promise<Response>
}

describe('buildCountQueryPayload', () => {
  it('builds the exact conditions the original report-tool sent', () => {
    const payload = buildCountQueryPayload(252434, FROM_EPOCH, TO_EPOCH, STATUS_TOTAL)

    expect(payload.operationName).toBe('getAdminSrList')
    expect(payload.query).toContain('getAdminSrList')
    expect(payload.query).toContain('totalCount')

    const rd = payload.variables.requestData
    expect(rd.requiredFields).toEqual(['id'])
    expect(rd.pagination).toEqual({ count: 1, page: 1 })
    expect(rd.sorting).toEqual([])
    expect(rd.conditions).toEqual([
      { field: 'date_filter', operation: 'equal', value: 'created_date' },
      { field: 'category', operation: 'in', value: ['252434'] },
      { field: 'from_date', operation: 'equal', value: String(FROM_EPOCH) },
      { field: 'to_date', operation: 'equal', value: String(TO_EPOCH) },
      { field: 'mygate_status', operation: 'in', value: STATUS_TOTAL },
    ])
  })

  it('uses the exact status arrays from the original', () => {
    expect(STATUS_TOTAL).toEqual(['open', 'hold', 're_opened', 'job_done', 'in_progress', 'closed'])
    expect(STATUS_OPEN).toEqual(['open', 'hold', 're_opened', 'job_done', 'in_progress'])
    expect(STATUS_RESOLVED).toEqual(['closed'])
  })
})

describe('fetchCategoryCount', () => {
  it('returns the numeric totalCount from a healthy response', async () => {
    const fetchMock = vi.fn(async (_url: string, init?: RequestInit) => {
      expect((init?.headers as Record<string, string>).authorization).toBe('tok')
      const sent = JSON.parse(init!.body as string) as { operationName: string }
      expect(sent.operationName).toBe('getAdminSrList')
      return jsonResponse({ data: { getAdminSrList: { dataResponse: { totalCount: 42 } } } })
    })

    const count = await fetchCategoryCount(
      252434,
      FROM_EPOCH,
      TO_EPOCH,
      STATUS_TOTAL,
      'tok',
      fetchMock as unknown as typeof fetch,
    )
    expect(count).toBe(42)
    expect(fetchMock).toHaveBeenCalledTimes(1)
  })

  it('omits the authorization header when the token is empty', async () => {
    const fetchMock = vi.fn(async (_url: string, init?: RequestInit) => {
      expect((init?.headers as Record<string, string>).authorization).toBeUndefined()
      return jsonResponse({ data: { getAdminSrList: { dataResponse: { totalCount: 0 } } } })
    })

    const count = await fetchCategoryCount(
      252434,
      FROM_EPOCH,
      TO_EPOCH,
      STATUS_TOTAL,
      '',
      fetchMock as unknown as typeof fetch,
    )
    expect(count).toBe(0)
  })

  it('maps 401/403 to the login prompt', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({}, 401, 'Unauthorized'))
    await expect(
      fetchCategoryCount(
        252434,
        FROM_EPOCH,
        TO_EPOCH,
        STATUS_TOTAL,
        null,
        fetchMock as unknown as typeof fetch,
      ),
    ).rejects.toThrow(
      'HTTP 401/403 Unauthorized: Please log into dashboard.mygate.com in Google Chrome.',
    )
  })

  it('maps other bad statuses to the generic HTTP error', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({}, 404, 'Not Found'))
    await expect(
      fetchCategoryCount(
        252434,
        FROM_EPOCH,
        TO_EPOCH,
        STATUS_TOTAL,
        null,
        fetchMock as unknown as typeof fetch,
      ),
    ).rejects.toThrow('HTTP Error 404: Not Found')
  })

  it('surfaces GraphQL-level errors', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({ errors: [{ message: 'nope' }] }))
    await expect(
      fetchCategoryCount(
        252434,
        FROM_EPOCH,
        TO_EPOCH,
        STATUS_TOTAL,
        null,
        fetchMock as unknown as typeof fetch,
      ),
    ).rejects.toThrow('MyGate API Error: nope')
  })

  it('raises the auth-verification message when totalCount is missing', async () => {
    const fetchMock = vi.fn(async () =>
      jsonResponse({ data: { getAdminSrList: { dataResponse: {} } } }),
    )
    await expect(
      fetchCategoryCount(
        252434,
        FROM_EPOCH,
        TO_EPOCH,
        STATUS_TOTAL,
        null,
        fetchMock as unknown as typeof fetch,
      ),
    ).rejects.toThrow(
      'Authentication Verification Failed: MyGate GraphQL returned empty data. Please ensure you are logged into dashboard.mygate.com in Google Chrome.',
    )
  })
})
