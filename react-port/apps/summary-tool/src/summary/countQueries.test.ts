import { describe, expect, it, vi } from 'vitest'
import {
  buildSummaryCountRequest,
  epochDDMMYYYY,
  fetchDayCount,
  STATUS_ALL,
  STATUS_CLOSED,
  STATUS_PREV_OPEN,
} from './countQueries'

function jsonResponse(body: unknown, status = 200, statusText = 'OK') {
  return Promise.resolve(
    new Response(JSON.stringify(body), { status, statusText }),
  ) as Promise<Response>
}

describe('epochDDMMYYYY', () => {
  it('converts dd-mm-yyyy to local-midnight epoch seconds (same as the original epoch())', () => {
    const expected = Math.floor(new Date(2024, 0, 1).getTime() / 1000)
    expect(epochDDMMYYYY('01-01-2024')).toBe(expected)
  })
})

describe('buildSummaryCountRequest', () => {
  it('builds the exact name/values conditions the original summary-tool sent', () => {
    const req = buildSummaryCountRequest('01-01-2024', '15-01-2024', STATUS_ALL)

    expect(req.operationName).toBe('getAdminSrList')
    expect(req.query).toContain('getAdminSrList')
    expect(req.query).toContain('totalCount')

    const rd = req.variables.requestData
    expect(rd.requiredFields).toEqual(['id'])
    expect(rd.pagination).toEqual({ count: 1, page: 1 })
    expect(rd.sorting).toEqual([])
    expect(rd.conditions).toEqual([
      { name: 'date_filter', operation: 'equal', values: ['created_date'] },
      { name: 'from_date', operation: 'gte', values: [epochDDMMYYYY('01-01-2024')] },
      { name: 'to_date', operation: 'lte', values: [epochDDMMYYYY('15-01-2024') + 86399] },
      { name: 'mygate_status', operation: 'in', values: STATUS_ALL },
    ])
  })

  it('uses the exact status arrays from the original', () => {
    expect(STATUS_PREV_OPEN).toEqual(['open', 'hold', 're_opened', 'in_progress', 'job_done'])
    expect(STATUS_ALL).toEqual(['open', 'hold', 're_opened', 'in_progress', 'job_done', 'closed'])
    expect(STATUS_CLOSED).toEqual(['closed'])
  })
})

describe('fetchDayCount', () => {
  it('returns the numeric totalCount from a healthy response, always sending the token', async () => {
    const fetchMock = vi.fn(async (_url: string, init?: RequestInit) => {
      expect((init?.headers as Record<string, string>).authorization).toBe('tok')
      expect(init?.credentials).toBe('same-origin')
      const sent = JSON.parse(init!.body as string) as { operationName: string }
      expect(sent.operationName).toBe('getAdminSrList')
      return jsonResponse({ data: { getAdminSrList: { dataResponse: { totalCount: 42 } } } })
    })

    const count = await fetchDayCount(
      '01-01-2024',
      '02-01-2024',
      STATUS_CLOSED,
      'tok',
      fetchMock as unknown as typeof fetch,
    )
    expect(count).toBe(42)
    expect(fetchMock).toHaveBeenCalledTimes(1)
  })

  it('returns 0 when totalCount is missing', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({ data: { getAdminSrList: {} } }))
    const count = await fetchDayCount(
      '01-01-2024',
      '02-01-2024',
      STATUS_ALL,
      'tok',
      fetchMock as unknown as typeof fetch,
    )
    expect(count).toBe(0)
  })

  it('maps non-2xx status to `HTTP <status>` (summary keeps HTTP status, no re-login text)', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({}, 500, 'Internal Server Error'))
    await expect(
      fetchDayCount(
        '01-01-2024',
        '02-01-2024',
        STATUS_ALL,
        'tok',
        fetchMock as unknown as typeof fetch,
      ),
    ).rejects.toThrow('HTTP 500')
  })

  it('maps GraphQL errors to the session-expired message', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({ errors: [{ message: 'oops' }] }))
    await expect(
      fetchDayCount(
        '01-01-2024',
        '02-01-2024',
        STATUS_ALL,
        'tok',
        fetchMock as unknown as typeof fetch,
      ),
    ).rejects.toThrow('Session expired! Please refresh dashboard.mygate.com and log in.')
  })

  it('maps a missing getAdminSrList to the session-expired message', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({ data: {} }))
    await expect(
      fetchDayCount(
        '01-01-2024',
        '02-01-2024',
        STATUS_ALL,
        'tok',
        fetchMock as unknown as typeof fetch,
      ),
    ).rejects.toThrow('Session expired! Please refresh dashboard.mygate.com and log in.')
  })
})
