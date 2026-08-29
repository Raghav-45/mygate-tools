import { DASHBOARD_HOME, DASHBOARD_ORIGIN, GRAPHQL_URL } from '../config'

/** GraphQL transport error raised when the HTTP response is not ok. */
export class GraphqlHttpError extends Error {
  readonly status: number
  readonly statusText: string

  constructor(status: number, statusText: string) {
    super(`HTTP ${status}${statusText ? `: ${statusText}` : ''}`)
    this.name = 'GraphqlHttpError'
    this.status = status
    this.statusText = statusText
  }
}

export interface GraphQLRequest {
  operationName: string
  variables: Record<string, unknown>
  query: string
}

export interface GraphQLRequestOptions {
  /** Sent as the `authorization` header when non-undefined. */
  authorization?: string
  /** Passed through to fetch; default `"include"` (the originals all send cookies). */
  credentials?: RequestCredentials
}

/** Request headers shared by all three tools. */
export function buildBaseHeaders(opts: GraphQLRequestOptions = {}): Record<string, string> {
  const headers: Record<string, string> = {
    'content-type': 'application/json',
    origin: DASHBOARD_ORIGIN,
    referer: DASHBOARD_HOME,
  }
  if (opts.authorization !== undefined) {
    headers.authorization = opts.authorization
  }
  return headers
}

/**
 * POST a GraphQL payload to the MyGate dashboard API and return the parsed JSON
 * body. Raises `GraphqlHttpError` for non-2xx responses; GraphQL-level errors are
 * NOT interpreted here — callers inspect `data.errors` / `data.data` themselves,
 * exactly like the original `background.js` files did.
 *
 * `fetchImpl` is injectable for tests.
 */
export async function postGraphQL<T = unknown>(
  payload: GraphQLRequest,
  opts: GraphQLRequestOptions = {},
  fetchImpl: typeof fetch = fetch,
): Promise<T> {
  const res = await fetchImpl(GRAPHQL_URL, {
    method: 'POST',
    headers: buildBaseHeaders(opts),
    credentials: opts.credentials ?? 'include',
    body: JSON.stringify(payload),
  })

  if (!res.ok) {
    throw new GraphqlHttpError(res.status, res.statusText)
  }

  return (await res.json()) as T
}
