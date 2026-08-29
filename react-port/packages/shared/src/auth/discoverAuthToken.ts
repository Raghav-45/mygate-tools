/**
 * Auth token discovery.
 *
 * All three original tools discover the dashboard's auth token the same way at a
 * high level (read values out of an open dashboard tab's `localStorage`, then —
 * for report/summary — browser cookies), but every detail differs: which
 * localStorage keys are considered, the minimum length, which characters reject a
 * candidate, whether JSON objects are unwrapped, and the cookie matcher.
 *
 * To avoid a shared abstraction that silently changes behavior, each tool's exact
 * scanning rules are passed in as plain data below and executed *inside the
 * injected tab script* (a bare function cannot close over worker scope).
 */

/** Rules used by the localStorage scan injected into the dashboard tab. */
export interface LocalStorageTokenSpec {
  /** Keys (case-insensitive substring) to consider. `undefined`/empty = any key. */
  keyContains?: string[]
  /** Minimum value length for the fast-path return. */
  minLength: number
  /** Substrings that disqualify a candidate (e.g. `{`, space). */
  reject: string[]
  /** Prefixes that disqualify the fast path (original code checks `startsWith`). */
  rejectPrefix: string[]
  /** When true, a value beginning with `{` is JSON-parsed and checked for known fields. */
  unwrap: boolean
  /** JSON fields to read when unwrapping (checked in order). */
  unwrapFields: string[]
}

/** Rules used by the cookie scan (runs in the worker, so closures are fine). */
export interface CookieTokenSpec {
  /** Exact cookie names (after lowercasing). */
  names?: string[]
  /** Cookie-name substrings (after lowercasing). */
  contains?: string[]
  /** Minimum cookie value length. */
  minLength: number
}

export interface TokenDiscoveryOptions {
  /** chrome.tabs.query URL pattern for candidate tabs. */
  tabQueryUrl: string
  localStorage: LocalStorageTokenSpec
  /** Omit to skip the cookie scan entirely (dump-tool never scans cookies). */
  cookie?: CookieTokenSpec
  /** Optional logger (report-tool logs its discovery to the console). */
  log?: (msg: string) => void
}

/**
 * Executed inside the dashboard tab. `spec` is plain serializable data.
 */
function scanLocalStorage(spec: LocalStorageTokenSpec): string | null {
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i)
    if (!key) continue

    if (spec.keyContains && spec.keyContains.length > 0) {
      const lowered = key.toLowerCase()
      if (!spec.keyContains.some((k) => lowered.includes(k))) continue
    }

    const val = localStorage.getItem(key)
    if (typeof val !== 'string' || val.length === 0) continue

    const prefixRejected = spec.rejectPrefix.some((p) => val.startsWith(p))

    if (val.length > spec.minLength && !prefixRejected) {
      const rejectable = spec.reject.some((r) => val.includes(r))
      if (!rejectable) return val
    }

    if (spec.unwrap && val.startsWith('{')) {
      try {
        const obj = JSON.parse(val)
        for (const field of spec.unwrapFields) {
          const found = obj?.[field]
          if (found) return found
        }
      } catch {
        // malformed JSON -> try the next localStorage entry
      }
    }
  }
  return null
}

/**
 * Returns a discovered token, or `null` when none is found. The caller decides
 * what to do with `null` (dump/report fall back to their sample token; summary
 * throws).
 */
export async function discoverAuthToken(opts: TokenDiscoveryOptions): Promise<string | null> {
  // 1. Inspect open dashboard tabs for localStorage tokens.
  try {
    const tabs = await chrome.tabs.query({ url: opts.tabQueryUrl })
    for (const tab of tabs) {
      if (tab.id == null) continue
      try {
        const results = await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: scanLocalStorage,
          args: [opts.localStorage],
        })
        if (results && results[0] && results[0].result) {
          opts.log?.('Discovered auth token from active tab')
          return results[0].result
        }
      } catch {
        // Contents of this tab were not inspectable; try the next one.
      }
    }
  } catch (e) {
    console.warn('Tab token discovery failed:', e)
  }

  // 2. Inspect browser cookies (report + summary only).
  if (opts.cookie) {
    try {
      const cookies = await chrome.cookies.getAll({ domain: 'mygate.com' })
      const names = (opts.cookie.names ?? []).map((n) => n.toLowerCase())
      const contains = (opts.cookie.contains ?? []).map((c) => c.toLowerCase())
      for (const c of cookies) {
        const name = c.name.toLowerCase()
        const matches = names.includes(name) || contains.some((part) => name.includes(part))
        if (!matches) continue
        if (c.value && c.value.length > opts.cookie.minLength) {
          opts.log?.('Discovered auth token from browser cookies')
          return c.value
        }
      }
    } catch (e) {
      console.warn('Cookie token discovery failed:', e)
    }
  }

  return null
}
