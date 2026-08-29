import '@testing-library/jest-dom/vitest'
import { afterEach, beforeEach } from 'vitest'

/** Shared in-memory chrome stub used by every popup/worker test. */
export function installChromeStub() {
  const messages: unknown[][] = []
  const listeners: Array<
    (msg: unknown, sender: unknown, sendResponse: (r?: unknown) => void) => void
  > = []

  const sendMessage = (message: unknown): Promise<unknown> => {
    messages.push([message])
    return Promise.resolve(undefined)
  }

  const broadcast = (message: unknown): void => {
    for (const listener of [...listeners]) {
      listener(message, {}, () => {})
    }
  }

  const runtime = {
    sendMessage,
    onMessage: {
      addListener: (fn: (m: unknown, s: unknown, r: () => void) => void) => listeners.push(fn),
    },
    lastError: undefined as undefined | { message: string },
  }

  const storage = {
    data: new Map<string, unknown>(),
    local: {
      get: async (keys: string[]) => {
        const out: Record<string, unknown> = {}
        for (const k of keys) {
          out[k] = storage.data.get(k)
        }
        return out
      },
      set: async (items: Record<string, unknown>) => {
        for (const [k, v] of Object.entries(items)) storage.data.set(k, v)
      },
    },
  }

  const downloads = {
    download: (_opts: unknown, cb?: (id?: number) => void) => {
      cb?.(123)
    },
  }

  ;(globalThis as Record<string, unknown>).chrome = {
    runtime,
    storage,
    downloads,
    tabs: { query: async () => [] },
    scripting: { executeScript: async () => [{ result: null }] },
    cookies: { getAll: async () => [] },
  }

  return { runtime, storage, downloads, messages, broadcast, sendMessage }
}

beforeEach(() => {
  installChromeStub()
})

afterEach(() => {
  delete (globalThis as Record<string, unknown>).chrome
})
