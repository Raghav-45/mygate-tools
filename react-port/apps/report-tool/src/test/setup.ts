import '@testing-library/jest-dom/vitest'
import { cleanup } from '@testing-library/react'
import { afterEach, beforeEach } from 'vitest'

// Deterministic dates for slice/parse tests (matches Node/worker UTC behavior).
process.env.TZ = 'UTC'

/** Shared in-memory chrome stub used by every popup/worker test. */
export function installChromeStub() {
  const messages: unknown[][] = []
  const listeners: Array<
    (msg: unknown, sender: unknown, sendResponse: (r?: unknown) => void) => void
  > = []
  const responses = new Map<string, unknown>()

  const setResponse = (type: string, value: unknown) => {
    responses.set(type, value)
  }

  const sendMessage = (message: unknown): Promise<unknown> => {
    messages.push([message])
    const type = (message as { type?: string }).type
    return Promise.resolve(type ? responses.get(type) : undefined)
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
      removeListener: (fn: unknown) => {
        const idx = listeners.indexOf(fn as (m: unknown, s: unknown, r: () => void) => void)
        if (idx >= 0) listeners.splice(idx, 1)
      },
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

  return { runtime, storage, downloads, messages, broadcast, sendMessage, setResponse }
}

let currentStub: ReturnType<typeof installChromeStub> | null = null

export function getChromeStub() {
  if (!currentStub) throw new Error('chrome stub not installed (beforeEach not run?)')
  return currentStub
}

beforeEach(() => {
  currentStub = installChromeStub()
})

afterEach(() => {
  cleanup()
  delete (globalThis as Record<string, unknown>).chrome
})
