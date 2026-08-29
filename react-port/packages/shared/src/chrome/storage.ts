/** Thin typed wrappers around the `chrome.*` storage API so tests can stub the global. */

export async function storageGet<K extends string>(keys: K[]): Promise<Record<K, unknown>> {
  return chrome.storage.local.get(keys) as Promise<Record<K, unknown>>
}

export async function storageSet(items: Record<string, unknown>): Promise<void> {
  return chrome.storage.local.set(items)
}
