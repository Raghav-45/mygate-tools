/** Thin typed wrappers around the `chrome.runtime` messaging API. */

export type RuntimeMessageListener = (
  message: unknown,
  sender: chrome.runtime.MessageSender,
  sendResponse: (response?: unknown) => void,
) => void | boolean

export function addRuntimeMessageListener(listener: RuntimeMessageListener): void {
  chrome.runtime.onMessage.addListener(listener)
}

/** Returns a promise that resolves to the (possibly undefined) response. */
export function sendRuntimeMessage<T = unknown>(message: unknown): Promise<T> {
  return chrome.runtime.sendMessage(message)
}
