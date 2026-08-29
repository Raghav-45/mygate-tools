import { render, screen, fireEvent } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { describe, expect, it } from 'vitest'
import App from './App'
import { getChromeStub } from '../test/setup'
import type { ChunkResult, StartDumpExport } from '../dump/dumpState'

async function seededStorage(items: Record<string, unknown>) {
  await chrome.storage.local.set(items)
}

function today() {
  return new Date().toISOString().split('T')[0]
}

describe('App', () => {
  it('renders default dates and the generate CTA', async () => {
    render(<App />)
    const fromInput = (await screen.findByLabelText('Dump From *')) as HTMLInputElement
    const toInput = (await screen.findByLabelText('Dump Till *')) as HTMLInputElement
    expect(fromInput.value).toBe('2024-01-01')
    expect(toInput.value).toBe(today())
    expect(screen.getByRole('button', { name: /Generate Master Dump/i })).toBeInTheDocument()
  })

  it('starts an export and sends the message with the ms delay', async () => {
    const user = userEvent.setup()
    render(<App />)
    await screen.findByLabelText('Dump From *')

    await user.click(screen.getByRole('button', { name: /Generate Master Dump/i }))

    const { messages } = getChromeStub()
    const start = messages[0]?.[0] as StartDumpExport
    expect(start.type).toBe('START_DUMP_EXPORT')
    expect(start.params.fromDate).toBe('2024-01-01')
    expect(start.params.requestDelayMs).toBe(2000)

    // CTA swaps to the abort button while scanning.
    expect(screen.getByRole('button', { name: /Stop Live Export/i })).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: /Generate Master Dump/i })).not.toBeInTheDocument()
  })

  it('persists the delay as milliseconds when the slider moves', async () => {
    const user = userEvent.setup()
    render(<App />)
    await screen.findByLabelText('Dump From *')

    await user.click(screen.getByRole('button', { name: /Polling Speed/i }))
    const slider = screen.getByLabelText('Delay Between Cloud Queries') as HTMLInputElement
    fireEvent.change(slider, { target: { value: '3' } })

    const stored = (await chrome.storage.local.get(['requestDelay'])) as { requestDelay: number }
    expect(stored.requestDelay).toBeCloseTo(3000)
    expect(screen.getByText('3.0s')).toBeInTheDocument()
  })

  it('replays an active scan from storage and finishes via DUMP_FINISHED', async () => {
    const chunk: ChunkResult = {
      rangeStr: '2025-01-01 to 2025-12-31',
      filterFrom: '1:1:2025',
      filterTo: '31:12:2025',
      rowsFound: '4',
      status: 'Merged ✅',
      downloadUrl: 'x',
    }
    await seededStorage({
      dumpScanState: {
        isScanning: true,
        pct: 12,
        statusText: 'Generating report for 2024-01-01 to 2024-12-31...',
        chunks: [chunk],
        totalRows: 4,
        isDone: false,
      },
    })

    render(<App />)
    await screen.findByText('Generating report for 2024-01-01 to 2024-12-31...')
    expect(screen.getByText('12%')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Stop Live Export/i })).toBeInTheDocument()
    expect(screen.getByText('Active')).toBeInTheDocument()

    const { broadcast } = getChromeStub()
    broadcast({ type: 'DUMP_FINISHED', chunks: [chunk], totalRows: 4 })

    expect(
      await screen.findByText('Master Dump exported and downloaded automatically!'),
    ).toBeInTheDocument()
    expect(screen.getByText('Completed')).toBeInTheDocument()
    expect(screen.getByText(/Master Dump Excel Downloaded Automatically/i)).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Generate Master Dump/i })).toBeInTheDocument()
  })

  it('applies the stored-ms quirk to the delay slider label', async () => {
    // The original stores milliseconds; reloading feeds the ms value straight
    // into the seconds slider and its label (kept as-is — see NOTES.md).
    const user = userEvent.setup()
    await seededStorage({ requestDelay: 2000 })
    render(<App />)
    await user.click(await screen.findByRole('button', { name: /Polling Speed/i }))
    expect(await screen.findByText('2000.0s')).toBeInTheDocument()
  })

  it('shows an error when the from date is after the till date', async () => {
    const user = userEvent.setup()
    render(<App />)
    const fromInput = await screen.findByLabelText('Dump From *')
    const toInput = screen.getByLabelText('Dump Till *')

    await user.clear(fromInput)
    await user.type(fromInput, '2025-06-01')
    await user.clear(toInput)
    await user.type(toInput, '2024-06-01')

    await user.click(screen.getByRole('button', { name: /Generate Master Dump/i }))

    expect(
      await screen.findByText('Dump From date cannot be after Dump Till date.'),
    ).toBeInTheDocument()
  })

  it('opens and closes the about modal', async () => {
    const user = userEvent.setup()
    render(<App />)
    await screen.findByLabelText('Dump From *')

    await user.click(screen.getByRole('button', { name: 'About & Credits' }))
    expect(await screen.findByText('Aditya Singh Khichi')).toBeInTheDocument()

    await user.click(screen.getByRole('button', { name: 'Close' }))
    expect(screen.queryByText('Aditya Singh Khichi')).not.toBeInTheDocument()
  })
})
