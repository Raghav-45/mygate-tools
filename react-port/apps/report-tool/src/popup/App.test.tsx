import { fireEvent, render, screen } from '@testing-library/react'
import { describe, expect, it, beforeEach } from 'vitest'
import App from './App'
import { getChromeStub } from '../test/setup'
import type { CategoryData, ReportSummary } from '../report/reportState'

const CATEGORIES = [
  ['Accounts Billing', 252434],
  ['Construction Or Project Related', 277747],
] as Array<[string, number]>

function seedCategories() {
  getChromeStub().setResponse('GET_CATEGORIES_LIST', { categories: CATEGORIES })
}

async function renderWithCategories() {
  seedCategories()
  render(<App />)
  // Wait until the categories round-trip resolves and pills render.
  await screen.findByRole('button', { name: /Accounts Billing/ })
}

function clickPill(name: string) {
  fireEvent.click(screen.getByRole('button', { name: new RegExp(name) }))
}

/** The stub records each sendMessage as `[message]`; flatten to raw messages. */
function sentMessages(): unknown[] {
  return getChromeStub().messages.map((m) => m[0])
}

describe('App', () => {
  beforeEach(() => {
    seedCategories()
  })

  it('renders the popup shell with default dates and header', () => {
    render(<App />)
    expect(document.querySelector('.app-wrapper')).toBeInTheDocument()
    expect(screen.getByText('Pending Tickets Report')).toBeInTheDocument()
    expect(screen.getByLabelText('Report From *')).toHaveValue('2024-01-01')
    expect(screen.getByLabelText('Report Till *')).toHaveValue(
      new Date().toISOString().split('T')[0],
    )
  })

  it('loads categories and selects all pills by default', async () => {
    await renderWithCategories()
    expect(screen.getByRole('button', { name: /Accounts Billing/ })).toHaveAttribute(
      'aria-pressed',
      'true',
    )
    expect(screen.getByRole('button', { name: /Construction Or Project Related/ })).toHaveAttribute(
      'aria-pressed',
      'true',
    )
  })

  it('toggles pills and persists selectedCatIds', async () => {
    await renderWithCategories()
    clickPill('Accounts Billing')
    expect(screen.getByRole('button', { name: /Accounts Billing/ })).toHaveAttribute(
      'aria-pressed',
      'false',
    )
    expect((await getChromeStub().storage.local.get(['selectedCatIds'])).selectedCatIds).toEqual([
      277747,
    ])
  })

  it('Toggle All deselects everything and back again', async () => {
    await renderWithCategories()
    fireEvent.click(screen.getByRole('button', { name: 'Toggle All' }))
    expect(screen.getByRole('button', { name: /Accounts Billing/ })).toHaveAttribute(
      'aria-pressed',
      'false',
    )
    expect((await getChromeStub().storage.local.get(['selectedCatIds'])).selectedCatIds).toEqual([])
    fireEvent.click(screen.getByRole('button', { name: 'Toggle All' }))
    expect(screen.getByRole('button', { name: /Accounts Billing/ })).toHaveAttribute(
      'aria-pressed',
      'true',
    )
  })

  it('requires dates in order and at least one category', async () => {
    await renderWithCategories()
    fireEvent.click(screen.getByRole('button', { name: 'Toggle All' }))
    fireEvent.click(screen.getByRole('button', { name: 'Generate Report' }))
    expect(screen.getByText('Please select at least one Ticket Category.')).toBeInTheDocument()

    fireEvent.change(screen.getByLabelText('Report From *'), {
      target: { value: '2099-01-01' },
    })
    expect(screen.getByLabelText('Report From *')).toHaveValue('2099-01-01')
    clickPill('Accounts Billing')
    fireEvent.click(screen.getByRole('button', { name: 'Generate Report' }))
    expect(screen.getByText('Report From date cannot be after Till date.')).toBeInTheDocument()
  })

  it('starts a scan with the selected categories and delays', async () => {
    await renderWithCategories()
    fireEvent.click(screen.getByRole('button', { name: 'Generate Report' }))

    const start = sentMessages().find(
      (m) => (m as { type?: string }).type === 'START_REPORT_SCAN',
    ) as unknown as {
      params: { requestDelayMs: number; selectedCategoryIds: number[]; fromDate: string }
    }
    expect(start.params.requestDelayMs).toBe(1000)
    expect(start.params.selectedCategoryIds).toEqual([252434, 277747])
    expect(screen.getByRole('button', { name: /Stop Live Scan/ })).toBeInTheDocument()
  })

  it('abort sends ABORT_REPORT_SCAN and disables the button', async () => {
    await renderWithCategories()
    fireEvent.click(screen.getByRole('button', { name: 'Generate Report' }))
    fireEvent.click(screen.getByRole('button', { name: /Stop Live Scan/ }))
    expect(screen.getByRole('button', { name: /Stopping\.\.\./ })).toBeDisabled()
    expect(sentMessages().some((m) => (m as { type?: string }).type === 'ABORT_REPORT_SCAN')).toBe(
      true,
    )
  })

  it('persists the delay in seconds via the settings drawer', async () => {
    await renderWithCategories()
    fireEvent.click(screen.getByRole('button', { name: /Report Speed/ }))
    fireEvent.change(screen.getByLabelText('Delay Between Cloud Queries'), {
      target: { value: '0.6' },
    })
    expect(screen.getByText('0.6s')).toBeInTheDocument()
    expect((await getChromeStub().storage.local.get(['requestDelay'])).requestDelay).toBe(0.6)
  })

  it('renders NaN% while the Downloading Excel progress update is shown (ported quirk)', async () => {
    await renderWithCategories()
    fireEvent.click(screen.getByRole('button', { name: 'Generate Report' }))
    getChromeStub().broadcast({ type: 'SCAN_PROGRESS_UPDATE', statusText: 'Downloading Excel...' })
    expect(await screen.findByText('NaN%')).toBeInTheDocument()
  })

  it('accumulates category results and shows the report on SCAN_FINISHED', async () => {
    await renderWithCategories()
    fireEvent.click(screen.getByRole('button', { name: 'Generate Report' }))

    const data: CategoryData = {
      name: 'IT WIFI Network',
      id: 277816,
      total: 10,
      resolved: 4,
      open: 6,
    }
    const summary: ReportSummary = { total: 10, resolved: 4, open: 6 }

    getChromeStub().broadcast({ type: 'CATEGORY_COMPLETED', data, summary })
    expect(await screen.findByText('IT WIFI Network')).toBeInTheDocument()
    expect(screen.getAllByText('4').length).toBeGreaterThan(0)

    getChromeStub().broadcast({
      type: 'SCAN_FINISHED',
      results: [data],
      summary,
      reportMeta: { fromDate: '2024-01-01', toDate: '2026-08-29' },
    })

    expect(
      await screen.findByText('Report downloaded automatically to your Downloads folder!'),
    ).toBeInTheDocument()
    expect(screen.getByText('🎉 Excel File Downloaded Automatically!')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Generate Report/ })).toBeInTheDocument()
  })

  it('surfaces category errors through the alert banner', async () => {
    await renderWithCategories()
    fireEvent.click(screen.getByRole('button', { name: 'Generate Report' }))
    getChromeStub().broadcast({
      type: 'SCAN_ERROR',
      category: 'IT WIFI Network',
      errorMessage: 'HTTP Error 500: Internal Server Error',
    })
    expect(await screen.findByText('HTTP Error 500: Internal Server Error')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Generate Report/ })).toBeInTheDocument()
  })

  it('shows Stopping alert after SCAN_ABORTED', async () => {
    await renderWithCategories()
    fireEvent.click(screen.getByRole('button', { name: 'Generate Report' }))
    getChromeStub().broadcast({ type: 'SCAN_ABORTED', results: [] })
    expect(await screen.findByText('Scan stopped by user.')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Generate Report/ })).toBeInTheDocument()
  })

  it('replays an in-flight scan from storage on reopen', async () => {
    await getChromeStub().storage.local.set({
      ticketsScanState: {
        isScanning: true,
        pct: 50,
        statusText: 'Completed Accounts Billing',
        results: [{ name: 'Accounts Billing', id: 252434, total: 5, resolved: 2, open: 3 }],
        summary: { total: 5, resolved: 2, open: 3 },
        isDone: false,
      },
    })
    render(<App />)
    expect(await screen.findByRole('button', { name: /Stop Live Scan/ })).toBeInTheDocument()
    expect(screen.getByText('Completed Accounts Billing')).toBeInTheDocument()
    expect(screen.getAllByText('3').length).toBeGreaterThan(0)
  })

  it('shows the auto-download banner and results when a finished scan is reopened', async () => {
    await getChromeStub().storage.local.set({
      ticketsScanState: {
        isScanning: false,
        pct: 100,
        statusText: 'Report Generated!',
        results: [{ name: 'Accounts Billing', id: 252434, total: 5, resolved: 2, open: 3 }],
        summary: { total: 5, resolved: 2, open: 3 },
        isDone: true,
      },
    })
    render(<App />)
    expect(await screen.findByText('🎉 Excel File Downloaded Automatically!')).toBeInTheDocument()
    expect(screen.getByText('Accounts Billing')).toBeInTheDocument()
  })
})
