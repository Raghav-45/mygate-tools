import { fireEvent, render, screen, within } from '@testing-library/react'
import { describe, expect, it } from 'vitest'
import App from './App'
import { getChromeStub } from '../test/setup'
import type { SummaryRow } from '../summary/summaryState'

function todayMonth(): string {
  const now = new Date()
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`
}

function setMonth(value: string) {
  fireEvent.change(screen.getByLabelText('Report Month *'), { target: { value } })
}

function clickGenerate() {
  fireEvent.click(screen.getByRole('button', { name: 'Generate Summary Sheet' }))
}

function sentMessages(): unknown[] {
  return getChromeStub().messages.map((m) => m[0])
}

/** Scope lookups to an individual KPI card so table cells can't collide. */
function kpiCard(label: string): HTMLElement {
  return screen.getByText(label).closest('div') as HTMLElement
}

const ROW_A: SummaryRow = {
  date: '01-Jan-2024',
  prevOpen: 100,
  received: 5,
  closed: 3,
  pending: 102,
}
const ROW_B: SummaryRow = {
  date: '02-Jan-2024',
  prevOpen: 102,
  received: 8,
  closed: 6,
  pending: 104,
}

describe('App', () => {
  it('renders the popup shell with today as the default report month', () => {
    render(<App />)
    expect(document.querySelector('.app-wrapper')).toBeInTheDocument()
    expect(screen.getByText('Complaint Summary Sheet')).toBeInTheDocument()
    expect(screen.getByText('Monthly Automation Tool')).toBeInTheDocument()
    expect(screen.getByLabelText('Report Month *')).toHaveValue(todayMonth())
  })

  it('requires a report month before starting a scan', () => {
    render(<App />)
    setMonth('')
    clickGenerate()
    expect(screen.getByText('Please select a report month.')).toBeInTheDocument()
    expect(sentMessages().some((m) => (m as { type?: string }).type === 'START_SUMMARY_SCAN')).toBe(
      false,
    )
  })

  it('starts a scan with year, month and the default 1.5s delay', () => {
    render(<App />)
    setMonth('2024-02')
    clickGenerate()

    const start = sentMessages().find(
      (m) => (m as { type?: string }).type === 'START_SUMMARY_SCAN',
    ) as { year: number; month: number; requestDelay: number }
    expect(start).toMatchObject({ year: 2024, month: 2, requestDelay: 1.5 })
    expect(screen.getByRole('button', { name: /Stop Live Scan/ })).toBeInTheDocument()
  })

  it('abort sends ABORT_SUMMARY_SCAN and disables the button', () => {
    render(<App />)
    setMonth('2024-02')
    clickGenerate()
    fireEvent.click(screen.getByRole('button', { name: /Stop Live Scan/ }))
    expect(screen.getByRole('button', { name: /Stopping\.\.\./ })).toBeDisabled()
    expect(sentMessages().some((m) => (m as { type?: string }).type === 'ABORT_SUMMARY_SCAN')).toBe(
      true,
    )
  })

  it('persists the delay in seconds via the settings drawer', async () => {
    render(<App />)
    fireEvent.click(screen.getByRole('button', { name: /Report Speed/ }))
    fireEvent.change(screen.getByLabelText('Delay Between Cloud Queries'), {
      target: { value: '2' },
    })
    expect(screen.getByText('2.0s')).toBeInTheDocument()
    expect((await getChromeStub().storage.local.get(['requestDelay'])).requestDelay).toBe(2)
  })

  it('accumulates rows, sorts them, and updates the KPIs on SUMMARY_PROGRESS', async () => {
    render(<App />)
    setMonth('2024-01')
    clickGenerate()

    getChromeStub().broadcast({
      type: 'SUMMARY_PROGRESS',
      stepText: 'Completed 02-01-2024',
      pct: 7,
      row: ROW_B,
    })
    expect(await screen.findByText('02-Jan-2024')).toBeInTheDocument()

    getChromeStub().broadcast({
      type: 'SUMMARY_PROGRESS',
      stepText: 'Completed 01-01-2024',
      pct: 3,
      row: ROW_A,
    })
    // Flush the batched re-render before inspecting the table order.
    await screen.findByText('01-Jan-2024')

    // Delivered out of order; the table and Final Pending use the chronological last row.
    const table = screen.getByRole('table')
    const dateCells = within(table).getAllByText(/Jan-2024/)
    expect(dateCells[0].textContent).toBe('01-Jan-2024')
    expect(dateCells[1].textContent).toBe('02-Jan-2024')

    expect(within(kpiCard('Total Received')).getByText('13')).toBeInTheDocument()
    expect(within(kpiCard('Total Closed')).getByText('9')).toBeInTheDocument()
    expect(within(kpiCard('Final Pending')).getByText('104')).toBeInTheDocument()
  })

  it('shows the download success flow on SUMMARY_DONE', async () => {
    render(<App />)
    setMonth('2024-01')
    clickGenerate()
    getChromeStub().broadcast({
      type: 'SUMMARY_PROGRESS',
      stepText: 'Downloading Excel...',
      pct: 100,
    })
    expect(await screen.findByText('100%')).toBeInTheDocument()

    getChromeStub().broadcast({ type: 'SUMMARY_DONE' })
    expect(
      await screen.findByText('Report downloaded automatically to your Downloads folder!'),
    ).toBeInTheDocument()
    expect(screen.getByText('🎉 Excel File Downloaded Automatically!')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Generate Summary Sheet/ })).toBeInTheDocument()
  })

  it('shows the stopped alert on SUMMARY_ABORTED', async () => {
    render(<App />)
    setMonth('2024-01')
    clickGenerate()
    getChromeStub().broadcast({ type: 'SUMMARY_ABORTED' })
    expect(await screen.findByText('Live scan stopped by user.')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Generate Summary Sheet/ })).toBeInTheDocument()
  })

  it('surfaces scan errors through the alert banner', async () => {
    render(<App />)
    setMonth('2024-01')
    clickGenerate()
    getChromeStub().broadcast({ type: 'SUMMARY_ERROR', error: 'HTTP 500' })
    expect(await screen.findByText('Scan Failed: HTTP 500')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Generate Summary Sheet/ })).toBeInTheDocument()
  })

  it('replays an in-flight scan from storage on reopen', async () => {
    await getChromeStub().storage.local.set({
      requestDelay: 2.5,
      summaryScanState: {
        isScanning: true,
        year: 2024,
        month: 1,
        pct: 50,
        stepText: 'Completed 15-01-2024',
        rows: [ROW_A],
        isDone: false,
      },
    })
    render(<App />)
    expect(await screen.findByRole('button', { name: /Stop Live Scan/ })).toBeInTheDocument()
    expect(screen.getByText('Completed 15-01-2024')).toBeInTheDocument()
    expect(screen.getByText('01-Jan-2024')).toBeInTheDocument()
    expect(within(kpiCard('Final Pending')).getByText('102')).toBeInTheDocument()

    fireEvent.click(screen.getByRole('button', { name: /Report Speed/ }))
    expect(screen.getByText('2.5s')).toBeInTheDocument()
  })

  it('shows the banner and a finished scan survived a popup reopen', async () => {
    await getChromeStub().storage.local.set({
      summaryScanState: {
        isScanning: false,
        year: 2024,
        month: 1,
        pct: 100,
        stepText: 'Report Generated!',
        rows: [ROW_A],
        isDone: true,
      },
    })
    render(<App />)
    expect(await screen.findByText('🎉 Excel File Downloaded Automatically!')).toBeInTheDocument()
    expect(screen.getByText('01-Jan-2024')).toBeInTheDocument()
  })
})
