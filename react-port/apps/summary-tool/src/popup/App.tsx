import { useCallback, useEffect, useState } from 'react'
import {
  AboutModal,
  AbortButton,
  AlertBanner,
  AutoDownloadBanner,
  BrandHeader,
  ExternalLink,
  GearIcon,
  KpiCard,
  KpiGrid,
  parseDateToTime,
  PrimaryButton,
  ProgressCard,
  SettingsDrawer,
  StopIcon,
  TextLinkButton,
  ZapIcon,
} from '@mygate/shared'
import type { SummaryRow, SummaryScanState, SummaryWorkerMessage } from '../summary/summaryState'

const MIN_DELAY = 0.5
const MAX_DELAY = 3.0
const STEP_DELAY = 0.5
const DEFAULT_DELAY = 1.5

function todayMonth(): string {
  const now = new Date()
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`
}

function MonthField({
  id,
  label,
  value,
  onChange,
}: {
  id: string
  label: string
  value: string
  onChange: (value: string) => void
}) {
  return (
    <div className="flex flex-col gap-1.5">
      <label htmlFor={id} className="text-[11px] font-semibold text-ink-dim">
        {label}
      </label>
      <input
        id={id}
        type="month"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="w-full h-[42px] px-3 text-[13px] font-medium text-ink-main bg-[#F8FAFC] border border-line-input rounded-md outline-none focus:border-teal focus:ring-2 focus:ring-teal/20 transition-colors"
      />
    </div>
  )
}

interface Alert {
  message: string
  tone: 'error' | 'success'
}

const EMPTY_ROWS: SummaryRow[] = []

export default function App() {
  const [reportMonth, setReportMonth] = useState(todayMonth)
  const [delaySeconds, setDelaySeconds] = useState(DEFAULT_DELAY)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [aboutOpen, setAboutOpen] = useState(false)

  const [scanning, setScanning] = useState(false)
  const [abortPending, setAbortPending] = useState(false)
  const [showProgress, setShowProgress] = useState(false)
  const [statusText, setStatusText] = useState('')
  const [pct, setPct] = useState(0)

  const [showResults, setShowResults] = useState(false)
  const [rows, setRows] = useState<SummaryRow[]>(EMPTY_ROWS)
  const [showBanner, setShowBanner] = useState(false)

  const [alert, setAlert] = useState<Alert | null>(null)

  // Load saved config + active scan state on mount.
  useEffect(() => {
    void (async () => {
      const stored = (await chrome.storage.local.get(['requestDelay', 'summaryScanState'])) as {
        requestDelay?: number
        summaryScanState?: SummaryScanState
      }

      if (stored.requestDelay) {
        setDelaySeconds(Number(stored.requestDelay))
      }

      const scan = stored.summaryScanState
      if (scan) {
        if (scan.rows && scan.rows.length) {
          setRows([...scan.rows])
        }
        if (scan.isScanning) {
          setScanning(true)
          setShowProgress(true)
          setShowResults(true)
          setStatusText(scan.stepText || 'Scanning...')
          setPct(scan.pct ?? 0)
        } else if (scan.isDone) {
          setShowResults(true)
          setShowBanner(true)
        }
      }
    })()
  }, [])

  // Live updates from the background worker.
  useEffect(() => {
    const listener = (msg: SummaryWorkerMessage) => {
      if (msg.type === 'SUMMARY_PROGRESS') {
        if (msg.stepText) setStatusText(msg.stepText)
        if (msg.pct !== undefined) setPct(msg.pct)
        setShowProgress(true)
        if (msg.row) {
          setRows((prev) => {
            const next = [...prev, msg.row as SummaryRow]
            next.sort((a, b) => parseDateToTime(a.date) - parseDateToTime(b.date))
            return next
          })
          setShowResults(true)
        }
      } else if (msg.type === 'SUMMARY_DONE' || msg.type === 'SUMMARY_ABORTED') {
        setScanning(false)
        setAbortPending(false)
        setShowProgress(false)

        if (msg.type === 'SUMMARY_DONE') {
          setShowBanner(true)
          setAlert({
            message: 'Report downloaded automatically to your Downloads folder!',
            tone: 'success',
          })
        } else {
          setAlert({ message: 'Live scan stopped by user.', tone: 'error' })
        }
      } else if (msg.type === 'SUMMARY_ERROR') {
        setScanning(false)
        setShowProgress(false)
        setAlert({ message: `Scan Failed: ${msg.error}`, tone: 'error' })
      }
    }
    chrome.runtime.onMessage.addListener(listener)
    return () => {
      chrome.runtime.onMessage.removeListener?.(listener)
    }
  }, [])

  const handleGenerate = useCallback(() => {
    setShowBanner(false)
    if (!reportMonth) {
      setAlert({ message: 'Please select a report month.', tone: 'error' })
      return
    }

    setAlert(null)
    setScanning(true)
    setShowProgress(true)
    setShowResults(true)
    setRows(EMPTY_ROWS)
    setStatusText('Connecting MyGate API...')
    setPct(0)

    void chrome.storage.local.set({ requestDelay: Number(delaySeconds) })

    const [year, month] = reportMonth.split('-').map(Number)
    void chrome.runtime.sendMessage({
      type: 'START_SUMMARY_SCAN',
      year,
      month,
      requestDelay: Number(delaySeconds),
    })
  }, [reportMonth, delaySeconds])

  const handleAbort = useCallback(() => {
    void chrome.runtime.sendMessage({ type: 'ABORT_SUMMARY_SCAN' })
    setAbortPending(true)
  }, [])

  const handleDelayChange = useCallback((val: number) => {
    setDelaySeconds(val)
    void chrome.storage.local.set({ requestDelay: Number(val.toFixed(1)) })
  }, [])

  const kpiReceived = rows.reduce((sum, r) => sum + r.received, 0)
  const kpiClosed = rows.reduce((sum, r) => sum + r.closed, 0)
  const kpiPending = rows.length ? rows[rows.length - 1].pending : 0

  return (
    <>
      <div className="top-accent-bar" />
      <main className="app-wrapper">
        <BrandHeader
          title="Complaint Summary Sheet"
          subtitle="Monthly Automation Tool"
          onAboutClick={() => setAboutOpen(true)}
        />

        <section className="bg-white border border-line-subtle rounded-lg shadow-card p-4 flex flex-col gap-4">
          <MonthField
            id="report-month"
            label="Report Month *"
            value={reportMonth}
            onChange={setReportMonth}
          />

          {settingsOpen && (
            <SettingsDrawer
              heading="Report Speed"
              queryLabel="Delay Between Cloud Queries"
              helperText="Interval between each GraphQL count query in seconds"
              delaySeconds={delaySeconds}
              min={MIN_DELAY}
              max={MAX_DELAY}
              step={STEP_DELAY}
              onChange={handleDelayChange}
            />
          )}

          <div className="flex flex-col gap-2">
            {scanning ? (
              <AbortButton onClick={handleAbort} disabled={abortPending}>
                <StopIcon /> {abortPending ? 'Stopping...' : 'Stop Live Scan'}
              </AbortButton>
            ) : (
              <PrimaryButton onClick={handleGenerate}>
                <ZapIcon /> Generate Summary Sheet
              </PrimaryButton>
            )}
          </div>

          <div className="flex items-center justify-between">
            <TextLinkButton onClick={() => setSettingsOpen((v) => !v)}>
              <GearIcon /> Report Speed
            </TextLinkButton>
            <ExternalLink href="https://dashboard.mygate.com/">Open Dashboard &rarr;</ExternalLink>
          </div>
        </section>

        <AlertBanner message={alert?.message ?? ''} tone={alert?.tone ?? null} />

        {showProgress && <ProgressCard stepText={statusText} pct={pct} />}

        {showResults && (
          <section className="flex flex-col gap-3.5">
            <KpiGrid>
              <KpiCard label="Total Received" value={kpiReceived.toLocaleString()} tone="total" />
              <KpiCard label="Total Closed" value={kpiClosed.toLocaleString()} tone="resolved" />
              <KpiCard label="Final Pending" value={kpiPending.toLocaleString()} tone="open" />
            </KpiGrid>

            <div className="bg-white border border-line-subtle rounded-lg shadow-card p-4">
              <div className="table-scroll" style={{ maxHeight: 180 }}>
                <table className="data-table">
                  <thead>
                    <tr>
                      <th style={{ width: '28%' }}>Date</th>
                      <th style={{ textAlign: 'right' }}>Prev Open</th>
                      <th style={{ textAlign: 'right' }}>Received</th>
                      <th style={{ textAlign: 'right' }}>Closed</th>
                    </tr>
                  </thead>
                  <tbody key={rows.map((r) => r.date).join('|')}>
                    {rows.map((r) => (
                      <tr key={r.date}>
                        <td style={{ fontWeight: 600 }}>{r.date}</td>
                        <td className="num">{r.prevOpen.toLocaleString()}</td>
                        <td className="num">{r.received.toLocaleString()}</td>
                        <td className="num">{r.closed.toLocaleString()}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {showBanner && <AutoDownloadBanner message="🎉 Excel File Downloaded Automatically!" />}
          </section>
        )}

        <AboutModal open={aboutOpen} onClose={() => setAboutOpen(false)} />
      </main>
    </>
  )
}
