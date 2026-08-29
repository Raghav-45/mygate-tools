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
  PrimaryButton,
  ProgressCard,
  SettingsDrawer,
  StopIcon,
  TextLinkButton,
  ZapIcon,
} from '@mygate/shared'
import type { ChunkResult, DumpWorkerMessage, StartDumpExport } from '../dump/dumpState'

const DEFAULT_FROM_DATE = '2024-01-01'
const MIN_DELAY = 1.0
const MAX_DELAY = 5.0
const STEP_DELAY = 0.5
const DEFAULT_DELAY = 2.0

function DateField({
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
        type="date"
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

export default function App() {
  const [fromDate, setFromDate] = useState(DEFAULT_FROM_DATE)
  const [toDate, setToDate] = useState('')
  const [delaySeconds, setDelaySeconds] = useState(DEFAULT_DELAY)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [aboutOpen, setAboutOpen] = useState(false)

  const [scanning, setScanning] = useState(false)
  const [abortPending, setAbortPending] = useState(false)
  const [showProgress, setShowProgress] = useState(false)
  const [statusText, setStatusText] = useState('')
  const [pct, setPct] = useState(0)

  const [showResults, setShowResults] = useState(false)
  const [chunks, setChunks] = useState<ChunkResult[]>([])
  const [totalRows, setTotalRows] = useState(0)
  const [kpiStatus, setKpiStatus] = useState('')
  const [showBanner, setShowBanner] = useState(false)

  const [alert, setAlert] = useState<Alert | null>(null)

  // Default "Dump Till" to today (ISO date, matching the original).
  useEffect(() => {
    setToDate(new Date().toISOString().split('T')[0])
  }, [])

  // Load saved config + active scan state on mount.
  useEffect(() => {
    void (async () => {
      const stored = (await chrome.storage.local.get(['requestDelay', 'dumpScanState'])) as {
        requestDelay?: number
        dumpScanState?: {
          pct?: number
          statusText?: string
          chunks?: ChunkResult[]
          totalRows?: number
          isScanning?: boolean
          isDone?: boolean
        }
      }

      if (stored.requestDelay) {
        // Original quirk (kept as-is): the value in storage is milliseconds but is
        // fed straight into the seconds slider and its label. See NOTES.md.
        setDelaySeconds(Number(stored.requestDelay))
      }

      const scan = stored.dumpScanState
      if (scan) {
        if (scan.chunks && scan.chunks.length) {
          setChunks(scan.chunks)
          setTotalRows(scan.totalRows ?? 0)
          setShowResults(true)
        }
        if (scan.isScanning) {
          setScanning(true)
          setShowProgress(true)
          setStatusText(scan.statusText || 'Exporting...')
          setPct(scan.pct || 0)
          setKpiStatus('Active')
        } else if (scan.isDone) {
          setShowBanner(true)
          setKpiStatus('Completed')
        }
      }
    })()
  }, [])

  // Live updates from the background worker.
  useEffect(() => {
    const listener = (msg: DumpWorkerMessage) => {
      if (msg.type === 'DUMP_PROGRESS_UPDATE') {
        setStatusText(msg.statusText)
        setPct(msg.pct)
        setKpiStatus('Exporting...')
        if (msg.chunks) {
          setChunks(msg.chunks)
          setTotalRows(msg.totalRows)
        }
      } else if (msg.type === 'DUMP_FINISHED' || msg.type === 'DUMP_ABORTED') {
        setScanning(false)
        setAbortPending(false)
        setShowProgress(false)

        if (msg.type === 'DUMP_FINISHED') {
          setKpiStatus('Completed')
          setShowBanner(true)
          if (msg.chunks) {
            setChunks(msg.chunks)
            setTotalRows(msg.totalRows)
          }
          setAlert({
            message: 'Master Dump exported and downloaded automatically!',
            tone: 'success',
          })
        } else {
          setKpiStatus('Stopped')
          setAlert({ message: 'Export stopped by user.', tone: 'error' })
        }
      }
    }
    chrome.runtime.onMessage.addListener(listener)
    return () => {
      chrome.runtime.onMessage.removeListener?.(listener)
    }
  }, [])

  const handleGenerate = useCallback(() => {
    setShowBanner(false)
    if (!fromDate || !toDate) {
      setAlert({ message: 'Please enter valid Dump From and Dump Till dates.', tone: 'error' })
      return
    }
    if (new Date(fromDate) > new Date(toDate)) {
      setAlert({ message: 'Dump From date cannot be after Dump Till date.', tone: 'error' })
      return
    }

    setAlert(null)
    setScanning(true)
    setShowProgress(true)
    setShowResults(true)
    setStatusText('Initiating multi-year slice...')
    setPct(0)
    setKpiStatus('Starting...')

    const start: StartDumpExport = {
      type: 'START_DUMP_EXPORT',
      params: {
        fromDate,
        toDate,
        requestDelayMs: Number(delaySeconds) * 1000,
      },
    }
    void chrome.runtime.sendMessage(start)
  }, [fromDate, toDate, delaySeconds])

  const handleAbort = useCallback(() => {
    void chrome.runtime.sendMessage({ type: 'ABORT_DUMP_EXPORT' })
    setAbortPending(true)
  }, [])

  const handleDelayChange = useCallback((val: number) => {
    setDelaySeconds(val)
    void chrome.storage.local.set({ requestDelay: val * 1000 })
  }, [])

  const readyCount = chunks.filter((c) => c.status && c.status.includes('Merged')).length

  return (
    <>
      <div className="top-accent-bar" />
      <main className="app-wrapper">
        <BrandHeader
          title="Multi-Year Ticket Dump"
          subtitle="Bypasses 1-Year Limit & Auto-Merges"
          onAboutClick={() => setAboutOpen(true)}
        />

        <section className="bg-white border border-line-subtle rounded-lg shadow-card p-4 flex flex-col gap-4">
          <div className="grid grid-cols-2 gap-3">
            <DateField id="dump-from" label="Dump From *" value={fromDate} onChange={setFromDate} />
            <DateField id="dump-till" label="Dump Till *" value={toDate} onChange={setToDate} />
          </div>

          <div className="text-[12px] text-ink-muted -mt-2 mb-1 text-center">
            ⚡ Select any multi-year range (e.g. 2024 to 2026). The tool slices &amp; merges it
            automatically!
          </div>

          {settingsOpen && (
            <SettingsDrawer
              heading="API Polling Speed"
              queryLabel="Delay Between Cloud Queries"
              helperText="Interval for checking if MyGate cloud generated your chunk"
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
                <StopIcon /> {abortPending ? 'Stopping...' : 'Stop Live Export'}
              </AbortButton>
            ) : (
              <PrimaryButton onClick={handleGenerate}>
                <ZapIcon /> Generate Master Dump
              </PrimaryButton>
            )}
          </div>

          <div className="flex items-center justify-between">
            <TextLinkButton onClick={() => setSettingsOpen((v) => !v)}>
              <GearIcon /> Polling Speed
            </TextLinkButton>
            <ExternalLink href="https://dashboard.mygate.com/home/society/generatedReports">
              Cloud Reports &rarr;
            </ExternalLink>
          </div>
        </section>

        <AlertBanner message={alert?.message ?? ''} tone={alert?.tone ?? null} />

        {showProgress && <ProgressCard stepText={statusText} pct={pct} />}

        {showResults && (
          <section className="flex flex-col gap-3.5">
            <KpiGrid>
              <KpiCard
                label="Yearly Chunks"
                value={`${readyCount} / ${chunks.length}`}
                tone="total"
              />
              <KpiCard
                label="Total Merged Rows"
                value={totalRows.toLocaleString()}
                tone="resolved"
              />
              <KpiCard label="Status" value={kpiStatus || 'Active'} tone="open" />
            </KpiGrid>

            <div className="bg-white border border-line-subtle rounded-lg shadow-card p-4">
              <div className="table-scroll" style={{ maxHeight: 180 }}>
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>Chunk Date Range</th>
                      <th style={{ textAlign: 'right' }}>Rows Found</th>
                      <th style={{ textAlign: 'right' }}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {chunks.map((c) => {
                      let statusStyle = { color: '#D97706', fontWeight: 600 }
                      if (c.status.includes('Merged'))
                        statusStyle = { color: '#10B981', fontWeight: 700 }
                      else if (c.status.includes('Error') || c.status.includes('Failed'))
                        statusStyle = { color: '#EF4444', fontWeight: 700 }
                      return (
                        <tr key={c.rangeStr}>
                          <td style={{ fontWeight: 600 }}>{c.rangeStr}</td>
                          <td className="num">{c.rowsFound || '-'}</td>
                          <td className="num" style={statusStyle}>
                            {c.status}
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            {showBanner && <AutoDownloadBanner />}
          </section>
        )}

        <AboutModal open={aboutOpen} onClose={() => setAboutOpen(false)} />
      </main>
    </>
  )
}
