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
import type { CategoryTuple } from '../report/categories'
import type {
  ReportWorkerMessage,
  ReportScanParams,
  ReportSummary,
  CategoryData,
} from '../report/reportState'

const DEFAULT_FROM_DATE = '2024-01-01'
const MIN_DELAY = 0.2
const MAX_DELAY = 3.0
const STEP_DELAY = 0.2
const DEFAULT_DELAY = 1.0

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

function CategoryPills({
  categories,
  selected,
  onToggle,
}: {
  categories: CategoryTuple[]
  selected: Set<number>
  onToggle: (id: number) => void
}) {
  return (
    <div className="flex flex-col gap-1.5">
      {categories.map(([name, id]) => {
        const isSelected = selected.has(id)
        return (
          <button
            key={id}
            type="button"
            aria-pressed={isSelected}
            onClick={() => onToggle(id)}
            className={`w-full h-[34px] px-3 text-left text-[12px] font-semibold rounded-full border transition-colors ${
              isSelected
                ? 'bg-[#10B981] border-[#10B981] text-white'
                : 'bg-white border-line-input text-ink-muted hover:border-teal hover:text-teal'
            }`}
          >
            {isSelected ? '✓ ' : ''}
            {name}
          </button>
        )
      })}
    </div>
  )
}

interface Alert {
  message: string
  tone: 'error' | 'success'
}

const EMPTY_SUMMARY: ReportSummary = { total: 0, resolved: 0, open: 0 }

export default function App() {
  const [fromDate, setFromDate] = useState(DEFAULT_FROM_DATE)
  const [toDate, setToDate] = useState('')
  const [delaySeconds, setDelaySeconds] = useState(DEFAULT_DELAY)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [aboutOpen, setAboutOpen] = useState(false)

  const [categories, setCategories] = useState<CategoryTuple[]>([])
  const [selected, setSelected] = useState<Set<number>>(new Set())
  const [savedSelection, setSavedSelection] = useState<number[] | null>(null)

  const [scanning, setScanning] = useState(false)
  const [abortPending, setAbortPending] = useState(false)
  const [showProgress, setShowProgress] = useState(false)
  const [statusText, setStatusText] = useState('')
  const [pct, setPct] = useState(0)

  const [showResults, setShowResults] = useState(false)
  const [results, setResults] = useState<CategoryData[]>([])
  const [summary, setSummary] = useState<ReportSummary>(EMPTY_SUMMARY)
  const [showBanner, setShowBanner] = useState(false)

  const [alert, setAlert] = useState<Alert | null>(null)

  // Default "Report Till" to today (ISO date, matching the original).
  useEffect(() => {
    setToDate(new Date().toISOString().split('T')[0])
  }, [])

  // Load saved config + active scan state on mount.
  useEffect(() => {
    void (async () => {
      const stored = (await chrome.storage.local.get([
        'requestDelay',
        'ticketsScanState',
        'selectedCatIds',
      ])) as {
        requestDelay?: number
        ticketsScanState?: {
          pct?: number
          statusText?: string
          results?: CategoryData[]
          summary?: ReportSummary
          isScanning?: boolean
          isDone?: boolean
        }
        selectedCatIds?: number[]
      }

      if (stored.requestDelay) {
        setDelaySeconds(Number(stored.requestDelay))
      }
      setSavedSelection(Array.isArray(stored.selectedCatIds) ? stored.selectedCatIds : null)

      const scan = stored.ticketsScanState
      if (scan) {
        if (scan.results && scan.results.length) {
          setResults(scan.results)
          setSummary(scan.summary ?? EMPTY_SUMMARY)
          setShowResults(true)
        }
        if (scan.isScanning) {
          setScanning(true)
          setShowProgress(true)
          setStatusText(scan.statusText || 'Connecting to MyGate API...')
          setPct(scan.pct ?? 0)
        } else if (scan.isDone) {
          setShowBanner(true)
        }
      }
    })()
  }, [])

  // Fetch the category list from the background worker.
  useEffect(() => {
    void (async () => {
      const res = (await chrome.runtime.sendMessage({ type: 'GET_CATEGORIES_LIST' })) as {
        categories?: CategoryTuple[]
      }
      setCategories(res?.categories ?? [])
    })()
  }, [])

  // Category pills default to selected; a saved list only deselects what it omits.
  useEffect(() => {
    if (!categories.length) return
    const ids = categories.map(([, id]) => id)
    const keep =
      savedSelection && savedSelection.length > 0
        ? savedSelection.filter((id) => ids.includes(id))
        : ids
    setSelected(new Set(keep))
  }, [categories, savedSelection])

  // Live updates from the background worker.
  useEffect(() => {
    const listener = (msg: ReportWorkerMessage) => {
      if (msg.type === 'SCAN_PROGRESS_UPDATE') {
        setStatusText(msg.statusText)
        // Original popup math: missing stepIndex/totalSteps (the "Downloading
        // Excel..." update) yields NaN and renders "NaN%".
        const step = msg.stepIndex
        const total = msg.totalSteps
        setPct(step !== undefined && total !== undefined ? Math.round((step / total) * 100) : NaN)
        setShowProgress(true)
      } else if (msg.type === 'CATEGORY_COMPLETED') {
        setResults((prev) => [...prev, msg.data])
        setSummary(msg.summary)
        setShowResults(true)
      } else if (msg.type === 'SCAN_FINISHED' || msg.type === 'SCAN_ABORTED') {
        setScanning(false)
        setAbortPending(false)
        setShowProgress(false)

        if (msg.type === 'SCAN_FINISHED') {
          setShowBanner(true)
          setAlert({
            message: 'Report downloaded automatically to your Downloads folder!',
            tone: 'success',
          })
        } else {
          setAlert({ message: 'Scan stopped by user.', tone: 'error' })
        }
      } else if (msg.type === 'SCAN_ERROR') {
        setScanning(false)
        setShowProgress(false)
        setAlert({
          message: msg.errorMessage || `Error querying category: ${msg.category}`,
          tone: 'error',
        })
      }
    }
    chrome.runtime.onMessage.addListener(listener)
    return () => {
      chrome.runtime.onMessage.removeListener?.(listener)
    }
  }, [])

  const handleGenerate = useCallback(() => {
    setShowBanner(false)
    if (new Date(fromDate) > new Date(toDate)) {
      setAlert({ message: 'Report From date cannot be after Till date.', tone: 'error' })
      return
    }
    if (selected.size === 0) {
      setAlert({ message: 'Please select at least one Ticket Category.', tone: 'error' })
      return
    }

    setAlert(null)
    setScanning(true)
    setShowProgress(true)
    setShowResults(false)
    setResults([])
    setSummary(EMPTY_SUMMARY)
    setStatusText('Connecting to MyGate API...')
    setPct(0)

    void chrome.storage.local.set({ requestDelay: Number(delaySeconds) })

    const start: ReportScanParams = {
      fromDate,
      toDate,
      requestDelayMs: Number(delaySeconds) * 1000,
      selectedCategoryIds: [...selected],
    }
    void chrome.runtime.sendMessage({ type: 'START_REPORT_SCAN', params: start })
  }, [fromDate, toDate, delaySeconds, selected])

  const handleAbort = useCallback(() => {
    void chrome.runtime.sendMessage({ type: 'ABORT_REPORT_SCAN' })
    setAbortPending(true)
  }, [])

  const handleDelayChange = useCallback((val: number) => {
    setDelaySeconds(val)
    void chrome.storage.local.set({ requestDelay: Number(val.toFixed(1)) })
  }, [])

  const handleToggleCategory = useCallback(
    (id: number) => {
      const next = new Set(selected)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      setSelected(next)
      void chrome.storage.local.set({ selectedCatIds: [...next] })
    },
    [selected],
  )

  const handleToggleAll = useCallback(() => {
    const ids = categories.map(([, id]) => id)
    const allSelected = ids.length > 0 && ids.every((id) => selected.has(id))
    const next = allSelected ? new Set<number>() : new Set(ids)
    setSelected(next)
    void chrome.storage.local.set({ selectedCatIds: [...next] })
  }, [categories, selected])

  return (
    <>
      <div className="top-accent-bar" />
      <main className="app-wrapper">
        <BrandHeader
          title="Pending Tickets Report"
          subtitle="Admin Dashboard Automation"
          onAboutClick={() => setAboutOpen(true)}
        />

        <section className="bg-white border border-line-subtle rounded-lg shadow-card p-4 flex flex-col gap-4">
          <div className="grid grid-cols-2 gap-3">
            <DateField
              id="report-from"
              label="Report From *"
              value={fromDate}
              onChange={setFromDate}
            />
            <DateField id="report-till" label="Report Till *" value={toDate} onChange={setToDate} />
          </div>

          <div className="flex flex-col gap-1.5">
            <div className="flex items-center justify-between">
              <label className="text-[11px] font-semibold text-ink-dim">Ticket Categories</label>
              <TextLinkButton onClick={handleToggleAll}>Toggle All</TextLinkButton>
            </div>
            <CategoryPills
              categories={categories}
              selected={selected}
              onToggle={handleToggleCategory}
            />
          </div>

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
                <ZapIcon /> Generate Report
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
              <KpiCard label="Total Tickets" value={summary.total.toLocaleString()} tone="total" />
              <KpiCard
                label="Resolved Tickets"
                value={summary.resolved.toLocaleString()}
                tone="resolved"
              />
              <KpiCard label="Open Tickets" value={summary.open.toLocaleString()} tone="open" />
            </KpiGrid>

            <div className="bg-white border border-line-subtle rounded-lg shadow-card p-4">
              <div className="table-scroll" style={{ maxHeight: 180 }}>
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>Category</th>
                      <th style={{ textAlign: 'right' }}>Total</th>
                      <th style={{ textAlign: 'right' }}>Resolved</th>
                      <th style={{ textAlign: 'right' }}>Open</th>
                    </tr>
                  </thead>
                  <tbody>
                    {results.map((c) => (
                      <tr key={c.id}>
                        <td style={{ fontWeight: 600 }}>{c.name}</td>
                        <td className="num">{c.total.toLocaleString()}</td>
                        <td className="num">{c.resolved.toLocaleString()}</td>
                        <td className="num">{c.open.toLocaleString()}</td>
                      </tr>
                    ))}
                    <tr className="bg-[#FEF08A] font-bold">
                      <td>Total</td>
                      <td className="num">{summary.total.toLocaleString()}</td>
                      <td className="num">{summary.resolved.toLocaleString()}</td>
                      <td className="num">{summary.open.toLocaleString()}</td>
                    </tr>
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
