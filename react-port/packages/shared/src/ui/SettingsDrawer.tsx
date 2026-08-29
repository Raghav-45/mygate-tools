interface SettingsDrawerProps {
  heading: string
  queryLabel: string
  helperText: string
  delaySeconds: number
  min: number
  max: number
  step: number
  onChange: (value: number) => void
}

/** The "API Polling/Throttling Speed" drawer shared by all three popups. */
export function SettingsDrawer({
  heading,
  queryLabel,
  helperText,
  delaySeconds,
  min,
  max,
  step,
  onChange,
}: SettingsDrawerProps) {
  return (
    <div className="p-4 bg-[#F8FAFC] border border-dashed border-line-input rounded-md flex flex-col gap-3 animate-slide-in">
      <div className="flex justify-between items-center">
        <h4 className="text-[12px] font-bold text-[#334155]">⚙️ {heading}</h4>
        <span className="text-[10px] font-semibold text-[#64748B] bg-[#E2E8F0] px-1.5 py-0.5 rounded">
          Optional
        </span>
      </div>
      <div className="flex flex-col gap-1.5">
        <div className="flex justify-between items-center text-[11px] font-semibold text-ink-muted">
          <label>{queryLabel}</label>
          <span className="text-teal font-bold tabular-nums">{delaySeconds.toFixed(1)}s</span>
        </div>
        <input
          type="range"
          className="accent-teal h-1.5 w-full"
          min={min}
          max={max}
          step={step}
          value={delaySeconds}
          aria-label={queryLabel}
          onChange={(e) => onChange(Number(e.target.value))}
        />
        <span className="text-[10px] text-ink-dim block">{helperText}</span>
      </div>
    </div>
  )
}
