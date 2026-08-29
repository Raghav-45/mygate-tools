interface ProgressCardProps {
  stepText: string
  pct: number
}

/** Live progress bar with a pulsing indicator and percentage. */
export function ProgressCard({ stepText, pct }: ProgressCardProps) {
  const finitePct = Number.isFinite(pct) ? pct : 0
  return (
    <section className="p-4 px-5 bg-white border border-line-subtle rounded-lg shadow-card flex flex-col gap-2.5">
      <div className="flex justify-between items-center">
        <div className="flex items-center gap-2">
          <span className="w-2 h-2 rounded-full bg-[#10B981] shadow-[0_0_10px_#10B981] animate-ping-slow" />
          <span className="text-[12px] font-semibold text-[#334155] whitespace-nowrap overflow-hidden text-ellipsis max-w-[310px]">
            {stepText}
          </span>
        </div>
        <span className="text-[13px] font-extrabold text-ink-main">{pct}%</span>
      </div>
      <div className="w-full h-2 bg-[#F1F5F9] rounded-full overflow-hidden">
        <div
          className="h-full rounded-full bg-gradient-to-r from-teal to-[#10B981] transition-[width] duration-300 ease-out"
          style={{ width: `${finitePct}%` }}
        />
      </div>
    </section>
  )
}
