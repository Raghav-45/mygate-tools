export type KpiTone = 'total' | 'resolved' | 'open'

const tones: Record<KpiTone, { bg: string; border: string; value: string }> = {
  total: { bg: 'bg-[#EEF2FF]', border: 'border-[#C7D2FE]', value: 'text-[#4F46E5]' },
  resolved: { bg: 'bg-[#ECFDF5]', border: 'border-[#A7F3D0]', value: 'text-[#059669]' },
  open: { bg: 'bg-[#FFFBEB]', border: 'border-[#FDE68A]', value: 'text-[#D97706]' },
}

interface KpiCardProps {
  label: string
  value: string
  tone: KpiTone
}

export function KpiCard({ label, value, tone }: KpiCardProps) {
  const t = tones[tone]
  return (
    <div
      className={`rounded-md py-3.5 px-2.5 flex flex-col items-center gap-1 shadow-sm transition-transform duration-200 hover:-translate-y-0.5 ${t.bg} ${t.border} border`}
    >
      <span className={`text-[22px] font-extrabold tabular-nums ${t.value}`}>{value}</span>
      <span className="text-[10px] font-bold text-ink-muted uppercase tracking-[0.5px]">
        {label}
      </span>
    </div>
  )
}

interface KpiGridProps {
  children: React.ReactNode
}

export function KpiGrid({ children }: KpiGridProps) {
  return <div className="grid grid-cols-3 gap-3">{children}</div>
}
