interface AlertBannerProps {
  message: string
  tone: 'error' | 'success' | null
}

/** Alert box shown as an error (red) or success (green) banner. */
export function AlertBanner({ message, tone }: AlertBannerProps) {
  if (!message || !tone) return null
  const classes =
    tone === 'error'
      ? 'bg-[#FEF2F2] border-[#FECACA] text-[#DC2626]'
      : 'bg-[#F0FDF4] border-[#BBF7D0] text-[#16A34A]'
  return (
    <div
      className={`px-4 py-3 rounded-md text-[12px] font-medium leading-normal text-center border ${classes}`}
    >
      {message}
    </div>
  )
}

/** Green "downloaded automatically" banner shown after a successful export. */
export function AutoDownloadBanner({
  message = '🎉 Master Dump Excel Downloaded Automatically!',
}: {
  message?: string
}) {
  return (
    <div className="mt-3.5 bg-[#ECFDF5] border-[1.5px] border-[#10B981] text-[#065F46] px-3.5 py-3.5 rounded-xl text-center text-[13px] font-bold flex items-center justify-center gap-2">
      <span>{message}</span>
    </div>
  )
}
