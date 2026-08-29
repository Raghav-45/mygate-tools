import { InfoIcon } from './icons'
import { Logo } from './Logo'

interface BrandHeaderProps {
  title: string
  subtitle: string
  onAboutClick: () => void
}

export function BrandHeader({ title, subtitle, onAboutClick }: BrandHeaderProps) {
  return (
    <header className="relative flex flex-col items-center text-center mt-1">
      <button
        type="button"
        onClick={onAboutClick}
        title="About & Credits"
        aria-label="About & Credits"
        className="absolute top-0 right-0 w-8 h-8 rounded-full bg-white border border-line-subtle text-ink-dim cursor-pointer inline-flex items-center justify-center shadow-sm transition-all duration-150 hover:bg-[#F1F5F9] hover:text-ink-main hover:border-line-input hover:scale-105"
      >
        <InfoIcon />
      </button>
      <div className="mb-3 transition-transform duration-200 hover:scale-[1.03]">
        <Logo />
      </div>
      <h2 className="text-[21px] font-extrabold text-[#1E293B] tracking-[-0.5px]">{title}</h2>
      <p className="text-xs text-ink-muted mb-1">{subtitle}</p>
    </header>
  )
}
