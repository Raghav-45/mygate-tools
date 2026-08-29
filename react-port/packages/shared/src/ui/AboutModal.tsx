import { GithubIcon, GlobeIcon } from './icons'

interface AboutModalProps {
  open: boolean
  onClose: () => void
}

/** Author/credits modal — identical across all three original popups. */
export function AboutModal({ open, onClose }: AboutModalProps) {
  if (!open) return null

  return (
    <div
      className="fixed inset-0 z-[100] flex items-center justify-center p-5 bg-[rgba(15,23,42,0.55)] backdrop-blur-[6px] animate-fade-in"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose()
      }}
    >
      <div className="relative w-full max-w-[340px] bg-white rounded-[28px] px-6 py-7 pb-6 flex flex-col items-center text-center shadow-modal border border-white/80 animate-scale-up">
        <button
          type="button"
          onClick={onClose}
          aria-label="Close"
          className="absolute top-[14px] right-[14px] w-7 h-7 rounded-full bg-[#F1F5F9] border-none text-[18px] leading-none text-ink-dim cursor-pointer flex items-center justify-center transition-all duration-150 hover:bg-[#E2E8F0] hover:text-ink-main"
        >
          &times;
        </button>

        <div className="mb-3">
          <img
            src="https://github.com/raghav-45.png?size=160"
            alt="Aditya Singh Khichi"
            className="w-[84px] h-[84px] rounded-full border-[3px] border-[#10B981] shadow-[0_8px_16px_rgba(16,185,129,0.25)]"
          />
        </div>
        <h3 className="text-[18px] font-extrabold text-ink-main">Aditya Singh Khichi</h3>
        <span className="text-[11px] font-bold text-[#059669] bg-[#ECFDF5] px-3 py-1 rounded-full mt-1 mb-3">
          Full Stack Engineer & AI Engineer
        </span>
        <p className="text-[12.5px] text-ink-muted leading-relaxed mb-[18px]">
          Crafted with ❤️ to make daily tasks effortless ⚡
        </p>

        <div className="flex w-full gap-2.5">
          <a
            href="https://github.com/raghav-45"
            target="_blank"
            rel="noreferrer"
            className="flex-1 h-[42px] rounded-[14px] text-[12px] font-bold no-underline inline-flex items-center justify-center gap-1.5 transition-all duration-150 bg-[#0F172A] text-white hover:bg-[#1E293B] hover:-translate-y-px"
          >
            <GithubIcon />
            raghav-45
          </a>
          <a
            href="https://aditya.is-a.dev"
            target="_blank"
            rel="noreferrer"
            className="flex-1 h-[42px] rounded-[14px] text-[12px] font-bold no-underline inline-flex items-center justify-center gap-1.5 transition-all duration-150 bg-[#ECFDF5] text-[#059669] border border-[#A7F3D0] hover:bg-[#10B981] hover:text-white hover:border-[#059669] hover:-translate-y-px"
          >
            <GlobeIcon />
            aditya.is-a.dev
          </a>
        </div>
      </div>
    </div>
  )
}
