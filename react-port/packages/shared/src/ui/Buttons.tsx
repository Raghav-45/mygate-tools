import type { ReactNode } from 'react'

/** Primary yellow CTA shared by all three popups. */
export function PrimaryButton({
  children,
  onClick,
  disabled,
}: {
  children: ReactNode
  onClick: () => void
  disabled?: boolean
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      className="w-full h-12 bg-gradient-to-r from-yellow to-yellow-light text-[#111827] text-[15px] font-bold border-none rounded-sm cursor-pointer inline-flex items-center justify-center gap-2 shadow-yellow-glow transition-all duration-200 hover:from-yellow-light hover:to-[#EAB308] hover:-translate-y-px hover:shadow-yellow-glow-lg active:translate-y-0 disabled:opacity-60 disabled:cursor-not-allowed"
    >
      {children}
    </button>
  )
}

/** Red abort/secondary button (label is overridden while stopping). */
export function AbortButton({
  children,
  onClick,
  disabled,
}: {
  children: ReactNode
  onClick: () => void
  disabled?: boolean
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      className="w-full h-12 bg-gradient-to-r from-[#EF4444] to-[#DC2626] text-white text-[15px] font-bold border-none rounded-sm cursor-pointer inline-flex items-center justify-center gap-2 shadow-red-glow"
    >
      {children}
    </button>
  )
}

/** Inline text link button (e.g. "Polling Speed", gear icon). */
export function TextLinkButton({
  children,
  onClick,
}: {
  children: ReactNode
  onClick: () => void
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="bg-transparent border-none text-[13px] font-semibold text-teal no-underline cursor-pointer inline-flex items-center gap-1.5 transition-colors duration-150 hover:text-teal-hover hover:underline"
    >
      {children}
    </button>
  )
}

/** External anchor styled like a TextLinkButton (e.g. "Cloud Reports →"). */
export function ExternalLink({ href, children }: { href: string; children: ReactNode }) {
  return (
    <a
      href={href}
      target="_blank"
      rel="noreferrer"
      className="bg-transparent border-none text-[13px] font-semibold text-teal no-underline cursor-pointer inline-flex items-center gap-1.5 transition-colors duration-150 hover:text-teal-hover hover:underline"
    >
      {children}
    </a>
  )
}
