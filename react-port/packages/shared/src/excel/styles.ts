import type { Borders, Fill, Font } from 'exceljs'

/**
 * Visual language of the dump-tool and report-tool workbooks (both originals use
 * the same blue fill `FF4D93D9`, Aptos 12 fonts and thin black borders). The
 * summary-tool deliberately uses its own Calibri/white style and keeps it local.
 */

export const blueFill: Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF4D93D9' },
}

export const whiteFill: Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFFFFF' },
}

export const aptosTitleFont: Partial<Font> = {
  name: 'Aptos',
  size: 12,
  bold: false,
  color: { argb: 'FF000000' },
}

export const thinSide = { style: 'thin' as const, color: { argb: 'FF000000' } }

export const cellBorder: Partial<Borders> = {
  top: thinSide,
  left: thinSide,
  bottom: thinSide,
  right: thinSide,
}
