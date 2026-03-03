/**
 * Color conversion utilities for PPTX export.
 *
 * PptxGenJS expects colors as 6-digit hex strings without the '#' prefix,
 * with transparency specified separately as a 0-100 percentage.
 * OpenPencil stores colors as '#RRGGBB' or '#RRGGBBAA'.
 */

export interface PptxColor {
  color: string // 6-digit hex, e.g. 'FF0000'
  transparency: number // 0-100
}

const DEFAULT_COLOR: PptxColor = { color: '000000', transparency: 0 }

/**
 * Parse a CSS hex color string into pptxgenjs color + transparency.
 * Supports: #RGB, #RRGGBB, #RRGGBBAA
 */
export function parseHexColor(hex: string | undefined): PptxColor {
  if (!hex) return DEFAULT_COLOR

  let cleaned = hex.replace(/^#/, '')

  // Expand shorthand #RGB → RRGGBB
  if (cleaned.length === 3) {
    cleaned = cleaned[0] + cleaned[0] + cleaned[1] + cleaned[1] + cleaned[2] + cleaned[2]
  }

  if (cleaned.length === 8) {
    const alpha = parseInt(cleaned.slice(6, 8), 16)
    const transparency = Math.round((1 - alpha / 255) * 100)
    return { color: cleaned.slice(0, 6).toUpperCase(), transparency }
  }

  if (cleaned.length === 6) {
    return { color: cleaned.toUpperCase(), transparency: 0 }
  }

  return DEFAULT_COLOR
}

/**
 * Convert an opacity value (0-1) to a PPTX transparency (0-100).
 */
export function opacityToTransparency(opacity: number | string | undefined): number {
  if (opacity === undefined) return 0
  const num = typeof opacity === 'number' ? opacity : parseFloat(String(opacity))
  if (isNaN(num)) return 0
  return Math.round((1 - Math.max(0, Math.min(1, num))) * 100)
}

/**
 * Combine fill opacity and node opacity into a single transparency value.
 */
export function combinedTransparency(
  fillTransparency: number,
  nodeOpacity: number | string | undefined,
): number {
  const opacityT = opacityToTransparency(nodeOpacity)
  // Combine: both are 0-100 percentages
  // Effective visibility = (1 - fillT/100) * (1 - opacityT/100)
  const effectiveVisible = (1 - fillTransparency / 100) * (1 - opacityT / 100)
  return Math.round((1 - effectiveVisible) * 100)
}
