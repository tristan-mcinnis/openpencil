/**
 * Pixel-to-inch coordinate conversion for PPTX export.
 *
 * PowerPoint uses inches as its coordinate system. OpenPencil canvas uses
 * pixels at 96 DPI (standard CSS pixels). This module handles all conversions.
 */

const PX_PER_INCH = 96

/** Convert pixels to inches. */
export function pxToInch(px: number): number {
  return px / PX_PER_INCH
}

/** Convert a point size (used by fonts) to inches. */
export function ptToInch(pt: number): number {
  return pt / 72
}

/**
 * Calculate slide dimensions from a frame's pixel size.
 * Returns inches, clamped to reasonable PPTX limits.
 */
export function frameDimensionsToSlide(widthPx: number, heightPx: number): {
  width: number
  height: number
} {
  const width = Math.max(1, pxToInch(widthPx))
  const height = Math.max(1, pxToInch(heightPx))
  return { width, height }
}
