/**
 * Maps PenNode types to pptxgenjs slide objects.
 *
 * Each handler converts a resolved PenNode into pptxgenjs API calls on a Slide.
 * Children of containers are flattened with absolute coordinates before mapping.
 */

import type PptxGenJS from 'pptxgenjs'
import type { PenNode, ContainerProps } from '@/types/pen'
import type { PenFill, PenStroke, ShadowEffect } from '@/types/styles'
import { pxToInch } from './coordinate-utils'
import { parseHexColor, opacityToTransparency, combinedTransparency } from './color-utils'

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

interface SlideContext {
  slide: PptxGenJS.Slide
  pptx: PptxGenJS
  /** Frame origin — subtracted to get slide-relative coordinates */
  originX: number
  originY: number
}

interface FlatNode {
  node: PenNode
  absX: number
  absY: number
}

// ---------------------------------------------------------------------------
// Fill conversion
// ---------------------------------------------------------------------------

function mapFill(
  fills: PenFill[] | undefined,
  nodeOpacity?: number | string,
): PptxGenJS.ShapeFillProps | undefined {
  if (!fills || fills.length === 0) return undefined

  // Use the first visible fill (PPTX doesn't support stacked fills)
  const fill = fills[0]
  if (fill.type === 'solid') {
    const { color, transparency } = parseHexColor(fill.color)
    const fillOpacity = fill.opacity !== undefined ? opacityToTransparency(fill.opacity) : transparency
    return {
      color,
      transparency: combinedTransparency(fillOpacity, nodeOpacity),
    }
  }
  if (fill.type === 'linear_gradient' || fill.type === 'radial_gradient') {
    // PPTX doesn't have great gradient support via pptxgenjs ShapeFillProps.
    // Fall back to the first gradient stop color.
    if (fill.stops.length > 0) {
      const { color } = parseHexColor(fill.stops[0].color)
      const fillOpacity = fill.opacity !== undefined ? opacityToTransparency(fill.opacity) : 0
      return {
        color,
        transparency: combinedTransparency(fillOpacity, nodeOpacity),
      }
    }
  }
  if (fill.type === 'image') {
    // Image fills are handled separately via addImage
    return undefined
  }
  return undefined
}

// ---------------------------------------------------------------------------
// Stroke / line conversion
// ---------------------------------------------------------------------------

function mapLine(
  stroke: PenStroke | undefined,
): PptxGenJS.ShapeLineProps | undefined {
  if (!stroke) return undefined

  const thickness = typeof stroke.thickness === 'number'
    ? stroke.thickness
    : Array.isArray(stroke.thickness) ? stroke.thickness[0] : 1

  if (thickness <= 0) return undefined

  const strokeColor = stroke.fill?.[0]
  let color = '333333'
  let transparency = 0

  if (strokeColor && strokeColor.type === 'solid') {
    const parsed = parseHexColor(strokeColor.color)
    color = parsed.color
    transparency = parsed.transparency
  }

  const line: PptxGenJS.ShapeLineProps = {
    color,
    width: thickness,
    transparency,
  }

  if (stroke.dashPattern && stroke.dashPattern.length > 0) {
    line.dashType = 'dash'
  }

  return line
}

// ---------------------------------------------------------------------------
// Shadow conversion
// ---------------------------------------------------------------------------

function mapShadow(
  effects: Array<{ type: string; [k: string]: unknown }> | undefined,
): PptxGenJS.ShadowProps | undefined {
  if (!effects) return undefined

  const shadow = effects.find((e) => e.type === 'shadow') as ShadowEffect | undefined
  if (!shadow) return undefined

  const { color, transparency } = parseHexColor(shadow.color)
  const offsetDist = Math.sqrt(shadow.offsetX ** 2 + shadow.offsetY ** 2)
  const angleDeg = Math.round(
    (Math.atan2(shadow.offsetY, shadow.offsetX) * 180) / Math.PI,
  )

  return {
    type: shadow.inner ? 'inner' : 'outer',
    color,
    blur: Math.round(shadow.blur * 0.75), // px to pt approximation
    offset: Math.round(offsetDist * 0.75),
    angle: ((angleDeg % 360) + 360) % 360,
    opacity: 1 - transparency / 100,
  }
}

// ---------------------------------------------------------------------------
// Corner radius conversion
// ---------------------------------------------------------------------------

function mapRectRadius(
  cornerRadius: number | [number, number, number, number] | undefined,
  widthPx: number,
): number {
  if (!cornerRadius) return 0
  const r = typeof cornerRadius === 'number' ? cornerRadius : cornerRadius[0]
  if (r <= 0 || widthPx <= 0) return 0
  // pptxgenjs rectRadius is 0.0-1.0 relative to half the width
  return Math.min(r / (widthPx / 2), 1)
}

// ---------------------------------------------------------------------------
// Node flattening — convert tree to absolute-position list
// ---------------------------------------------------------------------------

export function flattenNodeTree(
  nodes: PenNode[],
  parentX: number,
  parentY: number,
): FlatNode[] {
  const result: FlatNode[] = []

  for (const node of nodes) {
    if (node.visible === false) continue

    const absX = parentX + (node.x ?? 0)
    const absY = parentY + (node.y ?? 0)

    // For containers, add the container itself first, then its children
    if ('children' in node && node.children && node.children.length > 0) {
      result.push({ node, absX, absY })
      result.push(...flattenNodeTree(node.children, absX, absY))
    } else {
      result.push({ node, absX, absY })
    }
  }

  return result
}

// ---------------------------------------------------------------------------
// Node width/height helpers
// ---------------------------------------------------------------------------

function getNodeWidth(node: PenNode): number {
  if ('width' in node && typeof node.width === 'number') return node.width
  return 0
}

function getNodeHeight(node: PenNode): number {
  if ('height' in node && typeof node.height === 'number') return node.height
  return 0
}

// ---------------------------------------------------------------------------
// Individual node mappers
// ---------------------------------------------------------------------------

function addRectangle(ctx: SlideContext, flat: FlatNode): void {
  const { node, absX, absY } = flat
  const w = getNodeWidth(node)
  const h = getNodeHeight(node)
  if (w <= 0 || h <= 0) return

  const container = node as PenNode & ContainerProps
  const fill = mapFill(container.fill, node.opacity)
  const line = mapLine(container.stroke)
  const shadow = mapShadow(container.effects as Array<{ type: string; [k: string]: unknown }> | undefined)
  const radius = mapRectRadius(container.cornerRadius, w)

  const shapeName = radius > 0
    ? ctx.pptx.ShapeType.roundRect
    : ctx.pptx.ShapeType.rect

  ctx.slide.addShape(shapeName, {
    x: pxToInch(absX - ctx.originX),
    y: pxToInch(absY - ctx.originY),
    w: pxToInch(w),
    h: pxToInch(h),
    fill: fill ?? undefined,
    line: line ?? undefined,
    shadow: shadow ?? undefined,
    rectRadius: radius > 0 ? radius : undefined,
    rotate: node.rotation ?? undefined,
    flipH: node.flipX ?? undefined,
    flipV: node.flipY ?? undefined,
  })
}

function addEllipse(ctx: SlideContext, flat: FlatNode): void {
  const { node, absX, absY } = flat
  const w = getNodeWidth(node)
  const h = getNodeHeight(node)
  if (w <= 0 || h <= 0) return

  const ellipseNode = node as PenNode & { fill?: PenFill[]; stroke?: PenStroke; effects?: unknown[] }
  const fill = mapFill(ellipseNode.fill, node.opacity)
  const line = mapLine(ellipseNode.stroke)
  const shadow = mapShadow(ellipseNode.effects as Array<{ type: string; [k: string]: unknown }> | undefined)

  ctx.slide.addShape(ctx.pptx.ShapeType.ellipse, {
    x: pxToInch(absX - ctx.originX),
    y: pxToInch(absY - ctx.originY),
    w: pxToInch(w),
    h: pxToInch(h),
    fill: fill ?? undefined,
    line: line ?? undefined,
    shadow: shadow ?? undefined,
    rotate: node.rotation ?? undefined,
    flipH: node.flipX ?? undefined,
    flipV: node.flipY ?? undefined,
  })
}

function addLine(ctx: SlideContext, flat: FlatNode): void {
  const { node, absX, absY } = flat
  if (node.type !== 'line') return

  const x2 = node.x2 ?? 0
  const y2 = node.y2 ?? 0
  const lineW = Math.abs(x2)
  const lineH = Math.abs(y2)
  const line = mapLine(node.stroke)

  ctx.slide.addShape(ctx.pptx.ShapeType.line, {
    x: pxToInch(absX - ctx.originX),
    y: pxToInch(absY - ctx.originY),
    w: pxToInch(lineW || 1),
    h: pxToInch(lineH),
    line: line ?? { color: '333333', width: 1 },
    rotate: node.rotation ?? undefined,
    flipH: (x2 < 0) !== (node.flipX ?? false),
    flipV: (y2 < 0) !== (node.flipY ?? false),
  })
}

function addText(ctx: SlideContext, flat: FlatNode): void {
  const { node, absX, absY } = flat
  if (node.type !== 'text') return

  const w = getNodeWidth(node)
  const h = getNodeHeight(node)

  // Build text props for rich text segments
  const textProps: PptxGenJS.TextProps[] = typeof node.content === 'string'
    ? [{ text: node.content, options: buildTextSegmentOptions(node) }]
    : node.content.map((seg) => ({
      text: seg.text,
      options: buildTextSegmentOptions(node, seg),
    }))

  const fill = mapFill(node.fill, node.opacity)
  const shadow = mapShadow(node.effects as Array<{ type: string; [k: string]: unknown }> | undefined)

  const textOptions: PptxGenJS.TextPropsOptions = {
    x: pxToInch(absX - ctx.originX),
    y: pxToInch(absY - ctx.originY),
    w: w > 0 ? pxToInch(w) : undefined,
    h: h > 0 ? pxToInch(h) : undefined,
    fill: fill ?? undefined,
    shadow: shadow ?? undefined,
    align: mapTextAlign(node.textAlign),
    valign: mapTextValign(node.textAlignVertical),
    rotate: node.rotation ?? undefined,
    wrap: true,
    isTextBox: true,
  }

  if (node.lineHeight) {
    textOptions.lineSpacingMultiple = node.lineHeight
  }

  if (node.letterSpacing) {
    textOptions.charSpacing = node.letterSpacing
  }

  ctx.slide.addText(textProps, textOptions)
}

function buildTextSegmentOptions(
  textNode: PenNode & { type: 'text' },
  segment?: { fontFamily?: string; fontSize?: number; fontWeight?: number; fontStyle?: 'normal' | 'italic'; fill?: string; underline?: boolean; strikethrough?: boolean },
): PptxGenJS.TextPropsOptions {
  const fontSize = segment?.fontSize ?? textNode.fontSize ?? 16
  const fontFamily = segment?.fontFamily ?? textNode.fontFamily ?? 'Arial'
  const fontWeight = segment?.fontWeight ?? (typeof textNode.fontWeight === 'number' ? textNode.fontWeight : undefined)
  const fontStyle = segment?.fontStyle ?? textNode.fontStyle
  const underline = segment?.underline ?? textNode.underline
  const strikethrough = segment?.strikethrough ?? textNode.strikethrough

  // Determine text color
  let textColor = '000000'
  if (segment?.fill) {
    textColor = parseHexColor(segment.fill).color
  } else if (textNode.fill && textNode.fill.length > 0) {
    const firstFill = textNode.fill[0]
    if (firstFill.type === 'solid') {
      textColor = parseHexColor(firstFill.color).color
    }
  }

  const opts: PptxGenJS.TextPropsOptions = {
    fontSize,
    fontFace: fontFamily,
    color: textColor,
    bold: fontWeight !== undefined ? fontWeight >= 600 : false,
    italic: fontStyle === 'italic',
  }

  if (underline) {
    opts.underline = { style: 'sng' }
  }
  if (strikethrough) {
    opts.strike = 'sngStrike'
  }

  return opts
}

function mapTextAlign(align: string | undefined): PptxGenJS.HAlign | undefined {
  switch (align) {
    case 'left': return 'left'
    case 'center': return 'center'
    case 'right': return 'right'
    case 'justify': return 'justify'
    default: return undefined
  }
}

function mapTextValign(valign: string | undefined): PptxGenJS.VAlign | undefined {
  switch (valign) {
    case 'top': return 'top'
    case 'middle': return 'middle'
    case 'bottom': return 'bottom'
    default: return undefined
  }
}

function addImage(ctx: SlideContext, flat: FlatNode): void {
  const { node, absX, absY } = flat
  if (node.type !== 'image') return

  const w = getNodeWidth(node)
  const h = getNodeHeight(node)
  if (!node.src) return

  const imgOptions: PptxGenJS.ImageProps = {
    x: pxToInch(absX - ctx.originX),
    y: pxToInch(absY - ctx.originY),
    w: w > 0 ? pxToInch(w) : undefined,
    h: h > 0 ? pxToInch(h) : undefined,
    rotate: node.rotation ?? undefined,
    flipH: node.flipX ?? undefined,
    flipV: node.flipY ?? undefined,
  }

  if (node.src.startsWith('data:')) {
    imgOptions.data = node.src
  } else {
    imgOptions.path = node.src
  }

  const transparency = opacityToTransparency(node.opacity)
  if (transparency > 0) {
    imgOptions.transparency = transparency
  }

  ctx.slide.addImage(imgOptions)
}

function addPath(ctx: SlideContext, flat: FlatNode): void {
  const { node, absX, absY } = flat
  if (node.type !== 'path') return

  const w = getNodeWidth(node)
  const h = getNodeHeight(node)
  if (w <= 0 || h <= 0) return

  // SVG paths can't be directly mapped to PPTX shapes.
  // Rasterize the path as an SVG image embedded in the slide.
  const fill = node.fill?.[0]
  let fillColor = 'none'
  if (fill?.type === 'solid') {
    fillColor = fill.color
  }

  const strokeAttr = node.stroke
    ? `stroke="${node.stroke.fill?.[0]?.type === 'solid' ? (node.stroke.fill[0] as { color: string }).color : '#333'}" stroke-width="${typeof node.stroke.thickness === 'number' ? node.stroke.thickness : 1}"`
    : 'stroke="none"'

  const svgStr = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}"><path d="${node.d}" fill="${fillColor}" ${strokeAttr}/></svg>`

  const base64 = btoa(svgStr)
  const dataUri = `image/svg+xml;base64,${base64}`

  ctx.slide.addImage({
    data: dataUri,
    x: pxToInch(absX - ctx.originX),
    y: pxToInch(absY - ctx.originY),
    w: pxToInch(w),
    h: pxToInch(h),
    rotate: node.rotation ?? undefined,
  })
}

function addPolygon(ctx: SlideContext, flat: FlatNode): void {
  const { node, absX, absY } = flat
  if (node.type !== 'polygon') return

  const w = getNodeWidth(node)
  const h = getNodeHeight(node)
  if (w <= 0 || h <= 0) return

  const count = node.polygonCount ?? 3

  // Generate polygon SVG and embed as image
  const points: string[] = []
  for (let i = 0; i < count; i++) {
    const angle = (2 * Math.PI * i) / count - Math.PI / 2
    const px = w / 2 + (w / 2) * Math.cos(angle)
    const py = h / 2 + (h / 2) * Math.sin(angle)
    points.push(`${px},${py}`)
  }

  const fill = node.fill?.[0]
  let fillColor = 'none'
  if (fill?.type === 'solid') {
    fillColor = fill.color
  }

  const strokeAttr = node.stroke
    ? `stroke="${node.stroke.fill?.[0]?.type === 'solid' ? (node.stroke.fill[0] as { color: string }).color : '#333'}" stroke-width="${typeof node.stroke.thickness === 'number' ? node.stroke.thickness : 1}"`
    : 'stroke="none"'

  const svgStr = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}"><polygon points="${points.join(' ')}" fill="${fillColor}" ${strokeAttr}/></svg>`

  const base64 = btoa(svgStr)
  const dataUri = `image/svg+xml;base64,${base64}`

  ctx.slide.addImage({
    data: dataUri,
    x: pxToInch(absX - ctx.originX),
    y: pxToInch(absY - ctx.originY),
    w: pxToInch(w),
    h: pxToInch(h),
    rotate: node.rotation ?? undefined,
  })
}

// ---------------------------------------------------------------------------
// Public API — map a single node
// ---------------------------------------------------------------------------

export function mapNodeToSlide(ctx: SlideContext, flat: FlatNode): void {
  const { node } = flat

  switch (node.type) {
    case 'frame':
    case 'group':
    case 'rectangle':
      addRectangle(ctx, flat)
      break
    case 'ellipse':
      addEllipse(ctx, flat)
      break
    case 'line':
      addLine(ctx, flat)
      break
    case 'text':
      addText(ctx, flat)
      break
    case 'image':
      addImage(ctx, flat)
      break
    case 'path':
      addPath(ctx, flat)
      break
    case 'polygon':
      addPolygon(ctx, flat)
      break
    case 'ref':
      // Ref nodes should be resolved before export (not handled here)
      break
  }
}
