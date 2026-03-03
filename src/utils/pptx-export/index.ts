/**
 * PPTX export orchestrator.
 *
 * Traverses the PenDocument tree, creates slides from top-level frames,
 * and maps all child nodes to native PowerPoint objects.
 */

import PptxGenJS from 'pptxgenjs'
import type { PenDocument, PenNode, PenPage } from '@/types/pen'
import type { VariableDefinition } from '@/types/variables'
import { resolveNodeForCanvas, getDefaultTheme } from '@/variables/resolve-variables'
import { frameDimensionsToSlide } from './coordinate-utils'
import { parseHexColor } from './color-utils'
import { flattenNodeTree, mapNodeToSlide } from './node-mapper'

// ---------------------------------------------------------------------------
// Export options
// ---------------------------------------------------------------------------

export interface PptxExportOptions {
  /** Export all pages or only current page. Default: 'all' */
  scope?: 'all' | 'current'
  /** File name (without extension). Default: document name or 'design' */
  fileName?: string
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function resolveTreeVariables(
  nodes: PenNode[],
  variables: Record<string, VariableDefinition>,
  theme: Record<string, string>,
): PenNode[] {
  return nodes.map((node) => {
    const resolved = resolveNodeForCanvas(node, variables, theme)
    if ('children' in resolved && resolved.children) {
      return {
        ...resolved,
        children: resolveTreeVariables(resolved.children, variables, theme),
      } as PenNode
    }
    return resolved
  })
}

function getFrameDimensions(node: PenNode): { w: number; h: number } {
  const w = ('width' in node && typeof node.width === 'number') ? node.width : 1200
  const h = ('height' in node && typeof node.height === 'number') ? node.height : 800
  return { w, h }
}

function getFrameBackground(node: PenNode): string | undefined {
  if ('fill' in node && Array.isArray((node as { fill?: unknown }).fill)) {
    const fills = (node as { fill: Array<{ type: string; color?: string }> }).fill
    const solid = fills.find((f) => f.type === 'solid')
    if (solid?.color) return solid.color
  }
  return undefined
}

// ---------------------------------------------------------------------------
// Build slides from a page's children
// ---------------------------------------------------------------------------

function addPageSlides(
  pptx: PptxGenJS,
  children: PenNode[],
): void {
  // Each top-level frame becomes a slide
  const topLevelFrames = children.filter(
    (n) => (n.type === 'frame' || n.type === 'group') && n.visible !== false,
  )

  // If there are no frames, create a single slide with all nodes
  if (topLevelFrames.length === 0 && children.length > 0) {
    const slide = addSlideWithLayout(pptx, 1200, 800)
    const flatNodes = flattenNodeTree(children, 0, 0)
    const ctx = { slide, pptx, originX: 0, originY: 0 }
    for (const flat of flatNodes) {
      mapNodeToSlide(ctx, flat)
    }
    return
  }

  for (const frame of topLevelFrames) {
    const { w, h } = getFrameDimensions(frame)
    const slide = addSlideWithLayout(pptx, w, h)

    // Set slide background from frame fill
    const bgColor = getFrameBackground(frame)
    if (bgColor) {
      const { color } = parseHexColor(bgColor)
      slide.background = { color }
    }

    // Flatten the frame's children (not the frame itself) and map them
    const frameChildren = ('children' in frame && frame.children) ? frame.children : []
    const flatNodes = flattenNodeTree(frameChildren, 0, 0)

    const ctx = {
      slide,
      pptx,
      originX: 0,
      originY: 0,
    }

    for (const flat of flatNodes) {
      mapNodeToSlide(ctx, flat)
    }
  }
}

/** Track layout counter for unique layout names. */
let layoutCounter = 0

function addSlideWithLayout(
  pptx: PptxGenJS,
  widthPx: number,
  heightPx: number,
): PptxGenJS.Slide {
  const { width, height } = frameDimensionsToSlide(widthPx, heightPx)
  const layoutName = `Custom_${++layoutCounter}`
  pptx.defineLayout({ name: layoutName, width, height })
  pptx.layout = layoutName
  return pptx.addSlide()
}

// ---------------------------------------------------------------------------
// Main export function
// ---------------------------------------------------------------------------

/**
 * Export a PenDocument to a .pptx file.
 *
 * Each top-level frame becomes a PPTX slide. All child nodes are mapped
 * to native PowerPoint shapes, text boxes, and images.
 */
export async function exportToPptx(
  document: PenDocument,
  activePageId: string | null,
  options?: PptxExportOptions,
): Promise<void> {
  const pptx = new PptxGenJS()
  layoutCounter = 0

  // Metadata
  pptx.author = 'OpenPencil'
  pptx.title = document.name ?? 'Untitled'

  // Resolve variables
  const variables = document.variables ?? {}
  const theme = getDefaultTheme(document.themes)

  const scope = options?.scope ?? 'all'

  // Determine which pages to export
  let pagesToExport: PenPage[]
  if (document.pages && document.pages.length > 0) {
    if (scope === 'current' && activePageId) {
      const currentPage = document.pages.find((p) => p.id === activePageId)
      pagesToExport = currentPage ? [currentPage] : [document.pages[0]]
    } else {
      pagesToExport = document.pages
    }
  } else {
    // Legacy single-page format
    pagesToExport = [{ id: 'default', name: 'Page 1', children: document.children }]
  }

  // Build slides from each page
  for (const page of pagesToExport) {
    const resolvedChildren = resolveTreeVariables(page.children, variables, theme)
    addPageSlides(pptx, resolvedChildren)
  }

  // Generate and download
  const fileName = options?.fileName ?? document.name ?? 'design'
  await pptx.writeFile({ fileName: `${fileName}.pptx` })
}

/**
 * Export to PPTX and return as a Blob (for programmatic use).
 */
export async function exportToPptxBlob(
  document: PenDocument,
  activePageId: string | null,
  options?: PptxExportOptions,
): Promise<Blob> {
  const pptx = new PptxGenJS()
  layoutCounter = 0

  pptx.author = 'OpenPencil'
  pptx.title = document.name ?? 'Untitled'

  const variables = document.variables ?? {}
  const theme = getDefaultTheme(document.themes)

  const scope = options?.scope ?? 'all'

  let pagesToExport: PenPage[]
  if (document.pages && document.pages.length > 0) {
    if (scope === 'current' && activePageId) {
      const currentPage = document.pages.find((p) => p.id === activePageId)
      pagesToExport = currentPage ? [currentPage] : [document.pages[0]]
    } else {
      pagesToExport = document.pages
    }
  } else {
    pagesToExport = [{ id: 'default', name: 'Page 1', children: document.children }]
  }

  for (const page of pagesToExport) {
    const resolvedChildren = resolveTreeVariables(page.children, variables, theme)
    addPageSlides(pptx, resolvedChildren)
  }

  const result = await pptx.write({ outputType: 'blob' })
  return result as Blob
}
