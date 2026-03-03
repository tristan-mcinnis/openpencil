import { describe, it, expect } from 'vitest'
import { parseHexColor, opacityToTransparency, combinedTransparency } from '../pptx-export/color-utils'
import { pxToInch, frameDimensionsToSlide } from '../pptx-export/coordinate-utils'
import { flattenNodeTree } from '../pptx-export/node-mapper'
import type { PenNode } from '@/types/pen'

// ---------------------------------------------------------------------------
// Color utils
// ---------------------------------------------------------------------------

describe('parseHexColor', () => {
  it('parses 6-digit hex with #', () => {
    expect(parseHexColor('#FF0000')).toEqual({ color: 'FF0000', transparency: 0 })
  })

  it('parses 6-digit hex without #', () => {
    expect(parseHexColor('00FF00')).toEqual({ color: '00FF00', transparency: 0 })
  })

  it('parses 8-digit hex (with alpha)', () => {
    const result = parseHexColor('#FF000080')
    expect(result.color).toBe('FF0000')
    // 0x80 = 128, transparency ≈ 50%
    expect(result.transparency).toBeCloseTo(50, 0)
  })

  it('parses fully transparent color', () => {
    const result = parseHexColor('#FF000000')
    expect(result.color).toBe('FF0000')
    expect(result.transparency).toBe(100)
  })

  it('parses shorthand #RGB', () => {
    const result = parseHexColor('#F00')
    expect(result.color).toBe('FF0000')
    expect(result.transparency).toBe(0)
  })

  it('returns default for undefined', () => {
    expect(parseHexColor(undefined)).toEqual({ color: '000000', transparency: 0 })
  })

  it('returns default for invalid string', () => {
    expect(parseHexColor('not-a-color')).toEqual({ color: '000000', transparency: 0 })
  })
})

describe('opacityToTransparency', () => {
  it('converts 1.0 to 0%', () => {
    expect(opacityToTransparency(1)).toBe(0)
  })

  it('converts 0.0 to 100%', () => {
    expect(opacityToTransparency(0)).toBe(100)
  })

  it('converts 0.5 to 50%', () => {
    expect(opacityToTransparency(0.5)).toBe(50)
  })

  it('handles undefined as fully opaque', () => {
    expect(opacityToTransparency(undefined)).toBe(0)
  })

  it('clamps values', () => {
    expect(opacityToTransparency(2)).toBe(0)
    expect(opacityToTransparency(-1)).toBe(100)
  })
})

describe('combinedTransparency', () => {
  it('combines fill transparency with node opacity', () => {
    // 50% fill transparency + 50% node opacity
    // effective visible = (1-0.5) * (1-0.5) = 0.25
    // combined transparency = 75%
    expect(combinedTransparency(50, 0.5)).toBe(75)
  })

  it('returns fill transparency when node opacity is 1', () => {
    expect(combinedTransparency(30, 1)).toBe(30)
  })

  it('returns node opacity transparency when fill is fully opaque', () => {
    expect(combinedTransparency(0, 0.7)).toBe(30)
  })
})

// ---------------------------------------------------------------------------
// Coordinate utils
// ---------------------------------------------------------------------------

describe('pxToInch', () => {
  it('converts 96px to 1 inch', () => {
    expect(pxToInch(96)).toBe(1)
  })

  it('converts 0 to 0', () => {
    expect(pxToInch(0)).toBe(0)
  })

  it('converts 192px to 2 inches', () => {
    expect(pxToInch(192)).toBe(2)
  })
})

describe('frameDimensionsToSlide', () => {
  it('converts standard frame to inches', () => {
    const { width, height } = frameDimensionsToSlide(1200, 800)
    expect(width).toBeCloseTo(12.5, 1)
    expect(height).toBeCloseTo(8.33, 1)
  })

  it('clamps to minimum 1 inch', () => {
    const { width, height } = frameDimensionsToSlide(0, 0)
    expect(width).toBe(1)
    expect(height).toBe(1)
  })
})

// ---------------------------------------------------------------------------
// Node flattening
// ---------------------------------------------------------------------------

describe('flattenNodeTree', () => {
  it('flattens a simple node', () => {
    const nodes: PenNode[] = [
      { id: 'rect1', type: 'rectangle', x: 10, y: 20, width: 100, height: 50 } as PenNode,
    ]
    const flat = flattenNodeTree(nodes, 0, 0)
    expect(flat).toHaveLength(1)
    expect(flat[0].absX).toBe(10)
    expect(flat[0].absY).toBe(20)
  })

  it('flattens nested children with parent offset', () => {
    const nodes: PenNode[] = [
      {
        id: 'frame1',
        type: 'frame',
        x: 100,
        y: 200,
        width: 500,
        height: 400,
        children: [
          { id: 'rect1', type: 'rectangle', x: 10, y: 20, width: 50, height: 30 } as PenNode,
        ],
      } as PenNode,
    ]
    const flat = flattenNodeTree(nodes, 0, 0)
    expect(flat).toHaveLength(2) // frame + child rect
    expect(flat[0].node.id).toBe('frame1')
    expect(flat[0].absX).toBe(100)
    expect(flat[1].node.id).toBe('rect1')
    expect(flat[1].absX).toBe(110) // 100 + 10
    expect(flat[1].absY).toBe(220) // 200 + 20
  })

  it('skips invisible nodes', () => {
    const nodes: PenNode[] = [
      { id: 'rect1', type: 'rectangle', x: 0, y: 0, width: 100, height: 50, visible: false } as PenNode,
      { id: 'rect2', type: 'rectangle', x: 0, y: 0, width: 100, height: 50 } as PenNode,
    ]
    const flat = flattenNodeTree(nodes, 0, 0)
    expect(flat).toHaveLength(1)
    expect(flat[0].node.id).toBe('rect2')
  })

  it('handles deeply nested trees', () => {
    const nodes: PenNode[] = [
      {
        id: 'group1',
        type: 'group',
        x: 10,
        y: 10,
        width: 200,
        height: 200,
        children: [
          {
            id: 'frame1',
            type: 'frame',
            x: 5,
            y: 5,
            width: 100,
            height: 100,
            children: [
              { id: 'text1', type: 'text', x: 2, y: 3, content: 'hello' } as PenNode,
            ],
          } as PenNode,
        ],
      } as PenNode,
    ]
    const flat = flattenNodeTree(nodes, 0, 0)
    expect(flat).toHaveLength(3) // group + frame + text
    const textFlat = flat.find((f) => f.node.id === 'text1')!
    expect(textFlat.absX).toBe(17) // 10 + 5 + 2
    expect(textFlat.absY).toBe(18) // 10 + 5 + 3
  })
})
