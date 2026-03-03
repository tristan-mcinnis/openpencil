import { useState } from 'react'
import { useCanvasStore } from '@/stores/canvas-store'
import { useDocumentStore } from '@/stores/document-store'
import { exportLayerToRaster, type RasterFormat } from '@/utils/export'
import { exportToPptx } from '@/utils/pptx-export'
import { syncCanvasPositionsToStore } from '@/canvas/use-canvas-sync'
import SectionHeader from '@/components/shared/section-header'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Button } from '@/components/ui/button'

const SCALE_OPTIONS = [
  { value: '1', label: '1x' },
  { value: '2', label: '2x' },
  { value: '3', label: '3x' },
]

const FORMAT_OPTIONS = [
  { value: 'png', label: 'PNG' },
  { value: 'jpeg', label: 'JPEG' },
  { value: 'webp', label: 'WEBP' },
  { value: 'pptx', label: 'PPTX' },
]

interface ExportSectionProps {
  nodeId: string
  nodeName: string
}

export default function ExportSection({ nodeId, nodeName }: ExportSectionProps) {
  const [scale, setScale] = useState('1')
  const [format, setFormat] = useState('png')
  const [exporting, setExporting] = useState(false)
  const fabricCanvas = useCanvasStore((s) => s.fabricCanvas)

  const handleExport = async () => {
    if (format === 'pptx') {
      setExporting(true)
      try {
        syncCanvasPositionsToStore()
        const doc = useDocumentStore.getState().document
        const activePageId = useCanvasStore.getState().activePageId
        await exportToPptx(doc, activePageId, {
          scope: 'current',
          fileName: nodeName,
        })
      } finally {
        setExporting(false)
      }
      return
    }

    if (!fabricCanvas) return

    // Collect all descendant IDs for this node
    const { getFlatNodes, isDescendantOf } = useDocumentStore.getState()
    const allNodes = getFlatNodes()
    const descendantIds = new Set<string>()
    for (const n of allNodes) {
      if (n.id !== nodeId && isDescendantOf(n.id, nodeId)) {
        descendantIds.add(n.id)
      }
    }

    const ext = format === 'jpeg' ? 'jpg' : format
    exportLayerToRaster(fabricCanvas, nodeId, descendantIds, {
      format: format as RasterFormat,
      multiplier: Number(scale),
      filename: `${nodeName}.${ext}`,
    })
  }

  const isRaster = format !== 'pptx'

  return (
    <div className="space-y-1.5">
      <SectionHeader title="Export" />
      <div className="flex gap-1.5">
        {isRaster && (
          <Select value={scale} onValueChange={setScale}>
            <SelectTrigger className="flex-1 h-6 text-[11px]">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              {SCALE_OPTIONS.map((opt) => (
                <SelectItem key={opt.value} value={opt.value}>
                  {opt.label}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        )}
        <Select value={format} onValueChange={setFormat}>
          <SelectTrigger className="flex-1 h-6 text-[11px]">
            <SelectValue />
          </SelectTrigger>
          <SelectContent>
            {FORMAT_OPTIONS.map((opt) => (
              <SelectItem key={opt.value} value={opt.value}>
                {opt.label}
              </SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>
      <Button
        variant="outline"
        size="sm"
        className="w-full text-xs"
        onClick={handleExport}
        disabled={exporting}
      >
        {exporting ? 'Exporting...' : `Export ${format === 'pptx' ? 'PowerPoint' : 'layer'}`}
      </Button>
    </div>
  )
}
