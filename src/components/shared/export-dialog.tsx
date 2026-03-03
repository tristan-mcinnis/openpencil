import { useState, useRef, useEffect } from 'react'
import { X } from 'lucide-react'
import { cn } from '@/lib/utils'
import { Button } from '@/components/ui/button'
import { useCanvasStore } from '@/stores/canvas-store'
import { useDocumentStore } from '@/stores/document-store'
import { exportToPNG, exportToSVG } from '@/utils/export'
import { exportToPptx } from '@/utils/pptx-export'
import { syncCanvasPositionsToStore } from '@/canvas/use-canvas-sync'

type ExportFormat = 'png' | 'svg' | 'pptx'

interface ExportDialogProps {
  open: boolean
  onClose: () => void
}

export default function ExportDialog({ open, onClose }: ExportDialogProps) {
  const [format, setFormat] = useState<ExportFormat>('png')
  const [scale, setScale] = useState(2)
  const [selectedOnly, setSelectedOnly] = useState(false)
  const [pptxScope, setPptxScope] = useState<'all' | 'current'>('all')
  const [exporting, setExporting] = useState(false)
  const fabricCanvas = useCanvasStore((s) => s.fabricCanvas)
  const hasSelection = useCanvasStore(
    (s) => s.selection.selectedIds.length > 0,
  )
  const hasMultiplePages = useDocumentStore(
    (s) => (s.document.pages?.length ?? 0) > 1,
  )
  const dialogRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (!open) return
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    document.addEventListener('keydown', handleKeyDown)
    return () => document.removeEventListener('keydown', handleKeyDown)
  }, [open, onClose])

  if (!open) return null

  const handleExport = async () => {
    if (format === 'pptx') {
      setExporting(true)
      try {
        syncCanvasPositionsToStore()
        const doc = useDocumentStore.getState().document
        const activePageId = useCanvasStore.getState().activePageId
        await exportToPptx(doc, activePageId, { scope: pptxScope })
      } finally {
        setExporting(false)
      }
      onClose()
      return
    }

    if (!fabricCanvas) return
    if (format === 'png') {
      exportToPNG(fabricCanvas, {
        multiplier: scale,
        selectedOnly,
      })
    } else {
      exportToSVG(fabricCanvas, { selectedOnly })
    }
    onClose()
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center">
      <div className="absolute inset-0 bg-background/80" onClick={onClose} />
      <div
        ref={dialogRef}
        className="relative bg-card rounded-lg border border-border p-4 w-72 shadow-xl"
      >
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-sm font-medium text-foreground">Export</h3>
          <Button variant="ghost" size="icon-sm" onClick={onClose}>
            <X size={14} />
          </Button>
        </div>

        {/* Format */}
        <div className="mb-3">
          <label className="text-xs text-muted-foreground block mb-1">Format</label>
          <div className="flex gap-2">
            {(['png', 'svg', 'pptx'] as const).map((f) => (
              <button
                key={f}
                type="button"
                onClick={() => setFormat(f)}
                className={cn(
                  'flex-1 text-xs py-1.5 rounded transition-colors',
                  format === f
                    ? 'bg-primary text-primary-foreground'
                    : 'bg-secondary text-secondary-foreground hover:bg-secondary/80',
                )}
              >
                {f.toUpperCase()}
              </button>
            ))}
          </div>
        </div>

        {/* Scale (PNG only) */}
        {format === 'png' && (
          <div className="mb-3">
            <label className="text-xs text-muted-foreground block mb-1">Scale</label>
            <div className="flex gap-2">
              {[1, 2, 3].map((s) => (
                <button
                  key={s}
                  type="button"
                  onClick={() => setScale(s)}
                  className={cn(
                    'flex-1 text-xs py-1.5 rounded transition-colors',
                    scale === s
                      ? 'bg-primary text-primary-foreground'
                      : 'bg-secondary text-secondary-foreground hover:bg-secondary/80',
                  )}
                >
                  {s}x
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Page scope (PPTX only, when multiple pages exist) */}
        {format === 'pptx' && hasMultiplePages && (
          <div className="mb-3">
            <label className="text-xs text-muted-foreground block mb-1">Pages</label>
            <div className="flex gap-2">
              {([['all', 'All pages'], ['current', 'Current page']] as const).map(([value, label]) => (
                <button
                  key={value}
                  type="button"
                  onClick={() => setPptxScope(value)}
                  className={cn(
                    'flex-1 text-xs py-1.5 rounded transition-colors',
                    pptxScope === value
                      ? 'bg-primary text-primary-foreground'
                      : 'bg-secondary text-secondary-foreground hover:bg-secondary/80',
                  )}
                >
                  {label}
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Selected only (raster/SVG only) */}
        {format !== 'pptx' && hasSelection && (
          <label className="flex items-center gap-2 mb-4 cursor-pointer">
            <input
              type="checkbox"
              checked={selectedOnly}
              onChange={(e) => setSelectedOnly(e.target.checked)}
              className="rounded border-input bg-secondary text-primary focus:ring-ring focus:ring-offset-0"
            />
            <span className="text-xs text-foreground">Export selected only</span>
          </label>
        )}

        {/* Export button */}
        <Button
          onClick={handleExport}
          disabled={format === 'pptx' ? exporting : !fabricCanvas}
          className="w-full"
          size="sm"
        >
          {exporting ? 'Exporting...' : `Export ${format.toUpperCase()}`}
        </Button>
      </div>
    </div>
  )
}
