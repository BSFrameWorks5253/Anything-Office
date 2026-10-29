import React, { useState, useRef, useEffect } from 'react';
import type {
  AnnotationTool,
  PdfAnnotation,
  AnnotationStyleConfig,
  Point,
} from '../../types/annotations';
import {
  MessageSquare,
  X,
  Check,
  ShieldCheck,
} from 'lucide-react';

interface AnnotationOverlayProps {
  pageNumber: number;
  zoom: number;
  rotation: number;
  activeTool: AnnotationTool;
  styleConfig: AnnotationStyleConfig;
  annotations: PdfAnnotation[];
  onAddAnnotation: (annotation: PdfAnnotation) => void;
  onUpdateAnnotation: (id: string, updates: Partial<PdfAnnotation>) => void;
  onDeleteAnnotation: (id: string) => void;
  selectedAnnotationId: string | null;
  onSelectAnnotation: (id: string | null) => void;
  activeSignatureToPlace: {
    dataUrl: string;
    signerName: string;
    includeDate: boolean;
  } | null;
  onSignaturePlaced: () => void;
  onOpenCommentsSidebar: () => void;
}

export const AnnotationOverlay: React.FC<AnnotationOverlayProps> = ({
  pageNumber,
  zoom,
  activeTool,
  styleConfig,
  annotations,
  onAddAnnotation,
  onUpdateAnnotation,
  onDeleteAnnotation,
  selectedAnnotationId,
  onSelectAnnotation,
  activeSignatureToPlace,
  onSignaturePlaced,
  onOpenCommentsSidebar,
}) => {
  const containerRef = useRef<HTMLDivElement | null>(null);

  // Active drawing states
  const [currentInkPoints, setCurrentInkPoints] = useState<Point[]>([]);
  const [isDrawingInk, setIsDrawingInk] = useState<boolean>(false);

  // Drag-to-highlight state
  const [highlightStart, setHighlightStart] = useState<Point | null>(null);
  const [highlightCurrent, setHighlightCurrent] = useState<Point | null>(null);

  // Inline text creator
  const [inlineTextInput, setInlineTextInput] = useState<{
    x: number;
    y: number;
    text: string;
  } | null>(null);

  // Dragging existing annotation
  const [draggingAnnotationId, setDraggingAnnotationId] = useState<string | null>(null);
  const [dragOffset, setDragOffset] = useState<Point>({ x: 0, y: 0 });

  // Only filter annotations for this page
  const pageAnnotations = annotations.filter((a) => a.pageNumber === pageNumber);

  // Helper to convert mouse client coordinates to base A4 coordinate units (794x1123)
  const getPageCoordinates = (e: React.MouseEvent): Point => {
    if (!containerRef.current) return { x: 0, y: 0 };
    const rect = containerRef.current.getBoundingClientRect();
    const scale = zoom / 100;
    const x = (e.clientX - rect.left) / scale;
    const y = (e.clientY - rect.top) / scale;
    return { x, y };
  };

  // Global Delete / Backspace key listener
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (
        (e.key === 'Delete' || e.key === 'Backspace') &&
        selectedAnnotationId &&
        !inlineTextInput &&
        document.activeElement?.tagName !== 'INPUT' &&
        document.activeElement?.tagName !== 'TEXTAREA'
      ) {
        onDeleteAnnotation(selectedAnnotationId);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [selectedAnnotationId, inlineTextInput, onDeleteAnnotation]);

  // Handle overlay click & drag start
  const handleMouseDown = (e: React.MouseEvent) => {
    if (e.button !== 0) return; // only left click
    const coords = getPageCoordinates(e);

    // 1. Text tool -> spawn typewriter input
    if (activeTool === 'text') {
      setInlineTextInput({
        x: coords.x,
        y: coords.y,
        text: '',
      });
      return;
    }

    // 2. Sticky Note -> spawn comment pin
    if (activeTool === 'sticky') {
      const newSticky: PdfAnnotation = {
        id: `sticky-${Date.now()}`,
        type: 'sticky',
        pageNumber,
        x: coords.x,
        y: coords.y,
        width: 32,
        height: 32,
        color: '#f59e0b',
        bgColor: '#fef3c7',
        content: 'New review note...',
        author: 'Reviewer',
        timestamp: Date.now(),
        replies: [],
      };
      onAddAnnotation(newSticky);
      onSelectAnnotation(newSticky.id);
      onOpenCommentsSidebar();
      return;
    }

    // 3. Highlight / Underline / Strikethrough -> start drag box
    if (activeTool === 'highlight' || activeTool === 'underline' || activeTool === 'strikethrough') {
      setHighlightStart(coords);
      setHighlightCurrent(coords);
      return;
    }

    // 4. Freehand Ink -> start stroke
    if (activeTool === 'ink') {
      setIsDrawingInk(true);
      setCurrentInkPoints([coords]);
      return;
    }

    // 5. Place Digital Signature Stamp
    if (activeTool === 'signature' && activeSignatureToPlace) {
      const newSignature: PdfAnnotation = {
        id: `sig-${Date.now()}`,
        type: 'signature',
        pageNumber,
        x: coords.x - 110,
        y: coords.y - 45,
        width: 220,
        height: 90,
        color: '#1e3a8a',
        signatureDataUrl: activeSignatureToPlace.dataUrl,
        signerName: activeSignatureToPlace.signerName,
        timestamp: Date.now(),
      };
      onAddAnnotation(newSignature);
      onSelectAnnotation(newSignature.id);
      onSignaturePlaced();
      return;
    }

    // Deselect if clicking blank area with select tool
    if (activeTool === 'select') {
      onSelectAnnotation(null);
    }
  };

  // Mouse move handler
  const handleMouseMove = (e: React.MouseEvent) => {
    const coords = getPageCoordinates(e);

    // Ink drawing
    if (isDrawingInk && activeTool === 'ink') {
      setCurrentInkPoints((prev) => [...prev, coords]);
      return;
    }

    // Drag-to-highlight
    if (highlightStart) {
      setHighlightCurrent(coords);
      return;
    }

    // Dragging an existing annotation
    if (draggingAnnotationId) {
      const currentAnn = annotations.find((a) => a.id === draggingAnnotationId);
      if (currentAnn) {
        onUpdateAnnotation(draggingAnnotationId, {
          x: Math.max(0, coords.x - dragOffset.x),
          y: Math.max(0, coords.y - dragOffset.y),
        });
      }
    }
  };

  // Mouse up handler
  const handleMouseUp = () => {
    // Commit ink stroke
    if (isDrawingInk && currentInkPoints.length > 1) {
      const minX = Math.min(...currentInkPoints.map((p) => p.x));
      const minY = Math.min(...currentInkPoints.map((p) => p.y));
      const maxX = Math.max(...currentInkPoints.map((p) => p.x));
      const maxY = Math.max(...currentInkPoints.map((p) => p.y));

      const newInk: PdfAnnotation = {
        id: `ink-${Date.now()}`,
        type: 'ink',
        pageNumber,
        x: minX,
        y: minY,
        width: Math.max(20, maxX - minX),
        height: Math.max(20, maxY - minY),
        color: styleConfig.inkColor,
        strokeWidth: styleConfig.inkWidth,
        strokes: [
          {
            points: currentInkPoints,
            color: styleConfig.inkColor,
            width: styleConfig.inkWidth,
          },
        ],
        timestamp: Date.now(),
      };

      onAddAnnotation(newInk);
      onSelectAnnotation(newInk.id);
      setIsDrawingInk(false);
      setCurrentInkPoints([]);
      return;
    }
    setIsDrawingInk(false);

    // Commit highlight / markup
    if (highlightStart && highlightCurrent) {
      const x = Math.min(highlightStart.x, highlightCurrent.x);
      const y = Math.min(highlightStart.y, highlightCurrent.y);
      const width = Math.abs(highlightCurrent.x - highlightStart.x);
      const height = Math.abs(highlightCurrent.y - highlightStart.y);

      if (width > 8 || height > 8) {
        const type =
          activeTool === 'underline'
            ? 'underline'
            : activeTool === 'strikethrough'
            ? 'strikethrough'
            : 'highlight';

        const newHighlight: PdfAnnotation = {
          id: `markup-${Date.now()}`,
          type,
          pageNumber,
          x,
          y,
          width: Math.max(30, width),
          height: Math.max(14, height),
          color:
            type === 'highlight'
              ? styleConfig.highlightColor
              : type === 'underline'
              ? '#2563eb'
              : '#dc2626',
          opacity: 0.45,
          timestamp: Date.now(),
        };

        onAddAnnotation(newHighlight);
        onSelectAnnotation(newHighlight.id);
      }
      setHighlightStart(null);
      setHighlightCurrent(null);
    }

    // Stop dragging existing annotation
    if (draggingAnnotationId) {
      setDraggingAnnotationId(null);
    }
  };

  // Commit typewriter inline text
  const handleCommitInlineText = () => {
    if (inlineTextInput && inlineTextInput.text.trim()) {
      const newTextAnn: PdfAnnotation = {
        id: `text-${Date.now()}`,
        type: 'text',
        pageNumber,
        x: inlineTextInput.x,
        y: inlineTextInput.y,
        width: Math.max(120, inlineTextInput.text.length * 8),
        height: styleConfig.fontSize + 16,
        content: inlineTextInput.text,
        color: styleConfig.textColor,
        bgColor: styleConfig.bgColor,
        fontSize: styleConfig.fontSize,
        author: 'Reviewer',
        timestamp: Date.now(),
        replies: [],
      };
      onAddAnnotation(newTextAnn);
      onSelectAnnotation(newTextAnn.id);
    }
    setInlineTextInput(null);
  };

  // Convert points to smooth SVG path string
  const renderSvgPath = (points: Point[]): string => {
    if (points.length < 2) return '';
    let d = `M ${points[0].x} ${points[0].y}`;
    for (let i = 1; i < points.length; i++) {
      d += ` L ${points[i].x} ${points[i].y}`;
    }
    return d;
  };

  const scale = zoom / 100;

  return (
    <div
      ref={containerRef}
      onMouseDown={handleMouseDown}
      onMouseMove={handleMouseMove}
      onMouseUp={handleMouseUp}
      className={`absolute inset-0 z-20 ${
        activeTool === 'select'
          ? 'cursor-default'
          : activeTool === 'hand'
          ? 'cursor-grab active:cursor-grabbing pointer-events-none'
          : activeTool === 'text'
          ? 'cursor-text'
          : activeTool === 'signature'
          ? 'cursor-copy'
          : 'cursor-crosshair'
      }`}
      style={{
        width: '100%',
        height: '100%',
        pointerEvents: activeTool === 'hand' ? 'none' : 'auto',
      }}
    >
      {/* SVG Layer for Ink Drawings & Underlines */}
      <svg
        className="absolute inset-0 w-full h-full pointer-events-none"
        style={{
          width: '100%',
          height: '100%',
        }}
      >
        {/* Render Saved Ink Annotations */}
        {pageAnnotations
          .filter((a) => a.type === 'ink' && a.strokes)
          .map((a) => (
            <g key={a.id} className={selectedAnnotationId === a.id ? 'filter drop-shadow-md' : ''}>
              {a.strokes?.map((stroke, sIdx) => (
                <path
                  key={sIdx}
                  d={renderSvgPath(stroke.points)}
                  stroke={stroke.color}
                  strokeWidth={(stroke.width || 3) * scale}
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  fill="none"
                  transform={`scale(${scale})`}
                />
              ))}
            </g>
          ))}

        {/* Live Ink Stroke While Drawing */}
        {isDrawingInk && currentInkPoints.length > 1 && (
          <path
            d={renderSvgPath(currentInkPoints)}
            stroke={styleConfig.inkColor}
            strokeWidth={styleConfig.inkWidth * scale}
            strokeLinecap="round"
            strokeLinejoin="round"
            fill="none"
            transform={`scale(${scale})`}
          />
        )}
      </svg>

      {/* Render All Page Annotations */}
      {pageAnnotations.map((ann) => {
        const isSelected = ann.id === selectedAnnotationId;

        // Position coordinates scaled to zoom
        const style: React.CSSProperties = {
          position: 'absolute',
          left: `${ann.x * scale}px`,
          top: `${ann.y * scale}px`,
          width: `${ann.width * scale}px`,
          height: `${ann.height * scale}px`,
        };

        // 1. Text Annotation
        if (ann.type === 'text') {
          return (
            <div
              key={ann.id}
              style={style}
              onClick={(e) => {
                e.stopPropagation();
                onSelectAnnotation(ann.id);
              }}
              onMouseDown={(e) => {
                if (activeTool === 'select') {
                  e.stopPropagation();
                  onSelectAnnotation(ann.id);
                  setDraggingAnnotationId(ann.id);
                  const coords = getPageCoordinates(e);
                  setDragOffset({ x: coords.x - ann.x, y: coords.y - ann.y });
                }
              }}
              className={`p-1.5 rounded transition-shadow flex items-center ${
                activeTool === 'select' ? 'cursor-move' : 'cursor-pointer'
              } ${
                isSelected
                  ? 'ring-2 ring-blue-500 bg-blue-50/20 shadow-md'
                  : 'hover:ring-1 hover:ring-blue-400'
              }`}
            >
              <div
                style={{
                  color: ann.color,
                  fontSize: `${(ann.fontSize || 13) * scale}px`,
                  backgroundColor: ann.bgColor || 'transparent',
                }}
                className="font-sans font-medium whitespace-pre-wrap select-text leading-tight px-1 rounded"
              >
                {ann.content}
              </div>
            </div>
          );
        }

        // 2. Highlight Annotation
        if (ann.type === 'highlight') {
          return (
            <div
              key={ann.id}
              style={{
                ...style,
                backgroundColor: ann.color || '#fef08a',
                opacity: 0.5,
                mixBlendMode: 'multiply',
              }}
              onClick={(e) => {
                e.stopPropagation();
                onSelectAnnotation(ann.id);
              }}
              className={`rounded-xs transition-all ${
                isSelected ? 'ring-2 ring-amber-600 ring-offset-1' : 'hover:opacity-60'
              }`}
            />
          );
        }

        // 3. Underline Annotation
        if (ann.type === 'underline') {
          return (
            <div
              key={ann.id}
              style={{
                ...style,
                borderBottom: `2.5px solid ${ann.color || '#2563eb'}`,
              }}
              onClick={(e) => {
                e.stopPropagation();
                onSelectAnnotation(ann.id);
              }}
              className={isSelected ? 'ring-1 ring-blue-500' : ''}
            />
          );
        }

        // 4. Strikethrough Annotation
        if (ann.type === 'strikethrough') {
          return (
            <div
              key={ann.id}
              style={style}
              onClick={(e) => {
                e.stopPropagation();
                onSelectAnnotation(ann.id);
              }}
              className="flex items-center"
            >
              <div
                className="w-full"
                style={{
                  height: '2px',
                  backgroundColor: ann.color || '#dc2626',
                }}
              />
            </div>
          );
        }

        // 5. Sticky Note (Comment Pin)
        if (ann.type === 'sticky') {
          return (
            <div
              key={ann.id}
              style={{
                position: 'absolute',
                left: `${ann.x * scale}px`,
                top: `${ann.y * scale}px`,
              }}
              onClick={(e) => {
                e.stopPropagation();
                onSelectAnnotation(ann.id);
                onOpenCommentsSidebar();
              }}
              onMouseDown={(e) => {
                if (activeTool === 'select') {
                  e.stopPropagation();
                  onSelectAnnotation(ann.id);
                  setDraggingAnnotationId(ann.id);
                  const coords = getPageCoordinates(e);
                  setDragOffset({ x: coords.x - ann.x, y: coords.y - ann.y });
                }
              }}
              title={`Comment by ${ann.author}: "${ann.content}"`}
              className={`w-7 h-7 rounded-full shadow-lg flex items-center justify-center cursor-pointer transition-transform hover:scale-110 active:scale-95 ${
                isSelected
                  ? 'bg-amber-500 text-white ring-4 ring-amber-400/40 scale-110'
                  : 'bg-amber-400 text-amber-950 border border-amber-500'
              }`}
            >
              <MessageSquare className="w-3.5 h-3.5 fill-current" />
            </div>
          );
        }

        // 6. Digital Signature Stamp
        if (ann.type === 'signature') {
          return (
            <div
              key={ann.id}
              style={style}
              onClick={(e) => {
                e.stopPropagation();
                onSelectAnnotation(ann.id);
              }}
              onMouseDown={(e) => {
                if (activeTool === 'select') {
                  e.stopPropagation();
                  onSelectAnnotation(ann.id);
                  setDraggingAnnotationId(ann.id);
                  const coords = getPageCoordinates(e);
                  setDragOffset({ x: coords.x - ann.x, y: coords.y - ann.y });
                }
              }}
              className={`rounded-xl border p-2 flex flex-col justify-between transition-all backdrop-blur-xs select-none ${
                activeTool === 'select' ? 'cursor-move' : 'cursor-pointer'
              } ${
                isSelected
                  ? 'border-blue-600 bg-blue-50/60 dark:bg-blue-900/30 ring-2 ring-blue-500 shadow-xl'
                  : 'border-blue-300 dark:border-blue-800 bg-white/80 dark:bg-slate-900/80 shadow-md'
              }`}
            >
              {/* Signature Image */}
              <div className="flex-1 flex items-center justify-center overflow-hidden">
                {ann.signatureDataUrl ? (
                  <img
                    src={ann.signatureDataUrl}
                    alt="Digital Signature"
                    className="max-h-full max-w-full object-contain pointer-events-none"
                  />
                ) : (
                  <span className="font-serif italic text-blue-900 text-lg">
                    {ann.signerName || 'Signature'}
                  </span>
                )}
              </div>

              {/* Verified Audit Stamp */}
              <div className="pt-1 border-t border-blue-200 dark:border-blue-800/60 flex items-center justify-between text-[9px] text-blue-800 dark:text-blue-300 font-mono">
                <span className="flex items-center gap-1 font-bold">
                  <ShieldCheck className="w-3 h-3 text-emerald-600" />
                  <span>{ann.signerName}</span>
                </span>
                <span>{new Date(ann.timestamp).toLocaleDateString()}</span>
              </div>
            </div>
          );
        }

        return null;
      })}

      {/* Dragging Highlight Rectangle Preview */}
      {highlightStart && highlightCurrent && (
        <div
          style={{
            position: 'absolute',
            left: `${Math.min(highlightStart.x, highlightCurrent.x) * scale}px`,
            top: `${Math.min(highlightStart.y, highlightCurrent.y) * scale}px`,
            width: `${Math.abs(highlightCurrent.x - highlightStart.x) * scale}px`,
            height: `${Math.abs(highlightCurrent.y - highlightStart.y) * scale}px`,
            backgroundColor:
              activeTool === 'highlight' ? styleConfig.highlightColor : 'transparent',
            opacity: 0.5,
            border:
              activeTool === 'underline'
                ? '2px solid #2563eb'
                : activeTool === 'strikethrough'
                ? '2px solid #dc2626'
                : '1px dashed #f59e0b',
          }}
          className="pointer-events-none rounded-xs"
        />
      )}

      {/* Inline Typewriter Input Spawner */}
      {inlineTextInput && (
        <div
          style={{
            position: 'absolute',
            left: `${inlineTextInput.x * scale}px`,
            top: `${inlineTextInput.y * scale}px`,
          }}
          className="z-30 flex items-center gap-1 bg-white dark:bg-slate-900 shadow-2xl rounded-xl border border-blue-500 p-1.5 animate-fadeIn"
          onClick={(e) => e.stopPropagation()}
        >
          <input
            type="text"
            autoFocus
            value={inlineTextInput.text}
            onChange={(e) =>
              setInlineTextInput({ ...inlineTextInput, text: e.target.value })
            }
            onKeyDown={(e) => {
              if (e.key === 'Enter') handleCommitInlineText();
              if (e.key === 'Escape') setInlineTextInput(null);
            }}
            placeholder="Type note..."
            style={{
              color: styleConfig.textColor,
              fontSize: `${styleConfig.fontSize}px`,
            }}
            className="bg-transparent focus:outline-none px-2 py-0.5 min-w-[140px] font-sans"
          />
          <button
            onClick={handleCommitInlineText}
            className="p-1 rounded-lg bg-blue-600 text-white hover:bg-blue-700 transition-colors"
            title="Commit Note (Enter)"
          >
            <Check className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => setInlineTextInput(null)}
            className="p-1 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-400"
            title="Cancel (Esc)"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}
    </div>
  );
};
