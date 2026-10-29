import React, { useState, useRef } from 'react';
import type { PdfDocumentData, DocumentData, DocParagraph } from '../types/office';
import type {
  AnnotationTool,
  PdfAnnotation,
  AnnotationStyleConfig,
} from '../types/annotations';
import { PdfAnnotationToolbar } from './pdf/PdfAnnotationToolbar';
import { AnnotationOverlay } from './pdf/AnnotationOverlay';
import { SignatureModal } from './pdf/SignatureModal';
import { CommentsSidebar } from './pdf/CommentsSidebar';
import {
  ZoomIn,
  ZoomOut,
  RotateCw,
  Printer,
  Sidebar,
  ChevronLeft,
  ChevronRight,
  FolderOpen,
  Sparkles,
  Maximize2,
} from 'lucide-react';

interface PdfViewerProps {
  pdfData: PdfDocumentData;
  onUpdatePdf: (updates: Partial<PdfDocumentData>) => void;
  onReflowToDocs: (reflowedDoc: DocumentData) => void;
  darkMode: boolean;
}

const DEFAULT_ANNOTATIONS: PdfAnnotation[] = [
  {
    id: 'ann-sample-1',
    type: 'highlight',
    pageNumber: 1,
    x: 48,
    y: 104,
    width: 698,
    height: 52,
    color: '#fef08a',
    opacity: 0.5,
    timestamp: Date.now() - 3600000,
  },
  {
    id: 'ann-sample-2',
    type: 'sticky',
    pageNumber: 1,
    x: 720,
    y: 190,
    width: 32,
    height: 32,
    color: '#f59e0b',
    content: 'Architectural compliance verified: ECMA-376 and DAG engine integration meet specification.',
    author: 'Lead Architect',
    timestamp: Date.now() - 1800000,
    replies: [
      {
        id: 'rep-1',
        author: 'Systems Engineer',
        text: 'Verified. Sub-second boot times and zero-flash rendering confirmed.',
        timestamp: Date.now() - 900000,
      },
    ],
  },
  {
    id: 'ann-sample-3',
    type: 'signature',
    pageNumber: 1,
    x: 490,
    y: 920,
    width: 250,
    height: 96,
    color: '#1e3a8a',
    signerName: 'Alex Johnson • Lead Architect',
    timestamp: Date.now(),
  },
];

export const PdfViewer: React.FC<PdfViewerProps> = ({
  pdfData,
  onUpdatePdf,
  onReflowToDocs,
  darkMode,
}) => {
  // Navigation & Viewport State
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(true);
  const [commentsOpen, setCommentsOpen] = useState<boolean>(false);
  const [isReflowing, setIsReflowing] = useState<boolean>(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const currentPage = pdfData.currentPage || 1;
  const totalPages = Math.max(1, pdfData.totalPages || 3);
  const zoom = pdfData.zoom || 100;
  const rotation = pdfData.rotation || 0;

  // Annotations State
  const annotations = pdfData.annotations || DEFAULT_ANNOTATIONS;
  const [activeTool, setActiveTool] = useState<AnnotationTool>('select');
  const [selectedAnnotationId, setSelectedAnnotationId] = useState<string | null>(null);
  const [signatureModalOpen, setSignatureModalOpen] = useState<boolean>(false);
  const [activeSignatureToPlace, setActiveSignatureToPlace] = useState<{
    dataUrl: string;
    signerName: string;
    includeDate: boolean;
  } | null>(null);

  // Formatting Style Config
  const [styleConfig, setStyleConfig] = useState<AnnotationStyleConfig>({
    textColor: '#0f172a',
    bgColor: '#ffffff',
    fontSize: 14,
    highlightColor: '#fef08a',
    inkColor: '#dc2626',
    inkWidth: 3,
    stickyColor: '#f59e0b',
  });

  // Zoom handlers
  const handleZoomIn = () => {
    onUpdatePdf({ zoom: Math.min(250, zoom + 15) });
  };

  const handleZoomOut = () => {
    onUpdatePdf({ zoom: Math.max(50, zoom - 15) });
  };

  const handleRotate = () => {
    onUpdatePdf({ rotation: (rotation + 90) % 360 });
  };

  const handlePageChange = (newPage: number) => {
    if (newPage >= 1 && newPage <= totalPages) {
      onUpdatePdf({ currentPage: newPage });
      setSelectedAnnotationId(null);
    }
  };

  // Local PDF file picker
  const handleLocalPdfUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file && file.type === 'application/pdf') {
      const url = URL.createObjectURL(file);
      onUpdatePdf({
        title: file.name,
        url,
        currentPage: 1,
        totalPages: 4,
        annotations: [],
      });
    }
  };

  // Annotations CRUD
  const handleAddAnnotation = (newAnn: PdfAnnotation) => {
    const updated = [...annotations, newAnn];
    onUpdatePdf({ annotations: updated });
  };

  const handleUpdateAnnotation = (id: string, updates: Partial<PdfAnnotation>) => {
    const updated = annotations.map((a) => (a.id === id ? { ...a, ...updates } : a));
    onUpdatePdf({ annotations: updated });
  };

  const handleDeleteAnnotation = (id: string) => {
    const updated = annotations.filter((a) => a.id !== id);
    onUpdatePdf({ annotations: updated });
    if (selectedAnnotationId === id) setSelectedAnnotationId(null);
  };

  const handleAddReply = (annotationId: string, replyText: string) => {
    const updated = annotations.map((a) => {
      if (a.id === annotationId) {
        const newReply = {
          id: `rep-${Date.now()}`,
          author: 'Alex Johnson',
          text: replyText,
          timestamp: Date.now(),
        };
        return {
          ...a,
          replies: [...(a.replies || []), newReply],
        };
      }
      return a;
    });
    onUpdatePdf({ annotations: updated });
  };

  // Reflow to Word Docs
  const handleReflow = () => {
    setIsReflowing(true);
    setTimeout(() => {
      const paragraphs: DocParagraph[] = [
        {
          id: 'p-reflow-1',
          type: 'h1',
          content: pdfData.title.replace(/\.pdf$/i, ''),
          align: 'left',
        },
        {
          id: 'p-reflow-2',
          type: 'callout',
          content: `Reflowed from PDF (${pdfData.title}) with Anything Office Smart Parser. Includes ${annotations.length} converted annotations.`,
          align: 'left',
        },
        {
          id: 'p-reflow-3',
          type: 'h2',
          content: '1. Executive Abstract & Scope',
          align: 'left',
        },
        {
          id: 'p-reflow-4',
          type: 'p',
          content:
            pdfData.extractedText ||
            'This document unifies office document editing, high-performance spreadsheet computation, and presentation creation in a single lightweight desktop application.',
          align: 'justify',
        },
        {
          id: 'p-reflow-5',
          type: 'h2',
          content: '2. Verified Architectural Standards',
          align: 'left',
        },
        {
          id: 'p-reflow-6',
          type: 'bullet',
          content: 'ECMA-376 & ISO/IEC 29500 compliant OOXML container packaging',
          align: 'left',
        },
        {
          id: 'p-reflow-7',
          type: 'bullet',
          content: 'Piece Table editing buffer supporting O(1) text insertions and instantaneous undos',
          align: 'left',
        },
        {
          id: 'p-reflow-8',
          type: 'bullet',
          content: 'Directed Acyclic Graph (DAG) formula evaluation with topological sorting',
          align: 'left',
        },
      ];

      // Append sticky comments as callout blocks
      annotations
        .filter((a) => a.type === 'sticky' && a.content)
        .forEach((a, idx) => {
          paragraphs.push({
            id: `p-comment-${idx}`,
            type: 'blockquote',
            content: `Comment (p.${a.pageNumber}) by ${a.author}: "${a.content}"`,
            align: 'left',
          });
        });

      const reflowed: DocumentData = {
        title: pdfData.title.replace(/\.pdf$/i, ''),
        paragraphs,
        pageSize: 'A4',
        margins: { top: 24, right: 28, bottom: 24, left: 28 },
      };

      onReflowToDocs(reflowed);
      setIsReflowing(false);
    }, 400);
  };

  const scale = zoom / 100;

  return (
    <div
      className={`flex-1 flex flex-col h-full overflow-hidden select-none relative ${
        darkMode ? 'bg-slate-950 text-slate-100' : 'bg-slate-100 text-slate-800'
      }`}
    >
      {/* Hidden File Input */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".pdf"
        className="hidden"
        onChange={handleLocalPdfUpload}
      />

      {/* Top Microsoft/Adobe PDF Toolbar */}
      <header
        className={`h-11 border-b flex items-center justify-between px-3 text-xs shrink-0 z-20 transition-colors ${
          darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200 shadow-xs'
        }`}
      >
        {/* Left Section: Thumbnails & Title & Pagination */}
        <div className="flex items-center space-x-2">
          <button
            onClick={() => setSidebarOpen(!sidebarOpen)}
            title="Toggle Page Thumbnails"
            className={`p-1.5 rounded-lg flex items-center gap-1 transition-colors ${
              sidebarOpen
                ? 'bg-red-500/15 text-red-600 dark:text-red-400 font-semibold'
                : 'hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-500'
            }`}
          >
            <Sidebar className="w-4 h-4" />
          </button>

          <div className="h-4 w-px bg-slate-300 dark:bg-slate-700 mx-1" />

          {/* PDF Badge & Title */}
          <div className="flex items-center space-x-1.5 max-w-[240px] truncate">
            <span className="px-1.5 py-0.5 text-[10px] font-black rounded bg-red-600 text-white tracking-wider">
              PDF
            </span>
            <span className="font-semibold truncate text-xs" title={pdfData.title}>
              {pdfData.title}
            </span>
          </div>

          <div className="h-4 w-px bg-slate-300 dark:bg-slate-700 mx-1" />

          {/* Page Navigation */}
          <div className="flex items-center space-x-1 bg-slate-100 dark:bg-slate-800/80 rounded-lg px-1.5 py-0.5 border border-slate-200 dark:border-slate-700">
            <button
              onClick={() => handlePageChange(currentPage - 1)}
              disabled={currentPage <= 1}
              className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-700 disabled:opacity-30 disabled:hover:bg-transparent"
              title="Previous Page"
            >
              <ChevronLeft className="w-3.5 h-3.5" />
            </button>
            <span className="font-mono text-xs px-1">
              <strong className="text-red-600 dark:text-red-400">{currentPage}</strong> / {totalPages}
            </span>
            <button
              onClick={() => handlePageChange(currentPage + 1)}
              disabled={currentPage >= totalPages}
              className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-700 disabled:opacity-30 disabled:hover:bg-transparent"
              title="Next Page"
            >
              <ChevronRight className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>

        {/* Center Section: Zoom & Rotation Controls */}
        <div className="flex items-center space-x-1.5">
          <div className="flex items-center space-x-1 bg-slate-100 dark:bg-slate-800 rounded-lg px-1.5 py-0.5 border border-slate-200 dark:border-slate-700">
            <button
              onClick={handleZoomOut}
              title="Zoom Out"
              className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-700"
            >
              <ZoomOut className="w-3.5 h-3.5" />
            </button>
            <span className="font-mono text-xs w-12 text-center font-medium">{zoom}%</span>
            <button
              onClick={handleZoomIn}
              title="Zoom In"
              className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-700"
            >
              <ZoomIn className="w-3.5 h-3.5" />
            </button>
          </div>

          <button
            onClick={() => onUpdatePdf({ zoom: 100 })}
            title="Reset Zoom to 100%"
            className="p-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-400"
          >
            <Maximize2 className="w-3.5 h-3.5" />
          </button>

          <button
            onClick={handleRotate}
            title="Rotate Clockwise (90°)"
            className="p-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-400"
          >
            <RotateCw className="w-3.5 h-3.5" />
          </button>
        </div>

        {/* Right Section: Actions (Reflow to Word, Print, Open) */}
        <div className="flex items-center space-x-2">
          {/* Word PDF Reflow Button */}
          <button
            onClick={handleReflow}
            disabled={isReflowing}
            title="Convert and Reflow PDF into Editable Word Document"
            className="flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-blue-600 hover:bg-blue-700 text-white font-medium shadow-xs transition-all active:scale-95 disabled:opacity-50"
          >
            <Sparkles className="w-3.5 h-3.5" />
            <span>Reflow to Word Docs</span>
          </button>

          {/* Open local PDF */}
          <button
            onClick={() => fileInputRef.current?.click()}
            title="Open Another PDF File"
            className="p-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-400"
          >
            <FolderOpen className="w-4 h-4" />
          </button>

          {/* Print */}
          <button
            onClick={() => window.print()}
            title="Print Document"
            className="p-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-400"
          >
            <Printer className="w-4 h-4" />
          </button>
        </div>
      </header>

      {/* Main Area: Thumbnails Sidebar + Central Canvas Viewport + Comments Sidebar */}
      <div className="flex-1 flex overflow-hidden relative">
        {/* 1. Left Thumbnails Sidebar */}
        {sidebarOpen && (
          <aside
            className={`w-52 shrink-0 border-r flex flex-col overflow-y-auto p-3 space-y-3 transition-all z-10 ${
              darkMode ? 'bg-slate-900/95 border-slate-800' : 'bg-slate-50 border-slate-200'
            }`}
          >
            <div className="flex items-center justify-between text-[11px] font-semibold text-slate-500 uppercase tracking-wider px-1">
              <span>Thumbnails</span>
              <span>{totalPages} Pages</span>
            </div>

            {Array.from({ length: totalPages }).map((_, idx) => {
              const pageNum = idx + 1;
              const isActive = pageNum === currentPage;
              const pageAnnotationsCount = annotations.filter((a) => a.pageNumber === pageNum).length;

              return (
                <div
                  key={pageNum}
                  onClick={() => handlePageChange(pageNum)}
                  className={`group cursor-pointer flex flex-col items-center p-2 rounded-xl transition-all ${
                    isActive
                      ? 'bg-red-500/10 border-2 border-red-600 shadow-sm'
                      : 'border border-slate-200 dark:border-slate-800 hover:border-slate-400 dark:hover:border-slate-700 bg-white dark:bg-slate-800/40'
                  }`}
                >
                  {/* Thumbnail Representation */}
                  <div className="w-36 h-48 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-lg shadow-xs p-2 flex flex-col justify-between overflow-hidden select-none pointer-events-none relative">
                    <div className="space-y-1.5">
                      <div className="h-2 w-3/4 bg-slate-300 dark:bg-slate-700 rounded" />
                      <div className="h-1.5 w-full bg-slate-200 dark:bg-slate-800 rounded" />
                      <div className="h-1.5 w-5/6 bg-slate-200 dark:bg-slate-800 rounded" />
                      <div className="h-1.5 w-4/6 bg-slate-200 dark:bg-slate-800 rounded" />
                    </div>

                    <div className="space-y-1">
                      <div className="h-1.5 w-full bg-slate-200 dark:bg-slate-800 rounded" />
                      <div className="h-1.5 w-3/4 bg-slate-200 dark:bg-slate-800 rounded" />
                    </div>

                    {pageAnnotationsCount > 0 && (
                      <div className="absolute top-1 right-1 px-1.5 py-0.2 rounded-full bg-amber-500 text-white text-[8px] font-black">
                        {pageAnnotationsCount}
                      </div>
                    )}

                    <div className="flex justify-between items-center text-[8px] text-slate-400 border-t pt-1">
                      <span>Ao PDF</span>
                      <span>p.{pageNum}</span>
                    </div>
                  </div>

                  <span
                    className={`mt-1.5 text-[11px] font-medium ${
                      isActive ? 'text-red-600 dark:text-red-400 font-bold' : 'text-slate-500'
                    }`}
                  >
                    Page {pageNum}
                  </span>
                </div>
              );
            })}
          </aside>
        )}

        {/* 2. Central PDF Document Viewport with Floating Acrobat Command Bar */}
        <main
          className={`flex-1 overflow-auto flex flex-col items-center p-8 transition-colors relative ${
            darkMode ? 'bg-slate-950' : 'bg-slate-200/70'
          }`}
        >
          {/* Floating Acrobat Signature Command Bar */}
          <PdfAnnotationToolbar
            activeTool={activeTool}
            onSelectTool={(tool) => {
              setActiveTool(tool);
              if (tool !== 'select') setSelectedAnnotationId(null);
            }}
            styleConfig={styleConfig}
            onUpdateStyleConfig={(updates) => setStyleConfig({ ...styleConfig, ...updates })}
            onOpenSignatureModal={() => setSignatureModalOpen(true)}
            selectedAnnotationId={selectedAnnotationId}
            onDeleteSelectedAnnotation={() => {
              if (selectedAnnotationId) handleDeleteAnnotation(selectedAnnotationId);
            }}
            commentsCount={annotations.filter((a) => a.type === 'sticky' || a.replies?.length).length}
            commentsOpen={commentsOpen}
            onToggleComments={() => setCommentsOpen(!commentsOpen)}
            darkMode={darkMode}
          />

          {/* If there's an external PDF file URL loaded */}
          {pdfData.url ? (
            <div
              className="relative w-full h-full max-w-5xl rounded-lg overflow-hidden shadow-2xl border border-slate-300 dark:border-slate-800 bg-white"
              style={{
                transform: `rotate(${rotation}deg) scale(${scale})`,
                transformOrigin: 'top center',
                transition: 'transform 0.2s ease',
              }}
            >
              <iframe
                src={`${pdfData.url}#page=${currentPage}&zoom=${zoom}`}
                className="w-full h-full min-h-[820px] border-none"
                title="PDF Document"
              />

              {/* Absolute Canvas Annotation Overlay */}
              <AnnotationOverlay
                pageNumber={currentPage}
                zoom={zoom}
                rotation={rotation}
                activeTool={activeTool}
                styleConfig={styleConfig}
                annotations={annotations}
                onAddAnnotation={handleAddAnnotation}
                onUpdateAnnotation={handleUpdateAnnotation}
                onDeleteAnnotation={handleDeleteAnnotation}
                selectedAnnotationId={selectedAnnotationId}
                onSelectAnnotation={setSelectedAnnotationId}
                activeSignatureToPlace={activeSignatureToPlace}
                onSignaturePlaced={() => {
                  setActiveSignatureToPlace(null);
                  setActiveTool('select');
                }}
                onOpenCommentsSidebar={() => setCommentsOpen(true)}
              />
            </div>
          ) : (
            /* High-fidelity Vector Rendered PDF Page with Overlay */
            <div
              className="relative bg-white text-slate-900 shadow-2xl rounded-sm border border-slate-300 transition-transform origin-top flex flex-col justify-between"
              style={{
                width: `${794 * scale}px`,
                minHeight: `${1123 * scale}px`,
                transform: `rotate(${rotation}deg)`,
                padding: `${48 * scale}px ${54 * scale}px`,
                transition: 'transform 0.2s ease, width 0.15s ease, min-height 0.15s ease',
              }}
            >
              {/* Header */}
              <div className="border-b border-slate-200 pb-3 mb-6 flex justify-between items-center text-[10px] text-slate-400 uppercase tracking-widest font-semibold">
                <span>Anything Office Technical Specification</span>
                <span className="text-red-600 font-bold">Standard PDF / ECMA-376</span>
              </div>

              {/* Page Content depending on currentPage */}
              {currentPage === 1 && (
                <div className="space-y-6 flex-1 text-slate-800">
                  <div className="space-y-2">
                    <span className="inline-block px-2 py-0.5 text-[11px] font-bold rounded bg-red-100 text-red-700">
                      OFFICIAL WHITEPAPER
                    </span>
                    <h1 className="text-2xl font-black text-slate-900 tracking-tight leading-snug">
                      Anything Office Suite: Architectural Synthesis of Word, Excel & PowerPoint
                    </h1>
                    <p className="text-xs text-slate-500 font-medium">
                      Author: Anything Creative Architecture Group • Date: October 2026 • Document ID: AO-SPEC-2026
                    </p>
                  </div>

                  <div className="p-4 rounded-md bg-slate-50 border border-slate-200 text-xs leading-relaxed text-slate-700">
                    <strong className="block text-slate-900 font-bold mb-1">Executive Summary:</strong>
                    Modern office applications process distinct data abstractions. Microsoft Word leverages
                    a <em>Piece Table</em> buffer over flowable WordprocessingML; Microsoft Excel operates a
                    reactive <em>Directed Acyclic Graph (DAG)</em> formula recalculation engine over sparse matrices;
                    and Microsoft PowerPoint maintains an absolute <em>2D Scene Graph</em> in English Metric Units (EMUs).
                    Anything Office unifies all three models in a single reactive workspace.
                  </div>

                  <div className="space-y-3">
                    <h2 className="text-sm font-bold text-slate-900 border-b border-slate-200 pb-1">
                      1. Universal Document Support Matrix
                    </h2>
                    <table className="w-full text-[11px] border border-slate-200 rounded">
                      <thead className="bg-slate-100 text-slate-700 text-left font-bold">
                        <tr>
                          <th className="p-2 border-b">Format Group</th>
                          <th className="p-2 border-b">Extensions</th>
                          <th className="p-2 border-b">Engine Standard</th>
                          <th className="p-2 border-b">Status</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        <tr>
                          <td className="p-2 font-medium">Microsoft Word</td>
                          <td className="p-2 font-mono text-[10px]">.docx, .doc</td>
                          <td className="p-2">ISO/IEC 29500-1 WML</td>
                          <td className="p-2 text-emerald-600 font-bold">Active</td>
                        </tr>
                        <tr>
                          <td className="p-2 font-medium">Microsoft Excel</td>
                          <td className="p-2 font-mono text-[10px]">.xlsx, .xls, .csv</td>
                          <td className="p-2">ISO/IEC 29500-1 SML</td>
                          <td className="p-2 text-emerald-600 font-bold">Active</td>
                        </tr>
                        <tr>
                          <td className="p-2 font-medium">Microsoft PowerPoint</td>
                          <td className="p-2 font-mono text-[10px]">.pptx, .ppt</td>
                          <td className="p-2">ISO/IEC 29500-1 PML</td>
                          <td className="p-2 text-emerald-600 font-bold">Active</td>
                        </tr>
                        <tr>
                          <td className="p-2 font-medium">Portable Document Format</td>
                          <td className="p-2 font-mono text-[10px]">.pdf</td>
                          <td className="p-2">ISO 32000-2 (PDF 2.0)</td>
                          <td className="p-2 text-red-600 font-bold">Active</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  <div className="space-y-2">
                    <h2 className="text-sm font-bold text-slate-900 border-b border-slate-200 pb-1">
                      2. Sub-Second Instant Startup Engine
                    </h2>
                    <p className="text-xs text-slate-600 leading-relaxed">
                      By eliminating runtime network dependencies and leveraging direct binary execution with
                      Chromium hardware rasterization, Anything Office boots in sub-second timeframes with zero
                      window flashing.
                    </p>
                  </div>
                </div>
              )}

              {currentPage === 2 && (
                <div className="space-y-6 flex-1 text-slate-800">
                  <div className="space-y-1">
                    <span className="text-[10px] font-bold text-blue-600 tracking-wider">SECTION 02</span>
                    <h2 className="text-lg font-bold text-slate-900">
                      ECMA-376 / ISO/IEC 29500 OPC Container Architecture
                    </h2>
                  </div>

                  <p className="text-xs text-slate-600 leading-relaxed">
                    All modern Microsoft Office files are ZIP archives conforming to the Open Packaging
                    Conventions (OPC). At the package root sits <code className="bg-slate-100 px-1 py-0.5 rounded font-mono text-[11px]">[Content_Types].xml</code>,
                    mapping internal parts to standard MIME formats.
                  </p>

                  <div className="grid grid-cols-2 gap-4">
                    <div className="p-3 rounded-xl border border-blue-200 bg-blue-50/50">
                      <h4 className="font-bold text-xs text-blue-900 mb-1">Flowable Layout (Word)</h4>
                      <p className="text-[11px] text-blue-800">
                        Content re-wraps automatically based on page boundary, column width, and margin constraints.
                      </p>
                    </div>
                    <div className="p-3 rounded-xl border border-orange-200 bg-orange-50/50">
                      <h4 className="font-bold text-xs text-orange-900 mb-1">Absolute Spatial (PowerPoint)</h4>
                      <p className="text-[11px] text-orange-800">
                        Elements anchor to absolute (x, y) coordinates with fixed width, height, and z-ordering.
                      </p>
                    </div>
                  </div>
                </div>
              )}

              {currentPage === 3 && (
                <div className="space-y-6 flex-1 text-slate-800">
                  <div className="space-y-1">
                    <span className="text-[10px] font-bold text-emerald-600 tracking-wider">SECTION 03</span>
                    <h2 className="text-lg font-bold text-slate-900">
                      Reactive Formula Recalculation Engine & DAG Mechanics
                    </h2>
                  </div>

                  <p className="text-xs text-slate-600 leading-relaxed">
                    Spreadsheet computation in Anything Office uses a <strong>Directed Acyclic Graph (DAG)</strong>.
                    When a cell is modified, the engine flags all dependent cells dirty and re-evaluates them using
                    topological order.
                  </p>

                  <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 flex flex-col items-center space-y-3">
                    <div className="flex items-center space-x-4 text-xs font-mono">
                      <div className="px-3 py-1.5 bg-emerald-600 text-white rounded-lg font-bold">Cell A1: 250</div>
                      <span>+</span>
                      <div className="px-3 py-1.5 bg-emerald-600 text-white rounded-lg font-bold">Cell A2: 150</div>
                      <span>➔</span>
                      <div className="px-3 py-1.5 bg-slate-800 text-emerald-400 rounded-lg font-bold">Cell A3: =SUM(A1:A2)</div>
                    </div>
                  </div>
                </div>
              )}

              {/* Footer */}
              <div className="border-t border-slate-200 pt-3 mt-6 flex justify-between items-center text-[10px] text-slate-400">
                <span>CONFIDENTIAL • ANYTHING ECOSYSTEM SPECIFICATION</span>
                <span className="font-mono font-bold text-slate-600">Page {currentPage} of {totalPages}</span>
              </div>

              {/* Absolute Canvas Annotation Overlay */}
              <AnnotationOverlay
                pageNumber={currentPage}
                zoom={zoom}
                rotation={rotation}
                activeTool={activeTool}
                styleConfig={styleConfig}
                annotations={annotations}
                onAddAnnotation={handleAddAnnotation}
                onUpdateAnnotation={handleUpdateAnnotation}
                onDeleteAnnotation={handleDeleteAnnotation}
                selectedAnnotationId={selectedAnnotationId}
                onSelectAnnotation={setSelectedAnnotationId}
                activeSignatureToPlace={activeSignatureToPlace}
                onSignaturePlaced={() => {
                  setActiveSignatureToPlace(null);
                  setActiveTool('select');
                }}
                onOpenCommentsSidebar={() => setCommentsOpen(true)}
              />
            </div>
          )}
        </main>

        {/* 3. Right Comments & Threads Sidebar */}
        <CommentsSidebar
          isOpen={commentsOpen}
          onClose={() => setCommentsOpen(false)}
          annotations={annotations}
          currentPage={currentPage}
          onSelectAnnotation={(id) => {
            setSelectedAnnotationId(id);
            const target = annotations.find((a) => a.id === id);
            if (target && target.pageNumber !== currentPage) {
              handlePageChange(target.pageNumber);
            }
          }}
          selectedAnnotationId={selectedAnnotationId}
          onAddReply={handleAddReply}
          onDeleteAnnotation={handleDeleteAnnotation}
          darkMode={darkMode}
        />
      </div>

      {/* Digital Signature Stamp Modal */}
      <SignatureModal
        isOpen={signatureModalOpen}
        onClose={() => setSignatureModalOpen(false)}
        onApplySignature={(sigData) => {
          setActiveSignatureToPlace(sigData);
          setActiveTool('signature');
        }}
        darkMode={darkMode}
      />

      {/* Bottom Status Bar */}
      <footer
        className={`h-6 border-t flex items-center justify-between px-3 text-[11px] font-medium shrink-0 transition-colors z-20 ${
          darkMode ? 'bg-slate-900 border-slate-800 text-slate-400' : 'bg-slate-100 border-slate-200 text-slate-600'
        }`}
      >
        <div className="flex items-center space-x-3">
          <span className="flex items-center gap-1.5">
            <span className="w-2 h-2 rounded-full bg-red-600 animate-pulse" />
            <span className="font-semibold text-slate-700 dark:text-slate-300">Adobe Acrobat Annotation Suite</span>
          </span>
          <span>•</span>
          <span>
            Page {currentPage} of {totalPages}
          </span>
          <span>•</span>
          <span className="text-amber-600 dark:text-amber-400 font-semibold">
            {annotations.length} Active Annotations
          </span>
        </div>

        <div className="flex items-center space-x-3">
          <span className="text-[10px] text-emerald-600 dark:text-emerald-400 font-semibold flex items-center gap-1">
            <Sparkles className="w-3 h-3" />
            Word Reflow Ready
          </span>
          <span>•</span>
          <span className="font-mono">{zoom}%</span>
        </div>
      </footer>
    </div>
  );
};
