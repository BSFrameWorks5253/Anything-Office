import React, { useState } from 'react';
import type { AnnotationTool, AnnotationStyleConfig } from '../../types/annotations';
import {
  MousePointer,
  Hand,
  Type,
  Highlighter,
  Underline as UnderlineIcon,
  Strikethrough as StrikethroughIcon,
  MessageSquarePlus,
  PenTool,
  Stamp,
  Trash2,
  ChevronDown,
  MessageSquare,
} from 'lucide-react';

interface PdfAnnotationToolbarProps {
  activeTool: AnnotationTool;
  onSelectTool: (tool: AnnotationTool) => void;
  styleConfig: AnnotationStyleConfig;
  onUpdateStyleConfig: (updates: Partial<AnnotationStyleConfig>) => void;
  onOpenSignatureModal: () => void;
  selectedAnnotationId: string | null;
  onDeleteSelectedAnnotation: () => void;
  commentsCount: number;
  commentsOpen: boolean;
  onToggleComments: () => void;
  darkMode: boolean;
}

const HIGHLIGHT_COLORS = ['#fef08a', '#bbf7d0', '#bae6fd', '#fbcfe8', '#fed7aa'];
const INK_COLORS = ['#dc2626', '#2563eb', '#16a34a', '#1e293b', '#9333ea', '#ea580c'];
const FONT_SIZES = [11, 13, 15, 18, 22, 28];

export const PdfAnnotationToolbar: React.FC<PdfAnnotationToolbarProps> = ({
  activeTool,
  onSelectTool,
  styleConfig,
  onUpdateStyleConfig,
  onOpenSignatureModal,
  selectedAnnotationId,
  onDeleteSelectedAnnotation,
  commentsCount,
  commentsOpen,
  onToggleComments,
  darkMode,
}) => {
  const [showColorPicker, setShowColorPicker] = useState<boolean>(false);
  const [showMarkupMenu, setShowMarkupMenu] = useState<boolean>(false);

  return (
    <div
      className={`absolute top-4 left-1/2 -translate-x-1/2 z-30 flex items-center gap-1 px-2.5 py-1.5 rounded-2xl shadow-2xl border backdrop-blur-xl transition-all duration-200 select-none ${
        darkMode
          ? 'bg-slate-900/90 border-slate-700/80 text-slate-200 shadow-black/60'
          : 'bg-white/95 border-slate-200/90 text-slate-700 shadow-slate-900/15'
      }`}
    >
      {/* 1. Selection & Panning */}
      <div className="flex items-center gap-0.5 pr-1 border-r border-slate-200 dark:border-slate-800">
        <button
          onClick={() => onSelectTool('select')}
          title="Select & Move Annotations (V)"
          className={`p-2 rounded-xl transition-all ${
            activeTool === 'select'
              ? 'bg-red-600 text-white shadow-md shadow-red-500/25 font-bold scale-105'
              : 'hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400'
          }`}
        >
          <MousePointer className="w-4 h-4" />
        </button>

        <button
          onClick={() => onSelectTool('hand')}
          title="Hand Panning Tool (H)"
          className={`p-2 rounded-xl transition-all ${
            activeTool === 'hand'
              ? 'bg-red-600 text-white shadow-md shadow-red-500/25 font-bold scale-105'
              : 'hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400'
          }`}
        >
          <Hand className="w-4 h-4" />
        </button>
      </div>

      {/* 2. Text (Typewriter) */}
      <button
        onClick={() => onSelectTool('text')}
        title="Add Text Annotation / Typewriter (T)"
        className={`p-2 rounded-xl transition-all flex items-center gap-1 ${
          activeTool === 'text'
            ? 'bg-red-600 text-white shadow-md shadow-red-500/25 font-bold scale-105'
            : 'hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400'
        }`}
      >
        <Type className="w-4 h-4" />
      </button>

      {/* 3. Highlight & Markup Dropdown */}
      <div className="relative">
        <div className="flex items-center">
          <button
            onClick={() => onSelectTool('highlight')}
            title="Highlight Text / Area (U)"
            className={`p-2 rounded-l-xl transition-all ${
              activeTool === 'highlight'
                ? 'bg-amber-500 text-white shadow-md shadow-amber-500/25 font-bold'
                : 'hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400'
            }`}
          >
            <Highlighter className="w-4 h-4" />
          </button>
          <button
            onClick={() => setShowMarkupMenu(!showMarkupMenu)}
            title="More Markup Options (Underline, Strikethrough)"
            className={`p-1.5 rounded-r-xl border-l transition-colors ${
              activeTool === 'highlight' || activeTool === 'underline' || activeTool === 'strikethrough'
                ? 'bg-amber-600 text-white border-amber-400'
                : 'hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-400 border-slate-200 dark:border-slate-800'
            }`}
          >
            <ChevronDown className="w-3 h-3" />
          </button>
        </div>

        {/* Dropdown for Underline / Strikethrough */}
        {showMarkupMenu && (
          <div
            className={`absolute top-full mt-2 left-0 w-44 rounded-xl shadow-xl border p-1 z-50 text-xs ${
              darkMode ? 'bg-slate-900 border-slate-700 text-slate-200' : 'bg-white border-slate-200 text-slate-700'
            }`}
          >
            <button
              onClick={() => {
                onSelectTool('highlight');
                setShowMarkupMenu(false);
              }}
              className="w-full flex items-center gap-2 px-2.5 py-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800"
            >
              <Highlighter className="w-3.5 h-3.5 text-amber-500" />
              <span>Highlight Area</span>
            </button>
            <button
              onClick={() => {
                onSelectTool('underline');
                setShowMarkupMenu(false);
              }}
              className="w-full flex items-center gap-2 px-2.5 py-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800"
            >
              <UnderlineIcon className="w-3.5 h-3.5 text-blue-500" />
              <span>Underline Text</span>
            </button>
            <button
              onClick={() => {
                onSelectTool('strikethrough');
                setShowMarkupMenu(false);
              }}
              className="w-full flex items-center gap-2 px-2.5 py-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800"
            >
              <StrikethroughIcon className="w-3.5 h-3.5 text-red-500" />
              <span>Strikethrough</span>
            </button>
          </div>
        )}
      </div>

      {/* 4. Sticky Note / Comment Pin */}
      <button
        onClick={() => onSelectTool('sticky')}
        title="Add Sticky Note Comment (C)"
        className={`p-2 rounded-xl transition-all ${
          activeTool === 'sticky'
            ? 'bg-amber-400 text-slate-900 shadow-md shadow-amber-400/25 font-bold scale-105'
            : 'hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400'
        }`}
      >
        <MessageSquarePlus className="w-4 h-4" />
      </button>

      {/* 5. Freehand Drawing / Ink */}
      <button
        onClick={() => onSelectTool('ink')}
        title="Freehand Ink Drawing (D)"
        className={`p-2 rounded-xl transition-all ${
          activeTool === 'ink'
            ? 'bg-red-600 text-white shadow-md shadow-red-500/25 font-bold scale-105'
            : 'hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400'
        }`}
      >
        <PenTool className="w-4 h-4" />
      </button>

      {/* 6. Digital Signatures */}
      <button
        onClick={() => {
          onOpenSignatureModal();
          onSelectTool('signature');
        }}
        title="Place Digital Signature (S)"
        className={`p-2 rounded-xl transition-all flex items-center gap-1.5 ${
          activeTool === 'signature'
            ? 'bg-blue-600 text-white shadow-md shadow-blue-500/25 font-bold scale-105'
            : 'hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400'
        }`}
      >
        <Stamp className="w-4 h-4" />
        <span className="text-xs font-semibold pr-1 hidden sm:inline">Sign</span>
      </button>

      {/* 7. Style & Color Picker Dropdown */}
      <div className="relative border-l border-slate-200 dark:border-slate-800 pl-1 ml-0.5">
        <button
          onClick={() => setShowColorPicker(!showColorPicker)}
          title="Annotation Formatting & Palette"
          className="p-2 rounded-xl hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400 flex items-center gap-1"
        >
          <div
            className="w-4 h-4 rounded-full border border-white/60 shadow-xs"
            style={{
              backgroundColor:
                activeTool === 'highlight'
                  ? styleConfig.highlightColor
                  : activeTool === 'ink'
                  ? styleConfig.inkColor
                  : styleConfig.textColor,
            }}
          />
          <ChevronDown className="w-3 h-3 text-slate-400" />
        </button>

        {showColorPicker && (
          <div
            className={`absolute top-full mt-2 left-1/2 -translate-x-1/2 w-64 rounded-2xl shadow-2xl border p-3 z-50 text-xs space-y-3 ${
              darkMode ? 'bg-slate-900 border-slate-700 text-slate-200' : 'bg-white border-slate-200 text-slate-700'
            }`}
          >
            {/* Ink / Text Color */}
            <div>
              <span className="block font-semibold mb-1.5 text-[11px] text-slate-400 uppercase tracking-wider">
                Pen & Text Color
              </span>
              <div className="flex items-center gap-1.5">
                {INK_COLORS.map((c) => (
                  <button
                    key={c}
                    onClick={() => {
                      onUpdateStyleConfig({ inkColor: c, textColor: c });
                    }}
                    className={`w-6 h-6 rounded-full border transition-transform ${
                      styleConfig.inkColor === c ? 'scale-125 border-white ring-2 ring-blue-500' : 'border-black/10'
                    }`}
                    style={{ backgroundColor: c }}
                  />
                ))}
              </div>
            </div>

            {/* Highlight Colors */}
            <div>
              <span className="block font-semibold mb-1.5 text-[11px] text-slate-400 uppercase tracking-wider">
                Highlight Palette
              </span>
              <div className="flex items-center gap-1.5">
                {HIGHLIGHT_COLORS.map((c) => (
                  <button
                    key={c}
                    onClick={() => onUpdateStyleConfig({ highlightColor: c })}
                    className={`w-6 h-6 rounded-full border transition-transform ${
                      styleConfig.highlightColor === c ? 'scale-125 border-white ring-2 ring-amber-500' : 'border-black/10'
                    }`}
                    style={{ backgroundColor: c }}
                  />
                ))}
              </div>
            </div>

            {/* Brush Width & Font Size */}
            <div className="grid grid-cols-2 gap-2 pt-2 border-t border-slate-200 dark:border-slate-800">
              <div>
                <label className="block text-[10px] font-semibold text-slate-400 mb-1">BRUSH WIDTH</label>
                <div className="flex items-center gap-1">
                  {[2, 4, 6, 8].map((w) => (
                    <button
                      key={w}
                      onClick={() => onUpdateStyleConfig({ inkWidth: w })}
                      className={`px-2 py-0.5 rounded text-xs font-mono font-bold ${
                        styleConfig.inkWidth === w
                          ? 'bg-red-600 text-white'
                          : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300'
                      }`}
                    >
                      {w}px
                    </button>
                  ))}
                </div>
              </div>

              <div>
                <label className="block text-[10px] font-semibold text-slate-400 mb-1">FONT SIZE</label>
                <select
                  value={styleConfig.fontSize}
                  onChange={(e) => onUpdateStyleConfig({ fontSize: Number(e.target.value) })}
                  className="w-full bg-slate-100 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded px-1.5 py-0.5 text-xs font-mono"
                >
                  {FONT_SIZES.map((sz) => (
                    <option key={sz} value={sz}>
                      {sz} pt
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* 8. Delete selected annotation */}
      {selectedAnnotationId && (
        <button
          onClick={onDeleteSelectedAnnotation}
          title="Delete Selected Annotation (Del)"
          className="p-2 rounded-xl text-red-500 hover:bg-red-500/15 transition-colors ml-0.5"
        >
          <Trash2 className="w-4 h-4" />
        </button>
      )}

      {/* 9. Comments thread panel toggle */}
      <div className="border-l border-slate-200 dark:border-slate-800 pl-1 ml-0.5">
        <button
          onClick={onToggleComments}
          title="Open Comments & Threads Sidebar"
          className={`p-2 rounded-xl transition-all relative ${
            commentsOpen
              ? 'bg-slate-200 dark:bg-slate-700 text-slate-900 dark:text-slate-100 font-bold'
              : 'hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400'
          }`}
        >
          <MessageSquare className="w-4 h-4" />
          {commentsCount > 0 && (
            <span className="absolute -top-1 -right-1 px-1.5 py-0.2 rounded-full text-[9px] font-black bg-red-600 text-white shadow-xs">
              {commentsCount}
            </span>
          )}
        </button>
      </div>
    </div>
  );
};
