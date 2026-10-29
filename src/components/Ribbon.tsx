import React, { useRef, useState } from 'react';
import type { AppMode, CellStyle } from '../types/office';
import {
  Bold,
  Italic,
  Underline,
  AlignLeft,
  AlignCenter,
  AlignRight,
  Plus,
  Play,
  RotateCcw,
  RotateCw,
  Save,
  DollarSign,
  Percent,
  Hash,
  Upload,
  Trash2,
  LayoutTemplate,
} from 'lucide-react';

interface RibbonProps {
  appMode: AppMode;
  onUndo: () => void;
  onRedo: () => void;
  canUndo: boolean;
  canRedo: boolean;
  onSave: () => void;
  onOpenFile: (file: File) => void;
  onExport: (format: 'bundle' | 'html' | 'csv' | 'markdown' | 'print') => void;
  onOpenTemplates: () => void;
  darkMode: boolean;

  // Active formatting state
  activeFormat?: CellStyle;
  onFormatChange?: (format: Partial<CellStyle>) => void;

  // App specific actions
  onInsertDocParagraph?: (type: 'h1' | 'h2' | 'p' | 'bullet' | 'callout' | 'table') => void;
  onInsertFormula?: (formula: string) => void;
  onAddSheetRow?: () => void;
  onAddSheetCol?: () => void;
  onAddSlide?: (layout?: 'title' | 'title-bullets' | 'two-column' | 'metric-highlight') => void;
  onDeleteSlide?: () => void;
  onStartSlideshow?: () => void;
}

export const Ribbon: React.FC<RibbonProps> = ({
  appMode,
  onUndo,
  onRedo,
  canUndo,
  canRedo,
  onSave,
  onOpenFile,
  onExport,
  onOpenTemplates,
  darkMode,
  activeFormat = {},
  onFormatChange,
  onInsertDocParagraph,
  onInsertFormula,
  onAddSheetRow,
  onAddSheetCol,
  onAddSlide,
  onDeleteSlide,
  onStartSlideshow,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [activeTab, setActiveTab] = useState<'home' | 'insert' | 'layout' | 'view'>('home');
  const [showFileMenu, setShowFileMenu] = useState<boolean>(false);

  // Microsoft Brand Accent Color depending on app mode
  const themeColor =
    appMode === 'docs'
      ? '#185abd' // Microsoft Word Blue
      : appMode === 'sheets'
      ? '#107c41' // Microsoft Excel Green
      : appMode === 'slides'
      ? '#c43e1c' // Microsoft PowerPoint Orange
      : '#e11d48'; // Adobe Acrobat Red

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onOpenFile(file);
      setShowFileMenu(false);
    }
  };

  const openSampleUrl = async (samplePath: string, filename: string) => {
    try {
      const res = await fetch(samplePath);
      const blob = await res.blob();
      const file = new File([blob], filename, { type: blob.type });
      onOpenFile(file);
      setShowFileMenu(false);
    } catch (e) {
      console.error('Failed to load sample', e);
    }
  };

  return (
    <header
      className={`border-b select-none flex flex-col transition-colors z-30 shrink-0 ${
        darkMode ? 'bg-slate-900 border-slate-800 text-slate-200' : 'bg-slate-100 border-slate-300 text-slate-800'
      }`}
    >
      {/* 1. AUTHENTIC MICROSOFT RIBBON TABS */}
      <div className="flex items-center px-2 pt-0.5 border-b border-inherit bg-slate-100 dark:bg-slate-900 text-xs">
        {/* File Backstage Tab */}
        <div className="relative">
          <button
            onClick={() => setShowFileMenu(!showFileMenu)}
            className="px-3.5 py-1.5 font-bold text-white rounded-t-sm transition-colors shadow-xs"
            style={{ backgroundColor: themeColor }}
          >
            File
          </button>

          {/* Microsoft Backstage File Menu Dropdown */}
          {showFileMenu && (
            <div className="absolute left-0 top-full mt-0.5 w-64 bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-b-lg shadow-2xl py-2 z-50 animate-fadeIn text-xs">
              <input
                ref={fileInputRef}
                type="file"
                accept=".docx,.xlsx,.pptx,.pdf,.odt,.ods,.odp,.doc,.xls,.ppt,.csv,.tsv,.txt,.md,.html,.anything,.json"
                onChange={handleFileChange}
                className="hidden"
              />

              <button
                onClick={() => fileInputRef.current?.click()}
                className="w-full text-left px-4 py-2 hover:bg-slate-100 dark:hover:bg-slate-700 font-medium flex items-center justify-between"
              >
                <div className="flex items-center space-x-2">
                  <Upload className="w-3.5 h-3.5 text-blue-500" />
                  <span>Open from PC...</span>
                </div>
                <span className="text-[10px] text-slate-400">All Office Formats</span>
              </button>

              <button
                onClick={() => {
                  onSave();
                  setShowFileMenu(false);
                }}
                className="w-full text-left px-4 py-2 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center space-x-2 font-medium"
              >
                <Save className="w-3.5 h-3.5 text-blue-500" />
                <span>Save (Ctrl+S)</span>
              </button>

              <button
                onClick={() => {
                  onOpenTemplates();
                  setShowFileMenu(false);
                }}
                className="w-full text-left px-4 py-2 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center space-x-2 font-medium"
              >
                <LayoutTemplate className="w-3.5 h-3.5 text-indigo-500" />
                <span>New from Template...</span>
              </button>

              <div className="border-t border-slate-200 dark:border-slate-700 my-1" />

              <div className="px-4 py-1 text-[10px] uppercase font-bold text-slate-400">
                Load Sample Documents
              </div>
              <button
                onClick={() => openSampleUrl('/samples/sample_word.docx', 'sample_word.docx')}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between"
              >
                <span>Word Sample Document</span>
                <span className="font-mono text-[10px] text-blue-500">.docx</span>
              </button>
              <button
                onClick={() => openSampleUrl('/samples/sample_libreoffice.odt', 'sample_libreoffice.odt')}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between"
              >
                <span>LibreOffice Writer</span>
                <span className="font-mono text-[10px] text-blue-500">.odt</span>
              </button>
              <button
                onClick={() => openSampleUrl('/samples/sample_excel.xlsx', 'sample_excel.xlsx')}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between"
              >
                <span>Excel Spreadsheet</span>
                <span className="font-mono text-[10px] text-emerald-500">.xlsx</span>
              </button>
              <button
                onClick={() => openSampleUrl('/samples/sample_librecalc.ods', 'sample_librecalc.ods')}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between"
              >
                <span>LibreOffice Calc</span>
                <span className="font-mono text-[10px] text-emerald-500">.ods</span>
              </button>
              <button
                onClick={() => openSampleUrl('/samples/sample_powerpoint.pptx', 'sample_powerpoint.pptx')}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between"
              >
                <span>PowerPoint Presentation</span>
                <span className="font-mono text-[10px] text-orange-500">.pptx</span>
              </button>
              <button
                onClick={() => openSampleUrl('/samples/sample_whitepaper.pdf', 'sample_whitepaper.pdf')}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between"
              >
                <span>PDF Technical Whitepaper</span>
                <span className="font-mono text-[10px] text-red-500 font-bold">.pdf</span>
              </button>

              <div className="border-t border-slate-200 dark:border-slate-700 my-1" />

              <div className="px-4 py-1 text-[10px] uppercase font-bold text-slate-400">
                Export Options
              </div>
              <button
                onClick={() => {
                  onExport('bundle');
                  setShowFileMenu(false);
                }}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 font-medium text-blue-600 dark:text-blue-400"
              >
                Export Unified .anything Bundle
              </button>
              <button
                onClick={() => {
                  onExport('print');
                  setShowFileMenu(false);
                }}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700"
              >
                Print / Save as PDF
              </button>
            </div>
          )}
        </div>

        {/* Standard Microsoft Fluent Ribbon Tabs */}
        {[
          { id: 'home', label: 'Home' },
          { id: 'insert', label: 'Insert' },
          {
            id: 'layout',
            label:
              appMode === 'docs' ? 'Layout' : appMode === 'sheets' ? 'Formulas' : 'Slide Show',
          },
          { id: 'view', label: 'View' },
        ].map((tab) => (
          <button
            key={tab.id}
            onClick={() => setActiveTab(tab.id as any)}
            className={`px-3.5 py-1.5 font-semibold transition-colors relative rounded-t-sm ${
              activeTab === tab.id
                ? 'bg-white dark:bg-slate-800 text-slate-900 dark:text-white font-bold shadow-xs'
                : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200 hover:bg-slate-200/50 dark:hover:bg-slate-800/50'
            }`}
          >
            <span>{tab.label}</span>
            {activeTab === tab.id && (
              <span
                className="absolute bottom-0 left-0 right-0 h-0.5"
                style={{ backgroundColor: themeColor }}
              />
            )}
          </button>
        ))}
      </div>

      {/* 2. AUTHENTIC MICROSOFT RIBBON ACTION GROUPS (With Dividers and Bottom Labels) */}
      <div className="flex items-stretch px-3 py-1.5 bg-white dark:bg-slate-800 overflow-x-auto min-h-[58px]">
        {/* GROUP 1: UNDO / CLIPBOARD */}
        <div className="flex flex-col justify-between items-center pr-3 border-r border-slate-200 dark:border-slate-700">
          <div className="flex items-center space-x-1">
            <button
              onClick={onUndo}
              disabled={!canUndo}
              title="Undo (Ctrl+Z)"
              className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-700 disabled:opacity-30"
            >
              <RotateCcw className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={onRedo}
              disabled={!canRedo}
              title="Redo (Ctrl+Y)"
              className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-700 disabled:opacity-30"
            >
              <RotateCw className="w-3.5 h-3.5" />
            </button>
          </div>
          <span className="text-[9px] uppercase tracking-wider text-slate-400 font-semibold mt-1">
            Undo
          </span>
        </div>

        {/* GROUP 2: FONT FORMATTING (Bold, Italic, Underline) */}
        <div className="flex flex-col justify-between items-center px-3 border-r border-slate-200 dark:border-slate-700">
          <div className="flex items-center space-x-1">
            <button
              onClick={() => onFormatChange?.({ bold: !activeFormat.bold })}
              title="Bold (Ctrl+B)"
              className={`p-1.5 rounded transition-colors ${
                activeFormat.bold
                  ? 'bg-slate-200 dark:bg-slate-600 font-bold'
                  : 'hover:bg-slate-100 dark:hover:bg-slate-700'
              }`}
            >
              <Bold className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => onFormatChange?.({ italic: !activeFormat.italic })}
              title="Italic (Ctrl+I)"
              className={`p-1.5 rounded transition-colors ${
                activeFormat.italic
                  ? 'bg-slate-200 dark:bg-slate-600'
                  : 'hover:bg-slate-100 dark:hover:bg-slate-700'
              }`}
            >
              <Italic className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => onFormatChange?.({ underline: !activeFormat.underline })}
              title="Underline (Ctrl+U)"
              className={`p-1.5 rounded transition-colors ${
                activeFormat.underline
                  ? 'bg-slate-200 dark:bg-slate-600'
                  : 'hover:bg-slate-100 dark:hover:bg-slate-700'
              }`}
            >
              <Underline className="w-3.5 h-3.5" />
            </button>
          </div>
          <span className="text-[9px] uppercase tracking-wider text-slate-400 font-semibold mt-1">
            Font
          </span>
        </div>

        {/* GROUP 3: PARAGRAPH ALIGNMENT */}
        <div className="flex flex-col justify-between items-center px-3 border-r border-slate-200 dark:border-slate-700">
          <div className="flex items-center space-x-1">
            <button
              onClick={() => onFormatChange?.({ align: 'left' })}
              title="Align Left"
              className={`p-1.5 rounded transition-colors ${
                activeFormat.align === 'left' || !activeFormat.align
                  ? 'bg-slate-200 dark:bg-slate-600'
                  : 'hover:bg-slate-100 dark:hover:bg-slate-700'
              }`}
            >
              <AlignLeft className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => onFormatChange?.({ align: 'center' })}
              title="Align Center"
              className={`p-1.5 rounded transition-colors ${
                activeFormat.align === 'center'
                  ? 'bg-slate-200 dark:bg-slate-600'
                  : 'hover:bg-slate-100 dark:hover:bg-slate-700'
              }`}
            >
              <AlignCenter className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => onFormatChange?.({ align: 'right' })}
              title="Align Right"
              className={`p-1.5 rounded transition-colors ${
                activeFormat.align === 'right'
                  ? 'bg-slate-200 dark:bg-slate-600'
                  : 'hover:bg-slate-100 dark:hover:bg-slate-700'
              }`}
            >
              <AlignRight className="w-3.5 h-3.5" />
            </button>
          </div>
          <span className="text-[9px] uppercase tracking-wider text-slate-400 font-semibold mt-1">
            Paragraph
          </span>
        </div>

        {/* GROUP 4: APP MODE SPECIFIC COMMAND GROUPS */}
        {appMode === 'docs' && (
          <div className="flex flex-col justify-between items-center px-3 border-r border-slate-200 dark:border-slate-700">
            <div className="flex items-center space-x-1">
              <button
                onClick={() => onInsertDocParagraph?.('h1')}
                className="px-2 py-1 text-xs font-bold rounded hover:bg-slate-100 dark:hover:bg-slate-700"
              >
                Heading 1
              </button>
              <button
                onClick={() => onInsertDocParagraph?.('h2')}
                className="px-2 py-1 text-xs font-semibold rounded hover:bg-slate-100 dark:hover:bg-slate-700"
              >
                Heading 2
              </button>
              <button
                onClick={() => onInsertDocParagraph?.('bullet')}
                className="px-2 py-1 text-xs rounded hover:bg-slate-100 dark:hover:bg-slate-700"
              >
                Bullet
              </button>
              <button
                onClick={() => onInsertDocParagraph?.('table')}
                className="px-2 py-1 text-xs rounded hover:bg-slate-100 dark:hover:bg-slate-700 bg-blue-50 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400 font-semibold"
              >
                + Table
              </button>
            </div>
            <span className="text-[9px] uppercase tracking-wider text-slate-400 font-semibold mt-1">
              Styles & Content
            </span>
          </div>
        )}

        {appMode === 'sheets' && (
          <>
            <div className="flex flex-col justify-between items-center px-3 border-r border-slate-200 dark:border-slate-700">
              <div className="flex items-center space-x-1">
                <button
                  onClick={() => onFormatChange?.({ format: 'currency' })}
                  title="Format Currency ($)"
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-700"
                >
                  <DollarSign className="w-3.5 h-3.5" />
                </button>
                <button
                  onClick={() => onFormatChange?.({ format: 'percent' })}
                  title="Format Percentage (%)"
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-700"
                >
                  <Percent className="w-3.5 h-3.5" />
                </button>
                <button
                  onClick={() => onFormatChange?.({ format: 'number' })}
                  title="Format Number"
                  className="p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-700"
                >
                  <Hash className="w-3.5 h-3.5" />
                </button>
              </div>
              <span className="text-[9px] uppercase tracking-wider text-slate-400 font-semibold mt-1">
                Number
              </span>
            </div>

            <div className="flex flex-col justify-between items-center px-3 border-r border-slate-200 dark:border-slate-700">
              <div className="flex items-center space-x-1">
                <button
                  onClick={() => onInsertFormula?.('=SUM(A1:A5)')}
                  className="px-2 py-1 text-xs font-mono font-bold rounded hover:bg-slate-100 dark:hover:bg-slate-700 text-emerald-600 dark:text-emerald-400"
                >
                  ∑ AutoSum
                </button>
                <button
                  onClick={() => onInsertFormula?.('=AVERAGE(A1:A5)')}
                  className="px-2 py-1 text-xs font-mono rounded hover:bg-slate-100 dark:hover:bg-slate-700"
                >
                  Average
                </button>
              </div>
              <span className="text-[9px] uppercase tracking-wider text-slate-400 font-semibold mt-1">
                Formulas
              </span>
            </div>

            <div className="flex flex-col justify-between items-center px-3 border-r border-slate-200 dark:border-slate-700">
              <div className="flex items-center space-x-1">
                <button
                  onClick={onAddSheetRow}
                  className="px-2 py-1 text-xs rounded hover:bg-slate-100 dark:hover:bg-slate-700"
                >
                  + Row
                </button>
                <button
                  onClick={onAddSheetCol}
                  className="px-2 py-1 text-xs rounded hover:bg-slate-100 dark:hover:bg-slate-700"
                >
                  + Column
                </button>
              </div>
              <span className="text-[9px] uppercase tracking-wider text-slate-400 font-semibold mt-1">
                Cells
              </span>
            </div>
          </>
        )}

        {appMode === 'slides' && (
          <div className="flex flex-col justify-between items-center px-3 border-r border-slate-200 dark:border-slate-700">
            <div className="flex items-center space-x-1">
              <button
                onClick={() => onAddSlide?.('title-bullets')}
                className="flex items-center space-x-1 px-2.5 py-1 text-xs font-semibold rounded bg-orange-600 text-white hover:bg-orange-700"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>New Slide</span>
              </button>
              <button
                onClick={onStartSlideshow}
                className="flex items-center space-x-1 px-2.5 py-1 text-xs font-semibold rounded bg-emerald-600 text-white hover:bg-emerald-700"
              >
                <Play className="w-3.5 h-3.5" />
                <span>From Beginning (F5)</span>
              </button>
              <button
                onClick={onDeleteSlide}
                className="p-1 rounded text-red-500 hover:bg-red-50 dark:hover:bg-red-950/30"
                title="Delete Current Slide"
              >
                <Trash2 className="w-3.5 h-3.5" />
              </button>
            </div>
            <span className="text-[9px] uppercase tracking-wider text-slate-400 font-semibold mt-1">
              Slide Show
            </span>
          </div>
        )}
      </div>
    </header>
  );
};
