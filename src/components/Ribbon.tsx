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
  LayoutTemplate,
  FileText,
  Table as TableIcon,
  Presentation,
  Printer,
  Sparkles,
  List,
  ListOrdered,
  FolderOpen,
  Info,
  Trash2,
} from 'lucide-react';

interface RibbonProps {
  appMode: AppMode;
  onUndo: () => void;
  onRedo: () => void;
  canUndo: boolean;
  canRedo: boolean;
  onSave: () => void;
  onOpenFile: (file: File) => void;
  onExport: (format: 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'csv' | 'markdown' | 'html' | 'bundle' | 'print') => void;
  onOpenTemplates: () => void;
  darkMode: boolean;

  // Active formatting state
  activeFormat?: CellStyle;
  onFormatChange?: (format: Partial<CellStyle>) => void;

  // App specific actions
  onInsertDocParagraph?: (type: 'h1' | 'h2' | 'p' | 'bullet' | 'number' | 'callout' | 'table' | 'image' | 'hr') => void;
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
  const [activeTab, setActiveTab] = useState<'home' | 'insert' | 'layout' | 'formulas' | 'view'>('home');
  const [showFileMenu, setShowFileMenu] = useState<boolean>(false);
  const [selectedFont, setSelectedFont] = useState<string>('Calibri');
  const [selectedFontSize, setSelectedFontSize] = useState<number>(11);

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
            className="px-4 py-1.5 font-bold text-white rounded-t-sm transition-colors shadow-xs"
            style={{ backgroundColor: themeColor }}
          >
            File
          </button>

          {/* Microsoft Backstage File Menu Dropdown */}
          {showFileMenu && (
            <div className="absolute left-0 top-full mt-0.5 w-72 bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-b-lg shadow-2xl py-2 z-50 animate-fadeIn text-xs max-h-[85vh] overflow-y-auto">
              <input
                ref={fileInputRef}
                type="file"
                accept=".docx,.xlsx,.pptx,.pdf,.odt,.ods,.odp,.doc,.xls,.ppt,.csv,.tsv,.txt,.md,.html,.anything,.json"
                onChange={handleFileChange}
                className="hidden"
              />

              <button
                onClick={() => {
                  if (window.electronAPI?.openNativeFileDialog) {
                    window.electronAPI.openNativeFileDialog().then((result: any) => {
                      if (result && result.buffer) {
                        const uint8 = new Uint8Array(result.buffer);
                        const file = new File([uint8], result.name);
                        onOpenFile(file);
                        setShowFileMenu(false);
                      }
                    });
                  } else {
                    fileInputRef.current?.click();
                  }
                }}
                className="w-full text-left px-4 py-2 hover:bg-slate-100 dark:hover:bg-slate-700 font-medium flex items-center justify-between"
              >
                <div className="flex items-center space-x-2">
                  <FolderOpen className="w-3.5 h-3.5 text-blue-500" />
                  <span>Open Document...</span>
                </div>
                <span className="text-[10px] text-slate-400">Ctrl+O</span>
              </button>

              <button
                onClick={() => {
                  onSave();
                  setShowFileMenu(false);
                }}
                className="w-full text-left px-4 py-2 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between font-medium"
              >
                <div className="flex items-center space-x-2">
                  <Save className="w-3.5 h-3.5 text-blue-500" />
                  <span>Save</span>
                </div>
                <span className="text-[10px] text-slate-400">Ctrl+S</span>
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

              {/* SAVE AS / EXPORT OPTIONS (Parity with MS Office) */}
              <div className="px-4 py-1 text-[10px] uppercase font-bold text-slate-400">
                Save As / Export Formats
              </div>
              <button
                onClick={() => {
                  onExport('docx');
                  setShowFileMenu(false);
                }}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between font-semibold text-[#185abd] dark:text-blue-400"
              >
                <div className="flex items-center space-x-2">
                  <FileText className="w-3.5 h-3.5" />
                  <span>Microsoft Word (.docx)</span>
                </div>
                <span className="font-mono text-[10px]">OOXML</span>
              </button>
              <button
                onClick={() => {
                  onExport('xlsx');
                  setShowFileMenu(false);
                }}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between font-semibold text-[#107c41] dark:text-emerald-400"
              >
                <div className="flex items-center space-x-2">
                  <TableIcon className="w-3.5 h-3.5" />
                  <span>Microsoft Excel (.xlsx)</span>
                </div>
                <span className="font-mono text-[10px]">OOXML</span>
              </button>
              <button
                onClick={() => {
                  onExport('pptx');
                  setShowFileMenu(false);
                }}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between font-semibold text-[#c43e1c] dark:text-orange-400"
              >
                <div className="flex items-center space-x-2">
                  <Presentation className="w-3.5 h-3.5" />
                  <span>Microsoft PowerPoint (.pptx)</span>
                </div>
                <span className="font-mono text-[10px]">OOXML</span>
              </button>
              <button
                onClick={() => {
                  onExport('pdf');
                  setShowFileMenu(false);
                }}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between font-semibold text-rose-600 dark:text-rose-400"
              >
                <div className="flex items-center space-x-2">
                  <Printer className="w-3.5 h-3.5" />
                  <span>PDF Document (.pdf)</span>
                </div>
                <span className="font-mono text-[10px]">Print</span>
              </button>
              <button
                onClick={() => {
                  onExport('csv');
                  setShowFileMenu(false);
                }}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between"
              >
                <span>Comma-Separated Values (.csv)</span>
                <span className="font-mono text-[10px] text-slate-400">CSV</span>
              </button>
              <button
                onClick={() => {
                  onExport('markdown');
                  setShowFileMenu(false);
                }}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between"
              >
                <span>Markdown Document (.md)</span>
                <span className="font-mono text-[10px] text-slate-400">MD</span>
              </button>
              <button
                onClick={() => {
                  onExport('html');
                  setShowFileMenu(false);
                }}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between"
              >
                <span>Web Page (.html)</span>
                <span className="font-mono text-[10px] text-slate-400">HTML</span>
              </button>
              <button
                onClick={() => {
                  onExport('bundle');
                  setShowFileMenu(false);
                }}
                className="w-full text-left px-4 py-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between"
              >
                <span>Anything Unified Bundle (.anything)</span>
                <span className="font-mono text-[10px] text-violet-500">JSON</span>
              </button>

              <div className="border-t border-slate-200 dark:border-slate-700 my-1" />

              <div className="px-4 py-1 text-[10px] uppercase font-bold text-slate-400">
                Load Test Samples
              </div>
              <button
                onClick={() => openSampleUrl('/samples/sample_word.docx', 'sample_word.docx')}
                className="w-full text-left px-4 py-1 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between text-[11px]"
              >
                <span>Word Sample (.docx)</span>
                <span className="font-mono text-[10px] text-blue-500">Load</span>
              </button>
              <button
                onClick={() => openSampleUrl('/samples/sample_excel.xlsx', 'sample_excel.xlsx')}
                className="w-full text-left px-4 py-1 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between text-[11px]"
              >
                <span>Excel Sample (.xlsx)</span>
                <span className="font-mono text-[10px] text-emerald-500">Load</span>
              </button>
              <button
                onClick={() => openSampleUrl('/samples/sample_powerpoint.pptx', 'sample_powerpoint.pptx')}
                className="w-full text-left px-4 py-1 hover:bg-slate-100 dark:hover:bg-slate-700 flex items-center justify-between text-[11px]"
              >
                <span>PowerPoint Sample (.pptx)</span>
                <span className="font-mono text-[10px] text-orange-500">Load</span>
              </button>
            </div>
          )}
        </div>

        {/* Fluent Ribbon Tabs */}
        {[
          { id: 'home', label: 'Home' },
          { id: 'insert', label: 'Insert' },
          { id: 'layout', label: appMode === 'docs' ? 'Layout' : appMode === 'sheets' ? 'Formulas' : 'Slide Show' },
          { id: 'view', label: 'View' },
        ].map((tab) => (
          <button
            key={tab.id}
            onClick={() => setActiveTab(tab.id as any)}
            className={`px-3 py-1 font-medium transition-colors border-b-2 ${
              activeTab === tab.id
                ? 'border-b-2 text-slate-900 dark:text-white font-semibold'
                : 'border-transparent text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
            }`}
            style={{
              borderBottomColor: activeTab === tab.id ? themeColor : 'transparent',
            }}
          >
            {tab.label}
          </button>
        ))}
      </div>

      {/* 2. AUTHENTIC RIBBON ACTION TOOLBAR */}
      <div className="flex items-center justify-between px-3 py-1.5 overflow-x-auto min-h-[46px]">
        {/* TAB: HOME */}
        {activeTab === 'home' && (
          <div className="flex items-center space-x-3 shrink-0">
            {/* Clipboard group */}
            <div className="flex items-center space-x-1 border-r border-slate-300 dark:border-slate-700 pr-2">
              <button
                onClick={onUndo}
                disabled={!canUndo}
                title="Undo (Ctrl+Z)"
                className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 disabled:opacity-30"
              >
                <RotateCcw className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={onRedo}
                disabled={!canRedo}
                title="Redo (Ctrl+Y)"
                className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 disabled:opacity-30"
              >
                <RotateCw className="w-3.5 h-3.5" />
              </button>
            </div>

            {/* Font Family & Size Dropdowns */}
            <div className="flex items-center space-x-1 border-r border-slate-300 dark:border-slate-700 pr-2">
              <select
                value={selectedFont}
                onChange={(e) => setSelectedFont(e.target.value)}
                className="h-6 px-1.5 text-xs rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 focus:outline-none"
              >
                <option value="Calibri">Calibri</option>
                <option value="Aptos">Aptos</option>
                <option value="Arial">Arial</option>
                <option value="Times New Roman">Times New Roman</option>
                <option value="Segoe UI">Segoe UI</option>
                <option value="Georgia">Georgia</option>
                <option value="Courier New">Courier New</option>
              </select>

              <select
                value={selectedFontSize}
                onChange={(e) => setSelectedFontSize(Number(e.target.value))}
                className="h-6 px-1 text-xs rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 focus:outline-none"
              >
                <option value={9}>9</option>
                <option value={10}>10</option>
                <option value={11}>11</option>
                <option value={12}>12</option>
                <option value={14}>14</option>
                <option value={16}>16</option>
                <option value={18}>18</option>
                <option value={20}>20</option>
                <option value={24}>24</option>
                <option value={28}>28</option>
                <option value={36}>36</option>
              </select>
            </div>

            {/* Typography Formatting: Bold, Italic, Underline, Strike */}
            <div className="flex items-center space-x-1 border-r border-slate-300 dark:border-slate-700 pr-2">
              <button
                onClick={() => onFormatChange?.({ bold: !activeFormat?.bold })}
                className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
                  activeFormat?.bold ? 'bg-slate-300 dark:bg-slate-700 font-bold' : ''
                }`}
                title="Bold (Ctrl+B)"
              >
                <Bold className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => onFormatChange?.({ italic: !activeFormat?.italic })}
                className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
                  activeFormat?.italic ? 'bg-slate-300 dark:bg-slate-700' : ''
                }`}
                title="Italic (Ctrl+I)"
              >
                <Italic className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => onFormatChange?.({ underline: !activeFormat?.underline })}
                className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
                  activeFormat?.underline ? 'bg-slate-300 dark:bg-slate-700' : ''
                }`}
                title="Underline (Ctrl+U)"
              >
                <Underline className="w-3.5 h-3.5" />
              </button>
            </div>

            {/* Paragraph Alignment */}
            <div className="flex items-center space-x-1 border-r border-slate-300 dark:border-slate-700 pr-2">
              <button
                onClick={() => onFormatChange?.({ align: 'left' })}
                className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
                title="Align Left (Ctrl+L)"
              >
                <AlignLeft className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => onFormatChange?.({ align: 'center' })}
                className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
                title="Center (Ctrl+E)"
              >
                <AlignCenter className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => onFormatChange?.({ align: 'right' })}
                className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
                title="Align Right (Ctrl+R)"
              >
                <AlignRight className="w-3.5 h-3.5" />
              </button>
            </div>

            {/* Heading Styles (in Docs) */}
            {appMode === 'docs' && onInsertDocParagraph && (
              <div className="flex items-center space-x-1">
                <button
                  onClick={() => onInsertDocParagraph('h1')}
                  className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800 font-bold text-xs text-[#185abd] dark:text-blue-400"
                >
                  Heading 1
                </button>
                <button
                  onClick={() => onInsertDocParagraph('h2')}
                  className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800 font-semibold text-xs"
                >
                  Heading 2
                </button>
                <button
                  onClick={() => onInsertDocParagraph('bullet')}
                  className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
                  title="Bullet List"
                >
                  <List className="w-3.5 h-3.5" />
                </button>
                <button
                  onClick={() => onInsertDocParagraph('number')}
                  className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
                  title="Numbered List"
                >
                  <ListOrdered className="w-3.5 h-3.5" />
                </button>
              </div>
            )}

            {/* Quick Math Formats (in Sheets) */}
            {appMode === 'sheets' && (
              <div className="flex items-center space-x-1">
                <button
                  onClick={() => onInsertFormula?.('=SUM(B2:B10)')}
                  className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800 font-mono font-bold text-xs text-[#107c41] dark:text-emerald-400"
                >
                  ∑ AutoSum
                </button>
                <button
                  onClick={() => onInsertFormula?.('=AVERAGE(B2:B10)')}
                  className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800 font-mono text-xs"
                >
                  Average
                </button>
              </div>
            )}

            {/* Quick Actions (in Slides) */}
            {appMode === 'slides' && onAddSlide && (
              <div className="flex items-center space-x-1">
                <button
                  onClick={() => onAddSlide('title-bullets')}
                  className="px-2.5 py-0.5 rounded bg-orange-600 hover:bg-orange-700 text-white font-medium text-xs flex items-center space-x-1"
                >
                  <Plus className="w-3 h-3" />
                  <span>New Slide</span>
                </button>
                {onDeleteSlide && (
                  <button
                    onClick={onDeleteSlide}
                    className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-rose-500"
                    title="Delete Slide"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                )}
                {onStartSlideshow && (
                  <button
                    onClick={onStartSlideshow}
                    className="px-2.5 py-0.5 rounded border border-orange-500/50 hover:bg-orange-500/10 text-orange-600 dark:text-orange-400 font-semibold text-xs flex items-center space-x-1"
                  >
                    <Play className="w-3 h-3 fill-current" />
                    <span>From Beginning (F5)</span>
                  </button>
                )}
              </div>
            )}
          </div>
        )}

        {/* TAB: INSERT */}
        {activeTab === 'insert' && (
          <div className="flex items-center space-x-2 shrink-0">
            {appMode === 'docs' && onInsertDocParagraph && (
              <>
                <button
                  onClick={() => onInsertDocParagraph('table')}
                  className="flex items-center space-x-1 px-2.5 py-1 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800"
                >
                  <TableIcon className="w-3.5 h-3.5 text-blue-500" />
                  <span>Table</span>
                </button>
                <button
                  onClick={() => onInsertDocParagraph('image')}
                  className="flex items-center space-x-1 px-2.5 py-1 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800"
                >
                  <Sparkles className="w-3.5 h-3.5 text-amber-500" />
                  <span>Image Figure</span>
                </button>
                <button
                  onClick={() => onInsertDocParagraph('callout')}
                  className="flex items-center space-x-1 px-2.5 py-1 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800"
                >
                  <Info className="w-3.5 h-3.5 text-blue-500" />
                  <span>Callout Note</span>
                </button>
                <button
                  onClick={() => onInsertDocParagraph('hr')}
                  className="flex items-center space-x-1 px-2.5 py-1 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800"
                >
                  <span>Horizontal Line</span>
                </button>
              </>
            )}

            {appMode === 'sheets' && (
              <>
                <button
                  onClick={onAddSheetRow}
                  className="flex items-center space-x-1 px-2.5 py-1 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800"
                >
                  <Plus className="w-3.5 h-3.5 text-emerald-500" />
                  <span>Insert Rows</span>
                </button>
                <button
                  onClick={onAddSheetCol}
                  className="flex items-center space-x-1 px-2.5 py-1 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800"
                >
                  <Plus className="w-3.5 h-3.5 text-emerald-500" />
                  <span>Insert Columns</span>
                </button>
              </>
            )}

            {appMode === 'slides' && onAddSlide && (
              <>
                <button
                  onClick={() => onAddSlide('title')}
                  className="flex items-center space-x-1 px-2.5 py-1 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800"
                >
                  <span>Title Slide</span>
                </button>
                <button
                  onClick={() => onAddSlide('two-column')}
                  className="flex items-center space-x-1 px-2.5 py-1 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800"
                >
                  <span>Two Columns</span>
                </button>
                <button
                  onClick={() => onAddSlide('metric-highlight')}
                  className="flex items-center space-x-1 px-2.5 py-1 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800"
                >
                  <span>Metrics Highlight</span>
                </button>
              </>
            )}
          </div>
        )}

        {/* TAB: LAYOUT / FORMULAS / SLIDESHOW */}
        {activeTab === 'layout' && (
          <div className="flex items-center space-x-2 shrink-0">
            {appMode === 'docs' && (
              <div className="flex items-center space-x-2">
                <span className="text-[11px] font-semibold text-slate-500">Page Setup:</span>
                <button
                  onClick={() => onExport('print')}
                  className="px-2.5 py-1 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800 flex items-center space-x-1"
                >
                  <Printer className="w-3 h-3" />
                  <span>Print Setup</span>
                </button>
              </div>
            )}

            {appMode === 'sheets' && (
              <div className="flex items-center space-x-2">
                <span className="text-[11px] font-semibold text-slate-500">Formulas Library:</span>
                <button
                  onClick={() => onInsertFormula?.('=SUM(B2:B10)')}
                  className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-700 font-mono text-xs hover:bg-slate-200 dark:hover:bg-slate-800"
                >
                  SUM
                </button>
                <button
                  onClick={() => onInsertFormula?.('=IF(B2>100, "Target Reached", "Needs Work")')}
                  className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-700 font-mono text-xs hover:bg-slate-200 dark:hover:bg-slate-800"
                >
                  IF
                </button>
                <button
                  onClick={() => onInsertFormula?.('=VLOOKUP(A2, A1:D10, 2)')}
                  className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-700 font-mono text-xs hover:bg-slate-200 dark:hover:bg-slate-800"
                >
                  VLOOKUP
                </button>
                <button
                  onClick={() => onInsertFormula?.('=COUNTIF(B2:B20, ">100")')}
                  className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-700 font-mono text-xs hover:bg-slate-200 dark:hover:bg-slate-800"
                >
                  COUNTIF
                </button>
              </div>
            )}

            {appMode === 'slides' && onStartSlideshow && (
              <div className="flex items-center space-x-2">
                <button
                  onClick={onStartSlideshow}
                  className="px-3 py-1 rounded bg-[#c43e1c] hover:bg-orange-700 text-white font-medium flex items-center space-x-1.5"
                >
                  <Play className="w-3.5 h-3.5 fill-current" />
                  <span>Start Slide Show from First Slide (F5)</span>
                </button>
              </div>
            )}
          </div>
        )}

        {/* TAB: VIEW */}
        {activeTab === 'view' && (
          <div className="flex items-center space-x-3 shrink-0">
            <div className="flex items-center space-x-1.5 text-xs text-slate-600 dark:text-slate-300">
              <span className="font-semibold">Live Edit Integration:</span>
              <span className="px-2 py-0.5 rounded bg-emerald-100 dark:bg-emerald-950/60 text-emerald-600 dark:text-emerald-400 font-mono text-[10px]">
                Proxy available at localhost:4321
              </span>
            </div>
            <button
              onClick={() => onOpenTemplates()}
              className="px-2.5 py-1 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800 flex items-center space-x-1"
            >
              <LayoutTemplate className="w-3.5 h-3.5 text-indigo-500" />
              <span>Templates Gallery</span>
            </button>
          </div>
        )}

        {/* Quick Save Indicator Button on Right */}
        <div className="flex items-center space-x-2">
          <button
            onClick={() => onExport('docx')}
            title="Export as Microsoft Word (.docx)"
            className="hidden lg:flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-semibold text-blue-600 hover:bg-blue-50 dark:hover:bg-blue-950/40"
          >
            <span>Word (.docx)</span>
          </button>
          <button
            onClick={() => onExport('xlsx')}
            title="Export as Microsoft Excel (.xlsx)"
            className="hidden lg:flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-semibold text-emerald-600 hover:bg-emerald-50 dark:hover:bg-emerald-950/40"
          >
            <span>Excel (.xlsx)</span>
          </button>
          <button
            onClick={() => onExport('pptx')}
            title="Export as Microsoft PowerPoint (.pptx)"
            className="hidden lg:flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-semibold text-orange-600 hover:bg-orange-50 dark:hover:bg-orange-950/40"
          >
            <span>PowerPoint (.pptx)</span>
          </button>
        </div>
      </div>
    </header>
  );
};
