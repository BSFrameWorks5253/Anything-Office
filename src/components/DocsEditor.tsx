import React, { useMemo, useState } from 'react';
import type { DocumentData, DocParagraph } from '../types/office';
import {
  AlignCenter,
  Plus,
  Trash2,
  Info,
  ZoomIn,
  ZoomOut,
  BookOpen,
  Layout,
  CheckCircle2,
} from 'lucide-react';

interface DocsEditorProps {
  data: DocumentData;
  onChange: (updated: DocumentData) => void;
  darkMode: boolean;
}

export const DocsEditor: React.FC<DocsEditorProps> = ({ data, onChange, darkMode }) => {
  const [zoomLevel, setZoomLevel] = useState<number>(100);

  // Statistics
  const stats = useMemo(() => {
    let words = 0;
    let characters = 0;

    for (const p of data.paragraphs) {
      if (p.content) {
        const text = p.content.trim();
        characters += text.length;
        if (text.length > 0) {
          words += text.split(/\s+/).length;
        }
      }
      if (p.tableData) {
        for (const row of p.tableData) {
          for (const cell of row) {
            characters += cell.length;
            if (cell.trim().length > 0) words += cell.trim().split(/\s+/).length;
          }
        }
      }
    }

    const readingMinutes = Math.max(1, Math.ceil(words / 200));
    return { words, characters, readingMinutes };
  }, [data]);

  const updateParagraph = (id: string, updates: Partial<DocParagraph>) => {
    const newParagraphs = data.paragraphs.map((p) => (p.id === id ? { ...p, ...updates } : p));
    onChange({ ...data, paragraphs: newParagraphs });
  };

  const deleteParagraph = (id: string) => {
    if (data.paragraphs.length <= 1) return;
    const newParagraphs = data.paragraphs.filter((p) => p.id !== id);
    onChange({ ...data, paragraphs: newParagraphs });
  };

  const addParagraphAfter = (
    index: number,
    type: 'p' | 'h1' | 'h2' | 'bullet' | 'callout' | 'table' = 'p'
  ) => {
    const newId = `p-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`;
    const newP: DocParagraph = {
      id: newId,
      type,
      content:
        type === 'h1'
          ? 'New Heading 1'
          : type === 'h2'
          ? 'New Section Title'
          : type === 'bullet'
          ? 'New bullet point item'
          : type === 'callout'
          ? 'Important callout note'
          : type === 'table'
          ? 'Data Summary Table'
          : 'Type narrative paragraph text here...',
      align: 'left',
      tableData:
        type === 'table'
          ? [
              ['Column 1', 'Column 2', 'Column 3'],
              ['Item A1', 'Value B1', 'Target C1'],
              ['Item A2', 'Value B2', 'Target C2'],
            ]
          : undefined,
    };

    const newParagraphs = [...data.paragraphs];
    newParagraphs.splice(index + 1, 0, newP);
    onChange({ ...data, paragraphs: newParagraphs });
  };

  const updateTableCell = (pId: string, rowIdx: number, colIdx: number, val: string) => {
    const target = data.paragraphs.find((p) => p.id === pId);
    if (!target || !target.tableData) return;

    const newTable = target.tableData.map((row, rI) =>
      row.map((cell, cI) => (rI === rowIdx && cI === colIdx ? val : cell))
    );
    updateParagraph(pId, { tableData: newTable });
  };

  const addTableRow = (pId: string) => {
    const target = data.paragraphs.find((p) => p.id === pId);
    if (!target || !target.tableData || target.tableData.length === 0) return;

    const colCount = target.tableData[0].length;
    const emptyRow = new Array(colCount).fill('New Value');
    updateParagraph(pId, { tableData: [...target.tableData, emptyRow] });
  };

  return (
    <div className="flex flex-col h-full w-full select-text overflow-hidden">
      {/* 1. AUTHENTIC MICROSOFT WORD HORIZONTAL RULER */}
      <div
        className={`h-5 border-b select-none flex items-center justify-center text-[10px] font-mono shrink-0 transition-colors ${
          darkMode ? 'bg-slate-900 border-slate-800 text-slate-400' : 'bg-slate-100 border-slate-300 text-slate-500'
        }`}
      >
        <div className="w-full max-w-4xl px-8 flex items-center justify-between relative">
          {/* Left margin marker */}
          <div className="absolute left-6 w-2 h-2.5 bg-slate-400 dark:bg-slate-600 border border-slate-500 cursor-ew-resize clip-path-triangle" />
          
          {/* Ruler tick marks (1 to 7.5 inches) */}
          <div className="flex justify-between w-full text-[9px] text-slate-400">
            <span>•</span>
            <span>1</span>
            <span>•</span>
            <span>2</span>
            <span>•</span>
            <span>3</span>
            <span>•</span>
            <span>4</span>
            <span>•</span>
            <span>5</span>
            <span>•</span>
            <span>6</span>
            <span>•</span>
            <span>7</span>
            <span>•</span>
          </div>

          {/* Right margin marker */}
          <div className="absolute right-6 w-2 h-2.5 bg-slate-400 dark:bg-slate-600 border border-slate-500 cursor-ew-resize clip-path-triangle" />
        </div>
      </div>

      {/* 2. MAIN WORD CANVAS (A4 Physical Page on Word Gray Background) */}
      <div
        className={`flex-1 overflow-y-auto p-8 flex justify-center transition-colors ${
          darkMode ? 'bg-[#1e1e1e]' : 'bg-[#f3f2f1]'
        }`}
      >
        {/* Physical A4 Sheet Container */}
        <div
          className={`w-full max-w-4xl min-h-[1100px] rounded-xs p-12 transition-all relative ${
            darkMode
              ? 'bg-[#2b2b2b] text-slate-100 border border-neutral-700 shadow-2xl'
              : 'bg-white text-slate-900 border border-slate-300 shadow-[0_4px_24px_rgba(0,0,0,0.12)]'
          }`}
          style={{
            transform: `scale(${zoomLevel / 100})`,
            transformOrigin: 'top center',
            paddingTop: `${data.margins.top * 1.5}px`,
            paddingBottom: `${data.margins.bottom * 1.5}px`,
            paddingLeft: `${data.margins.left * 1.5}px`,
            paddingRight: `${data.margins.right * 1.5}px`,
          }}
        >
          {/* Header Bar */}
          <div className="border-b border-slate-200 dark:border-neutral-700 pb-3 mb-8 flex justify-between items-center text-xs text-slate-400 select-none">
            <span className="font-semibold tracking-wide uppercase text-[10px]">
              {data.title || 'Document 1'}
            </span>
            <span className="text-[10px]">A4 Format • Normal Margins (1")</span>
          </div>

          {/* Paragraph Blocks */}
          <div className="space-y-4">
            {data.paragraphs.map((p, idx) => (
              <div
                key={p.id}
                id={p.id}
                className="group relative -mx-4 px-4 py-1 rounded hover:bg-slate-50 dark:hover:bg-neutral-800/40 transition-colors"
              >
                {/* Floating Block Controls on Hover */}
                <div className="absolute right-2 -top-3 hidden group-hover:flex items-center space-x-1 bg-white dark:bg-neutral-800 border border-slate-300 dark:border-neutral-700 rounded-md shadow-md px-1.5 py-0.5 z-20 select-none text-[11px]">
                  <button
                    onClick={() => addParagraphAfter(idx, 'p')}
                    title="Add paragraph below"
                    className="p-1 hover:text-blue-500 text-slate-500"
                  >
                    <Plus className="w-3.5 h-3.5" />
                  </button>
                  <button
                    onClick={() =>
                      updateParagraph(p.id, {
                        align: p.align === 'center' ? 'left' : p.align === 'left' ? 'right' : 'center',
                      })
                    }
                    title="Cycle alignment"
                    className="p-1 hover:text-blue-500 text-slate-500"
                  >
                    <AlignCenter className="w-3.5 h-3.5" />
                  </button>
                  <button
                    onClick={() => deleteParagraph(p.id)}
                    title="Delete block"
                    className="p-1 hover:text-rose-500 text-slate-500"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>

                {/* Render Paragraph by Type */}
                {p.type === 'h1' && (
                  <input
                    type="text"
                    value={p.content}
                    onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                    className={`w-full text-3xl font-extrabold tracking-tight bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded px-1 text-[#185abd] dark:text-blue-400 ${
                      p.align === 'center' ? 'text-center' : p.align === 'right' ? 'text-right' : 'text-left'
                    }`}
                  />
                )}

                {p.type === 'h2' && (
                  <input
                    type="text"
                    value={p.content}
                    onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                    className={`w-full text-xl font-bold tracking-tight text-slate-900 dark:text-white bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded px-1 mt-3 ${
                      p.align === 'center' ? 'text-center' : p.align === 'right' ? 'text-right' : 'text-left'
                    }`}
                  />
                )}

                {p.type === 'h3' && (
                  <input
                    type="text"
                    value={p.content}
                    onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                    className={`w-full text-base font-semibold text-slate-800 dark:text-slate-100 bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded px-1 mt-2 ${
                      p.align === 'center' ? 'text-center' : p.align === 'right' ? 'text-right' : 'text-left'
                    }`}
                  />
                )}

                {p.type === 'p' && (
                  <textarea
                    rows={Math.max(2, Math.ceil(p.content.length / 85))}
                    value={p.content}
                    onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                    className={`w-full leading-relaxed text-sm text-slate-900 dark:text-slate-100 bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded p-1 resize-none ${
                      p.align === 'center'
                        ? 'text-center'
                        : p.align === 'right'
                        ? 'text-right'
                        : p.align === 'justify'
                        ? 'text-justify'
                        : 'text-left'
                    }`}
                  />
                )}

                {p.type === 'bullet' && (
                  <div className="flex items-start space-x-2.5">
                    <span className="text-[#185abd] dark:text-blue-400 font-bold select-none text-base leading-relaxed">•</span>
                    <textarea
                      rows={Math.max(1, Math.ceil(p.content.length / 80))}
                      value={p.content}
                      onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                      className="w-full text-sm leading-relaxed text-slate-900 dark:text-slate-100 bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded p-0.5 resize-none"
                    />
                  </div>
                )}

                {p.type === 'callout' && (
                  <div className="p-3.5 rounded-md border-l-4 border-[#185abd] bg-blue-50/70 dark:bg-blue-900/40 text-blue-950 dark:text-blue-50 flex items-start space-x-3">
                    <Info className="w-4 h-4 text-[#185abd] dark:text-blue-300 shrink-0 mt-0.5" />
                    <textarea
                      rows={2}
                      value={p.content}
                      onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                      className="w-full text-xs font-semibold leading-relaxed text-blue-950 dark:text-blue-50 bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded resize-none"
                    />
                  </div>
                )}

                {/* Word Data Table */}
                {p.type === 'table' && p.tableData && (
                  <div className="my-3 overflow-x-auto rounded border border-slate-300 dark:border-neutral-700">
                    <table className="w-full text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-100 dark:bg-neutral-800 border-b border-slate-300 dark:border-neutral-700">
                          {p.tableData[0]?.map((colHeader, cIdx) => (
                            <th key={cIdx} className="p-2 text-left font-bold border-r border-inherit last:border-r-0">
                              <input
                                type="text"
                                value={colHeader}
                                onChange={(e) => updateTableCell(p.id, 0, cIdx, e.target.value)}
                                className="w-full bg-transparent font-bold focus:outline-none"
                              />
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {p.tableData.slice(1).map((row, rIdx) => (
                          <tr
                            key={rIdx}
                            className="border-b border-slate-200 dark:border-neutral-800 last:border-b-0 hover:bg-slate-50 dark:hover:bg-neutral-800/40"
                          >
                            {row.map((cell, cIdx) => (
                              <td
                                key={cIdx}
                                className="p-2 border-r border-slate-200 dark:border-neutral-700 last:border-r-0"
                              >
                                <input
                                  type="text"
                                  value={cell}
                                  onChange={(e) =>
                                    updateTableCell(p.id, rIdx + 1, cIdx, e.target.value)
                                  }
                                  className="w-full bg-transparent focus:outline-none"
                                />
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    <div className="p-1 bg-slate-50 dark:bg-neutral-900 border-t border-inherit flex justify-end">
                      <button
                        onClick={() => addTableRow(p.id)}
                        className="text-[10px] text-[#185abd] dark:text-blue-400 hover:underline flex items-center space-x-1 px-2 py-0.5 font-medium"
                      >
                        <Plus className="w-3 h-3" />
                        <span>Insert Row Below</span>
                      </button>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* 3. AUTHENTIC MICROSOFT WORD STATUS BAR */}
      <footer
        className={`h-6 border-t select-none px-3 flex items-center justify-between text-[11px] shrink-0 transition-colors ${
          darkMode ? 'bg-slate-900 border-slate-800 text-slate-400' : 'bg-[#0078d4] text-white border-blue-700'
        }`}
      >
        {/* Left Status Information */}
        <div className="flex items-center space-x-3">
          <span>Page 1 of 1</span>
          <span>•</span>
          <span>{stats.words.toLocaleString()} words</span>
          <span>•</span>
          <span>English (United States)</span>
          <span className="hidden md:inline flex items-center space-x-1">
            <CheckCircle2 className="w-3 h-3 text-emerald-400 inline mr-1" />
            <span>Accessibility: Good to go</span>
          </span>
        </div>

        {/* Right Status Controls & Zoom Slider */}
        <div className="flex items-center space-x-4">
          <div className="flex items-center space-x-1">
            <button title="Read Mode" className="p-0.5 hover:bg-white/10 rounded">
              <BookOpen className="w-3.5 h-3.5" />
            </button>
            <button title="Print Layout" className="p-0.5 bg-white/20 rounded">
              <Layout className="w-3.5 h-3.5" />
            </button>
          </div>

          <div className="flex items-center space-x-1.5 font-mono text-[10px]">
            <button onClick={() => setZoomLevel((z) => Math.max(50, z - 10))} title="Zoom Out">
              <ZoomOut className="w-3 h-3" />
            </button>
            <span>{zoomLevel}%</span>
            <button onClick={() => setZoomLevel((z) => Math.min(200, z + 10))} title="Zoom In">
              <ZoomIn className="w-3 h-3" />
            </button>
          </div>
        </div>
      </footer>
    </div>
  );
};
