import React, { useState, useEffect, useRef, useMemo } from 'react';
import type { SpreadsheetData, CellStyle } from '../types/office';
import { FormulaEngine } from '../engine/formulaEngine';
import { Plus, X, Check, Sheet, ZoomIn, ZoomOut } from 'lucide-react';

interface SheetsEditorProps {
  data: SpreadsheetData;
  onChange: (updated: SpreadsheetData) => void;
  darkMode: boolean;
  activeFormat?: CellStyle;
}

export const SheetsEditor: React.FC<SheetsEditorProps> = ({
  data,
  onChange,
  darkMode,
}) => {
  const activeSheet =
    data.sheets.find((s) => s.id === data.activeSheetId) || data.sheets[0];

  const [selectedCell, setSelectedCell] = useState<string>('A1');
  const [isEditing, setIsEditing] = useState<boolean>(false);
  const [editValue, setEditValue] = useState<string>('');
  const [zoomLevel, setZoomLevel] = useState<number>(100);

  const inputRef = useRef<HTMLInputElement>(null);

  // When selected cell changes, update the formula bar text
  useEffect(() => {
    const rawVal = activeSheet.cells[selectedCell]?.raw || '';
    setEditValue(rawVal);
    setIsEditing(false);
  }, [selectedCell, activeSheet]);

  // Focus input when editing starts
  useEffect(() => {
    if (isEditing && inputRef.current) {
      inputRef.current.focus();
    }
  }, [isEditing]);

  // Evaluated cache of all cells in the active sheet
  const computedSheet = useMemo(() => {
    const computed: Record<string, { value: string | number; error?: string }> = {};
    for (const key of Object.keys(activeSheet.cells)) {
      computed[key] = FormulaEngine.evaluateCell(key, activeSheet.cells);
    }
    return computed;
  }, [activeSheet]);

  // Live Excel Status Bar Calculations (Sum, Average, Count of numeric values)
  const gridStats = useMemo(() => {
    let sum = 0;
    let count = 0;

    for (const k of Object.keys(computedSheet)) {
      const v = computedSheet[k]?.value;
      if (typeof v === 'number' && !isNaN(v)) {
        sum += v;
        count++;
      }
    }

    const avg = count > 0 ? Math.round((sum / count) * 100) / 100 : 0;
    return { sum, count, avg };
  }, [computedSheet]);

  const updateCellRaw = (cellKey: string, newRaw: string) => {
    const currentCell = activeSheet.cells[cellKey] || { raw: '' };
    const updatedCells = {
      ...activeSheet.cells,
      [cellKey]: {
        ...currentCell,
        raw: newRaw,
      },
    };

    const newSheets = data.sheets.map((s) =>
      s.id === activeSheet.id ? { ...s, cells: updatedCells } : s
    );
    onChange({ ...data, sheets: newSheets });
  };

  const commitEdit = () => {
    updateCellRaw(selectedCell, editValue);
    setIsEditing(false);
  };

  const cancelEdit = () => {
    setEditValue(activeSheet.cells[selectedCell]?.raw || '');
    setIsEditing(false);
  };

  // Keyboard navigation across the grid
  const handleKeyDown = (e: React.KeyboardEvent) => {
    const pos = FormulaEngine.parseCellId(selectedCell);
    if (!pos) return;

    if (isEditing) {
      if (e.key === 'Enter') {
        e.preventDefault();
        commitEdit();
        setSelectedCell(FormulaEngine.toCellId(pos.col, pos.row + 1));
      } else if (e.key === 'Escape') {
        cancelEdit();
      }
      return;
    }

    if (e.key === 'ArrowUp' && pos.row > 0) {
      e.preventDefault();
      setSelectedCell(FormulaEngine.toCellId(pos.col, pos.row - 1));
    } else if (e.key === 'ArrowDown' && pos.row < activeSheet.rowCount - 1) {
      e.preventDefault();
      setSelectedCell(FormulaEngine.toCellId(pos.col, pos.row + 1));
    } else if (e.key === 'ArrowLeft' && pos.col > 0) {
      e.preventDefault();
      setSelectedCell(FormulaEngine.toCellId(pos.col - 1, pos.row));
    } else if (e.key === 'ArrowRight' && pos.col < activeSheet.colCount - 1) {
      e.preventDefault();
      setSelectedCell(FormulaEngine.toCellId(pos.col + 1, pos.row));
    } else if (e.key === 'Enter') {
      e.preventDefault();
      setIsEditing(true);
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      updateCellRaw(selectedCell, '');
      setEditValue('');
    } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
      setIsEditing(true);
      setEditValue(e.key);
    }
  };

  const addSheet = () => {
    const newId = `sheet-${Date.now()}`;
    const newSheet = {
      id: newId,
      name: `Sheet${data.sheets.length + 1}`,
      rowCount: 30,
      colCount: 15,
      cells: {},
    };
    onChange({
      ...data,
      sheets: [...data.sheets, newSheet],
      activeSheetId: newId,
    });
  };

  const columns = useMemo(() => {
    return Array.from({ length: activeSheet.colCount }, (_, i) => String.fromCharCode(65 + i));
  }, [activeSheet.colCount]);

  const rows = useMemo(() => {
    return Array.from({ length: activeSheet.rowCount }, (_, i) => i + 1);
  }, [activeSheet.rowCount]);

  const selectedPos = FormulaEngine.parseCellId(selectedCell);

  return (
    <div
      className={`flex flex-col h-full w-full select-none outline-none text-xs ${
        darkMode ? 'bg-slate-900 text-slate-200' : 'bg-slate-100 text-slate-800'
      }`}
      onKeyDown={handleKeyDown}
      tabIndex={0}
    >
      {/* 1. AUTHENTIC MICROSOFT EXCEL FORMULA BAR */}
      <div
        className={`flex items-center px-2 py-1 border-b space-x-1.5 text-xs transition-colors shrink-0 ${
          darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-300'
        }`}
      >
        {/* Real Excel Name Box with drop arrow */}
        <div className="w-18 px-2 py-0.5 flex items-center justify-between font-mono font-bold rounded-sm border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-slate-800 dark:text-slate-200 text-xs shadow-2xs">
          <span>{selectedCell}</span>
          <span className="text-[9px] text-slate-400">▼</span>
        </div>

        {/* Formula Bar Cancel (✕) and Enter (✓) Buttons */}
        <div className="flex items-center space-x-0.5">
          <button
            onClick={cancelEdit}
            disabled={!isEditing}
            title="Cancel (Esc)"
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 disabled:opacity-30 text-rose-600"
          >
            <X className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={commitEdit}
            disabled={!isEditing}
            title="Enter (Return)"
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 disabled:opacity-30 text-emerald-600"
          >
            <Check className="w-3.5 h-3.5" />
          </button>
        </div>

        {/* Function Icon fx */}
        <div className="px-1 text-slate-500 font-serif italic font-bold text-sm">
          fx
        </div>

        {/* Real Excel Formula Input Box */}
        <input
          type="text"
          value={editValue}
          onChange={(e) => {
            setEditValue(e.target.value);
            updateCellRaw(selectedCell, e.target.value);
          }}
          onFocus={() => setIsEditing(true)}
          onBlur={commitEdit}
          placeholder="Enter a value or formula (e.g., =SUM(B2:E2))"
          className={`flex-1 px-2.5 py-0.5 font-mono text-xs rounded-sm border focus:outline-none focus:ring-1 focus:ring-[#107c41] ${
            darkMode
              ? 'bg-slate-800/90 border-slate-700 text-emerald-400'
              : 'bg-white border-slate-300 text-slate-900'
          }`}
        />
      </div>

      {/* 2. AUTHENTIC MICROSOFT EXCEL 2D GRID */}
      <div className="flex-1 overflow-auto bg-white dark:bg-slate-900">
        <table className="border-collapse table-fixed w-full">
          <thead>
            <tr className="sticky top-0 z-10 bg-[#f8f9fa] dark:bg-slate-800 border-b border-slate-300 dark:border-slate-700">
              {/* Corner Select All button */}
              <th className="w-10 min-w-10 p-1 border-r border-slate-300 dark:border-slate-700 text-center text-slate-400 font-mono text-[9px]">
                ◢
              </th>
              {columns.map((col, cIdx) => (
                <th
                  key={col}
                  className={`w-32 min-w-32 p-1 border-r border-slate-300 dark:border-slate-700 text-center font-semibold text-xs transition-colors ${
                    selectedPos?.col === cIdx
                      ? 'bg-[#107c41]/15 text-[#107c41] dark:text-emerald-400 font-bold'
                      : 'text-slate-600 dark:text-slate-400'
                  }`}
                >
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, rIdx) => (
              <tr key={row} className="border-b border-slate-200 dark:border-slate-800">
                {/* Excel Row Header */}
                <td
                  className={`sticky left-0 z-10 p-1 border-r border-slate-300 dark:border-slate-700 text-center font-mono text-[11px] font-medium transition-colors ${
                    selectedPos?.row === rIdx
                      ? 'bg-[#107c41]/15 text-[#107c41] dark:text-emerald-400 font-bold'
                      : 'bg-[#f8f9fa] dark:bg-slate-800 text-slate-500 dark:text-slate-400'
                  }`}
                >
                  {row}
                </td>

                {/* Cells */}
                {columns.map((col) => {
                  const cellKey = `${col}${row}`;
                  const cell = activeSheet.cells[cellKey];
                  const isSelected = selectedCell === cellKey;
                  const computed = computedSheet[cellKey];
                  const formatted = FormulaEngine.formatValue(computed?.value, cell?.style);

                  const isError =
                    typeof computed?.value === 'string' && computed.value.startsWith('#');

                  return (
                    <td
                      key={cellKey}
                      onClick={() => setSelectedCell(cellKey)}
                      onDoubleClick={() => {
                        setSelectedCell(cellKey);
                        setIsEditing(true);
                      }}
                      className={`relative p-1 border-r border-slate-200 dark:border-slate-800 select-none cursor-cell h-6 font-sans text-xs ${
                        isSelected
                          ? 'ring-2 ring-[#107c41] z-10'
                          : 'hover:bg-slate-50 dark:hover:bg-slate-800/40'
                      }`}
                      style={{
                        backgroundColor: cell?.style?.bgColor,
                        color: cell?.style?.textColor,
                        fontWeight: cell?.style?.bold ? 'bold' : 'normal',
                        fontStyle: cell?.style?.italic ? 'italic' : 'normal',
                        textAlign: cell?.style?.align || (typeof computed?.value === 'number' ? 'right' : 'left'),
                      }}
                    >
                      {/* Active cell small square drag handle (Signature Excel feature) */}
                      {isSelected && !isEditing && (
                        <div className="absolute -bottom-1 -right-1 w-2 h-2 bg-[#107c41] border border-white cursor-crosshair z-20" />
                      )}

                      {isSelected && isEditing ? (
                        <input
                          ref={inputRef}
                          type="text"
                          value={editValue}
                          onChange={(e) => setEditValue(e.target.value)}
                          onBlur={commitEdit}
                          className="absolute inset-0 w-full h-full px-1 font-mono text-xs bg-white dark:bg-slate-900 focus:outline-none ring-2 ring-[#107c41]"
                        />
                      ) : (
                        <span
                          className={`truncate block ${
                            isError ? 'text-rose-500 font-mono font-bold' : ''
                          }`}
                        >
                          {formatted}
                        </span>
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* 3. EXCEL WORKSHEET TABS BAR */}
      <div
        className={`flex items-center justify-between px-2 py-0.5 border-t text-xs select-none shrink-0 ${
          darkMode ? 'bg-slate-900 border-slate-800' : 'bg-[#f3f2f1] border-slate-300'
        }`}
      >
        <div className="flex items-center space-x-1 overflow-x-auto">
          {data.sheets.map((s) => (
            <button
              key={s.id}
              onClick={() => {
                onChange({ ...data, activeSheetId: s.id });
                setSelectedCell('A1');
              }}
              className={`flex items-center space-x-1.5 px-3 py-1 text-xs font-semibold rounded-t-sm transition-colors border-t border-x ${
                s.id === activeSheet.id
                  ? 'bg-white dark:bg-slate-800 text-[#107c41] dark:text-emerald-400 border-slate-300 dark:border-slate-700 shadow-2xs relative'
                  : 'bg-transparent text-slate-600 dark:text-slate-400 border-transparent hover:bg-slate-200/50'
              }`}
            >
              {s.id === activeSheet.id && (
                <span className="absolute bottom-0 left-0 right-0 h-0.5 bg-[#107c41]" />
              )}
              <Sheet className="w-3.5 h-3.5" />
              <span>{s.name}</span>
            </button>
          ))}

          <button
            onClick={addSheet}
            className="p-1 rounded hover:bg-slate-300 dark:hover:bg-slate-700 text-slate-600 dark:text-slate-400"
            title="New sheet (+)"
          >
            <Plus className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* 4. AUTHENTIC MICROSOFT EXCEL STATUS BAR */}
      <footer
        className={`h-6 border-t select-none px-3 flex items-center justify-between text-[11px] shrink-0 transition-colors ${
          darkMode ? 'bg-slate-900 border-slate-800 text-slate-400' : 'bg-[#107c41] text-white border-emerald-800'
        }`}
      >
        {/* Left: Mode */}
        <div className="flex items-center space-x-4">
          <span className="font-semibold">{isEditing ? 'Edit' : 'Ready'}</span>
          <span>•</span>
          <span className="text-[10px]">Active: {selectedCell}</span>
        </div>

        {/* Right: Live Calculation Summaries (Signature Excel Status Bar feature) */}
        <div className="flex items-center space-x-4">
          {gridStats.count > 0 && (
            <div className="flex items-center space-x-3 text-[11px] font-mono">
              <span>Average: ${gridStats.avg.toLocaleString()}</span>
              <span>Count: {gridStats.count}</span>
              <span className="font-bold">Sum: ${gridStats.sum.toLocaleString()}</span>
            </div>
          )}

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
