import React, { useState, useEffect, useRef, useMemo } from 'react';
import type {
  SpreadsheetData,
  CellStyle,
  ConditionalFormatRule,
  DataValidationRule,
} from '../types/office';
import { FormulaEngine } from '../engine/formulaEngine';
import {
  Plus,
  X,
  Check,
  Sheet,
  ZoomIn,
  ZoomOut,
  Bold,
  Italic,
  Underline,
  AlignLeft,
  AlignCenter,
  AlignRight,
  DollarSign,
  Percent,
  BarChart2,
  PieChart,
  TrendingUp,
  ArrowDownAZ,
  ArrowUpZA,
  Trash2,
  Copy,
  Table,
  Filter,
  CheckSquare,
  Lock,
} from 'lucide-react';

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
  const [selectionRange, setSelectionRange] = useState<{ start: string; end: string } | null>(null);
  const [isSelectingRange, setIsSelectingRange] = useState<boolean>(false);
  const [isAutofilling, setIsAutofilling] = useState<boolean>(false);
  const [colWidths, setColWidths] = useState<Record<number, number>>({});
  const [rowHeights, setRowHeights] = useState<Record<number, number>>({});
  const [resizingCol, setResizingCol] = useState<{ colIndex: number; startX: number; startWidth: number } | null>(null);
  const [resizingRow, setResizingRow] = useState<{ rowIndex: number; startY: number; startHeight: number } | null>(null);

  const [isEditing, setIsEditing] = useState<boolean>(false);
  const [editValue, setEditValue] = useState<string>('');
  const [zoomLevel, setZoomLevel] = useState<number>(100);

  // Modals state
  const [showChartModal, setShowChartModal] = useState<boolean>(false);
  const [chartType, setChartType] = useState<'bar' | 'line' | 'pie' | 'area'>('bar');
  const [showPivotModal, setShowPivotModal] = useState<boolean>(false);
  const [showConditionalModal, setShowConditionalModal] = useState<boolean>(false);
  const [showValidationModal, setShowValidationModal] = useState<boolean>(false);

  // Pivot Table state
  const [pivotRowCol, setPivotRowCol] = useState<string>('A');
  const [pivotValCol, setPivotValCol] = useState<string>('B');
  const [pivotAgg, setPivotAgg] = useState<'SUM' | 'COUNT' | 'AVERAGE' | 'MIN' | 'MAX'>('SUM');

  // Conditional Formatting State
  const [condRule, setCondRule] = useState<'greaterThan' | 'lessThan' | 'equals' | 'contains'>('greaterThan');
  const [condVal, setCondVal] = useState<string>('100');
  const [condBg, setCondBg] = useState<string>('#dcfce7');
  const [condText, setCondText] = useState<string>('#166534');

  // Data Validation State
  const [validationOptions, setValidationOptions] = useState<string>('Approved, In Review, Rejected');

  // Sheet rename state
  const [editingSheetNameId, setEditingSheetNameId] = useState<string | null>(null);
  const [sheetNameInput, setSheetNameInput] = useState<string>('');

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

  // Helper to determine if a cell is within active marquee range
  const isCellInRange = (cellKey: string): boolean => {
    if (!selectionRange) return cellKey === selectedCell;
    const p1 = FormulaEngine.parseCellId(selectionRange.start);
    const p2 = FormulaEngine.parseCellId(selectionRange.end);
    const cur = FormulaEngine.parseCellId(cellKey);
    if (!p1 || !p2 || !cur) return cellKey === selectedCell;

    const minRow = Math.min(p1.row, p2.row);
    const maxRow = Math.max(p1.row, p2.row);
    const minCol = Math.min(p1.col, p2.col);
    const maxCol = Math.max(p1.col, p2.col);

    return cur.row >= minRow && cur.row <= maxRow && cur.col >= minCol && cur.col <= maxCol;
  };

  // Live Excel Status Bar Calculations (Sum, Average, Count, Min, Max over selected range or sheet)
  const gridStats = useMemo(() => {
    let sum = 0;
    let count = 0;
    let min = Infinity;
    let max = -Infinity;

    let targetKeys: string[] = [];
    if (selectionRange) {
      const p1 = FormulaEngine.parseCellId(selectionRange.start);
      const p2 = FormulaEngine.parseCellId(selectionRange.end);
      if (p1 && p2) {
        const minR = Math.min(p1.row, p2.row);
        const maxR = Math.max(p1.row, p2.row);
        const minC = Math.min(p1.col, p2.col);
        const maxC = Math.max(p1.col, p2.col);
        for (let r = minR; r <= maxR; r++) {
          for (let c = minC; c <= maxC; c++) {
            targetKeys.push(FormulaEngine.toCellId(c, r));
          }
        }
      }
    } else {
      targetKeys = Object.keys(computedSheet);
    }

    for (const k of targetKeys) {
      const v = computedSheet[k]?.value;
      if (typeof v === 'number' && !isNaN(v)) {
        sum += v;
        count++;
        if (v < min) min = v;
        if (v > max) max = v;
      }
    }

    const avg = count > 0 ? Math.round((sum / count) * 100) / 100 : 0;
    return {
      sum,
      count,
      avg,
      min: min === Infinity ? 0 : min,
      max: max === -Infinity ? 0 : max,
    };
  }, [computedSheet, selectionRange]);

  // Autofill Execution: propagates values, incremental numbers, or formulas with relative row references
  const executeAutofill = (targetCellKey: string) => {
    const srcPos = FormulaEngine.parseCellId(selectedCell);
    const tgtPos = FormulaEngine.parseCellId(targetCellKey);
    if (!srcPos || !tgtPos) return;

    const srcRaw = activeSheet.cells[selectedCell]?.raw || '';
    const newCells = { ...activeSheet.cells };

    // Fill downwards
    if (tgtPos.col === srcPos.col && tgtPos.row > srcPos.row) {
      const numMatch = srcRaw.match(/^(\d+)$/);
      let baseNum = numMatch ? parseInt(numMatch[1], 10) : null;

      for (let r = srcPos.row + 1; r <= tgtPos.row; r++) {
        const k = FormulaEngine.toCellId(srcPos.col, r);
        if (baseNum !== null) {
          baseNum++;
          newCells[k] = { raw: String(baseNum) };
        } else if (srcRaw.startsWith('=')) {
          const rowDiff = r - srcPos.row;
          const adjusted = srcRaw.replace(
            /([A-Z]+)(\d+)/g,
            (_, c, rowNum) => `${c}${parseInt(rowNum, 10) + rowDiff}`
          );
          newCells[k] = { raw: adjusted };
        } else {
          newCells[k] = { raw: srcRaw };
        }
      }
    }
    const newSheets = data.sheets.map((s) => (s.id === activeSheet.id ? { ...s, cells: newCells } : s));
    onChange({ ...data, sheets: newSheets });
  };

  // Global mouse listeners for Column/Row resizing & Range selection release
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (resizingCol) {
        const diff = e.clientX - resizingCol.startX;
        const newW = Math.max(50, resizingCol.startWidth + diff);
        setColWidths((prev) => ({ ...prev, [resizingCol.colIndex]: newW }));
      } else if (resizingRow) {
        const diff = e.clientY - resizingRow.startY;
        const newH = Math.max(20, resizingRow.startHeight + diff);
        setRowHeights((prev) => ({ ...prev, [resizingRow.rowIndex]: newH }));
      }
    };

    const handleMouseUp = () => {
      setResizingCol(null);
      setResizingRow(null);
      setIsSelectingRange(false);
      setIsAutofilling(false);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [resizingCol, resizingRow]);

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

  const updateCellStyle = (cellKey: string, styleUpdates: Partial<CellStyle>) => {
    const currentCell = activeSheet.cells[cellKey] || { raw: '' };
    const currentStyle = currentCell.style || {};
    const updatedCells = {
      ...activeSheet.cells,
      [cellKey]: {
        ...currentCell,
        style: { ...currentStyle, ...styleUpdates },
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

  // Clipboard TSV Copy & Paste
  const copySelectedToClipboard = async () => {
    let tsv = '';
    if (selectionRange) {
      const p1 = FormulaEngine.parseCellId(selectionRange.start);
      const p2 = FormulaEngine.parseCellId(selectionRange.end);
      if (p1 && p2) {
        const minR = Math.min(p1.row, p2.row);
        const maxR = Math.max(p1.row, p2.row);
        const minC = Math.min(p1.col, p2.col);
        const maxC = Math.max(p1.col, p2.col);
        const rows: string[] = [];
        for (let r = minR; r <= maxR; r++) {
          const rowVals: string[] = [];
          for (let c = minC; c <= maxC; c++) {
            const k = FormulaEngine.toCellId(c, r);
            const cell = activeSheet.cells[k];
            rowVals.push(cell?.raw ?? '');
          }
          rows.push(rowVals.join('\t'));
        }
        tsv = rows.join('\n');
      }
    } else {
      tsv = activeSheet.cells[selectedCell]?.raw ?? '';
    }
    try {
      await navigator.clipboard.writeText(tsv);
    } catch (err) {
      console.warn('Clipboard write failed', err);
    }
  };

  const pasteFromClipboard = async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (!text) return;
      const startPos = FormulaEngine.parseCellId(selectedCell);
      if (!startPos) return;

      const rows = text.replace(/\r/g, '').split('\n');
      if (rows.length === 0) return;

      const newCells = { ...activeSheet.cells };
      let maxR = activeSheet.rowCount;
      let maxC = activeSheet.colCount;

      rows.forEach((rowStr, rOffset) => {
        if (rowStr === '' && rOffset === rows.length - 1) return;
        const cols = rowStr.split('\t');
        cols.forEach((val, cOffset) => {
          const targetR = startPos.row + rOffset;
          const targetC = startPos.col + cOffset;
          if (targetR + 1 > maxR) maxR = targetR + 1;
          if (targetC + 1 > maxC) maxC = targetC + 1;
          const key = FormulaEngine.toCellId(targetC, targetR);
          newCells[key] = { ...(newCells[key] || {}), raw: val };
        });
      });

      const newSheets = data.sheets.map((s) =>
        s.id === activeSheet.id ? { ...s, rowCount: maxR, colCount: maxC, cells: newCells } : s
      );
      onChange({ ...data, sheets: newSheets });
    } catch (err) {
      console.warn('Clipboard read failed', err);
    }
  };

  // Instant AutoSum (Alt+= or Sigma button)
  const insertAutoSum = () => {
    const pos = FormulaEngine.parseCellId(selectedCell);
    if (!pos) return;

    // 1. Check numbers directly above
    let startRowAbove = -1;
    for (let r = pos.row - 1; r >= 0; r--) {
      const k = FormulaEngine.toCellId(pos.col, r);
      const val = computedSheet[k]?.value;
      if (typeof val === 'number') {
        startRowAbove = r;
      } else {
        break;
      }
    }

    if (startRowAbove !== -1) {
      const formula = `=SUM(${FormulaEngine.toCellId(pos.col, startRowAbove)}:${FormulaEngine.toCellId(pos.col, pos.row - 1)})`;
      updateCellRaw(selectedCell, formula);
      setEditValue(formula);
      return;
    }

    // 2. Check numbers to the left
    let startColLeft = -1;
    for (let c = pos.col - 1; c >= 0; c--) {
      const k = FormulaEngine.toCellId(c, pos.row);
      const val = computedSheet[k]?.value;
      if (typeof val === 'number') {
        startColLeft = c;
      } else {
        break;
      }
    }

    if (startColLeft !== -1) {
      const formula = `=SUM(${FormulaEngine.toCellId(startColLeft, pos.row)}:${FormulaEngine.toCellId(pos.col - 1, pos.row)})`;
      updateCellRaw(selectedCell, formula);
      setEditValue(formula);
      return;
    }

    // Fallback default
    const defaultFormula = `=SUM(${FormulaEngine.toCellId(pos.col, 0)}:${FormulaEngine.toCellId(pos.col, Math.max(0, pos.row - 1))})`;
    updateCellRaw(selectedCell, defaultFormula);
    setEditValue(defaultFormula);
  };

  // Keyboard navigation across the grid
  const handleKeyDown = (e: React.KeyboardEvent) => {
    const pos = FormulaEngine.parseCellId(selectedCell);
    if (!pos) return;

    if (isEditing) {
      if (e.key === 'Enter') {
        e.preventDefault();
        commitEdit();
        if (e.shiftKey) {
          if (pos.row > 0) setSelectedCell(FormulaEngine.toCellId(pos.col, pos.row - 1));
        } else {
          if (pos.row < activeSheet.rowCount - 1) setSelectedCell(FormulaEngine.toCellId(pos.col, pos.row + 1));
        }
      } else if (e.key === 'Tab') {
        e.preventDefault();
        commitEdit();
        if (e.shiftKey) {
          if (pos.col > 0) setSelectedCell(FormulaEngine.toCellId(pos.col - 1, pos.row));
        } else {
          if (pos.col < activeSheet.colCount - 1) setSelectedCell(FormulaEngine.toCellId(pos.col + 1, pos.row));
        }
      } else if (e.key === 'Escape') {
        cancelEdit();
      }
      return;
    }

    // Shortcuts: Copy, Paste, AutoSum
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
      e.preventDefault();
      copySelectedToClipboard();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
      e.preventDefault();
      pasteFromClipboard();
      return;
    }
    if (e.altKey && e.key === '=') {
      e.preventDefault();
      insertAutoSum();
      return;
    }

    // Tab Navigation
    if (e.key === 'Tab') {
      e.preventDefault();
      if (e.shiftKey) {
        if (pos.col > 0) setSelectedCell(FormulaEngine.toCellId(pos.col - 1, pos.row));
      } else {
        if (pos.col < activeSheet.colCount - 1) setSelectedCell(FormulaEngine.toCellId(pos.col + 1, pos.row));
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
      if (e.shiftKey) {
        if (pos.row > 0) setSelectedCell(FormulaEngine.toCellId(pos.col, pos.row - 1));
      } else {
        setIsEditing(true);
      }
    } else if (e.key === 'F2') {
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

  // Grid Structure Manipulations: Insert/Delete Row & Col
  const insertRow = (above: boolean = false) => {
    const pos = FormulaEngine.parseCellId(selectedCell);
    if (!pos) return;

    const targetRow = above ? pos.row + 1 : pos.row + 2;
    const newCells: Record<string, any> = {};

    for (const [key, cell] of Object.entries(activeSheet.cells)) {
      const p = FormulaEngine.parseCellId(key);
      if (!p) continue;
      const r = p.row + 1;
      if (r >= targetRow) {
        newCells[FormulaEngine.toCellId(p.col, r)] = cell;
      } else {
        newCells[key] = cell;
      }
    }

    const newSheets = data.sheets.map((s) =>
      s.id === activeSheet.id ? { ...s, rowCount: s.rowCount + 1, cells: newCells } : s
    );
    onChange({ ...data, sheets: newSheets });
  };

  const deleteRow = () => {
    const pos = FormulaEngine.parseCellId(selectedCell);
    if (!pos || activeSheet.rowCount <= 5) return;

    const targetRow = pos.row + 1;
    const newCells: Record<string, any> = {};

    for (const [key, cell] of Object.entries(activeSheet.cells)) {
      const p = FormulaEngine.parseCellId(key);
      if (!p) continue;
      const r = p.row + 1;
      if (r === targetRow) continue;
      if (r > targetRow) {
        newCells[FormulaEngine.toCellId(p.col, r - 2)] = cell;
      } else {
        newCells[key] = cell;
      }
    }

    const newSheets = data.sheets.map((s) =>
      s.id === activeSheet.id ? { ...s, rowCount: s.rowCount - 1, cells: newCells } : s
    );
    onChange({ ...data, sheets: newSheets });
  };

  const insertCol = (left: boolean = false) => {
    const pos = FormulaEngine.parseCellId(selectedCell);
    if (!pos) return;

    const targetCol = left ? pos.col : pos.col + 1;
    const newCells: Record<string, any> = {};

    for (const [key, cell] of Object.entries(activeSheet.cells)) {
      const p = FormulaEngine.parseCellId(key);
      if (!p) continue;
      if (p.col >= targetCol) {
        newCells[FormulaEngine.toCellId(p.col + 1, p.row)] = cell;
      } else {
        newCells[key] = cell;
      }
    }

    const newSheets = data.sheets.map((s) =>
      s.id === activeSheet.id ? { ...s, colCount: s.colCount + 1, cells: newCells } : s
    );
    onChange({ ...data, sheets: newSheets });
  };

  const deleteCol = () => {
    const pos = FormulaEngine.parseCellId(selectedCell);
    if (!pos || activeSheet.colCount <= 3) return;

    const targetCol = pos.col;
    const newCells: Record<string, any> = {};

    for (const [key, cell] of Object.entries(activeSheet.cells)) {
      const p = FormulaEngine.parseCellId(key);
      if (!p) continue;
      if (p.col === targetCol) continue;
      if (p.col > targetCol) {
        newCells[FormulaEngine.toCellId(p.col - 1, p.row)] = cell;
      } else {
        newCells[key] = cell;
      }
    }

    const newSheets = data.sheets.map((s) =>
      s.id === activeSheet.id ? { ...s, colCount: s.colCount - 1, cells: newCells } : s
    );
    onChange({ ...data, sheets: newSheets });
  };

  // Sort rows based on selected column
  const sortColumn = (ascending: boolean = true) => {
    const pos = FormulaEngine.parseCellId(selectedCell);
    if (!pos) return;

    const colLetter = String.fromCharCode(65 + pos.col);
    const rowIndices: number[] = [];
    for (let r = 2; r <= activeSheet.rowCount; r++) {
      rowIndices.push(r);
    }

    rowIndices.sort((a, b) => {
      const valA = computedSheet[`${colLetter}${a}`]?.value ?? '';
      const valB = computedSheet[`${colLetter}${b}`]?.value ?? '';
      if (typeof valA === 'number' && typeof valB === 'number') {
        return ascending ? valA - valB : valB - valA;
      }
      return ascending
        ? String(valA).localeCompare(String(valB))
        : String(valB).localeCompare(String(valA));
    });

    const newCells: Record<string, any> = {};
    for (let c = 0; c < activeSheet.colCount; c++) {
      const key = `${String.fromCharCode(65 + c)}1`;
      if (activeSheet.cells[key]) newCells[key] = activeSheet.cells[key];
    }

    rowIndices.forEach((origRow, newIdx) => {
      const targetRow = newIdx + 2;
      for (let c = 0; c < activeSheet.colCount; c++) {
        const origKey = `${String.fromCharCode(65 + c)}${origRow}`;
        const targetKey = `${String.fromCharCode(65 + c)}${targetRow}`;
        if (activeSheet.cells[origKey]) {
          newCells[targetKey] = activeSheet.cells[origKey];
        }
      }
    });

    const newSheets = data.sheets.map((s) =>
      s.id === activeSheet.id ? { ...s, cells: newCells } : s
    );
    onChange({ ...data, sheets: newSheets });
  };

  // Multi-sheet operations
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

  const duplicateActiveSheet = () => {
    const newId = `sheet-${Date.now()}`;
    const clone = {
      ...activeSheet,
      id: newId,
      name: `${activeSheet.name} (Copy)`,
      cells: { ...activeSheet.cells },
    };
    onChange({
      ...data,
      sheets: [...data.sheets, clone],
      activeSheetId: newId,
    });
  };

  const deleteActiveSheet = () => {
    if (data.sheets.length <= 1) return;
    const remaining = data.sheets.filter((s) => s.id !== activeSheet.id);
    onChange({
      ...data,
      sheets: remaining,
      activeSheetId: remaining[0].id,
    });
  };

  const commitSheetRename = () => {
    if (editingSheetNameId && sheetNameInput.trim()) {
      const newSheets = data.sheets.map((s) =>
        s.id === editingSheetNameId ? { ...s, name: sheetNameInput.trim() } : s
      );
      onChange({ ...data, sheets: newSheets });
    }
    setEditingSheetNameId(null);
  };

  // Dynamic Conditional Formatting check
  const getConditionalStyle = (_cellKey: string, val: string | number | undefined) => {
    if (!activeSheet.conditionalFormats || val === undefined || val === '') return null;
    const num = Number(val);
    const isNum = !isNaN(num);

    for (const rule of activeSheet.conditionalFormats) {
      if (rule.rule === 'greaterThan' && isNum && num > Number(rule.value)) {
        return { backgroundColor: rule.bgColor, color: rule.textColor };
      }
      if (rule.rule === 'lessThan' && isNum && num < Number(rule.value)) {
        return { backgroundColor: rule.bgColor, color: rule.textColor };
      }
      if (rule.rule === 'equals' && String(val).toLowerCase() === String(rule.value).toLowerCase()) {
        return { backgroundColor: rule.bgColor, color: rule.textColor };
      }
      if (rule.rule === 'contains' && String(val).toLowerCase().includes(String(rule.value).toLowerCase())) {
        return { backgroundColor: rule.bgColor, color: rule.textColor };
      }
    }
    return null;
  };

  const addConditionalRule = () => {
    const newRule: ConditionalFormatRule = {
      id: `cf-${Date.now()}`,
      range: 'All',
      rule: condRule,
      value: condVal,
      bgColor: condBg,
      textColor: condText,
    };
    const current = activeSheet.conditionalFormats || [];
    const newSheets = data.sheets.map((s) =>
      s.id === activeSheet.id ? { ...s, conditionalFormats: [...current, newRule] } : s
    );
    onChange({ ...data, sheets: newSheets });
    setShowConditionalModal(false);
  };

  // Add Data Validation
  const addValidation = () => {
    const options = validationOptions.split(',').map((o) => o.trim()).filter(Boolean);
    const newRule: DataValidationRule = {
      range: selectedCell,
      type: 'list',
      options,
    };
    const current = activeSheet.validations || [];
    const newSheets = data.sheets.map((s) =>
      s.id === activeSheet.id ? { ...s, validations: [...current, newRule] } : s
    );
    onChange({ ...data, sheets: newSheets });
    setShowValidationModal(false);
  };

  const cellValidation = (activeSheet.validations || []).find((v) => v.range === selectedCell);

  // Dynamic Pivot Table Engine
  const pivotResults = useMemo(() => {
    const groups: Record<string, number[]> = {};

    for (let r = 2; r <= activeSheet.rowCount; r++) {
      const rowKey = String(computedSheet[`${pivotRowCol}${r}`]?.value || '').trim();
      const valRaw = computedSheet[`${pivotValCol}${r}`]?.value;
      const num = typeof valRaw === 'number' ? valRaw : Number(valRaw);

      if (rowKey && !isNaN(num)) {
        if (!groups[rowKey]) groups[rowKey] = [];
        groups[rowKey].push(num);
      }
    }

    const rowsSummary: { key: string; val: number }[] = [];
    for (const [k, nums] of Object.entries(groups)) {
      let aggVal = 0;
      if (pivotAgg === 'SUM') aggVal = nums.reduce((a, b) => a + b, 0);
      else if (pivotAgg === 'COUNT') aggVal = nums.length;
      else if (pivotAgg === 'AVERAGE') aggVal = Math.round((nums.reduce((a, b) => a + b, 0) / nums.length) * 100) / 100;
      else if (pivotAgg === 'MIN') aggVal = Math.min(...nums);
      else if (pivotAgg === 'MAX') aggVal = Math.max(...nums);

      rowsSummary.push({ key: k, val: aggVal });
    }

    return rowsSummary;
  }, [computedSheet, pivotRowCol, pivotValCol, pivotAgg, activeSheet.rowCount]);

  const insertPivotAsNewSheet = () => {
    const newId = `sheet-pivot-${Date.now()}`;
    const newCells: Record<string, any> = {
      A1: { raw: `${pivotRowCol} (Row Labels)` },
      B1: { raw: `${pivotAgg} of ${pivotValCol}` },
    };

    pivotResults.forEach((row, i) => {
      newCells[`A${i + 2}`] = { raw: row.key };
      newCells[`B${i + 2}`] = { raw: String(row.val) };
    });

    const newSheet = {
      id: newId,
      name: `Pivot - ${activeSheet.name}`,
      rowCount: Math.max(30, pivotResults.length + 5),
      colCount: 10,
      cells: newCells,
    };

    onChange({
      ...data,
      sheets: [...data.sheets, newSheet],
      activeSheetId: newId,
    });
    setShowPivotModal(false);
  };

  // Extract chart data from columns A & B
  const chartData = useMemo(() => {
    const items: { label: string; value: number }[] = [];
    for (let r = 2; r <= Math.min(activeSheet.rowCount, 25); r++) {
      const label = String(computedSheet[`A${r}`]?.value || `Row ${r}`);
      const rawVal = computedSheet[`B${r}`]?.value;
      const num = typeof rawVal === 'number' ? rawVal : Number(rawVal);
      if (!isNaN(num) && num > 0) {
        items.push({ label, value: num });
      }
    }
    return items;
  }, [computedSheet, activeSheet.rowCount]);

  const columns = useMemo(() => {
    return Array.from({ length: activeSheet.colCount }, (_, i) => String.fromCharCode(65 + i));
  }, [activeSheet.colCount]);

  const rows = useMemo(() => {
    return Array.from({ length: activeSheet.rowCount }, (_, i) => i + 1);
  }, [activeSheet.rowCount]);

  const selectedPos = FormulaEngine.parseCellId(selectedCell);
  const currentCellStyle = activeSheet.cells[selectedCell]?.style || {};

  return (
    <div
      className={`flex flex-col h-full w-full select-none outline-none text-xs relative ${
        darkMode ? 'bg-slate-900 text-slate-200' : 'bg-slate-100 text-slate-800'
      }`}
      onKeyDown={handleKeyDown}
      tabIndex={0}
    >
      {/* 1. EXCEL QUICK FORMATTING & ENTERPRISE TOOLS BAR */}
      <div
        className={`flex items-center justify-between px-2 py-1 border-b text-xs shrink-0 select-none ${
          darkMode ? 'bg-slate-950 border-slate-800' : 'bg-slate-100 border-slate-300'
        }`}
      >
        <div className="flex items-center space-x-1">
          {/* Font Styles */}
          <button
            onClick={() => updateCellStyle(selectedCell, { bold: !currentCellStyle.bold })}
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              currentCellStyle.bold ? 'bg-slate-200 dark:bg-slate-800 text-blue-600 font-bold' : ''
            }`}
            title="Bold"
          >
            <Bold className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => updateCellStyle(selectedCell, { italic: !currentCellStyle.italic })}
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              currentCellStyle.italic ? 'bg-slate-200 dark:bg-slate-800 text-blue-600' : ''
            }`}
            title="Italic"
          >
            <Italic className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => updateCellStyle(selectedCell, { underline: !currentCellStyle.underline })}
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              currentCellStyle.underline ? 'bg-slate-200 dark:bg-slate-800 text-blue-600' : ''
            }`}
            title="Underline"
          >
            <Underline className="w-3.5 h-3.5" />
          </button>

          <div className="h-4 w-px bg-slate-300 dark:bg-slate-700 mx-1" />

          {/* Number Formats */}
          <button
            onClick={() =>
              updateCellStyle(selectedCell, {
                format: currentCellStyle.format === 'currency' ? 'text' : 'currency',
              })
            }
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              currentCellStyle.format === 'currency' ? 'bg-emerald-100 dark:bg-emerald-950/60 text-emerald-600 font-bold' : ''
            }`}
            title="Currency Format ($)"
          >
            <DollarSign className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() =>
              updateCellStyle(selectedCell, {
                format: currentCellStyle.format === 'percent' ? 'text' : 'percent',
              })
            }
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              currentCellStyle.format === 'percent' ? 'bg-emerald-100 dark:bg-emerald-950/60 text-emerald-600 font-bold' : ''
            }`}
            title="Percent Format (%)"
          >
            <Percent className="w-3.5 h-3.5" />
          </button>

          <div className="h-4 w-px bg-slate-300 dark:bg-slate-700 mx-1" />

          {/* Alignment */}
          <button
            onClick={() => updateCellStyle(selectedCell, { align: 'left' })}
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
            title="Align Left"
          >
            <AlignLeft className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => updateCellStyle(selectedCell, { align: 'center' })}
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
            title="Align Center"
          >
            <AlignCenter className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => updateCellStyle(selectedCell, { align: 'right' })}
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
            title="Align Right"
          >
            <AlignRight className="w-3.5 h-3.5" />
          </button>

          <div className="h-4 w-px bg-slate-300 dark:bg-slate-700 mx-1" />

          {/* Insert / Delete Rows & Cols */}
          <button
            onClick={() => insertRow(false)}
            className="px-2 py-0.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-[11px] font-medium"
            title="Insert Row Below"
          >
            + Row
          </button>
          <button
            onClick={deleteRow}
            className="px-2 py-0.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-[11px] text-rose-500 font-medium"
            title="Delete Current Row"
          >
            - Row
          </button>
          <button
            onClick={() => insertCol(false)}
            className="px-2 py-0.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-[11px] font-medium"
            title="Insert Column Right"
          >
            + Col
          </button>
          <button
            onClick={deleteCol}
            className="px-2 py-0.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-[11px] text-rose-500 font-medium"
            title="Delete Current Column"
          >
            - Col
          </button>

          <div className="h-4 w-px bg-slate-300 dark:bg-slate-700 mx-1" />

          {/* Freeze Header Toggle */}
          <button
            onClick={() => {
              const newSheets = data.sheets.map((s) =>
                s.id === activeSheet.id ? { ...s, freezeHeader: !s.freezeHeader } : s
              );
              onChange({ ...data, sheets: newSheets });
            }}
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 flex items-center space-x-1 ${
              activeSheet.freezeHeader ? 'text-emerald-600 font-bold' : ''
            }`}
            title="Freeze Header Row"
          >
            <Lock className="w-3 h-3" />
            <span className="text-[10px]">Freeze Row 1</span>
          </button>

          <div className="h-4 w-px bg-slate-300 dark:bg-slate-700 mx-1" />

          {/* Sort */}
          <button
            onClick={() => sortColumn(true)}
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
            title="Sort Ascending (A-Z)"
          >
            <ArrowDownAZ className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => sortColumn(false)}
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
            title="Sort Descending (Z-A)"
          >
            <ArrowUpZA className="w-3.5 h-3.5" />
          </button>

          <div className="h-4 w-px bg-slate-300 dark:bg-slate-700 mx-1" />

          {/* Microsoft AutoSum */}
          <button
            onClick={insertAutoSum}
            className="px-2 py-0.5 rounded bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-[11px] flex items-center space-x-1 shadow-xs"
            title="AutoSum (Alt+=) - Automatically sums contiguous numbers above or left"
          >
            <span className="font-serif italic text-xs">∑</span>
            <span>AutoSum</span>
          </button>

          {/* Quick Copy / Paste */}
          <button
            onClick={copySelectedToClipboard}
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-400"
            title="Copy to Clipboard (Ctrl+C)"
          >
            <Copy className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={pasteFromClipboard}
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-400"
            title="Paste from Clipboard (Ctrl+V)"
          >
            <CheckSquare className="w-3.5 h-3.5" />
          </button>
        </div>

        {/* Enterprise Tools: Pivot Table, Conditional Formatting, Data Validation, Charts */}
        <div className="flex items-center space-x-1.5">
          <button
            onClick={() => setShowConditionalModal(true)}
            className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800 text-[11px] flex items-center space-x-1"
            title="Conditional Formatting"
          >
            <Filter className="w-3 h-3 text-amber-500" />
            <span>Conditional Rules</span>
          </button>

          <button
            onClick={() => setShowValidationModal(true)}
            className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800 text-[11px] flex items-center space-x-1"
            title="Data Validation Dropdowns"
          >
            <CheckSquare className="w-3 h-3 text-blue-500" />
            <span>Dropdown</span>
          </button>

          <button
            onClick={() => setShowPivotModal(true)}
            className="px-2.5 py-0.5 rounded bg-blue-600 hover:bg-blue-700 text-white font-medium text-[11px] flex items-center space-x-1 shadow-2xs"
            title="Pivot Table Builder"
          >
            <Table className="w-3 h-3" />
            <span>Pivot Table</span>
          </button>

          <button
            onClick={() => setShowChartModal(true)}
            className="flex items-center space-x-1 px-2.5 py-0.5 rounded bg-[#107c41] hover:bg-emerald-700 text-white font-medium text-[11px] shadow-2xs"
            title="Visualize Data Range as a Chart"
          >
            <BarChart2 className="w-3 h-3" />
            <span>Quick Chart</span>
          </button>
        </div>
      </div>

      {/* 2. AUTHENTIC MICROSOFT EXCEL FORMULA BAR */}
      <div
        className={`flex items-center px-2 py-1 border-b space-x-1.5 text-xs transition-colors shrink-0 ${
          darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-300'
        }`}
      >
        {/* Name Box */}
        <div className="w-18 px-2 py-0.5 flex items-center justify-between font-mono font-bold rounded-sm border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-slate-800 dark:text-slate-200 text-xs shadow-2xs">
          <span>{selectedCell}</span>
          <span className="text-[9px] text-slate-400">▼</span>
        </div>

        {/* Cancel & Enter */}
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
        <div className="px-1 text-slate-500 font-serif italic font-bold text-sm select-none">
          fx
        </div>

        {/* Formula Input / Dropdown Selector if Validation exists */}
        {cellValidation ? (
          <select
            value={editValue}
            onChange={(e) => {
              setEditValue(e.target.value);
              updateCellRaw(selectedCell, e.target.value);
            }}
            className="flex-1 px-2.5 py-0.5 text-xs rounded-sm border focus:outline-none ring-1 ring-blue-500 bg-white dark:bg-slate-800"
          >
            <option value="">-- Select Valid Option --</option>
            {cellValidation.options.map((opt) => (
              <option key={opt} value={opt}>
                {opt}
              </option>
            ))}
          </select>
        ) : (
          <input
            type="text"
            value={editValue}
            onChange={(e) => {
              setEditValue(e.target.value);
              updateCellRaw(selectedCell, e.target.value);
            }}
            onFocus={() => setIsEditing(true)}
            onBlur={commitEdit}
            placeholder="Enter a value or formula (e.g. =SUM(B2:B10) or =VLOOKUP(A2, A1:D10, 2))"
            className={`flex-1 px-2.5 py-0.5 font-mono text-xs rounded-sm border focus:outline-none focus:ring-1 focus:ring-[#107c41] ${
              darkMode
                ? 'bg-slate-800/90 border-slate-700 text-emerald-400'
                : 'bg-white border-slate-300 text-slate-900'
            }`}
          />
        )}
      </div>

      {/* 3. AUTHENTIC MICROSOFT EXCEL 2D GRID */}
      <div className="flex-1 overflow-auto bg-white dark:bg-slate-900">
        <table className="border-collapse table-fixed w-full">
          <thead>
            <tr className="sticky top-0 z-10 bg-[#f8f9fa] dark:bg-slate-800 border-b border-slate-300 dark:border-slate-700 shadow-2xs">
              <th className="w-10 min-w-10 p-1 border-r border-slate-300 dark:border-slate-700 text-center text-slate-400 font-mono text-[9px]">
                ◢
              </th>
              {columns.map((col, cIdx) => (
                <th
                  key={col}
                  style={{ width: colWidths[cIdx] || 128, minWidth: colWidths[cIdx] || 128 }}
                  className={`relative p-1 border-r border-slate-300 dark:border-slate-700 text-center font-semibold text-xs transition-colors select-none ${
                    selectedPos?.col === cIdx
                      ? 'bg-[#107c41]/15 text-[#107c41] dark:text-emerald-400 font-bold'
                      : 'text-slate-600 dark:text-slate-400'
                  }`}
                >
                  {col}
                  {/* Column Resizing Divider */}
                  <div
                    onMouseDown={(e) => {
                      e.stopPropagation();
                      setResizingCol({ colIndex: cIdx, startX: e.clientX, startWidth: colWidths[cIdx] || 128 });
                    }}
                    className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-[#107c41] z-20"
                    title="Drag to resize column"
                  />
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, rIdx) => {
              const isFrozenRow = activeSheet.freezeHeader && row === 1;

              return (
                <tr
                  key={row}
                  className={`border-b border-slate-200 dark:border-slate-800 ${
                    isFrozenRow ? 'sticky top-6 z-10 bg-slate-50 dark:bg-slate-800/90 shadow-xs' : ''
                  }`}
                >
                  <td
                    style={{ height: rowHeights[rIdx] || 24 }}
                    className={`relative sticky left-0 z-10 p-1 border-r border-slate-300 dark:border-slate-700 text-center font-mono text-[11px] font-medium transition-colors select-none ${
                      selectedPos?.row === rIdx
                        ? 'bg-[#107c41]/15 text-[#107c41] dark:text-emerald-400 font-bold'
                        : 'bg-[#f8f9fa] dark:bg-slate-800 text-slate-500 dark:text-slate-400'
                    }`}
                  >
                    {row}
                    {/* Row Resizing Divider */}
                    <div
                      onMouseDown={(e) => {
                        e.stopPropagation();
                        setResizingRow({ rowIndex: rIdx, startY: e.clientY, startHeight: rowHeights[rIdx] || 24 });
                      }}
                      className="absolute bottom-0 left-0 right-0 h-1.5 cursor-row-resize hover:bg-[#107c41] z-20"
                      title="Drag to resize row"
                    />
                  </td>

                  {columns.map((col) => {
                    const cellKey = `${col}${row}`;
                    const cell = activeSheet.cells[cellKey];
                    const isSelected = selectedCell === cellKey;
                    const inRange = isCellInRange(cellKey);
                    const computed = computedSheet[cellKey];
                    const formatted = FormulaEngine.formatValue(computed?.value, cell?.style);

                    const isError =
                      typeof computed?.value === 'string' && computed.value.startsWith('#');

                    const condStyle = getConditionalStyle(cellKey, computed?.value);

                    return (
                      <td
                        key={cellKey}
                        onMouseDown={(e) => {
                          if (e.button === 0) {
                            setSelectedCell(cellKey);
                            setSelectionRange({ start: cellKey, end: cellKey });
                            setIsSelectingRange(true);
                          }
                        }}
                        onMouseEnter={() => {
                          if (isSelectingRange && selectionRange) {
                            setSelectionRange({ start: selectionRange.start, end: cellKey });
                          }
                          if (isAutofilling) {
                            executeAutofill(cellKey);
                          }
                        }}
                        onDoubleClick={() => {
                          setSelectedCell(cellKey);
                          setIsEditing(true);
                        }}
                        className={`relative p-1 border-r border-slate-200 dark:border-slate-800 select-none cursor-cell h-6 font-sans text-xs ${
                          isSelected
                            ? 'ring-2 ring-[#107c41] z-10'
                            : inRange
                            ? 'bg-emerald-500/15 border border-emerald-500/40 z-5'
                            : 'hover:bg-slate-50 dark:hover:bg-slate-800/40'
                        }`}
                        style={{
                          backgroundColor: condStyle?.backgroundColor || cell?.style?.bgColor,
                          color: condStyle?.color || cell?.style?.textColor,
                          fontWeight: cell?.style?.bold ? 'bold' : 'normal',
                          fontStyle: cell?.style?.italic ? 'italic' : 'normal',
                          textDecoration: cell?.style?.underline ? 'underline' : 'none',
                          textAlign: cell?.style?.align || (typeof computed?.value === 'number' ? 'right' : 'left'),
                        }}
                      >
                        {isSelected && !isEditing && (
                          <div
                            onMouseDown={(e) => {
                              e.stopPropagation();
                              setIsAutofilling(true);
                            }}
                            className="absolute -bottom-1 -right-1 w-2.5 h-2.5 bg-[#107c41] border border-white cursor-crosshair z-20 shadow-xs"
                            title="Drag to autofill formula or series"
                          />
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
              );
            })}
          </tbody>
        </table>
      </div>

      {/* 4. EXCEL WORKSHEET TABS BAR */}
      <div
        className={`flex items-center justify-between px-2 py-0.5 border-t text-xs select-none shrink-0 ${
          darkMode ? 'bg-slate-900 border-slate-800' : 'bg-[#f3f2f1] border-slate-300'
        }`}
      >
        <div className="flex items-center space-x-1 overflow-x-auto">
          {data.sheets.map((s) => (
            <div key={s.id} className="flex items-center">
              {editingSheetNameId === s.id ? (
                <input
                  type="text"
                  value={sheetNameInput}
                  onChange={(e) => setSheetNameInput(e.target.value)}
                  onBlur={commitSheetRename}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') commitSheetRename();
                    if (e.key === 'Escape') setEditingSheetNameId(null);
                  }}
                  autoFocus
                  className="px-2 py-0.5 text-xs rounded border border-emerald-500 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100"
                />
              ) : (
                <button
                  onClick={() => {
                    onChange({ ...data, activeSheetId: s.id });
                    setSelectedCell('A1');
                  }}
                  onDoubleClick={() => {
                    setEditingSheetNameId(s.id);
                    setSheetNameInput(s.name);
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
              )}
            </div>
          ))}

          <button
            onClick={addSheet}
            className="p-1 rounded hover:bg-slate-300 dark:hover:bg-slate-700 text-slate-600 dark:text-slate-400"
            title="New Sheet (+)"
          >
            <Plus className="w-3.5 h-3.5" />
          </button>
        </div>

        <div className="flex items-center space-x-1">
          <button
            onClick={duplicateActiveSheet}
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-500"
            title="Duplicate Current Sheet"
          >
            <Copy className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={deleteActiveSheet}
            disabled={data.sheets.length <= 1}
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 disabled:opacity-30 text-rose-500"
            title="Delete Current Sheet"
          >
            <Trash2 className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* 5. AUTHENTIC MICROSOFT EXCEL STATUS BAR */}
      <footer
        className={`h-6 border-t select-none px-3 flex items-center justify-between text-[11px] shrink-0 transition-colors ${
          darkMode ? 'bg-slate-900 border-slate-800 text-slate-400' : 'bg-[#107c41] text-white border-emerald-800'
        }`}
      >
        <div className="flex items-center space-x-4">
          <span className="font-semibold">{isEditing ? 'Edit' : 'Ready'}</span>
          <span>•</span>
          <span className="text-[10px] font-mono">
            {selectionRange && selectionRange.start !== selectionRange.end
              ? `Range: ${selectionRange.start}:${selectionRange.end}`
              : `Cell: ${selectedCell}`}
          </span>
          {activeSheet.freezeHeader && <span className="text-[10px] text-amber-300">• Freeze Row 1 ON</span>}
        </div>

        <div className="flex items-center space-x-4">
          {gridStats.count > 0 && (
            <div className="flex items-center space-x-3 text-[11px] font-mono">
              <span>Avg: ${gridStats.avg.toLocaleString()}</span>
              <span>Count: {gridStats.count}</span>
              <span>Min: ${gridStats.min.toLocaleString()}</span>
              <span>Max: ${gridStats.max.toLocaleString()}</span>
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

      {/* 6. MODAL: PIVOT TABLE BUILDER */}
      {showPivotModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4">
          <div
            className={`w-full max-w-xl rounded-xl shadow-2xl border p-6 flex flex-col space-y-4 ${
              darkMode ? 'bg-slate-900 border-slate-700 text-slate-100' : 'bg-white border-slate-200 text-slate-900'
            }`}
          >
            <div className="flex items-center justify-between border-b pb-3 dark:border-slate-800">
              <div className="flex items-center space-x-2">
                <Table className="w-5 h-5 text-blue-500" />
                <h3 className="font-bold text-base">Microsoft Excel Pivot Table Wizard</h3>
              </div>
              <button onClick={() => setShowPivotModal(false)}>
                <X className="w-5 h-5 text-slate-400" />
              </button>
            </div>

            <div className="grid grid-cols-3 gap-3 text-xs">
              <div>
                <label className="font-semibold text-slate-500">Row Labels Field:</label>
                <select
                  value={pivotRowCol}
                  onChange={(e) => setPivotRowCol(e.target.value)}
                  className="w-full mt-1 p-2 rounded border dark:bg-slate-800 dark:border-slate-700"
                >
                  {columns.map((c) => (
                    <option key={c} value={c}>
                      Column {c} ({computedSheet[`${c}1`]?.value || `Header ${c}`})
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="font-semibold text-slate-500">Values Field:</label>
                <select
                  value={pivotValCol}
                  onChange={(e) => setPivotValCol(e.target.value)}
                  className="w-full mt-1 p-2 rounded border dark:bg-slate-800 dark:border-slate-700"
                >
                  {columns.map((c) => (
                    <option key={c} value={c}>
                      Column {c} ({computedSheet[`${c}1`]?.value || `Header ${c}`})
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="font-semibold text-slate-500">Aggregation:</label>
                <select
                  value={pivotAgg}
                  onChange={(e) => setPivotAgg(e.target.value as any)}
                  className="w-full mt-1 p-2 rounded border dark:bg-slate-800 dark:border-slate-700 font-bold"
                >
                  <option value="SUM">SUM</option>
                  <option value="COUNT">COUNT</option>
                  <option value="AVERAGE">AVERAGE</option>
                  <option value="MIN">MIN</option>
                  <option value="MAX">MAX</option>
                </select>
              </div>
            </div>

            {/* Pivot Preview */}
            <div className="border rounded-lg p-3 max-h-56 overflow-y-auto bg-slate-50 dark:bg-slate-950/60 dark:border-slate-800">
              <table className="w-full text-xs">
                <thead>
                  <tr className="border-b font-bold text-slate-600 dark:text-slate-300">
                    <th className="text-left pb-1.5">Row Labels</th>
                    <th className="text-right pb-1.5">{pivotAgg} of Column {pivotValCol}</th>
                  </tr>
                </thead>
                <tbody>
                  {pivotResults.length === 0 ? (
                    <tr>
                      <td colSpan={2} className="py-4 text-center text-slate-400">
                        No grouped data found. Make sure your values column has numbers.
                      </td>
                    </tr>
                  ) : (
                    pivotResults.map((r, i) => (
                      <tr key={i} className="border-b dark:border-slate-800/60">
                        <td className="py-1">{r.key}</td>
                        <td className="py-1 text-right font-mono font-bold text-blue-600 dark:text-blue-400">
                          ${r.val.toLocaleString()}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            <div className="flex justify-end space-x-2 pt-2">
              <button
                onClick={() => setShowPivotModal(false)}
                className="px-3 py-1.5 rounded text-xs border dark:border-slate-700"
              >
                Cancel
              </button>
              <button
                onClick={insertPivotAsNewSheet}
                disabled={pivotResults.length === 0}
                className="px-4 py-1.5 rounded text-xs bg-blue-600 hover:bg-blue-700 text-white font-medium"
              >
                Insert Pivot Table as New Sheet
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 7. MODAL: CONDITIONAL FORMATTING */}
      {showConditionalModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4">
          <div
            className={`w-full max-w-md rounded-xl shadow-2xl border p-6 flex flex-col space-y-4 ${
              darkMode ? 'bg-slate-900 border-slate-700 text-slate-100' : 'bg-white border-slate-200 text-slate-900'
            }`}
          >
            <div className="flex items-center justify-between border-b pb-2 dark:border-slate-800">
              <h3 className="font-bold text-base flex items-center space-x-2">
                <Filter className="w-4 h-4 text-amber-500" />
                <span>Conditional Formatting Rule</span>
              </h3>
              <button onClick={() => setShowConditionalModal(false)}>
                <X className="w-4 h-4 text-slate-400" />
              </button>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="font-semibold text-slate-500">Condition Rule:</label>
                <select
                  value={condRule}
                  onChange={(e) => setCondRule(e.target.value as any)}
                  className="w-full mt-1 p-2 rounded border dark:bg-slate-800 dark:border-slate-700"
                >
                  <option value="greaterThan">Cell Value Greater Than (&gt;)</option>
                  <option value="lessThan">Cell Value Less Than (&lt;)</option>
                  <option value="equals">Cell Value Equals (=)</option>
                  <option value="contains">Text Contains</option>
                </select>
              </div>

              <div>
                <label className="font-semibold text-slate-500">Target Value:</label>
                <input
                  type="text"
                  value={condVal}
                  onChange={(e) => setCondVal(e.target.value)}
                  className="w-full mt-1 p-2 rounded border dark:bg-slate-800 dark:border-slate-700 font-mono"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="font-semibold text-slate-500">Background Color:</label>
                  <input
                    type="color"
                    value={condBg}
                    onChange={(e) => setCondBg(e.target.value)}
                    className="w-full h-8 mt-1 rounded cursor-pointer border"
                  />
                </div>
                <div>
                  <label className="font-semibold text-slate-500">Text Color:</label>
                  <input
                    type="color"
                    value={condText}
                    onChange={(e) => setCondText(e.target.value)}
                    className="w-full h-8 mt-1 rounded cursor-pointer border"
                  />
                </div>
              </div>

              <button
                onClick={addConditionalRule}
                className="w-full mt-2 py-2 rounded bg-emerald-600 hover:bg-emerald-700 text-white font-medium"
              >
                Apply Conditional Rule
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 8. MODAL: DATA VALIDATION DROPDOWN */}
      {showValidationModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4">
          <div
            className={`w-full max-w-md rounded-xl shadow-2xl border p-6 flex flex-col space-y-4 ${
              darkMode ? 'bg-slate-900 border-slate-700 text-slate-100' : 'bg-white border-slate-200 text-slate-900'
            }`}
          >
            <div className="flex items-center justify-between border-b pb-2 dark:border-slate-800">
              <h3 className="font-bold text-base flex items-center space-x-2">
                <CheckSquare className="w-4 h-4 text-blue-500" />
                <span>Cell Dropdown Validation ({selectedCell})</span>
              </h3>
              <button onClick={() => setShowValidationModal(false)}>
                <X className="w-4 h-4 text-slate-400" />
              </button>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="font-semibold text-slate-500">
                  Allowed Options (comma separated list):
                </label>
                <input
                  type="text"
                  value={validationOptions}
                  onChange={(e) => setValidationOptions(e.target.value)}
                  placeholder="e.g. Approved, In Review, Rejected"
                  className="w-full mt-1 p-2 rounded border dark:bg-slate-800 dark:border-slate-700"
                />
              </div>

              <div className="flex flex-wrap gap-1">
                {['High, Medium, Low', 'Complete, In Progress, Blocked', 'Yes, No, N/A'].map((preset) => (
                  <button
                    key={preset}
                    onClick={() => setValidationOptions(preset)}
                    className="px-2 py-1 rounded bg-slate-100 dark:bg-slate-800 text-[10px]"
                  >
                    {preset}
                  </button>
                ))}
              </div>

              <button
                onClick={addValidation}
                className="w-full mt-2 py-2 rounded bg-blue-600 hover:bg-blue-700 text-white font-medium"
              >
                Create Dropdown List on {selectedCell}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 9. MODAL: REAL-TIME EXCEL QUICK CHART MODAL */}
      {showChartModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4">
          <div
            className={`w-full max-w-2xl rounded-xl shadow-2xl border p-6 flex flex-col space-y-4 ${
              darkMode ? 'bg-slate-900 border-slate-700 text-slate-100' : 'bg-white border-slate-200 text-slate-900'
            }`}
          >
            <div className="flex items-center justify-between border-b pb-3 dark:border-slate-800">
              <div className="flex items-center space-x-2">
                <BarChart2 className="w-5 h-5 text-emerald-500" />
                <h3 className="font-bold text-base">Quick Chart Visualizer (Columns A & B)</h3>
              </div>
              <button
                onClick={() => setShowChartModal(false)}
                className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Chart Type Selector */}
            <div className="flex items-center space-x-2">
              <button
                onClick={() => setChartType('bar')}
                className={`flex items-center space-x-1 px-3 py-1.5 rounded-lg text-xs font-semibold ${
                  chartType === 'bar'
                    ? 'bg-emerald-600 text-white'
                    : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300'
                }`}
              >
                <BarChart2 className="w-3.5 h-3.5" />
                <span>Bar Chart</span>
              </button>
              <button
                onClick={() => setChartType('line')}
                className={`flex items-center space-x-1 px-3 py-1.5 rounded-lg text-xs font-semibold ${
                  chartType === 'line'
                    ? 'bg-emerald-600 text-white'
                    : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300'
                }`}
              >
                <TrendingUp className="w-3.5 h-3.5" />
                <span>Line Chart</span>
              </button>
              <button
                onClick={() => setChartType('pie')}
                className={`flex items-center space-x-1 px-3 py-1.5 rounded-lg text-xs font-semibold ${
                  chartType === 'pie'
                    ? 'bg-emerald-600 text-white'
                    : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300'
                }`}
              >
                <PieChart className="w-3.5 h-3.5" />
                <span>Pie Chart</span>
              </button>
            </div>

            {/* SVG Visual Chart */}
            <div className="w-full h-72 bg-slate-50 dark:bg-slate-950/50 rounded-lg p-4 flex items-center justify-center border border-slate-200 dark:border-slate-800">
              {chartData.length === 0 ? (
                <div className="text-center text-slate-400 text-xs">
                  No numeric values found in Column B. Enter labels in Column A and numbers in Column B.
                </div>
              ) : chartType === 'bar' ? (
                <div className="w-full h-full flex items-end justify-between space-x-3 pt-6 px-2">
                  {chartData.map((item, idx) => {
                    const maxVal = Math.max(...chartData.map((d) => d.value), 1);
                    const heightPercent = Math.max(8, Math.round((item.value / maxVal) * 100));
                    return (
                      <div key={idx} className="flex-1 flex flex-col items-center h-full justify-end group">
                        <span className="text-[10px] font-mono text-emerald-600 dark:text-emerald-400 font-semibold mb-1 opacity-0 group-hover:opacity-100 transition-opacity">
                          ${item.value.toLocaleString()}
                        </span>
                        <div
                          className="w-full rounded-t-md bg-linear-to-t from-emerald-600 to-teal-400 transition-all hover:brightness-110 shadow-sm"
                          style={{ height: `${heightPercent}%` }}
                        />
                        <span className="text-[10px] text-slate-500 mt-2 truncate w-full text-center">
                          {item.label}
                        </span>
                      </div>
                    );
                  })}
                </div>
              ) : chartType === 'line' ? (
                <svg className="w-full h-full" viewBox="0 0 500 200">
                  <polyline
                    fill="none"
                    stroke="#10b981"
                    strokeWidth="3"
                    points={chartData
                      .map((d, i) => {
                        const maxVal = Math.max(...chartData.map((x) => x.value), 1);
                        const x = (i / Math.max(1, chartData.length - 1)) * 460 + 20;
                        const y = 180 - (d.value / maxVal) * 150;
                        return `${x},${y}`;
                      })
                      .join(' ')}
                  />
                  {chartData.map((d, i) => {
                    const maxVal = Math.max(...chartData.map((x) => x.value), 1);
                    const x = (i / Math.max(1, chartData.length - 1)) * 460 + 20;
                    const y = 180 - (d.value / maxVal) * 150;
                    return (
                      <g key={i}>
                        <circle cx={x} cy={y} r="5" fill="#047857" stroke="#fff" strokeWidth="2" />
                        <text x={x} y={y - 8} fontSize="10" textAnchor="middle" fill="#64748b">
                          {d.label}
                        </text>
                      </g>
                    );
                  })}
                </svg>
              ) : (
                <div className="flex items-center justify-around w-full">
                  <div className="relative w-48 h-48 rounded-full border-8 border-emerald-500 flex items-center justify-center font-bold text-sm text-emerald-600">
                    {chartData.length} Data Points
                  </div>
                  <div className="space-y-1.5 text-xs">
                    {chartData.slice(0, 5).map((d, i) => (
                      <div key={i} className="flex items-center space-x-2">
                        <span className="w-2.5 h-2.5 rounded-full bg-emerald-500" />
                        <span className="font-medium text-slate-700 dark:text-slate-300">{d.label}:</span>
                        <span className="font-mono text-emerald-600 dark:text-emerald-400">
                          ${d.value.toLocaleString()}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            <div className="flex justify-end">
              <button
                onClick={() => setShowChartModal(false)}
                className="px-4 py-1.5 text-xs rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-medium shadow-xs"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
