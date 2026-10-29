import type { CellData, CellStyle } from '../types/office';

export class FormulaEngine {
  // Convert "A1" to { col: 0, row: 0 }
  public static parseCellId(cellId: string): { col: number; row: number } | null {
    const match = cellId.toUpperCase().match(/^([A-Z]+)([0-9]+)$/);
    if (!match) return null;

    const colStr = match[1];
    const rowNum = parseInt(match[2], 10) - 1;

    let colIndex = 0;
    for (let i = 0; i < colStr.length; i++) {
      colIndex = colIndex * 26 + (colStr.charCodeAt(i) - 64);
    }
    colIndex -= 1;

    return { col: colIndex, row: rowNum };
  }

  // Convert { col: 0, row: 0 } to "A1"
  public static toCellId(col: number, row: number): string {
    let colName = '';
    let temp = col + 1;
    while (temp > 0) {
      const rem = (temp - 1) % 26;
      colName = String.fromCharCode(65 + rem) + colName;
      temp = Math.floor((temp - 1) / 26);
    }
    return `${colName}${row + 1}`;
  }

  // Expand "A1:B3" to ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']
  public static expandRange(rangeStr: string): string[] {
    const parts = rangeStr.toUpperCase().split(':');
    if (parts.length === 1) return [parts[0].trim()];
    if (parts.length !== 2) return [];

    const start = this.parseCellId(parts[0].trim());
    const end = this.parseCellId(parts[1].trim());
    if (!start || !end) return [];

    const minCol = Math.min(start.col, end.col);
    const maxCol = Math.max(start.col, end.col);
    const minRow = Math.min(start.row, end.row);
    const maxRow = Math.max(start.row, end.row);

    const cells: string[] = [];
    for (let c = minCol; c <= maxCol; c++) {
      for (let r = minRow; r <= maxRow; r++) {
        cells.push(this.toCellId(c, r));
      }
    }
    return cells;
  }

  // Evaluate a formula or cell value against the spreadsheet cell dictionary
  public static evaluateCell(
    cellKey: string,
    cells: Record<string, CellData>,
    visited: Set<string> = new Set()
  ): { value: string | number; error?: string } {
    const cell = cells[cellKey];
    if (!cell || !cell.raw) return { value: '' };

    const raw = cell.raw.trim();
    if (!raw.startsWith('=')) {
      // Literal number or text
      const num = Number(raw);
      if (!isNaN(num) && raw !== '') {
        return { value: num };
      }
      return { value: raw };
    }

    if (visited.has(cellKey)) {
      return { value: '#CYCLE!', error: 'Circular dependency detected' };
    }
    visited.add(cellKey);

    const formula = raw.substring(1).trim();

    try {
      const result = this.evalExpression(formula, cells, visited);
      visited.delete(cellKey);
      return { value: result };
    } catch (err: unknown) {
      visited.delete(cellKey);
      const errMsg = err instanceof Error ? err.message : '#ERROR!';
      return { value: errMsg.startsWith('#') ? errMsg : '#VALUE!', error: errMsg };
    }
  }

  private static evalExpression(
    expr: string,
    cells: Record<string, CellData>,
    visited: Set<string>
  ): string | number {
    expr = expr.trim();

    // Check for function calls like SUM(A1:A5) or IF(A1>10, 100, 0)
    const funcMatch = expr.match(/^([A-Z_]+)\s*\((.*)\)$/i);
    if (funcMatch) {
      const funcName = funcMatch[1].toUpperCase();
      const argsContent = funcMatch[2];
      const args = this.splitArgs(argsContent);
      return this.executeFunction(funcName, args, cells, visited);
    }

    // Replace cell references (e.g., A1, B2) with their values
    const parsedExpr = expr.replace(/\b([A-Z]+[0-9]+)\b/gi, (match) => {
      const val = this.getNumericValueOfCell(match.toUpperCase(), cells, visited);
      return val.toString();
    });

    // Safely evaluate standard math expressions (+, -, *, /, %, parentheses)
    // Only allow digits, decimals, basic operators, parentheses, spaces
    if (/^[\d\s+\-*/%.()^]+$/.test(parsedExpr)) {
      const sanitized = parsedExpr.replace(/\^/g, '**');
      // eslint-disable-next-line no-new-func
      const calc = new Function(`return (${sanitized});`);
      const res = calc();
      if (typeof res === 'number') {
        if (!isFinite(res)) return '#DIV/0!';
        return Math.round(res * 1000000) / 1000000;
      }
      return res;
    }

    // Fallback: evaluate single cell ref
    if (/^[A-Z]+[0-9]+$/i.test(expr)) {
      const evalRes = this.evaluateCell(expr.toUpperCase(), cells, visited);
      return evalRes.value;
    }

    return expr;
  }

  private static splitArgs(argsContent: string): string[] {
    const args: string[] = [];
    let depth = 0;
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < argsContent.length; i++) {
      const char = argsContent[i];
      if (char === '"' || char === "'") inQuotes = !inQuotes;
      if (!inQuotes) {
        if (char === '(') depth++;
        else if (char === ')') depth--;
        else if (char === ',' && depth === 0) {
          args.push(current.trim());
          current = '';
          continue;
        }
      }
      current += char;
    }
    if (current.trim().length > 0) {
      args.push(current.trim());
    }
    return args;
  }

  private static executeFunction(
    name: string,
    args: string[],
    cells: Record<string, CellData>,
    visited: Set<string>
  ): string | number {
    switch (name) {
      case 'SUM': {
        const numbers = this.collectNumbersFromArgs(args, cells, visited);
        return numbers.reduce((acc, n) => acc + n, 0);
      }
      case 'AVERAGE':
      case 'AVG': {
        const numbers = this.collectNumbersFromArgs(args, cells, visited);
        if (numbers.length === 0) return 0;
        const sum = numbers.reduce((acc, n) => acc + n, 0);
        return Math.round((sum / numbers.length) * 1000) / 1000;
      }
      case 'COUNT': {
        const numbers = this.collectNumbersFromArgs(args, cells, visited);
        return numbers.length;
      }
      case 'MIN': {
        const numbers = this.collectNumbersFromArgs(args, cells, visited);
        if (numbers.length === 0) return 0;
        return Math.min(...numbers);
      }
      case 'MAX': {
        const numbers = this.collectNumbersFromArgs(args, cells, visited);
        if (numbers.length === 0) return 0;
        return Math.max(...numbers);
      }
      case 'PRODUCT': {
        const numbers = this.collectNumbersFromArgs(args, cells, visited);
        if (numbers.length === 0) return 0;
        return numbers.reduce((acc, n) => acc * n, 1);
      }
      case 'ROUND': {
        if (args.length < 1) return '#N/A';
        const num = Number(this.evalExpression(args[0], cells, visited)) || 0;
        const decimals = args[1] ? Number(this.evalExpression(args[1], cells, visited)) || 0 : 0;
        const factor = Math.pow(10, decimals);
        return Math.round(num * factor) / factor;
      }
      case 'ABS': {
        if (args.length < 1) return '#N/A';
        const num = Number(this.evalExpression(args[0], cells, visited)) || 0;
        return Math.abs(num);
      }
      case 'UPPER': {
        if (args.length < 1) return '';
        const val = String(this.evalExpression(args[0], cells, visited));
        return val.toUpperCase();
      }
      case 'LOWER': {
        if (args.length < 1) return '';
        const val = String(this.evalExpression(args[0], cells, visited));
        return val.toLowerCase();
      }
      case 'CONCAT':
      case 'CONCATENATE': {
        return args
          .map((arg) => {
            const raw = arg.replace(/^["']|["']$/g, '');
            if (this.parseCellId(arg)) {
              return String(this.evaluateCell(arg.toUpperCase(), cells, visited).value);
            }
            return raw;
          })
          .join('');
      }
      case 'IF': {
        if (args.length < 2) return '#N/A';
        const condition = args[0];
        const trueVal = args[1];
        const falseVal = args[2] ?? '""';

        const isTrue = this.evalCondition(condition, cells, visited);
        return this.evalExpression(isTrue ? trueVal : falseVal, cells, visited);
      }
      default:
        return `#NAME? (${name})`;
    }
  }

  private static evalCondition(
    cond: string,
    cells: Record<string, CellData>,
    visited: Set<string>
  ): boolean {
    const operators = ['>=', '<=', '!=', '<>', '>', '<', '='];
    for (const op of operators) {
      if (cond.includes(op)) {
        const [leftStr, rightStr] = cond.split(op);
        const leftVal = this.evalExpression(leftStr.trim(), cells, visited);
        const rightVal = this.evalExpression(rightStr.trim(), cells, visited);

        const leftNum = Number(leftVal);
        const rightNum = Number(rightVal);
        const isNum = !isNaN(leftNum) && !isNaN(rightNum);

        if (op === '=') return isNum ? leftNum === rightNum : leftVal === rightVal;
        if (op === '!=' || op === '<>') return isNum ? leftNum !== rightNum : leftVal !== rightVal;
        if (op === '>') return isNum ? leftNum > rightNum : String(leftVal) > String(rightVal);
        if (op === '<') return isNum ? leftNum < rightNum : String(leftVal) < String(rightVal);
        if (op === '>=') return isNum ? leftNum >= rightNum : String(leftVal) >= String(rightVal);
        if (op === '<=') return isNum ? leftNum <= rightNum : String(leftVal) <= String(rightVal);
      }
    }
    return Boolean(this.evalExpression(cond, cells, visited));
  }

  private static collectNumbersFromArgs(
    args: string[],
    cells: Record<string, CellData>,
    visited: Set<string>
  ): number[] {
    const numbers: number[] = [];

    for (const arg of args) {
      if (arg.includes(':')) {
        const rangeCells = this.expandRange(arg);
        for (const cKey of rangeCells) {
          const val = this.getNumericValueOfCell(cKey, cells, visited);
          if (val !== 0 || (cells[cKey] && cells[cKey].raw !== '')) {
            numbers.push(val);
          }
        }
      } else if (this.parseCellId(arg.trim())) {
        numbers.push(this.getNumericValueOfCell(arg.trim().toUpperCase(), cells, visited));
      } else {
        const evaluated = this.evalExpression(arg, cells, visited);
        const num = Number(evaluated);
        if (!isNaN(num)) numbers.push(num);
      }
    }

    return numbers;
  }

  private static getNumericValueOfCell(
    cellKey: string,
    cells: Record<string, CellData>,
    visited: Set<string>
  ): number {
    const cellVal = this.evaluateCell(cellKey, cells, visited).value;
    if (typeof cellVal === 'number') return cellVal;
    const parsed = Number(cellVal);
    return isNaN(parsed) ? 0 : parsed;
  }

  // Format cell output value with its style
  public static formatValue(value: string | number | undefined, style?: CellStyle): string {
    if (value === undefined || value === null || value === '') return '';
    if (typeof value === 'string' && value.startsWith('#')) return value; // Return error as is

    const num = typeof value === 'number' ? value : Number(value);

    if (style?.format === 'currency') {
      if (isNaN(num)) return String(value);
      return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(num);
    }
    if (style?.format === 'percent') {
      if (isNaN(num)) return String(value);
      return `${(num * 100).toFixed(1)}%`;
    }
    if (style?.format === 'number') {
      if (isNaN(num)) return String(value);
      return num.toLocaleString('en-US', { maximumFractionDigits: 2 });
    }

    return String(value);
  }
}
