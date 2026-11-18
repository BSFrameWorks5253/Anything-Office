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
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
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
    if (!cell || cell.raw === undefined || cell.raw === null || cell.raw === '') {
      return { value: '' };
    }

    const raw = String(cell.raw).trim();
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
    if (!expr) return '';

    // Handle function calls: NAME(...)
    const funcMatch = expr.match(/^([A-Z_]+)\s*\(([\s\S]*)\)$/i);
    if (funcMatch) {
      const funcName = funcMatch[1].toUpperCase();
      const argsContent = funcMatch[2];
      const args = this.splitArgs(argsContent);
      return this.executeFunction(funcName, args, cells, visited);
    }

    // Replace cell references (e.g. A1, B2) with evaluated numeric values when surrounded by operators
    const parsedExpr = expr.replace(/\b([A-Z]+[0-9]+)\b/gi, (match) => {
      const val = this.getNumericValueOfCell(match.toUpperCase(), cells, visited);
      return val.toString();
    });

    // Evaluate basic math expressions
    if (/^[\d\s+\-*/%.()^]+$/.test(parsedExpr)) {
      try {
        const sanitized = parsedExpr.replace(/\^/g, '**');
        // eslint-disable-next-line no-new-func
        const calc = new Function(`return (${sanitized});`);
        const res = calc();
        if (typeof res === 'number') {
          if (!isFinite(res)) return '#DIV/0!';
          return Math.round(res * 1000000) / 1000000;
        }
        return res;
      } catch {
        return '#ERROR!';
      }
    }

    // Fallback: single cell ref without formula
    if (/^[A-Z]+[0-9]+$/i.test(expr)) {
      const evalRes = this.evaluateCell(expr.toUpperCase(), cells, visited);
      return evalRes.value;
    }

    // Quoted string
    if ((expr.startsWith('"') && expr.endsWith('"')) || (expr.startsWith("'") && expr.endsWith("'"))) {
      return expr.slice(1, -1);
    }

    return expr;
  }

  private static splitArgs(argsContent: string): string[] {
    const args: string[] = [];
    let depth = 0;
    let current = '';
    let inQuotes = false;
    let quoteChar = '';

    for (let i = 0; i < argsContent.length; i++) {
      const char = argsContent[i];
      if ((char === '"' || char === "'") && (!inQuotes || quoteChar === char)) {
        inQuotes = !inQuotes;
        quoteChar = inQuotes ? char : '';
      }
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
      // 1. MATH & STATISTICAL
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
      case 'COUNTA': {
        let count = 0;
        for (const arg of args) {
          if (arg.includes(':')) {
            const range = this.expandRange(arg);
            for (const key of range) {
              const val = this.evaluateCell(key, cells, visited).value;
              if (val !== '' && val !== undefined && val !== null) count++;
            }
          } else {
            const val = this.evalExpression(arg, cells, visited);
            if (val !== '' && val !== undefined && val !== null) count++;
          }
        }
        return count;
      }
      case 'COUNTBLANK': {
        let blankCount = 0;
        for (const arg of args) {
          if (arg.includes(':')) {
            const range = this.expandRange(arg);
            for (const key of range) {
              const val = this.evaluateCell(key, cells, visited).value;
              if (val === '' || val === undefined || val === null) blankCount++;
            }
          }
        }
        return blankCount;
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
      case 'MEDIAN': {
        const numbers = this.collectNumbersFromArgs(args, cells, visited).sort((a, b) => a - b);
        if (numbers.length === 0) return 0;
        const mid = Math.floor(numbers.length / 2);
        return numbers.length % 2 !== 0 ? numbers[mid] : (numbers[mid - 1] + numbers[mid]) / 2;
      }
      case 'ROUND': {
        if (args.length < 1) return '#N/A';
        const num = Number(this.evalExpression(args[0], cells, visited)) || 0;
        const decimals = args[1] ? Number(this.evalExpression(args[1], cells, visited)) || 0 : 0;
        const factor = Math.pow(10, decimals);
        return Math.round(num * factor) / factor;
      }
      case 'ROUNDUP': {
        if (args.length < 1) return '#N/A';
        const num = Number(this.evalExpression(args[0], cells, visited)) || 0;
        const decimals = args[1] ? Number(this.evalExpression(args[1], cells, visited)) || 0 : 0;
        const factor = Math.pow(10, decimals);
        return Math.ceil(num * factor) / factor;
      }
      case 'ROUNDDOWN': {
        if (args.length < 1) return '#N/A';
        const num = Number(this.evalExpression(args[0], cells, visited)) || 0;
        const decimals = args[1] ? Number(this.evalExpression(args[1], cells, visited)) || 0 : 0;
        const factor = Math.pow(10, decimals);
        return Math.floor(num * factor) / factor;
      }
      case 'ABS': {
        if (args.length < 1) return '#N/A';
        const num = Number(this.evalExpression(args[0], cells, visited)) || 0;
        return Math.abs(num);
      }
      case 'SQRT': {
        if (args.length < 1) return '#N/A';
        const num = Number(this.evalExpression(args[0], cells, visited)) || 0;
        if (num < 0) return '#NUM!';
        return Math.sqrt(num);
      }
      case 'POWER': {
        if (args.length < 2) return '#N/A';
        const base = Number(this.evalExpression(args[0], cells, visited)) || 0;
        const exp = Number(this.evalExpression(args[1], cells, visited)) || 0;
        return Math.pow(base, exp);
      }
      case 'MOD': {
        if (args.length < 2) return '#N/A';
        const num = Number(this.evalExpression(args[0], cells, visited)) || 0;
        const div = Number(this.evalExpression(args[1], cells, visited)) || 0;
        if (div === 0) return '#DIV/0!';
        return num % div;
      }
      case 'INT': {
        if (args.length < 1) return '#N/A';
        const num = Number(this.evalExpression(args[0], cells, visited)) || 0;
        return Math.floor(num);
      }

      // 2. CONDITIONALS & LOGIC
      case 'IF': {
        if (args.length < 2) return '#N/A';
        const condition = args[0];
        const trueVal = args[1];
        const falseVal = args[2] ?? '""';

        const isTrue = this.evalCondition(condition, cells, visited);
        return this.evalExpression(isTrue ? trueVal : falseVal, cells, visited);
      }
      case 'IFS': {
        for (let i = 0; i < args.length; i += 2) {
          const cond = args[i];
          const val = args[i + 1];
          if (this.evalCondition(cond, cells, visited)) {
            return this.evalExpression(val, cells, visited);
          }
        }
        return '#N/A';
      }
      case 'AND': {
        for (const arg of args) {
          if (!this.evalCondition(arg, cells, visited)) return 'FALSE';
        }
        return 'TRUE';
      }
      case 'OR': {
        for (const arg of args) {
          if (this.evalCondition(arg, cells, visited)) return 'TRUE';
        }
        return 'FALSE';
      }
      case 'NOT': {
        if (args.length < 1) return '#N/A';
        return this.evalCondition(args[0], cells, visited) ? 'FALSE' : 'TRUE';
      }

      // 3. TEXT FUNCTIONS
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
      case 'PROPER': {
        if (args.length < 1) return '';
        const val = String(this.evalExpression(args[0], cells, visited));
        return val.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
      }
      case 'LEN': {
        if (args.length < 1) return 0;
        const val = String(this.evalExpression(args[0], cells, visited));
        return val.length;
      }
      case 'TRIM': {
        if (args.length < 1) return '';
        const val = String(this.evalExpression(args[0], cells, visited));
        return val.trim();
      }
      case 'LEFT': {
        if (args.length < 1) return '';
        const val = String(this.evalExpression(args[0], cells, visited));
        const num = args[1] ? Number(this.evalExpression(args[1], cells, visited)) || 1 : 1;
        return val.substring(0, num);
      }
      case 'RIGHT': {
        if (args.length < 1) return '';
        const val = String(this.evalExpression(args[0], cells, visited));
        const num = args[1] ? Number(this.evalExpression(args[1], cells, visited)) || 1 : 1;
        return val.substring(Math.max(0, val.length - num));
      }
      case 'MID': {
        if (args.length < 3) return '';
        const val = String(this.evalExpression(args[0], cells, visited));
        const start = Math.max(0, (Number(this.evalExpression(args[1], cells, visited)) || 1) - 1);
        const len = Number(this.evalExpression(args[2], cells, visited)) || 0;
        return val.substring(start, start + len);
      }

      // 4. LOOKUP & REFERENCE
      case 'VLOOKUP': {
        if (args.length < 3) return '#N/A';
        const lookupVal = String(this.evalExpression(args[0], cells, visited)).toLowerCase();
        const tableRange = args[1];
        const colIdx = Number(this.evalExpression(args[2], cells, visited));
        if (isNaN(colIdx) || colIdx < 1) return '#VALUE!';

        const rangeParts = tableRange.split(':');
        if (rangeParts.length !== 2) return '#REF!';
        const start = this.parseCellId(rangeParts[0].trim());
        const end = this.parseCellId(rangeParts[1].trim());
        if (!start || !end) return '#REF!';

        const minCol = Math.min(start.col, end.col);
        const maxCol = Math.max(start.col, end.col);
        const minRow = Math.min(start.row, end.row);
        const maxRow = Math.max(start.row, end.row);

        const targetCol = minCol + colIdx - 1;
        if (targetCol > maxCol) return '#REF!';

        for (let r = minRow; r <= maxRow; r++) {
          const firstCellKey = this.toCellId(minCol, r);
          const cellVal = String(this.evaluateCell(firstCellKey, cells, visited).value).toLowerCase();
          if (cellVal === lookupVal) {
            const targetCellKey = this.toCellId(targetCol, r);
            return this.evaluateCell(targetCellKey, cells, visited).value;
          }
        }
        return '#N/A';
      }
      case 'INDEX': {
        if (args.length < 2) return '#N/A';
        const range = args[0];
        const rowNum = Number(this.evalExpression(args[1], cells, visited));
        const colNum = args[2] ? Number(this.evalExpression(args[2], cells, visited)) : 1;

        const rangeParts = range.split(':');
        if (rangeParts.length !== 2) return '#REF!';
        const start = this.parseCellId(rangeParts[0].trim());
        const end = this.parseCellId(rangeParts[1].trim());
        if (!start || !end) return '#REF!';

        const minCol = Math.min(start.col, end.col);
        const minRow = Math.min(start.row, end.row);

        const targetCol = minCol + (colNum - 1);
        const targetRow = minRow + (rowNum - 1);

        const cellKey = this.toCellId(targetCol, targetRow);
        return this.evaluateCell(cellKey, cells, visited).value;
      }
      case 'MATCH': {
        if (args.length < 2) return '#N/A';
        const lookupVal = String(this.evalExpression(args[0], cells, visited)).toLowerCase();
        const rangeCells = this.expandRange(args[1]);
        for (let i = 0; i < rangeCells.length; i++) {
          const cellVal = String(this.evaluateCell(rangeCells[i], cells, visited).value).toLowerCase();
          if (cellVal === lookupVal) {
            return i + 1; // 1-based index
          }
        }
        return '#N/A';
      }

      // 5. STATISTICAL CONDITIONALS
      case 'COUNTIF': {
        if (args.length < 2) return 0;
        const range = this.expandRange(args[0]);
        const criteria = String(this.evalExpression(args[1], cells, visited));
        let count = 0;
        for (const cellKey of range) {
          const cellVal = this.evaluateCell(cellKey, cells, visited).value;
          if (this.matchesCriteria(cellVal, criteria)) count++;
        }
        return count;
      }
      case 'SUMIF': {
        if (args.length < 2) return 0;
        const range = this.expandRange(args[0]);
        const criteria = String(this.evalExpression(args[1], cells, visited));
        const sumRange = args[2] ? this.expandRange(args[2]) : range;
        let sum = 0;
        for (let i = 0; i < range.length; i++) {
          const cellVal = this.evaluateCell(range[i], cells, visited).value;
          if (this.matchesCriteria(cellVal, criteria)) {
            const sumVal = Number(this.evaluateCell(sumRange[i], cells, visited).value);
            if (!isNaN(sumVal)) sum += sumVal;
          }
        }
        return sum;
      }

      // 6. DATE & TIME
      case 'TODAY': {
        return new Date().toISOString().split('T')[0];
      }
      case 'NOW': {
        return new Date().toLocaleString();
      }
      case 'YEAR': {
        if (args.length < 1) return new Date().getFullYear();
        const d = new Date(String(this.evalExpression(args[0], cells, visited)));
        return isNaN(d.getTime()) ? '#VALUE!' : d.getFullYear();
      }
      case 'MONTH': {
        if (args.length < 1) return new Date().getMonth() + 1;
        const d = new Date(String(this.evalExpression(args[0], cells, visited)));
        return isNaN(d.getTime()) ? '#VALUE!' : d.getMonth() + 1;
      }
      case 'DAY': {
        if (args.length < 1) return new Date().getDate();
        const d = new Date(String(this.evalExpression(args[0], cells, visited)));
        return isNaN(d.getTime()) ? '#VALUE!' : d.getDate();
      }

      default:
        return `#NAME? (${name})`;
    }
  }

  private static matchesCriteria(cellVal: string | number, criteria: string): boolean {
    criteria = criteria.trim().replace(/^["']|["']$/g, '');
    const numVal = Number(cellVal);
    const isNum = !isNaN(numVal) && cellVal !== '';

    if (criteria.startsWith('>=')) {
      const target = Number(criteria.slice(2));
      return isNum && numVal >= target;
    }
    if (criteria.startsWith('<=')) {
      const target = Number(criteria.slice(2));
      return isNum && numVal <= target;
    }
    if (criteria.startsWith('<>')) {
      const target = criteria.slice(2);
      return String(cellVal) !== target;
    }
    if (criteria.startsWith('>')) {
      const target = Number(criteria.slice(1));
      return isNum && numVal > target;
    }
    if (criteria.startsWith('<')) {
      const target = Number(criteria.slice(1));
      return isNum && numVal < target;
    }
    if (criteria.startsWith('=')) {
      const target = criteria.slice(1);
      return String(cellVal).toLowerCase() === target.toLowerCase();
    }

    return String(cellVal).toLowerCase() === criteria.toLowerCase();
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
    const decimals = style?.decimals ?? 2;

    if (style?.format === 'currency') {
      if (isNaN(num)) return String(value);
      return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
      }).format(num);
    }
    if (style?.format === 'percent') {
      if (isNaN(num)) return String(value);
      return `${(num * 100).toFixed(decimals)}%`;
    }
    if (style?.format === 'number') {
      if (isNaN(num)) return String(value);
      return num.toLocaleString('en-US', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
      });
    }

    return String(value);
  }
}
