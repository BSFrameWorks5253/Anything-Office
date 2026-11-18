import JSZip from 'jszip';
import type {
  AnythingProject,
  DocumentData,
  DocParagraph,
  SpreadsheetData,
  SheetData,
  CellData,
  PresentationData,
  SlideData,
  SlideElement,
} from '../types/office';

export interface ImportResult {
  mode: 'docs' | 'sheets' | 'slides' | 'pdf' | 'all';
  projectUpdates: Partial<AnythingProject>;
  fileName: string;
  detectedFormat: string;
}

export class DocumentImporter {
  public static async parseFile(file: File): Promise<ImportResult> {
    const ext = file.name.split('.').pop()?.toLowerCase() || '';

    // 1. JSON or .anything bundle
    if (ext === 'anything' || ext === 'json') {
      const text = await file.text();
      const parsed = JSON.parse(text);
      if (parsed.docs || parsed.sheets || parsed.slides) {
        return {
          mode: 'all',
          projectUpdates: parsed,
          fileName: file.name,
          detectedFormat: 'Anything Unified Project Bundle (.anything)',
        };
      }
    }

    // 2. Portable Document Format (PDF)
    if (ext === 'pdf') {
      const fileUrl = URL.createObjectURL(file);
      const arrayBuffer = await file.arrayBuffer();
      let textSample = '';
      let pageCountEstimate = 1;
      try {
        const uint8 = new Uint8Array(arrayBuffer);
        const decoder = new TextDecoder('latin1');
        const raw = decoder.decode(uint8);
        const pageMatches = raw.match(/\/Type\s*\/Page\b/g);
        if (pageMatches) {
          pageCountEstimate = Math.max(1, pageMatches.length);
        } else {
          const countMatch = raw.match(/\/Count\s+(\d+)/);
          if (countMatch) {
            pageCountEstimate = Math.max(1, parseInt(countMatch[1], 10));
          }
        }
        const textMatches = raw.match(/\(([^)]+)\)\s*Tj/g);
        if (textMatches && textMatches.length > 0) {
          textSample = textMatches
            .map(m => m.replace(/^\(/, '').replace(/\)\s*Tj$/, ''))
            .filter(t => t.trim().length > 1)
            .join(' ')
            .slice(0, 3000);
        } else {
          const arrayMatches = raw.match(/\[([^\]]+)\]\s*TJ/gi);
          if (arrayMatches && arrayMatches.length > 0) {
            textSample = arrayMatches
              .map(m => (m.match(/\(([^)]+)\)/g) || []).map(t => t.replace(/^\(/, '').replace(/\)$/, '')).join(''))
              .filter(t => t.trim().length > 1)
              .join(' ')
              .slice(0, 3000);
          }
        }
      } catch (e) {
        console.warn('PDF stream extraction fallback', e);
      }

      return {
        mode: 'pdf',
        projectUpdates: {
          pdf: {
            title: file.name,
            url: fileUrl,
            totalPages: pageCountEstimate,
            currentPage: 1,
            zoom: 100,
            rotation: 0,
            extractedText: textSample,
          },
        },
        fileName: file.name,
        detectedFormat: 'Portable Document Format (.pdf)',
      };
    }

    // 3. CSV or TSV (Spreadsheets)
    if (ext === 'csv' || ext === 'tsv') {
      const text = await file.text();
      const delimiter = ext === 'tsv' ? '\t' : ',';
      const sheet = this.parseCsvToSheet(text, file.name.replace(/\.[^/.]+$/, ''), delimiter);
      return {
        mode: 'sheets',
        projectUpdates: {
          sheets: {
            activeSheetId: sheet.id,
            sheets: [sheet],
          },
        },
        fileName: file.name,
        detectedFormat: `${ext.toUpperCase()} Delimited Spreadsheet`,
      };
    }

    // 3. Plain Text or Markdown (Docs)
    if (ext === 'txt' || ext === 'md' || ext === 'markdown') {
      const text = await file.text();
      const doc = this.parseTextToDoc(text, file.name.replace(/\.[^/.]+$/, ''));
      return {
        mode: 'docs',
        projectUpdates: { docs: doc },
        fileName: file.name,
        detectedFormat: ext === 'md' ? 'Markdown Document (.md)' : 'Plain Text Document (.txt)',
      };
    }

    // 4. HTML (Docs)
    if (ext === 'html' || ext === 'htm') {
      const text = await file.text();
      const doc = this.parseHtmlToDoc(text, file.name.replace(/\.[^/.]+$/, ''));
      return {
        mode: 'docs',
        projectUpdates: { docs: doc },
        fileName: file.name,
        detectedFormat: 'HTML Document (.html)',
      };
    }

    // ZIP-based Office Formats (DOCX, XLSX, PPTX, ODT, ODS, ODP)
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    // 5. Microsoft Word (.docx)
    if (ext === 'docx' || zip.file('word/document.xml')) {
      const doc = await this.parseDocx(zip, file.name.replace(/\.[^/.]+$/, ''));
      return {
        mode: 'docs',
        projectUpdates: { docs: doc },
        fileName: file.name,
        detectedFormat: 'Microsoft Word Document (.docx)',
      };
    }

    // 6. LibreOffice Writer (.odt)
    if (ext === 'odt' || (zip.file('content.xml') && (await this.isOdfType(zip, 'text')))) {
      const doc = await this.parseOdt(zip, file.name.replace(/\.[^/.]+$/, ''));
      return {
        mode: 'docs',
        projectUpdates: { docs: doc },
        fileName: file.name,
        detectedFormat: 'LibreOffice Writer Document (.odt)',
      };
    }

    // 7. Microsoft Excel (.xlsx)
    if (ext === 'xlsx' || zip.file('xl/workbook.xml')) {
      const sheetsData = await this.parseXlsx(zip);
      return {
        mode: 'sheets',
        projectUpdates: { sheets: sheetsData },
        fileName: file.name,
        detectedFormat: 'Microsoft Excel Spreadsheet (.xlsx)',
      };
    }

    // 8. LibreOffice Calc (.ods)
    if (ext === 'ods' || (zip.file('content.xml') && (await this.isOdfType(zip, 'spreadsheet')))) {
      const sheetsData = await this.parseOds(zip);
      return {
        mode: 'sheets',
        projectUpdates: { sheets: sheetsData },
        fileName: file.name,
        detectedFormat: 'LibreOffice Calc Spreadsheet (.ods)',
      };
    }

    // 9. Microsoft PowerPoint (.pptx)
    if (ext === 'pptx' || zip.file('ppt/presentation.xml')) {
      const slidesData = await this.parsePptx(zip);
      return {
        mode: 'slides',
        projectUpdates: { slides: slidesData },
        fileName: file.name,
        detectedFormat: 'Microsoft PowerPoint Presentation (.pptx)',
      };
    }

    // 10. LibreOffice Impress (.odp)
    if (ext === 'odp' || (zip.file('content.xml') && (await this.isOdfType(zip, 'presentation')))) {
      const slidesData = await this.parseOdp(zip);
      return {
        mode: 'slides',
        projectUpdates: { slides: slidesData },
        fileName: file.name,
        detectedFormat: 'LibreOffice Impress Presentation (.odp)',
      };
    }

    // Legacy binary .doc, .xls, .ppt formats
    if (ext === 'doc' || ext === 'xls' || ext === 'ppt') {
      if (ext === 'doc') {
        const doc = this.parseLegacyWordDoc(new Uint8Array(arrayBuffer), file.name);
        return {
          mode: 'docs',
          projectUpdates: { docs: doc },
          fileName: file.name,
          detectedFormat: 'Legacy Microsoft Word Document (.doc 97-2003)',
        };
      }
      if (ext === 'xls') {
        const sheetsData = this.parseLegacyBiffXls(new Uint8Array(arrayBuffer), file.name);
        return {
          mode: 'sheets',
          projectUpdates: { sheets: sheetsData },
          fileName: file.name,
          detectedFormat: 'Legacy Microsoft Excel Spreadsheet (.xls BIFF5/BIFF8)',
        };
      }
      if (ext === 'ppt') {
        const text = this.extractStringsFromBinary(new Uint8Array(arrayBuffer));
        const slides = this.parseTextToSlides(text);
        return {
          mode: 'slides',
          projectUpdates: { slides },
          fileName: file.name,
          detectedFormat: 'Legacy Microsoft PowerPoint (.ppt)',
        };
      }
    }

    throw new Error(`Unsupported file format .${ext}`);
  }

  // Check ODF mimetype
  private static async isOdfType(
    zip: JSZip,
    expected: 'text' | 'spreadsheet' | 'presentation'
  ): Promise<boolean> {
    const mimeFile = zip.file('mimetype');
    if (!mimeFile) return false;
    const mime = await mimeFile.async('text');
    return mime.includes(expected);
  }

  // --- DOCX Parser (Word) ---
  private static async parseDocx(zip: JSZip, title: string): Promise<DocumentData> {
    const docXmlFile = zip.file('word/document.xml') || zip.file('word/Document.xml');
    if (!docXmlFile) throw new Error('Invalid DOCX: missing word/document.xml');

    const xmlStr = await docXmlFile.async('text');
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');

    const paragraphs: DocParagraph[] = [];

    // Helper: Extract text from any element preserving tabs, breaks, and child runs
    const extractRunText = (element: Element): string => {
      let result = '';
      const walker = (node: Node) => {
        if (node.nodeType === 1) {
          const el = node as Element;
          const localName = (el.localName || el.nodeName.split(':').pop() || '').toLowerCase();
          if (localName === 't') {
            result += el.textContent || '';
          } else if (localName === 'tab') {
            result += '\t';
          } else if (localName === 'br' || localName === 'cr') {
            result += '\n';
          }
        }
        for (let i = 0; i < node.childNodes.length; i++) {
          walker(node.childNodes[i]);
        }
      };
      walker(element);
      return result;
    };

    // Find all paragraphs (prefixed or unprefixed)
    const pElements = xmlDoc.getElementsByTagName('w:p').length > 0
      ? xmlDoc.getElementsByTagName('w:p')
      : xmlDoc.getElementsByTagName('p');

    for (let i = 0; i < pElements.length; i++) {
      const pNode = pElements[i];

      // Check if this paragraph is inside a table (handled separately)
      if (pNode.closest('w:tbl') || pNode.closest('tbl')) continue;

      let pType: DocParagraph['type'] = 'p';
      let align: DocParagraph['align'] = 'left';

      // Check heading style
      const pStyle = pNode.getElementsByTagName('w:pStyle')[0] || pNode.getElementsByTagName('pStyle')[0];
      if (pStyle) {
        const val = (pStyle.getAttribute('w:val') || pStyle.getAttribute('val') || '').toLowerCase();
        if (val.includes('heading1') || val.includes('heading 1') || val.includes('title')) pType = 'h1';
        else if (val.includes('heading2') || val.includes('heading 2')) pType = 'h2';
        else if (val.includes('heading3') || val.includes('heading 3')) pType = 'h3';
      }

      // Check alignment
      const jc = pNode.getElementsByTagName('w:jc')[0] || pNode.getElementsByTagName('jc')[0];
      if (jc) {
        const jcVal = jc.getAttribute('w:val') || jc.getAttribute('val');
        if (jcVal === 'center') align = 'center';
        else if (jcVal === 'right') align = 'right';
        else if (jcVal === 'both') align = 'justify';
      }

      const text = extractRunText(pNode).trim();

      if (text.length > 0) {
        paragraphs.push({
          id: `docx-p-${i}-${Math.random().toString(36).substr(2, 4)}`,
          type: pType,
          content: text,
          align,
        });
      }
    }

    // Extract Tables (prefixed or unprefixed)
    const tblElements = xmlDoc.getElementsByTagName('w:tbl').length > 0
      ? xmlDoc.getElementsByTagName('w:tbl')
      : xmlDoc.getElementsByTagName('tbl');

    for (let t = 0; t < tblElements.length; t++) {
      const tblNode = tblElements[t];
      const rows = tblNode.getElementsByTagName('w:tr').length > 0
        ? tblNode.getElementsByTagName('w:tr')
        : tblNode.getElementsByTagName('tr');
      const tableData: string[][] = [];

      for (let r = 0; r < rows.length; r++) {
        const rowNode = rows[r];
        const cells = rowNode.getElementsByTagName('w:tc').length > 0
          ? rowNode.getElementsByTagName('w:tc')
          : rowNode.getElementsByTagName('tc');
        const rowValues: string[] = [];

        for (let c = 0; c < cells.length; c++) {
          const cellNode = cells[c];
          rowValues.push(extractRunText(cellNode).trim());
        }
        if (rowValues.length > 0) tableData.push(rowValues);
      }

      if (tableData.length > 0) {
        paragraphs.push({
          id: `docx-tbl-${t}-${Math.random().toString(36).substr(2, 4)}`,
          type: 'table',
          content: `Imported Table #${t + 1}`,
          tableData,
        });
      }
    }

    if (paragraphs.length === 0) {
      paragraphs.push({
        id: 'docx-empty',
        type: 'p',
        content: 'Document contains no readable text content.',
      });
    }

    return {
      title,
      paragraphs,
      pageSize: 'A4',
      margins: { top: 24, right: 28, bottom: 24, left: 28 },
    };
  }

  // --- ODT Parser (LibreOffice Writer) ---
  private static async parseOdt(zip: JSZip, title: string): Promise<DocumentData> {
    const contentXml = zip.file('content.xml');
    if (!contentXml) throw new Error('Invalid ODT: missing content.xml');

    const xmlStr = await contentXml.async('text');
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');

    const paragraphs: DocParagraph[] = [];
    const textNodes = xmlDoc.querySelectorAll('text\\:h, text\\:p, table\\:table, h, p, table');

    let counter = 0;
    textNodes.forEach((node) => {
      counter++;
      const tagName = node.tagName.toLowerCase();

      if (tagName.includes('h')) {
        const level = node.getAttribute('text:outline-level') || '1';
        paragraphs.push({
          id: `odt-h-${counter}`,
          type: level === '1' ? 'h1' : level === '2' ? 'h2' : 'h3',
          content: node.textContent?.trim() || '',
        });
      } else if (tagName.includes('table')) {
        const rows = node.querySelectorAll('table\\:table-row, table-row');
        const tableData: string[][] = [];
        rows.forEach((r) => {
          const cells = r.querySelectorAll('table\\:table-cell, table-cell');
          const rowVals: string[] = [];
          cells.forEach((c) => rowVals.push(c.textContent?.trim() || ''));
          if (rowVals.length > 0) tableData.push(rowVals);
        });
        if (tableData.length > 0) {
          paragraphs.push({
            id: `odt-tbl-${counter}`,
            type: 'table',
            content: 'LibreOffice Table',
            tableData,
          });
        }
      } else if (tagName.includes('p')) {
        const text = node.textContent?.trim() || '';
        if (text.length > 0) {
          paragraphs.push({
            id: `odt-p-${counter}`,
            type: 'p',
            content: text,
          });
        }
      }
    });

    return {
      title,
      paragraphs:
        paragraphs.length > 0
          ? paragraphs
          : [{ id: 'odt-1', type: 'p', content: 'Empty document' }],
      pageSize: 'A4',
      margins: { top: 24, right: 28, bottom: 24, left: 28 },
    };
  }

  // --- XLSX Parser (Excel) ---
  private static async parseXlsx(zip: JSZip): Promise<SpreadsheetData> {
    // 1. Read shared strings table if present
    const sharedStrings: string[] = [];
    const ssFile = zip.file('xl/sharedStrings.xml') || zip.file('xl/SharedStrings.xml');
    if (ssFile) {
      const ssXml = await ssFile.async('text');
      const parser = new DOMParser();
      const ssDoc = parser.parseFromString(ssXml, 'application/xml');
      const siNodes = ssDoc.getElementsByTagName('si');
      for (let i = 0; i < siNodes.length; i++) {
        const tNodes = siNodes[i].getElementsByTagName('t');
        let str = '';
        for (let j = 0; j < tNodes.length; j++) {
          str += tNodes[j].textContent || '';
        }
        sharedStrings.push(str);
      }
    }

    // Helper: Column letters (e.g. A, Z, AA, BC) to 1-based number
    const colLetterToNumber = (col: string): number => {
      let num = 0;
      const upper = col.toUpperCase();
      for (let i = 0; i < upper.length; i++) {
        num = num * 26 + (upper.charCodeAt(i) - 64);
      }
      return num;
    };

    // 2. Discover sheets from workbook.xml + relationships
    interface SheetMeta {
      name: string;
      filePath: string;
    }
    const sheetMetaList: SheetMeta[] = [];
    const wbFile = zip.file('xl/workbook.xml') || zip.file('xl/Workbook.xml');
    const wbRelsFile = zip.file('xl/_rels/workbook.xml.rels') || zip.file('xl/_rels/Workbook.xml.rels');
    const relsMap: Record<string, string> = {};

    if (wbRelsFile) {
      const relsXml = await wbRelsFile.async('text');
      const parser = new DOMParser();
      const relsDoc = parser.parseFromString(relsXml, 'application/xml');
      const relNodes = relsDoc.getElementsByTagName('Relationship');
      for (let i = 0; i < relNodes.length; i++) {
        const id = relNodes[i].getAttribute('Id');
        let target = relNodes[i].getAttribute('Target');
        if (id && target) {
          target = target.replace(/^\/xl\//i, 'xl/').replace(/^\//, '');
          if (!target.toLowerCase().startsWith('xl/')) {
            target = 'xl/' + target;
          }
          relsMap[id] = target;
        }
      }
    }

    if (wbFile) {
      const wbXml = await wbFile.async('text');
      const parser = new DOMParser();
      const wbDoc = parser.parseFromString(wbXml, 'application/xml');
      const sheetNodes = wbDoc.getElementsByTagName('sheet');
      for (let i = 0; i < sheetNodes.length; i++) {
        const name = sheetNodes[i].getAttribute('name') || `Sheet ${i + 1}`;
        const rId = sheetNodes[i].getAttribute('r:id') || sheetNodes[i].getAttribute('id');
        const targetPath = rId && relsMap[rId] ? relsMap[rId] : `xl/worksheets/sheet${i + 1}.xml`;
        sheetMetaList.push({ name, filePath: targetPath });
      }
    }

    // Fallback if workbook discovery had no sheets: search zip for all worksheets
    if (sheetMetaList.length === 0) {
      const wsFiles = Object.keys(zip.files)
        .filter((f) => /^xl\/worksheets\/sheet.*\.xml$/i.test(f))
        .sort((a, b) => {
          const numA = parseInt(a.replace(/\D/g, '') || '0', 10);
          const numB = parseInt(b.replace(/\D/g, '') || '0', 10);
          return numA - numB;
        });

      wsFiles.forEach((filePath, idx) => {
        sheetMetaList.push({ name: `Sheet ${idx + 1}`, filePath });
      });
    }

    const sheets: SheetData[] = [];

    for (let sIdx = 0; sIdx < sheetMetaList.length; sIdx++) {
      const { name, filePath } = sheetMetaList[sIdx];
      const sheetFile = zip.file(filePath) || zip.file(filePath.toLowerCase());
      if (!sheetFile) continue;

      const sheetXml = await sheetFile.async('text');
      const parser = new DOMParser();
      const sheetDoc = parser.parseFromString(sheetXml, 'application/xml');

      const cells: Record<string, CellData> = {};
      const cElements = sheetDoc.getElementsByTagName('c');

      let maxRow = 15;
      let maxCol = 10;

      for (let i = 0; i < cElements.length; i++) {
        const cNode = cElements[i];
        const cellCoord = cNode.getAttribute('r');
        if (!cellCoord) continue;

        const cellType = cNode.getAttribute('t');
        const fNode = cNode.getElementsByTagName('f')[0];
        const vNode = cNode.getElementsByTagName('v')[0];
        const isNode = cNode.getElementsByTagName('is')[0];

        let rawVal = '';

        // Priority 1: Formula
        if (fNode && fNode.textContent) {
          rawVal = `=${fNode.textContent.trim()}`;
        }
        // Priority 2: Inline String (<is><t>...</t></is> or <is><r><t>...</t></r></is>)
        else if (cellType === 'inlineStr' || isNode) {
          const targetNode = isNode || cNode;
          const tNodes = targetNode.getElementsByTagName('t');
          let inlineText = '';
          for (let k = 0; k < tNodes.length; k++) {
            inlineText += tNodes[k].textContent || '';
          }
          rawVal = inlineText;
        }
        // Priority 3: Cell Value (<v>)
        else if (vNode) {
          const valText = vNode.textContent?.trim() || '';
          if (cellType === 's') {
            const ssIndex = parseInt(valText, 10);
            rawVal = sharedStrings[ssIndex] !== undefined ? sharedStrings[ssIndex] : valText;
          } else if (cellType === 'b') {
            rawVal = valText === '1' ? 'TRUE' : 'FALSE';
          } else {
            rawVal = valText;
          }
        }
        // Priority 4: Direct <t> inside <c>
        else {
          const directT = cNode.getElementsByTagName('t')[0];
          if (directT && directT.textContent) {
            rawVal = directT.textContent;
          }
        }

        // Track max row and max column
        const match = cellCoord.match(/^([A-Z]+)([0-9]+)$/i);
        if (match) {
          const colLetters = match[1];
          const rowNum = parseInt(match[2], 10);
          if (rowNum > maxRow) maxRow = rowNum;
          const colNum = colLetterToNumber(colLetters);
          if (colNum > maxCol) maxCol = colNum;
        }

        if (rawVal !== '') {
          cells[cellCoord] = { raw: rawVal };
        }
      }

      sheets.push({
        id: `xlsx-sheet-${sIdx + 1}`,
        name,
        rowCount: Math.max(50, maxRow + 10),
        colCount: Math.max(26, maxCol + 2),
        cells,
      });
    }

    if (sheets.length === 0) {
      sheets.push({
        id: 'xlsx-empty',
        name: 'Sheet 1',
        rowCount: 50,
        colCount: 26,
        cells: {},
      });
    }

    return {
      sheets,
      activeSheetId: sheets[0].id,
    };
  }

  // --- ODS Parser (LibreOffice Calc) ---
  private static async parseOds(zip: JSZip): Promise<SpreadsheetData> {
    const contentXml = zip.file('content.xml');
    if (!contentXml) throw new Error('Invalid ODS: missing content.xml');

    const xmlStr = await contentXml.async('text');
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');

    const sheets: SheetData[] = [];
    const tableNodes = xmlDoc.querySelectorAll('table\\:table, table');

    // Helper: Number to Column letter (1 -> A, 27 -> AA, etc.)
    const numberToColLetter = (n: number): string => {
      let s = '';
      while (n > 0) {
        const m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - m) / 26);
      }
      return s;
    };

    let sheetCounter = 1;
    tableNodes.forEach((tableNode) => {
      const sheetName =
        tableNode.getAttribute('table:name') || `Sheet ${sheetCounter}`;
      const cells: Record<string, CellData> = {};

      const rowNodes = tableNode.querySelectorAll('table\\:table-row, table-row');
      let rIdx = 1;
      let maxCol = 10;

      rowNodes.forEach((rowNode) => {
        let cIdx = 1;
        const cellNodes = rowNode.querySelectorAll('table\\:table-cell, table-cell');

        cellNodes.forEach((cellNode) => {
          const repeat = parseInt(
            cellNode.getAttribute('table:number-columns-repeated') || '1',
            10
          );
          const formula = cellNode.getAttribute('table:formula');
          const pNode = cellNode.querySelector('text\\:p, p');
          const textVal =
            pNode?.textContent?.trim() ||
            cellNode.getAttribute('office:string-value') ||
            cellNode.getAttribute('office:value') ||
            '';

          if (formula || textVal) {
            const colLetter = numberToColLetter(cIdx);
            const cellCoord = `${colLetter}${rIdx}`;

            let rawVal = textVal;
            if (formula) {
              rawVal = formula
                .replace(/^of:=/, '=')
                .replace(/\[\./g, '')
                .replace(/\]/g, '');
            }

            cells[cellCoord] = { raw: rawVal };
            if (cIdx > maxCol) maxCol = cIdx;
          }
          cIdx += Math.min(repeat, 250);
        });
        rIdx++;
      });

      sheets.push({
        id: `ods-sheet-${sheetCounter}`,
        name: sheetName,
        rowCount: Math.max(50, rIdx + 5),
        colCount: Math.max(26, maxCol + 2),
        cells,
      });
      sheetCounter++;
    });

    if (sheets.length === 0) {
      sheets.push({
        id: 'ods-empty',
        name: 'Sheet 1',
        rowCount: 50,
        colCount: 26,
        cells: {},
      });
    }

    return {
      sheets,
      activeSheetId: sheets[0].id,
    };
  }

  // --- PPTX Parser (PowerPoint) ---
  private static async parsePptx(zip: JSZip): Promise<PresentationData> {
    const slides: SlideData[] = [];

    // Discover slides in presentation order
    const slidePaths: string[] = [];
    const presFile = zip.file('ppt/presentation.xml') || zip.file('ppt/Presentation.xml');
    const presRelsFile =
      zip.file('ppt/_rels/presentation.xml.rels') || zip.file('ppt/_rels/Presentation.xml.rels');
    const relsMap: Record<string, string> = {};

    if (presRelsFile) {
      const relsXml = await presRelsFile.async('text');
      const parser = new DOMParser();
      const relsDoc = parser.parseFromString(relsXml, 'application/xml');
      const relNodes = relsDoc.getElementsByTagName('Relationship');
      for (let i = 0; i < relNodes.length; i++) {
        const id = relNodes[i].getAttribute('Id');
        let target = relNodes[i].getAttribute('Target');
        if (id && target) {
          target = target.replace(/^\/ppt\//i, 'ppt/').replace(/^\//, '');
          if (!target.toLowerCase().startsWith('ppt/')) target = 'ppt/' + target;
          relsMap[id] = target;
        }
      }
    }

    if (presFile) {
      const presXml = await presFile.async('text');
      const parser = new DOMParser();
      const presDoc = parser.parseFromString(presXml, 'application/xml');
      const sldIdNodes =
        presDoc.getElementsByTagName('p:sldId').length > 0
          ? presDoc.getElementsByTagName('p:sldId')
          : presDoc.getElementsByTagName('sldId');

      for (let i = 0; i < sldIdNodes.length; i++) {
        const rId = sldIdNodes[i].getAttribute('r:id') || sldIdNodes[i].getAttribute('id');
        if (rId && relsMap[rId]) {
          slidePaths.push(relsMap[rId]);
        }
      }
    }

    // Fallback: search zip for all slide files
    if (slidePaths.length === 0) {
      const foundFiles = Object.keys(zip.files)
        .filter((f) => /^ppt\/slides\/slide.*\.xml$/i.test(f))
        .sort((a, b) => {
          const numA = parseInt(a.replace(/\D/g, '') || '0', 10);
          const numB = parseInt(b.replace(/\D/g, '') || '0', 10);
          return numA - numB;
        });
      slidePaths.push(...foundFiles);
    }

    for (let sIdx = 0; sIdx < slidePaths.length; sIdx++) {
      const slidePath = slidePaths[sIdx];
      const slideFile = zip.file(slidePath) || zip.file(slidePath.toLowerCase());
      if (!slideFile) continue;

      const slideXml = await slideFile.async('text');
      const parser = new DOMParser();
      const slideDoc = parser.parseFromString(slideXml, 'application/xml');

      const elements: SlideElement[] = [];

      // Extract shapes (p:sp or sp)
      const spElements =
        slideDoc.getElementsByTagName('p:sp').length > 0
          ? slideDoc.getElementsByTagName('p:sp')
          : slideDoc.getElementsByTagName('sp');

      let titleFound = false;

      for (let i = 0; i < spElements.length; i++) {
        const spNode = spElements[i];
        const tNodes =
          spNode.getElementsByTagName('a:t').length > 0
            ? spNode.getElementsByTagName('a:t')
            : spNode.getElementsByTagName('t');

        let text = '';
        for (let j = 0; j < tNodes.length; j++) {
          text += (tNodes[j].textContent || '') + ' ';
        }
        text = text.trim();
        if (text.length === 0) continue;

        if (!titleFound) {
          titleFound = true;
          elements.push({
            id: `pptx-el-${sIdx + 1}-${i}`,
            type: 'heading',
            content: text,
            x: 8,
            y: 15,
            width: 84,
            style: { fontSize: 34, bold: true, color: '#ffffff', align: 'left' },
          });
        } else {
          elements.push({
            id: `pptx-el-${sIdx + 1}-${i}`,
            type: 'shape',
            content: text,
            x: 8,
            y: 35 + elements.length * 15,
            width: 84,
            style: { bgColor: '#1e293b', color: '#f1f5f9', align: 'left' },
          });
        }
      }

      slides.push({
        id: `pptx-slide-${sIdx + 1}`,
        title: elements[0]?.content || `Slide ${sIdx + 1}`,
        layout: 'title-bullets',
        bgColor: '#0f172a',
        elements:
          elements.length > 0
            ? elements
            : [
                {
                  id: `empty-el-${sIdx + 1}`,
                  type: 'heading',
                  content: `Slide ${sIdx + 1}`,
                  x: 8,
                  y: 40,
                  width: 84,
                  style: { fontSize: 32, bold: true, color: '#ffffff', align: 'center' },
                },
              ],
      });
    }

    return {
      theme: 'modern-dark',
      aspectRatio: '16:9',
      slides:
        slides.length > 0
          ? slides
          : [{ id: 'empty-1', title: 'Slide 1', layout: 'title', elements: [] }],
    };
  }

  // --- ODP Parser (LibreOffice Impress) ---
  private static async parseOdp(zip: JSZip): Promise<PresentationData> {
    const contentXml = zip.file('content.xml');
    if (!contentXml) throw new Error('Invalid ODP: missing content.xml');

    const xmlStr = await contentXml.async('text');
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');

    const slides: SlideData[] = [];
    const pageNodes = xmlDoc.querySelectorAll('draw\\:page, page');

    let pageIdx = 1;
    pageNodes.forEach((page) => {
      const elements: SlideElement[] = [];
      const textNodes = page.querySelectorAll('text\\:p, p');

      let isFirst = true;
      textNodes.forEach((tNode, idx) => {
        const text = tNode.textContent?.trim();
        if (!text) return;

        if (isFirst) {
          isFirst = false;
          elements.push({
            id: `odp-el-${pageIdx}-${idx}`,
            type: 'heading',
            content: text,
            x: 8,
            y: 15,
            width: 84,
            style: { fontSize: 34, bold: true, color: '#ffffff', align: 'left' },
          });
        } else {
          elements.push({
            id: `odp-el-${pageIdx}-${idx}`,
            type: 'shape',
            content: text,
            x: 8,
            y: 32 + elements.length * 14,
            width: 84,
            style: { bgColor: '#1e293b', color: '#f1f5f9' },
          });
        }
      });

      slides.push({
        id: `odp-slide-${pageIdx}`,
        title: elements[0]?.content || `Slide ${pageIdx}`,
        layout: 'title-bullets',
        bgColor: '#0f172a',
        elements,
      });
      pageIdx++;
    });

    return {
      theme: 'modern-dark',
      aspectRatio: '16:9',
      slides,
    };
  }

  // --- Utilities: CSV, Text, HTML ---
  private static parseCsvToSheet(text: string, sheetName: string, delimiter: string): SheetData {
    const lines = text.split(/\r?\n/);
    const cells: Record<string, CellData> = {};
    let maxCol = 10;

    lines.forEach((line, rIdx) => {
      if (!line.trim()) return;
      const parts = line.split(delimiter);
      if (parts.length > maxCol) maxCol = parts.length;

      parts.forEach((val, cIdx) => {
        if (cIdx < 26) {
          const colLetter = String.fromCharCode(65 + cIdx);
          const cellCoord = `${colLetter}${rIdx + 1}`;
          const clean = val.replace(/^["']|["']$/g, '').trim();
          if (clean !== '') {
            cells[cellCoord] = { raw: clean };
          }
        }
      });
    });

    return {
      id: `csv-${Date.now()}`,
      name: sheetName || 'Imported Data',
      rowCount: Math.max(25, lines.length + 5),
      colCount: Math.min(26, Math.max(12, maxCol + 2)),
      cells,
    };
  }

  private static parseTextToDoc(text: string, title: string): DocumentData {
    const lines = text.split(/\r?\n/);
    const paragraphs: DocParagraph[] = [];

    lines.forEach((line, idx) => {
      const trimmed = line.trim();
      if (!trimmed) return;

      if (trimmed.startsWith('# ')) {
        paragraphs.push({ id: `txt-${idx}`, type: 'h1', content: trimmed.replace(/^#\s+/, '') });
      } else if (trimmed.startsWith('## ')) {
        paragraphs.push({ id: `txt-${idx}`, type: 'h2', content: trimmed.replace(/^##\s+/, '') });
      } else if (trimmed.startsWith('* ') || trimmed.startsWith('- ')) {
        paragraphs.push({ id: `txt-${idx}`, type: 'bullet', content: trimmed.replace(/^[-*]\s+/, '') });
      } else {
        paragraphs.push({ id: `txt-${idx}`, type: 'p', content: trimmed });
      }
    });

    return {
      title,
      paragraphs:
        paragraphs.length > 0
          ? paragraphs
          : [{ id: 'txt-1', type: 'p', content: text }],
      pageSize: 'A4',
      margins: { top: 24, right: 28, bottom: 24, left: 28 },
    };
  }

  private static parseHtmlToDoc(html: string, title: string): DocumentData {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const paragraphs: DocParagraph[] = [];

    const nodes = doc.body.querySelectorAll('h1, h2, h3, p, li, table');
    let idx = 0;
    nodes.forEach((node) => {
      idx++;
      const tag = node.tagName.toLowerCase();
      if (tag === 'h1') {
        paragraphs.push({ id: `html-${idx}`, type: 'h1', content: node.textContent?.trim() || '' });
      } else if (tag === 'h2') {
        paragraphs.push({ id: `html-${idx}`, type: 'h2', content: node.textContent?.trim() || '' });
      } else if (tag === 'h3') {
        paragraphs.push({ id: `html-${idx}`, type: 'h3', content: node.textContent?.trim() || '' });
      } else if (tag === 'li') {
        paragraphs.push({ id: `html-${idx}`, type: 'bullet', content: node.textContent?.trim() || '' });
      } else if (tag === 'table') {
        const rows = node.querySelectorAll('tr');
        const tableData: string[][] = [];
        rows.forEach((r) => {
          const cells = r.querySelectorAll('th, td');
          const rowVals: string[] = [];
          cells.forEach((c) => rowVals.push(c.textContent?.trim() || ''));
          if (rowVals.length > 0) tableData.push(rowVals);
        });
        if (tableData.length > 0) {
          paragraphs.push({ id: `html-${idx}`, type: 'table', content: 'HTML Table', tableData });
        }
      } else {
        const text = node.textContent?.trim() || '';
        if (text) {
          paragraphs.push({ id: `html-${idx}`, type: 'p', content: text });
        }
      }
    });

    return {
      title,
      paragraphs:
        paragraphs.length > 0
          ? paragraphs
          : [{ id: 'html-1', type: 'p', content: doc.body.textContent || '' }],
      pageSize: 'A4',
      margins: { top: 24, right: 28, bottom: 24, left: 28 },
    };
  }

  private static parseTextToSlides(text: string): PresentationData {
    const lines = text.split(/\r?\n/).filter((l) => l.trim().length > 0);
    const slides: SlideData[] = [];

    for (let i = 0; i < lines.length; i += 4) {
      const chunk = lines.slice(i, i + 4);
      slides.push({
        id: `slide-${i}`,
        title: chunk[0] || 'Presentation Slide',
        layout: 'title-bullets',
        bgColor: '#0f172a',
        elements: [
          {
            id: `el-${i}-1`,
            type: 'heading',
            content: chunk[0] || 'Slide Title',
            x: 8,
            y: 15,
            width: 84,
            style: { fontSize: 32, bold: true, color: '#fff' },
          },
          {
            id: `el-${i}-2`,
            type: 'bullet-list',
            content: chunk.slice(1).join('\n'),
            x: 8,
            y: 40,
            width: 84,
            style: { fontSize: 16, color: '#cbd5e1' },
          },
        ],
      });
    }

    return {
      theme: 'modern-dark',
      aspectRatio: '16:9',
      slides:
        slides.length > 0
          ? slides
          : [{ id: 'empty-1', title: 'Slide 1', layout: 'title', elements: [] }],
    };
  }

  // --- Binary BIFF Parser (.xls) ---
  private static parseLegacyBiffXls(bytes: Uint8Array, fileName: string): SpreadsheetData {
    // Attempt BIFF record extraction
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const sheets: { name: string; cells: Record<string, CellData>; maxRow: number; maxCol: number }[] = [];
    const sharedStrings: string[] = [];

    let currentSheetIndex = -1;
    let offset = 0;
    const len = bytes.byteLength;

    // Helper to decode cell key (e.g., A1, B2)
    const toCellKey = (row: number, col: number) => {
      let colName = '';
      let c = col;
      while (c >= 0) {
        colName = String.fromCharCode(65 + (c % 26)) + colName;
        c = Math.floor(c / 26) - 1;
      }
      return `${colName}${row + 1}`;
    };

    try {
      while (offset + 4 <= len) {
        const recordType = view.getUint16(offset, true);
        const recordLen = view.getUint16(offset + 2, true);
        const dataOffset = offset + 4;
        if (dataOffset + recordLen > len) break;

        // 0x0085 = BOUNDSHEET (Sheet name)
        if (recordType === 0x0085 && recordLen >= 8) {
          const sheetNameLen = view.getUint8(dataOffset + 6);
          const isUnicode = recordLen > 7 && (view.getUint8(dataOffset + 7) & 0x01) !== 0;
          let sheetName = '';
          const strStart = isUnicode ? dataOffset + 8 : dataOffset + 7;
          if (isUnicode) {
            for (let i = 0; i < sheetNameLen && strStart + i * 2 + 1 < dataOffset + recordLen; i++) {
              sheetName += String.fromCharCode(view.getUint16(strStart + i * 2, true));
            }
          } else {
            for (let i = 0; i < sheetNameLen && strStart + i < dataOffset + recordLen; i++) {
              sheetName += String.fromCharCode(view.getUint8(strStart + i));
            }
          }
          if (sheetName.trim()) {
            sheets.push({
              name: sheetName.trim(),
              cells: {},
              maxRow: 0,
              maxCol: 0,
            });
          }
        }

        // 0x0809 = BOF (Beginning of Worksheet / Workbook)
        if (recordType === 0x0809 && recordLen >= 4) {
          const subType = view.getUint16(dataOffset + 2, true);
          if (subType === 0x0010) {
            // Worksheet
            currentSheetIndex++;
            if (currentSheetIndex >= sheets.length) {
              sheets.push({
                name: `Sheet${currentSheetIndex + 1}`,
                cells: {},
                maxRow: 0,
                maxCol: 0,
              });
            }
          }
        }

        // 0x00FC = SST (Shared String Table)
        if (recordType === 0x00FC && recordLen >= 8) {
          const totalStrings = view.getUint32(dataOffset + 4, true);
          let sstOffset = dataOffset + 8;
          for (let i = 0; i < Math.min(totalStrings, 2000) && sstOffset + 3 < dataOffset + recordLen; i++) {
            const strLen = view.getUint16(sstOffset, true);
            const flag = view.getUint8(sstOffset + 2);
            const isUni = (flag & 0x01) !== 0;
            let str = '';
            const charOffset = sstOffset + 3;
            if (isUni) {
              for (let j = 0; j < strLen && charOffset + j * 2 + 1 < dataOffset + recordLen; j++) {
                str += String.fromCharCode(view.getUint16(charOffset + j * 2, true));
              }
              sstOffset = charOffset + strLen * 2;
            } else {
              for (let j = 0; j < strLen && charOffset + j < dataOffset + recordLen; j++) {
                str += String.fromCharCode(view.getUint8(charOffset + j));
              }
              sstOffset = charOffset + strLen;
            }
            sharedStrings.push(str);
          }
        }

        const activeSheet = sheets[Math.max(0, currentSheetIndex)];

        if (activeSheet) {
          // 0x00FD = LABELSST (Cell string from Shared String Table)
          if (recordType === 0x00FD && recordLen >= 10) {
            const row = view.getUint16(dataOffset, true);
            const col = view.getUint16(dataOffset + 2, true);
            const sstIdx = view.getUint32(dataOffset + 6, true);
            const val = sharedStrings[sstIdx] || '';
            const key = toCellKey(row, col);
            activeSheet.cells[key] = { raw: val, computed: val };
            activeSheet.maxRow = Math.max(activeSheet.maxRow, row);
            activeSheet.maxCol = Math.max(activeSheet.maxCol, col);
          }

          // 0x0204 = LABEL (Cell string inline)
          if (recordType === 0x0204 && recordLen >= 8) {
            const row = view.getUint16(dataOffset, true);
            const col = view.getUint16(dataOffset + 2, true);
            const strLen = view.getUint16(dataOffset + 6, true);
            let val = '';
            for (let i = 0; i < strLen && dataOffset + 8 + i < dataOffset + recordLen; i++) {
              val += String.fromCharCode(view.getUint8(dataOffset + 8 + i));
            }
            const key = toCellKey(row, col);
            activeSheet.cells[key] = { raw: val, computed: val };
            activeSheet.maxRow = Math.max(activeSheet.maxRow, row);
            activeSheet.maxCol = Math.max(activeSheet.maxCol, col);
          }

          // 0x0203 = NUMBER (IEEE 754 64-bit float)
          if (recordType === 0x0203 && recordLen >= 14) {
            const row = view.getUint16(dataOffset, true);
            const col = view.getUint16(dataOffset + 2, true);
            const numVal = view.getFloat64(dataOffset + 6, true);
            const key = toCellKey(row, col);
            activeSheet.cells[key] = { raw: String(numVal), computed: numVal };
            activeSheet.maxRow = Math.max(activeSheet.maxRow, row);
            activeSheet.maxCol = Math.max(activeSheet.maxCol, col);
          }

          // 0x027E = RK (Packed 30-bit integer / float)
          if (recordType === 0x027E && recordLen >= 10) {
            const row = view.getUint16(dataOffset, true);
            const col = view.getUint16(dataOffset + 2, true);
            const rkRaw = view.getUint32(dataOffset + 6, true);
            let val: number;
            if ((rkRaw & 0x02) !== 0) {
              // integer
              val = rkRaw >> 2;
            } else {
              // 30-bit float
              const tempView = new DataView(new ArrayBuffer(8));
              tempView.setUint32(4, rkRaw & 0xfffffffc, false);
              val = tempView.getFloat64(0, false);
            }
            if ((rkRaw & 0x01) !== 0) val /= 100;
            const key = toCellKey(row, col);
            activeSheet.cells[key] = { raw: String(val), computed: val };
            activeSheet.maxRow = Math.max(activeSheet.maxRow, row);
            activeSheet.maxCol = Math.max(activeSheet.maxCol, col);
          }

          // 0x0006 = FORMULA
          if (recordType === 0x0006 && recordLen >= 14) {
            const row = view.getUint16(dataOffset, true);
            const col = view.getUint16(dataOffset + 2, true);
            const numVal = view.getFloat64(dataOffset + 6, true);
            const key = toCellKey(row, col);
            activeSheet.cells[key] = {
              raw: '=SUM()',
              computed: isNaN(numVal) ? 0 : numVal,
            };
            activeSheet.maxRow = Math.max(activeSheet.maxRow, row);
            activeSheet.maxCol = Math.max(activeSheet.maxCol, col);
          }
        }

        offset += 4 + recordLen;
      }
    } catch (err) {
      console.warn('BIFF record parsing interrupted, synthesizing recovered cells', err);
    }

    // Filter sheets that actually contain cells
    const validSheets = sheets.filter((s) => Object.keys(s.cells).length > 0);

    if (validSheets.length > 0) {
      const parsedSheets: SheetData[] = validSheets.map((s, idx) => ({
        id: `biff-sheet-${idx + 1}`,
        name: s.name || `Sheet${idx + 1}`,
        cells: s.cells,
        rowCount: Math.max(50, s.maxRow + 10),
        colCount: Math.max(26, s.maxCol + 5),
      }));

      return {
        activeSheetId: parsedSheets[0].id,
        sheets: parsedSheets,
      };
    }

    // Fallback: extract strings and parse tabular structure
    const rawText = this.extractStringsFromBinary(bytes);
    const fallbackSheet = this.parseCsvToSheet(rawText, fileName.replace(/\.[^/.]+$/, ''), '\t');
    return {
      activeSheetId: fallbackSheet.id,
      sheets: [fallbackSheet],
    };
  }

  // --- Binary Word Document Parser (.doc) ---
  private static parseLegacyWordDoc(bytes: Uint8Array, fileName: string): DocumentData {
    const title = fileName.replace(/\.[^/.]+$/, '');
    const paragraphs: DocParagraph[] = [];

    // Word 97-2003 stores text in UTF-16LE and 8-bit ANSI runs.
    const extractedParagraphs: string[] = [];

    // 1. Try UTF-16LE extraction (standard for Word 97-2003)
    let utf16Str = '';
    for (let i = 0; i < bytes.byteLength - 1; i += 2) {
      const code = bytes[i] | (bytes[i + 1] << 8);
      if (code === 0x0d || code === 0x0a) {
        if (utf16Str.trim().length > 0) {
          extractedParagraphs.push(utf16Str.trim());
          utf16Str = '';
        }
      } else if ((code >= 32 && code <= 126) || (code >= 160 && code <= 65533)) {
        utf16Str += String.fromCharCode(code);
      } else {
        if (utf16Str.trim().length >= 4) {
          extractedParagraphs.push(utf16Str.trim());
        }
        utf16Str = '';
      }
    }
    if (utf16Str.trim().length >= 4) {
      extractedParagraphs.push(utf16Str.trim());
    }

    // 2. If UTF-16LE didn't find substantial text, try 8-bit ANSI runs
    if (extractedParagraphs.length < 3) {
      const ansiText = this.extractStringsFromBinary(bytes);
      ansiText
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter((l) => l.length > 2)
        .forEach((l) => extractedParagraphs.push(l));
    }

    // Filter out internal binary noise like "WordDocument", "SummaryInformation", font names
    const filtered = extractedParagraphs.filter((line) => {
      if (
        /^(Root Entry|WordDocument|1Table|0Table|Data|SummaryInformation|DocumentSummaryInformation|CompObj)/i.test(
          line
        )
      ) {
        return false;
      }
      if (/^(Times New Roman|Arial|Calibri|Symbol|Wingdings|Courier New|Normal\.dotm?)$/i.test(line)) {
        return false;
      }
      return line.length >= 2;
    });

    if (filtered.length === 0) {
      return {
        title,
        paragraphs: [
          { id: 'doc-title', type: 'h1', content: title, align: 'left' },
          {
            id: 'doc-empty',
            type: 'p',
            content: 'No readable text content extracted from legacy .doc binary stream.',
            align: 'left',
          },
        ],
        pageSize: 'A4',
        margins: { top: 24, right: 28, bottom: 24, left: 28 },
      };
    }

    // Classify paragraphs hierarchically
    filtered.forEach((text, idx) => {
      let type: DocParagraph['type'] = 'p';
      let align: DocParagraph['align'] = 'left';

      if (idx === 0) {
        type = 'h1';
      } else if (/^(\d+\.|\bSection\b|\bChapter\b|\bPart\b)/i.test(text) && text.length < 80) {
        type = 'h2';
      } else if (/^(\d+\.\d+|\bAppendix\b)/i.test(text) && text.length < 100) {
        type = 'h3';
      } else if (/^[•\-*\u2022\u2023]\s/.test(text)) {
        type = 'bullet';
      } else if (/^(Note|Warning|Important|Caution|Notice):/i.test(text)) {
        type = 'callout';
      } else if (text.length > 250) {
        align = 'justify';
      }

      paragraphs.push({
        id: `doc-p-${idx + 1}`,
        type,
        content: text.replace(/^[•\-*\u2022\u2023]\s*/, ''),
        align,
      });
    });

    return {
      title,
      paragraphs,
      pageSize: 'A4',
      margins: { top: 24, right: 28, bottom: 24, left: 28 },
    };
  }

  private static extractStringsFromBinary(bytes: Uint8Array): string {
    let str = '';
    let current = '';
    for (let i = 0; i < bytes.length; i++) {
      const code = bytes[i];
      if (code >= 32 && code <= 126) {
        current += String.fromCharCode(code);
      } else {
        if (current.length >= 4) {
          str += current + '\n';
        }
        current = '';
      }
    }
    return str;
  }
}
