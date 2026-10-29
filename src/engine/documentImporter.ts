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
        }
        const textMatches = raw.match(/\(([^)]+)\)\s*Tj/g);
        if (textMatches && textMatches.length > 0) {
          textSample = textMatches
            .map(m => m.replace(/^\(/, '').replace(/\)\s*Tj$/, ''))
            .filter(t => t.trim().length > 1)
            .join(' ')
            .slice(0, 3000);
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

    // Fallback for older .doc, .xls, .ppt binary strings
    if (ext === 'doc' || ext === 'xls' || ext === 'ppt') {
      const text = this.extractStringsFromBinary(new Uint8Array(arrayBuffer));
      if (ext === 'doc') {
        const doc = this.parseTextToDoc(text, file.name);
        return {
          mode: 'docs',
          projectUpdates: { docs: doc },
          fileName: file.name,
          detectedFormat: 'Legacy Microsoft Word Document (.doc)',
        };
      }
      if (ext === 'xls') {
        const sheet = this.parseCsvToSheet(text, file.name, '\t');
        return {
          mode: 'sheets',
          projectUpdates: { sheets: { activeSheetId: sheet.id, sheets: [sheet] } },
          fileName: file.name,
          detectedFormat: 'Legacy Microsoft Excel (.xls)',
        };
      }
      if (ext === 'ppt') {
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
    const docXmlFile = zip.file('word/document.xml');
    if (!docXmlFile) throw new Error('Invalid DOCX: missing word/document.xml');

    const xmlStr = await docXmlFile.async('text');
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');

    const paragraphs: DocParagraph[] = [];
    const pElements = xmlDoc.getElementsByTagName('w:p');

    for (let i = 0; i < pElements.length; i++) {
      const pNode = pElements[i];

      // Check if this paragraph is inside a table (we handle tables separately)
      if (pNode.closest('w:tbl')) continue;

      let pType: DocParagraph['type'] = 'p';
      let align: DocParagraph['align'] = 'left';

      // Check heading style
      const pStyle = pNode.getElementsByTagName('w:pStyle')[0];
      if (pStyle) {
        const val = pStyle.getAttribute('w:val')?.toLowerCase() || '';
        if (val.includes('heading1') || val.includes('title')) pType = 'h1';
        else if (val.includes('heading2')) pType = 'h2';
        else if (val.includes('heading3')) pType = 'h3';
      }

      // Check alignment
      const jc = pNode.getElementsByTagName('w:jc')[0];
      if (jc) {
        const jcVal = jc.getAttribute('w:val');
        if (jcVal === 'center') align = 'center';
        else if (jcVal === 'right') align = 'right';
        else if (jcVal === 'both') align = 'justify';
      }

      // Extract text from w:t nodes
      const tNodes = pNode.getElementsByTagName('w:t');
      let text = '';
      for (let j = 0; j < tNodes.length; j++) {
        text += tNodes[j].textContent || '';
      }

      if (text.trim().length > 0) {
        paragraphs.push({
          id: `docx-p-${i}-${Math.random().toString(36).substr(2, 4)}`,
          type: pType,
          content: text.trim(),
          align,
        });
      }
    }

    // Extract Tables
    const tblElements = xmlDoc.getElementsByTagName('w:tbl');
    for (let t = 0; t < tblElements.length; t++) {
      const tblNode = tblElements[t];
      const rows = tblNode.getElementsByTagName('w:tr');
      const tableData: string[][] = [];

      for (let r = 0; r < rows.length; r++) {
        const rowNode = rows[r];
        const cells = rowNode.getElementsByTagName('w:tc');
        const rowValues: string[] = [];

        for (let c = 0; c < cells.length; c++) {
          const cellNode = cells[c];
          const tNodes = cellNode.getElementsByTagName('w:t');
          let cellText = '';
          for (let k = 0; k < tNodes.length; k++) {
            cellText += tNodes[k].textContent || '';
          }
          rowValues.push(cellText.trim());
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
    const ssFile = zip.file('xl/sharedStrings.xml');
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

    // 2. Discover sheets
    const sheets: SheetData[] = [];
    let sheetIdx = 1;

    while (true) {
      const sheetFile = zip.file(`xl/worksheets/sheet${sheetIdx}.xml`);
      if (!sheetFile) break;

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

        let rawVal = '';
        if (fNode && fNode.textContent) {
          rawVal = `=${fNode.textContent.trim()}`;
        } else if (vNode) {
          const valText = vNode.textContent?.trim() || '';
          if (cellType === 's') {
            const ssIndex = parseInt(valText, 10);
            rawVal = sharedStrings[ssIndex] || valText;
          } else {
            rawVal = valText;
          }
        }

        const match = cellCoord.match(/^([A-Z]+)([0-9]+)$/);
        if (match) {
          const rowNum = parseInt(match[2], 10);
          if (rowNum > maxRow) maxRow = rowNum;
        }

        if (rawVal !== '') {
          cells[cellCoord] = { raw: rawVal };
        }
      }

      sheets.push({
        id: `xlsx-sheet-${sheetIdx}`,
        name: `Sheet ${sheetIdx}`,
        rowCount: Math.min(100, Math.max(25, maxRow + 5)),
        colCount: Math.min(26, Math.max(12, maxCol + 2)),
        cells,
      });

      sheetIdx++;
    }

    if (sheets.length === 0) {
      sheets.push({
        id: 'xlsx-empty',
        name: 'Sheet 1',
        rowCount: 25,
        colCount: 12,
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

    let sheetCounter = 1;
    tableNodes.forEach((tableNode) => {
      const sheetName =
        tableNode.getAttribute('table:name') || `Sheet ${sheetCounter}`;
      const cells: Record<string, CellData> = {};

      const rowNodes = tableNode.querySelectorAll('table\\:table-row, table-row');
      let rIdx = 1;

      rowNodes.forEach((rowNode) => {
        let cIdx = 0;
        const cellNodes = rowNode.querySelectorAll('table\\:table-cell, table-cell');

        cellNodes.forEach((cellNode) => {
          const repeat = parseInt(
            cellNode.getAttribute('table:number-columns-repeated') || '1',
            10
          );
          const formula = cellNode.getAttribute('table:formula');
          const pNode = cellNode.querySelector('text\\:p, p');
          const textVal = pNode?.textContent?.trim() || '';

          if ((formula || textVal) && cIdx < 26) {
            const colLetter = String.fromCharCode(65 + cIdx);
            const cellCoord = `${colLetter}${rIdx}`;

            let rawVal = textVal;
            if (formula) {
              // Normalize of:=SUM([.A1:.A5]) to =SUM(A1:A5)
              rawVal = formula
                .replace(/^of:=/, '=')
                .replace(/\[\./g, '')
                .replace(/\]/g, '');
            }

            cells[cellCoord] = { raw: rawVal };
          }
          cIdx += Math.min(repeat, 10);
        });
        rIdx++;
      });

      sheets.push({
        id: `ods-sheet-${sheetCounter}`,
        name: sheetName,
        rowCount: Math.max(25, rIdx + 2),
        colCount: 15,
        cells,
      });
      sheetCounter++;
    });

    return {
      sheets:
        sheets.length > 0
          ? sheets
          : [
              {
                id: 'ods-empty',
                name: 'Sheet 1',
                rowCount: 25,
                colCount: 12,
                cells: {},
              },
            ],
      activeSheetId: sheets[0]?.id || 'ods-empty',
    };
  }

  // --- PPTX Parser (PowerPoint) ---
  private static async parsePptx(zip: JSZip): Promise<PresentationData> {
    const slides: SlideData[] = [];
    let slideIdx = 1;

    while (true) {
      const slideFile = zip.file(`ppt/slides/slide${slideIdx}.xml`);
      if (!slideFile) break;

      const slideXml = await slideFile.async('text');
      const parser = new DOMParser();
      const slideDoc = parser.parseFromString(slideXml, 'application/xml');

      const elements: SlideElement[] = [];

      // Extract text boxes & shapes from p:sp
      const spElements = slideDoc.getElementsByTagName('p:sp');
      let titleFound = false;

      for (let i = 0; i < spElements.length; i++) {
        const spNode = spElements[i];
        const tNodes = spNode.getElementsByTagName('a:t');
        let text = '';
        for (let j = 0; j < tNodes.length; j++) {
          text += (tNodes[j].textContent || '') + ' ';
        }
        text = text.trim();
        if (text.length === 0) continue;

        if (!titleFound) {
          titleFound = true;
          elements.push({
            id: `pptx-el-${slideIdx}-${i}`,
            type: 'heading',
            content: text,
            x: 8,
            y: 15,
            width: 84,
            style: { fontSize: 34, bold: true, color: '#ffffff', align: 'left' },
          });
        } else {
          elements.push({
            id: `pptx-el-${slideIdx}-${i}`,
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
        id: `pptx-slide-${slideIdx}`,
        title: elements[0]?.content || `Slide ${slideIdx}`,
        layout: 'title-bullets',
        bgColor: '#0f172a',
        elements:
          elements.length > 0
            ? elements
            : [
                {
                  id: `pptx-el-empty`,
                  type: 'heading',
                  content: `Slide ${slideIdx}`,
                  x: 10,
                  y: 40,
                  width: 80,
                  style: { fontSize: 32, color: '#fff', align: 'center' },
                },
              ],
      });

      slideIdx++;
    }

    return {
      theme: 'modern-dark',
      aspectRatio: '16:9',
      slides:
        slides.length > 0
          ? slides
          : [
              {
                id: 'pptx-empty',
                title: 'Empty Slide',
                layout: 'title',
                elements: [],
              },
            ],
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
