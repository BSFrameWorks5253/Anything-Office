import JSZip from 'jszip';
import type {
  AnythingProject,
  DocumentData,
  SpreadsheetData,
  PresentationData,
  SheetData,
} from '../types/office';
import { FormulaEngine } from './formulaEngine';

export class DocumentExporter {
  /**
   * Helper to download or save a blob natively or via browser
   */
  public static async downloadBlob(blob: Blob, filename: string): Promise<void> {
    if (window.electronAPI?.saveNativeFileDialog) {
      try {
        const arrayBuffer = await blob.arrayBuffer();
        const uint8Array = new Uint8Array(arrayBuffer);
        const success = await window.electronAPI.saveNativeFileDialog(
          uint8Array as any,
          filename
        );
        if (success) return;
      } catch (err) {
        console.warn('Native save dialog failed, falling back to browser download', err);
      }
    }

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 100);
  }

  /**
   * 1. Export Microsoft Word (.docx) Open XML Archive
   */
  public static async exportDocx(doc: DocumentData): Promise<Blob> {
    const zip = new JSZip();

    // [Content_Types].xml
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`
    );

    // _rels/.rels
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
    );

    // word/_rels/document.xml.rels
    zip.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
    );

    // word/styles.xml
    zip.file(
      'word/styles.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:color w:val="1E293B"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="48"/>
      <w:color w:val="0F172A"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="36"/>
      <w:color w:val="1E293B"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="28"/>
      <w:color w:val="334155"/>
    </w:rPr>
  </w:style>
</w:styles>`
    );

    // Helper for XML escaping
    const escapeXml = (unsafe: string) => {
      return unsafe
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
    };

    // Build document.xml body paragraphs and tables
    let bodyXml = '';

    for (const p of doc.paragraphs) {
      if (p.type === 'table' && p.tableData && p.tableData.length > 0) {
        bodyXml += '<w:tbl>';
        bodyXml += '<w:tblPr><w:tblBorders>';
        bodyXml += '<w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>';
        bodyXml += '<w:left w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>';
        bodyXml += '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>';
        bodyXml += '<w:right w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>';
        bodyXml += '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>';
        bodyXml += '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>';
        bodyXml += '</w:tblBorders></w:tblPr>';

        p.tableData.forEach((row, rowIdx) => {
          bodyXml += '<w:tr>';
          row.forEach((cell) => {
            bodyXml += '<w:tc>';
            if (rowIdx === 0) {
              bodyXml += '<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="F1F5F9"/></w:tcPr>';
            }
            bodyXml += '<w:p><w:r>';
            if (rowIdx === 0) bodyXml += '<w:rPr><w:b/></w:rPr>';
            bodyXml += `<w:t>${escapeXml(cell)}</w:t></w:r></w:p>`;
            bodyXml += '</w:tc>';
          });
          bodyXml += '</w:tr>';
        });
        bodyXml += '</w:tbl>';
        continue;
      }

      // Paragraph styling
      let pStyle = 'Normal';
      let sz = '22';
      let isBold = false;
      let isItalic = false;

      if (p.type === 'h1') {
        pStyle = 'Heading1';
        sz = '48';
        isBold = true;
      } else if (p.type === 'h2') {
        pStyle = 'Heading2';
        sz = '36';
        isBold = true;
      } else if (p.type === 'h3') {
        pStyle = 'Heading3';
        sz = '28';
        isBold = true;
      } else if (p.type === 'blockquote' || p.type === 'callout') {
        isItalic = true;
      }

      let jc = '';
      if (p.align === 'center') jc = '<w:jc w:val="center"/>';
      else if (p.align === 'right') jc = '<w:jc w:val="right"/>';
      else if (p.align === 'justify') jc = '<w:jc w:val="both"/>';

      bodyXml += `<w:p><w:pPr><w:pStyle w:val="${pStyle}"/>${jc}</w:pPr>`;
      bodyXml += '<w:r><w:rPr>';
      if (isBold) bodyXml += '<w:b/>';
      if (isItalic) bodyXml += '<w:i/>';
      bodyXml += `<w:sz w:val="${sz}"/></w:rPr>`;
      bodyXml += `<w:t xml:space="preserve">${escapeXml(p.content || '')}</w:t></w:r></w:p>`;
    }

    // Section Properties (A4: 11906 x 16838 dxa, Letter: 12240 x 15840 dxa)
    const isLandscape = doc.orientation === 'landscape';
    const width = isLandscape ? (doc.pageSize === 'A4' ? 16838 : 15840) : (doc.pageSize === 'A4' ? 11906 : 12240);
    const height = isLandscape ? (doc.pageSize === 'A4' ? 11906 : 12240) : (doc.pageSize === 'A4' ? 16838 : 15840);

    bodyXml += `<w:sectPr>
      <w:pgSz w:w="${width}" w:h="${height}" w:orient="${isLandscape ? 'landscape' : 'portrait'}"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
    </w:sectPr>`;

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${bodyXml}
  </w:body>
</w:document>`;

    zip.file('word/document.xml', documentXml);

    return await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  }

  /**
   * 2. Export Microsoft Excel (.xlsx) Open XML Workbook
   */
  public static async exportXlsx(spreadsheet: SpreadsheetData): Promise<Blob> {
    const zip = new JSZip();

    const escapeXml = (unsafe: string) => {
      return unsafe
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
    };

    // [Content_Types].xml
    let contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`;

    spreadsheet.sheets.forEach((_, idx) => {
      contentTypesXml += `\n  <Override PartName="/xl/worksheets/sheet${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    });
    contentTypesXml += '\n</Types>';
    zip.file('[Content_Types].xml', contentTypesXml);

    // _rels/.rels
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
    );

    // xl/_rels/workbook.xml.rels
    let wbRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;

    spreadsheet.sheets.forEach((_, idx) => {
      wbRelsXml += `\n  <Relationship Id="rIdSheet${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${idx + 1}.xml"/>`;
    });
    wbRelsXml += '\n</Relationships>';
    zip.file('xl/_rels/workbook.xml.rels', wbRelsXml);

    // xl/workbook.xml
    let wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>`;

    spreadsheet.sheets.forEach((sheet, idx) => {
      wbXml += `\n    <sheet name="${escapeXml(sheet.name)}" sheetId="${idx + 1}" r:id="rIdSheet${idx + 1}"/>`;
    });
    wbXml += '\n  </sheets>\n</workbook>';
    zip.file('xl/workbook.xml', wbXml);

    // xl/styles.xml
    zip.file(
      'xl/styles.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>`
    );

    // xl/worksheets/sheetN.xml
    spreadsheet.sheets.forEach((sheet, sheetIdx) => {
      let sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>`;

      for (let r = 1; r <= Math.max(sheet.rowCount, 30); r++) {
        let rowContent = '';
        let hasCells = false;

        for (let c = 0; c < Math.max(sheet.colCount, 15); c++) {
          const colLetter = String.fromCharCode(65 + c);
          const cellKey = `${colLetter}${r}`;
          const cell = sheet.cells[cellKey];

          if (cell && cell.raw) {
            hasCells = true;
            const isFormula = cell.raw.startsWith('=');
            const evalResult = FormulaEngine.evaluateCell(cellKey, sheet.cells);
            const val = evalResult.value;
            const isNum = typeof val === 'number' && !isNaN(val);

            if (isFormula) {
              const formulaText = escapeXml(cell.raw.substring(1));
              const valText = escapeXml(String(val));
              rowContent += `<c r="${cellKey}" ${isNum ? '' : 't="str"'}><f>${formulaText}</f><v>${valText}</v></c>`;
            } else if (isNum) {
              rowContent += `<c r="${cellKey}"><v>${val}</v></c>`;
            } else {
              rowContent += `<c r="${cellKey}" t="inlineStr"><is><t>${escapeXml(String(val))}</t></is></c>`;
            }
          }
        }

        if (hasCells) {
          sheetXml += `\n    <row r="${r}">${rowContent}</row>`;
        }
      }

      sheetXml += '\n  </sheetData>\n</worksheet>';
      zip.file(`xl/worksheets/sheet${sheetIdx + 1}.xml`, sheetXml);
    });

    return await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  }

  /**
   * 3. Export Microsoft PowerPoint (.pptx) Presentation Archive
   */
  public static async exportPptx(presentation: PresentationData): Promise<Blob> {
    const zip = new JSZip();

    const escapeXml = (unsafe: string) => {
      return unsafe
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
    };

    // [Content_Types].xml
    let contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>`;

    presentation.slides.forEach((_, idx) => {
      contentTypesXml += `\n  <Override PartName="/ppt/slides/slide${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
    });
    contentTypesXml += '\n</Types>';
    zip.file('[Content_Types].xml', contentTypesXml);

    // _rels/.rels
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`
    );

    // ppt/_rels/presentation.xml.rels
    let presRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    presentation.slides.forEach((_, idx) => {
      presRelsXml += `\n  <Relationship Id="rIdSlide${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${idx + 1}.xml"/>`;
    });
    presRelsXml += '\n</Relationships>';
    zip.file('ppt/_rels/presentation.xml.rels', presRelsXml);

    // ppt/presentation.xml (16:9 widescreen = 12192000 x 6858000 EMU)
    let presXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst/>
  <p:sldIdLst>`;
    presentation.slides.forEach((_, idx) => {
      presXml += `\n    <p:sldId id="${256 + idx}" r:id="rIdSlide${idx + 1}"/>`;
    });
    presXml += `\n  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;
    zip.file('ppt/presentation.xml', presXml);

    // ppt/slides/slideN.xml
    presentation.slides.forEach((slide, idx) => {
      const slideWidthEmu = 12192000;
      const slideHeightEmu = 6858000;

      let shapesXml = '';
      let shapeId = 2;

      // Background rect shape
      const bgHex = (slide.bgColor || '#0f172a').replace('#', '').toUpperCase();
      shapesXml += `
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="1" name="Background"/>
            <p:cNvSpPr/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:off x="0" y="0"/>
              <a:ext cx="${slideWidthEmu}" cy="${slideHeightEmu}"/>
            </a:xfrm>
            <a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="${bgHex}"/></a:solidFill>
          </p:spPr>
        </p:sp>`;

      for (const el of slide.elements) {
        shapeId++;
        const xEmu = Math.round((el.x / 100) * slideWidthEmu);
        const yEmu = Math.round((el.y / 100) * slideHeightEmu);
        const wEmu = Math.round((el.width / 100) * slideWidthEmu);
        const hEmu = Math.round(((el.height || 20) / 100) * slideHeightEmu);

        const fontColor = (el.style?.color || '#FFFFFF').replace('#', '').toUpperCase();
        const fontSize = Math.round((el.style?.fontSize || 24) * 100);

        shapesXml += `
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="${shapeId}" name="${escapeXml(el.type)}"/>
            <p:cNvSpPr/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:off x="${xEmu}" y="${yEmu}"/>
              <a:ext cx="${wEmu}" cy="${hEmu}"/>
            </a:xfrm>
            <a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect"><a:avLst/></a:prstGeom>
          </p:spPr>
          <p:txBody>
            <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
            <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:r>
                <a:rPr sz="${fontSize}" b="${el.style?.bold ? '1' : '0'}">
                  <a:solidFill><a:srgbClr val="${fontColor}"/></a:solidFill>
                </a:rPr>
                <a:t>${escapeXml(el.content || '')}</a:t>
              </a:r>
            </a:p>
          </p:txBody>
        </p:sp>`;
      }

      const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      ${shapesXml}
    </p:spTree>
  </p:cSld>
</p:sld>`;

      zip.file(`ppt/slides/slide${idx + 1}.xml`, slideXml);
    });

    return await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });
  }

  /**
   * 4. Export Delimited CSV for active spreadsheet sheet
   */
  public static exportCsv(sheet: SheetData): Blob {
    let csv = '';
    for (let r = 1; r <= sheet.rowCount; r++) {
      const rowVals: string[] = [];
      let hasContent = false;
      for (let c = 0; c < sheet.colCount; c++) {
        const colLetter = String.fromCharCode(65 + c);
        const cellKey = `${colLetter}${r}`;
        const evaluated = FormulaEngine.evaluateCell(cellKey, sheet.cells).value;
        const str = evaluated !== undefined ? String(evaluated) : '';
        if (str.trim().length > 0) hasContent = true;
        rowVals.push(`"${str.replace(/"/g, '""')}"`);
      }
      if (hasContent || r <= 20) {
        csv += rowVals.join(',') + '\n';
      }
    }
    return new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  }

  /**
   * 5. Export Markdown
   */
  public static exportMarkdown(doc: DocumentData): Blob {
    let md = `# ${doc.title}\n\n`;
    for (const p of doc.paragraphs) {
      if (p.type === 'h1') md += `# ${p.content}\n\n`;
      else if (p.type === 'h2') md += `## ${p.content}\n\n`;
      else if (p.type === 'h3') md += `### ${p.content}\n\n`;
      else if (p.type === 'bullet') md += `* ${p.content}\n`;
      else if (p.type === 'callout') md += `> **Note:** ${p.content}\n\n`;
      else if (p.type === 'blockquote') md += `> ${p.content}\n\n`;
      else if (p.type === 'table' && p.tableData && p.tableData.length > 0) {
        md += `\n| ${p.tableData[0]?.join(' | ')} |\n`;
        md += `| ${p.tableData[0]?.map(() => '---').join(' | ')} |\n`;
        p.tableData.slice(1).forEach((r) => {
          md += `| ${r.join(' | ')} |\n`;
        });
        md += '\n';
      } else md += `${p.content}\n\n`;
    }
    return new Blob([md], { type: 'text/markdown;charset=utf-8;' });
  }

  /**
   * 6. Export HTML
   */
  public static exportHtml(project: AnythingProject): Blob {
    let html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${project.name}</title>
<style>
  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; max-width: 860px; margin: 40px auto; padding: 24px; line-height: 1.6; color: #1e293b; background: #fff; }
  h1 { color: #0f172a; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; }
  h2 { color: #1e293b; margin-top: 24px; }
  table { width: 100%; border-collapse: collapse; margin: 24px 0; }
  th, td { border: 1px solid #cbd5e1; padding: 10px 14px; text-align: left; }
  th { background: #f8fafc; font-weight: 600; }
  blockquote { border-left: 4px solid #3b82f6; margin: 16px 0; padding: 8px 16px; background: #f8fafc; color: #475569; }
</style></head><body><h1>${project.docs.title}</h1>`;

    for (const p of project.docs.paragraphs) {
      if (p.type === 'h1') html += `<h1>${p.content}</h1>`;
      else if (p.type === 'h2') html += `<h2>${p.content}</h2>`;
      else if (p.type === 'h3') html += `<h3>${p.content}</h3>`;
      else if (p.type === 'bullet') html += `<li>${p.content}</li>`;
      else if (p.type === 'table' && p.tableData) {
        html += '<table><thead><tr>';
        p.tableData[0]?.forEach((th) => (html += `<th>${th}</th>`));
        html += '</tr></thead><tbody>';
        p.tableData.slice(1).forEach((tr) => {
          html += '<tr>' + tr.map((td) => `<td>${td}</td>`).join('') + '</tr>';
        });
        html += '</tbody></table>';
      } else html += `<p>${p.content}</p>`;
    }
    html += '</body></html>';
    return new Blob([html], { type: 'text/html;charset=utf-8;' });
  }

  /**
   * 7. Export .anything Unified Bundle
   */
  public static exportBundle(project: AnythingProject): Blob {
    const jsonStr = JSON.stringify(project, null, 2);
    return new Blob([jsonStr], { type: 'application/json;charset=utf-8;' });
  }
}
