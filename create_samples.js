import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';

async function makeSamples() {
  const outDir = path.resolve('./sample_documents');
  const pubDir = path.resolve('./public/samples');
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  if (!fs.existsSync(pubDir)) fs.mkdirSync(pubDir, { recursive: true });

  // 1. DOCX
  const docxZip = new JSZip();
  docxZip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  docxZip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Quarterly Business Analysis Report</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>1. Executive Summary</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>This document was exported as an authentic Microsoft Word (.docx) archive and opened inside Anything Office. The piece table and paragraph flow engines seamlessly render it.</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Product Line</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Q1 Revenue</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Growth</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Cloud Core</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>$450,000</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>+28%</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`
  );
  const docxBuf = await docxZip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(path.join(outDir, 'sample_word.docx'), docxBuf);
  fs.writeFileSync(path.join(pubDir, 'sample_word.docx'), docxBuf);

  // 2. ODT (LibreOffice Writer)
  const odtZip = new JSZip();
  odtZip.file('mimetype', 'application/vnd.oasis.opendocument.text');
  odtZip.file(
    'content.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body>
    <office:text>
      <text:h text:outline-level="1">LibreOffice Writer Test Document</text:h>
      <text:p>This document originates from LibreOffice (.odt format). Anything Office parses ODF content.xml with full heading and table support.</text:p>
      <table:table table:name="Table1">
        <table:table-row>
          <table:table-cell><text:p>Department</text:p></table:table-cell>
          <table:table-cell><text:p>Headcount</text:p></table:table-cell>
        </table:table-row>
        <table:table-row>
          <table:table-cell><text:p>Engineering</text:p></table:table-cell>
          <table:table-cell><text:p>42</text:p></table:table-cell>
        </table:table-row>
      </table:table>
    </office:text>
  </office:body>
</office:document-content>`
  );
  const odtBuf = await odtZip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(path.join(outDir, 'sample_libreoffice.odt'), odtBuf);
  fs.writeFileSync(path.join(pubDir, 'sample_libreoffice.odt'), odtBuf);

  // 3. XLSX (Microsoft Excel)
  const xlsxZip = new JSZip();
  xlsxZip.file(
    'xl/sharedStrings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si><t>Department</t></si>
  <si><t>Budget</t></si>
  <si><t>Actual</t></si>
  <si><t>R&amp;D</t></si>
</sst>`
  );
  xlsxZip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>3</v></c>
      <c r="B2"><v>150000</v></c>
      <c r="C2"><v>138000</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>Variance</v></c>
      <c r="B3"><f>B2-C2</f><v>12000</v></c>
    </row>
  </sheetData>
</worksheet>`
  );
  xlsxZip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets><sheet name="Budget Sheet" sheetId="1"/></sheets>
</workbook>`
  );
  const xlsxBuf = await xlsxZip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(path.join(outDir, 'sample_excel.xlsx'), xlsxBuf);
  fs.writeFileSync(path.join(pubDir, 'sample_excel.xlsx'), xlsxBuf);

  // 4. ODS (LibreOffice Calc)
  const odsZip = new JSZip();
  odsZip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet');
  odsZip.file(
    'content.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Calc Model">
        <table:table-row>
          <table:table-cell><text:p>Region</text:p></table:table-cell>
          <table:table-cell><text:p>Sales</text:p></table:table-cell>
        </table:table-row>
        <table:table-row>
          <table:table-cell><text:p>North America</text:p></table:table-cell>
          <table:table-cell><text:p>85000</text:p></table:table-cell>
        </table:table-row>
        <table:table-row>
          <table:table-cell><text:p>Europe</text:p></table:table-cell>
          <table:table-cell><text:p>62000</text:p></table:table-cell>
        </table:table-row>
        <table:table-row>
          <table:table-cell><text:p>Total</text:p></table:table-cell>
          <table:table-cell table:formula="of:=SUM([.B2:.B3])"><text:p>147000</text:p></table:table-cell>
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>`
  );
  const odsBuf = await odsZip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(path.join(outDir, 'sample_librecalc.ods'), odsBuf);
  fs.writeFileSync(path.join(pubDir, 'sample_librecalc.ods'), odsBuf);

  // 5. PPTX (Microsoft PowerPoint)
  const pptxZip = new JSZip();
  pptxZip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst><p:sldId id="256"/></p:sldIdLst>
</p:presentation>`
  );
  pptxZip.file(
    'ppt/slides/slide1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody><a:p><a:r><a:t>PowerPoint Deck Imported</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:txBody><a:p><a:r><a:t>• Slide shapes parsed into Anything Office scene graph\n• 16:9 widescreen canvas rendering\n• Fullscreen presentation ready</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`
  );
  const pptxBuf = await pptxZip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(path.join(outDir, 'sample_powerpoint.pptx'), pptxBuf);
  fs.writeFileSync(path.join(pubDir, 'sample_powerpoint.pptx'), pptxBuf);

  // 6. ODP (LibreOffice Impress)
  const odpZip = new JSZip();
  odpZip.file('mimetype', 'application/vnd.oasis.opendocument.presentation');
  odpZip.file(
    'content.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1">
        <draw:frame>
          <draw:text-box>
            <text:p>LibreOffice Impress Presentation (.odp)</text:p>
            <text:p>Successfully unpacked and loaded into Anything Office slides canvas!</text:p>
          </draw:text-box>
        </draw:frame>
      </draw:page>
    </office:presentation>
  </office:body>
</office:document-content>`
  );
  const odpBuf = await odpZip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(path.join(outDir, 'sample_libreimpress.odp'), odpBuf);
  fs.writeFileSync(path.join(pubDir, 'sample_libreimpress.odp'), odpBuf);

  // 7. CSV
  fs.writeFileSync(
    path.join(outDir, 'sample_data.csv'),
    `Quarter,Target ARR,Actual Revenue,Variance\nQ1,$1200000,$1340000,$140000\nQ2,$1600000,$1720000,$120000\nQ3,$2100000,$2250000,$150000\n`
  );
  fs.writeFileSync(
    path.join(pubDir, 'sample_data.csv'),
    `Quarter,Target ARR,Actual Revenue,Variance\nQ1,$1200000,$1340000,$140000\nQ2,$1600000,$1720000,$120000\nQ3,$2100000,$2250000,$150000\n`
  );

  console.log('Successfully generated all sample test documents in sample_documents/ and public/samples/');
}

makeSamples().catch(console.error);
