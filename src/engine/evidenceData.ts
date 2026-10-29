export interface EvidenceSection {
  id: string;
  title: string;
  badge: string;
  summary: string;
  specs: {
    standard: string;
    isoStandard: string;
    containerFormat: string;
    coreSchemaNamespace: string;
  };
  fileStructure: {
    name: string;
    role: string;
    contentPreview: string;
  }[];
  engineDetails: {
    title: string;
    mechanism: string;
    algorithmicComplexity: string;
    description: string;
  }[];
  comparisons: {
    feature: string;
    msOffice: string;
    anythingApp: string;
  }[];
}

export const OFFICE_EVIDENCE: Record<'word' | 'excel' | 'powerpoint' | 'unified', EvidenceSection> = {
  word: {
    id: 'word',
    title: 'Microsoft Word Internal Architecture',
    badge: 'DOCX / ECMA-376 Part 1 §17',
    summary:
      'Microsoft Word operates on a hierarchical tree of document elements serialized inside an Open Packaging Conventions (OPC) ZIP archive. Internally, text editing is driven by a Piece Table, while layout rendering uses an incremental layout pipeline.',
    specs: {
      standard: 'ECMA-376 (Office Open XML) Part 1',
      isoStandard: 'ISO/IEC 29500-1:2016',
      containerFormat: 'ZIP Archive containing Open Packaging Conventions (OPC) Parts',
      coreSchemaNamespace: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    },
    fileStructure: [
      {
        name: '[Content_Types].xml',
        role: 'MIME Registry',
        contentPreview: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
      },
      {
        name: 'word/document.xml',
        role: 'Primary Document Body Tree',
        contentPreview: `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="left"/></w:pPr>
      <w:r>
        <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
        <w:t>Quarterly Strategic Report</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`,
      },
      {
        name: '_rels/.rels',
        role: 'Package Relationship Target Map',
        contentPreview: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`,
      },
    ],
    engineDetails: [
      {
        title: 'Piece Table Memory Structure',
        mechanism: 'Original Buffer + Append Buffer + Indexed Piece Descriptors',
        algorithmicComplexity: 'O(1) insert/delete, O(log N) positional lookup',
        description:
          'Rather than modifying a huge contiguous character array or string upon every keystroke, Word maintains an append-only memory buffer with a piece table linking ranges. This enables instant undo/redo with zero data copying.',
      },
      {
        title: 'Pagination & Flow Layout Pipeline',
        mechanism: 'Break Iterator & Glyphs to Line Boxes to Page Columns',
        algorithmicComplexity: 'O(K) where K is visible page paragraphs',
        description:
          'Documents are flowable (not fixed canvas). Text wraps into line boxes, paragraphs collect line boxes, and section properties (w:sectPr) compute page heights (measured in twips, 1/1440 inch).',
      },
    ],
    comparisons: [
      {
        feature: 'Data Model',
        msOffice: 'w:document -> w:p -> w:r -> w:t hierarchy in XML',
        anythingApp: 'Unified DocumentData model with styled paragraphs, runs, and table blocks',
      },
      {
        feature: 'Undo / Redo',
        msOffice: 'Piece Table transactions with state pointers',
        anythingApp: 'Immutable snapshot history stack with instant state restoration',
      },
      {
        feature: 'Layout Unit',
        msOffice: 'Twips (1/20 pt = 1/1440 inch)',
        anythingApp: 'CSS typographic pixels with print-calibrated A4/Letter margins',
      },
    ],
  },
  excel: {
    id: 'excel',
    title: 'Microsoft Excel Internal Architecture',
    badge: 'XLSX / ECMA-376 Part 1 §18',
    summary:
      'Excel employs a sparse matrix storage engine for 2D cell grids, backed by a Directed Acyclic Graph (DAG) recalculation engine that evaluates formulas via topological sorting of dirty dependency chains.',
    specs: {
      standard: 'ECMA-376 (Office Open XML) Part 1',
      isoStandard: 'ISO/IEC 29500-1:2016',
      containerFormat: 'ZIP Archive containing Open Packaging Conventions (OPC) Parts',
      coreSchemaNamespace: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    },
    fileStructure: [
      {
        name: 'xl/workbook.xml',
        role: 'Workbook Structure & Sheet Registry',
        contentPreview: `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Financials" sheetId="1" r:id="rId1"/>
    <sheet name="Projections" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`,
      },
      {
        name: 'xl/worksheets/sheet1.xml',
        role: 'Sparse Cell Rows & Formulas',
        contentPreview: `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1"><v>125000</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>1</v></c>
      <c r="B2"><f>B1*1.15</f><v>143750</v></c>
    </row>
  </sheetData>
</worksheet>`,
      },
      {
        name: 'xl/sharedStrings.xml',
        role: 'String Deduplication Table',
        contentPreview: `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Revenue</t></si>
  <si><t>Projected Q2</t></si>
</sst>`,
      },
    ],
    engineDetails: [
      {
        title: 'DAG Dependency Graph & Recalculation',
        mechanism: 'Precedent/Dependent trees with topological dirty-propagation',
        algorithmicComplexity: 'O(V + E) for affected subgraph only',
        description:
          'When cell B1 is mutated, Excel does not recalculate the million possible cells. It traverses the dependency tree starting from B1, marks child nodes (B2, B3) as dirty, and evaluates in topological sequence.',
      },
      {
        title: 'Sparse Matrix Cell Indexing',
        mechanism: 'Compressed chunk arrays indexed by (Row, Column)',
        algorithmicComplexity: 'O(1) direct cell coordinate lookup',
        description:
          'Empty cells consume 0 memory. Only instantiated cells appear in the sheet data structure, conserving hundreds of megabytes on sparse data sheets.',
      },
    ],
    comparisons: [
      {
        feature: 'Grid Model',
        msOffice: 'Sparse rows containing only populated `<c r="A1">` nodes',
        anythingApp: 'Key-value map `Record<string, CellData>` with instant O(1) cell lookup',
      },
      {
        feature: 'Calculation Engine',
        msOffice: 'C++ bytecode interpreter with multi-threaded DAG evaluation',
        anythingApp: 'TypeScript FormulaEngine with AST parsing, cycle detection, and range math',
      },
      {
        feature: 'String Pool',
        msOffice: 'SharedStrings table (indexed by integer string ID)',
        anythingApp: 'Direct reactive string references with automatic currency/percent formatting',
      },
    ],
  },
  powerpoint: {
    id: 'powerpoint',
    title: 'Microsoft PowerPoint Internal Architecture',
    badge: 'PPTX / ECMA-376 Part 1 §19',
    summary:
      'PowerPoint is a 2D scene graph engine. Unlike Word’s flowable text, PowerPoint slides are absolute spatial coordinate systems with parent/child shape trees (p:spTree), master layout inheritance, and a timeline-driven presentation runtime.',
    specs: {
      standard: 'ECMA-376 (Office Open XML) Part 1',
      isoStandard: 'ISO/IEC 29500-1:2016',
      containerFormat: 'ZIP Archive containing Open Packaging Conventions (OPC) Parts',
      coreSchemaNamespace: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    },
    fileStructure: [
      {
        name: 'ppt/presentation.xml',
        role: 'Slide Deck Registry & Dimensions',
        contentPreview: `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
</p:presentation>`,
      },
      {
        name: 'ppt/slides/slide1.xml',
        role: 'Slide Scene Graph & Shape Tree',
        contentPreview: `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="1524000" y="1143000"/><a:ext cx="9144000" cy="1828800"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:p><a:r><a:t>Quarterly Vision 2026</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
      },
      {
        name: 'DrawingML Units (EMU)',
        role: 'Coordinate Standard',
        contentPreview: `English Metric Units (EMU):
1 Inch = 914,400 EMUs
1 Point = 12,700 EMUs
1 cm = 360,000 EMUs
Eliminates rounding errors between inches, cm, and pixels across high-DPI displays.`,
      },
    ],
    engineDetails: [
      {
        title: 'Hierarchical 2D Scene Graph (spTree)',
        mechanism: 'Hierarchical transform tree: Group Shapes -> Shapes -> Text Bodies',
        algorithmicComplexity: 'O(N) rendering where N is shapes on active slide',
        description:
          'Every slide owns an independent shape tree (p:spTree) where objects have spatial bounding boxes (off x,y and ext cx,cy), rotation, z-index layering, and fill/stroke shaders.',
      },
      {
        title: 'Slide Master & Layout Inheritance',
        mechanism: 'Prototype-chain resolution of styles, placeholders, and backgrounds',
        algorithmicComplexity: 'O(1) layout template binding',
        description:
          'Slides inherit styling from slideLayouts, which inherit from slideMasters. Placeholders define where titles, subtitles, and content automatically bind.',
      },
    ],
    comparisons: [
      {
        feature: 'Coordinate System',
        msOffice: 'EMUs (English Metric Units: 914,400 per inch)',
        anythingApp: 'Percentage-based spatial geometry (0-100% width/height) for responsive rendering',
      },
      {
        feature: 'Scene Graph',
        msOffice: 'DrawingML `<a:xfrm>` and `<p:sp>` trees',
        anythingApp: 'SlideElement hierarchy with interactive dragging, live typography, and layout presets',
      },
      {
        feature: 'Presenter Engine',
        msOffice: 'DirectX / Metal hardware-accelerated dual-window presenter display',
        anythingApp: 'Full-screen Web Presentation API mode with live timer and slide-deck keyboard navigation',
      },
    ],
  },
  unified: {
    id: 'unified',
    title: 'The "Anything Office" Unified Architecture',
    badge: 'Universal All-in-One Engine',
    summary:
      'Rather than 3 disconnected heavy legacy programs, our Anything Office app unifies Docs, Sheets, and Slides into a single composable runtime sharing a unified command ribbon, state store, clipboard, and instant project serialization.',
    specs: {
      standard: 'Anything Unified Document Model (AUDM)',
      isoStandard: 'Compatible with standard Web & Open Standards (HTML5, CSV, Markdown, JSON)',
      containerFormat: 'Universal Single-File JSON Bundle (`.anything`) + Standard Exports',
      coreSchemaNamespace: 'https://anything.internal/office/v1',
    },
    fileStructure: [
      {
        name: 'Unified Project Bundle (.anything)',
        role: 'Single file containing all Docs, Sheets, and Slides',
        contentPreview: `{
  "id": "proj-2026",
  "name": "Acme Strategic Suite",
  "version": "1.0.0",
  "docs": { "title": "Project Brief", "paragraphs": [...] },
  "sheets": { "sheets": [{ "name": "Budget", "cells": {...} }] },
  "slides": { "slides": [{ "title": "Pitch Deck", "elements": [...] }] }
}`,
      },
      {
        name: 'Command Action Bus',
        role: 'Unified Ribbon & Shortcut Handler',
        contentPreview: `Command Bus:
- Ctrl+B / Format -> Contextually toggles text bold in Docs, cell bold in Sheets, shape text bold in Slides.
- Ctrl+S -> Autosaves all three modules in one unified transaction.
- F5 -> Triggers live presentation mode.`,
      },
    ],
    engineDetails: [
      {
        title: 'Composable Cross-App Bridge',
        mechanism: 'Shared data pipeline between spreadsheet cells, document tables, and presentation slides',
        algorithmicComplexity: 'O(1) in-memory data transfer',
        description:
          'Spreadsheet calculations can be copied directly into rich document tables or presentation metric cards with zero formatting loss.',
      },
      {
        title: 'Zero Latency Multi-App Switching',
        mechanism: 'Single-page unified virtual DOM without process boundaries',
        algorithmicComplexity: 'Zero startup time vs multi-second legacy app launches',
        description:
          'Switching between Word, Excel, and PowerPoint in Microsoft 365 requires launching separate 500MB+ processes. Anything Office switches in 0ms within a unified desktop shell.',
      },
    ],
    comparisons: [
      {
        feature: 'Application Model',
        msOffice: '3 separate 500MB binaries running in isolated OS processes',
        anythingApp: '1 lightweight unified desktop app with instant module switching',
      },
      {
        feature: 'File Management',
        msOffice: 'Disconnected .docx, .xlsx, and .pptx files',
        anythingApp: 'Unified `.anything` multi-suite project + separate standard exports (HTML/CSV/TXT)',
      },
      {
        feature: 'Platform Footprint',
        msOffice: '4 GB+ installation footprint',
        anythingApp: 'Instant zero-dependency client that runs anywhere with high-speed interactivity',
      },
    ],
  },
};
