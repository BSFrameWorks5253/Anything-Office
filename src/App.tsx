import React, { useState, useEffect } from 'react';
import type {
  AppMode,
  AnythingProject,
  CellStyle,
  DocumentData,
  SpreadsheetData,
  PresentationData,
  PdfDocumentData,
} from './types/office';
import { DEFAULT_PROJECT } from './engine/sampleProjects';
import { FormulaEngine } from './engine/formulaEngine';
import { DocumentImporter } from './engine/documentImporter';
import { TitleBar } from './components/TitleBar';
import { AppRail } from './components/AppRail';
import { Ribbon } from './components/Ribbon';
import { DocsEditor } from './components/DocsEditor';
import { SheetsEditor } from './components/SheetsEditor';
import { SlidesEditor } from './components/SlidesEditor';
import { PdfViewer } from './components/PdfViewer';
import { EvidenceModal } from './components/EvidenceModal';
import { TemplateGalleryModal } from './components/TemplateGalleryModal';
import { Check, UploadCloud } from 'lucide-react';

export const App: React.FC = () => {
  // Theme state
  const [darkMode, setDarkMode] = useState<boolean>(() => {
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  });

  // Active module mode: 'docs' | 'sheets' | 'slides'
  const [appMode, setAppMode] = useState<AppMode>('docs');

  // Presentation fullscreen state
  const [isPresenting, setIsPresenting] = useState<boolean>(false);

  // Drag and drop overlay state
  const [isDragging, setIsDragging] = useState<boolean>(false);

  // AutoSave state
  const [autoSaveEnabled, setAutoSaveEnabled] = useState<boolean>(true);

  // Active Project State (Docs + Sheets + Slides)
  const [project, setProject] = useState<AnythingProject>(() => {
    const saved = localStorage.getItem('anything_office_project');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        console.error('Failed to parse saved project', e);
      }
    }
    return { ...DEFAULT_PROJECT, name: 'Anything Workspace' };
  });

  // Undo / Redo history stack
  const [history, setHistory] = useState<AnythingProject[]>([project]);
  const [historyIndex, setHistoryIndex] = useState<number>(0);

  // Modals
  const [showEvidence, setShowEvidence] = useState<boolean>(false);
  const [showTemplates, setShowTemplates] = useState<boolean>(false);
  const [saveToast, setSaveToast] = useState<string | null>(null);

  // Active text formatting
  const [activeFormat, setActiveFormat] = useState<CellStyle>({
    bold: false,
    italic: false,
    underline: false,
    align: 'left',
  });

  // Update HTML root class for dark mode
  useEffect(() => {
    if (darkMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [darkMode]);

  // Push new state to history
  const pushState = (newProject: AnythingProject) => {
    const newHistory = history.slice(0, historyIndex + 1);
    newHistory.push(newProject);
    setHistory(newHistory);
    setHistoryIndex(newHistory.length - 1);
    setProject(newProject);
  };

  const handleUndo = () => {
    if (historyIndex > 0) {
      const target = history[historyIndex - 1];
      setHistoryIndex(historyIndex - 1);
      setProject(target);
    }
  };

  const handleRedo = () => {
    if (historyIndex < history.length - 1) {
      const target = history[historyIndex + 1];
      setHistoryIndex(historyIndex + 1);
      setProject(target);
    }
  };

  const handleSave = () => {
    localStorage.setItem('anything_office_project', JSON.stringify(project));
    setSaveToast('Saved changes to Anything Cloud');
    setTimeout(() => setSaveToast(null), 2500);
  };

  // Open & Parse Any Microsoft or LibreOffice file
  const handleOpenFile = async (file: File) => {
    try {
      const result = await DocumentImporter.parseFile(file);

      const updatedProject: AnythingProject = {
        ...project,
        name: file.name.replace(/\.[^/.]+$/, ''),
        lastModified: Date.now(),
        ...result.projectUpdates,
      };

      pushState(updatedProject);

      if (result.mode !== 'all') {
        setAppMode(result.mode);
      }

      setSaveToast(`Opened "${file.name}" (${result.detectedFormat})`);
      setTimeout(() => setSaveToast(null), 4000);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Unknown error parsing file';
      setSaveToast(`Error opening file: ${msg}`);
      setTimeout(() => setSaveToast(null), 4500);
    }
  };

  // Drag and drop listeners
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    const file = e.dataTransfer.files?.[0];
    if (file) {
      await handleOpenFile(file);
    }
  };

  // Keyboard shortcut listener
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        handleSave();
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        if (e.shiftKey) {
          handleRedo();
        } else {
          handleUndo();
        }
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
        e.preventDefault();
        handleRedo();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [historyIndex, history, project]);

  // Export handlers
  const handleExport = (format: 'bundle' | 'html' | 'csv' | 'markdown' | 'print') => {
    if (format === 'bundle') {
      const jsonStr = JSON.stringify(project, null, 2);
      const blob = new Blob([jsonStr], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${project.name.toLowerCase().replace(/\s+/g, '-')}.anything`;
      a.click();
      URL.revokeObjectURL(url);
    } else if (format === 'markdown') {
      let md = `# ${project.docs.title}\n\n`;
      for (const p of project.docs.paragraphs) {
        if (p.type === 'h1') md += `# ${p.content}\n\n`;
        else if (p.type === 'h2') md += `## ${p.content}\n\n`;
        else if (p.type === 'h3') md += `### ${p.content}\n\n`;
        else if (p.type === 'bullet') md += `* ${p.content}\n`;
        else if (p.type === 'callout') md += `> **Note:** ${p.content}\n\n`;
        else if (p.type === 'blockquote') md += `> ${p.content}\n\n`;
        else if (p.type === 'table' && p.tableData) {
          md += `\n| ${p.tableData[0]?.join(' | ')} |\n`;
          md += `| ${p.tableData[0]?.map(() => '---').join(' | ')} |\n`;
          p.tableData.slice(1).forEach((r) => {
            md += `| ${r.join(' | ')} |\n`;
          });
          md += '\n';
        } else md += `${p.content}\n\n`;
      }
      const blob = new Blob([md], { type: 'text/markdown' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${project.docs.title.toLowerCase().replace(/\s+/g, '-')}.md`;
      a.click();
      URL.revokeObjectURL(url);
    } else if (format === 'csv') {
      const sheet =
        project.sheets.sheets.find((s) => s.id === project.sheets.activeSheetId) ||
        project.sheets.sheets[0];

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
        if (hasContent || r <= 15) {
          csv += rowVals.join(',') + '\n';
        }
      }
      const blob = new Blob([csv], { type: 'text/csv' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${sheet.name.toLowerCase().replace(/\s+/g, '-')}.csv`;
      a.click();
      URL.revokeObjectURL(url);
    } else if (format === 'html') {
      let html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${project.name}</title><style>body{font-family:system-ui,sans-serif;max-width:800px;margin:40px auto;padding:20px;line-height:1.6;color:#1e293b}h1{color:#0f172a}table{width:100%;border-collapse:collapse;margin:20px 0}th,td{border:1px solid #cbd5e1;padding:8px;text-align:left}th{background:#f1f5f9}</style></head><body><h1>${project.docs.title}</h1>`;
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
      const blob = new Blob([html], { type: 'text/html' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${project.name.toLowerCase().replace(/\s+/g, '-')}.html`;
      a.click();
      URL.revokeObjectURL(url);
    } else if (format === 'print') {
      window.print();
    }
  };

  // Specific Actions from Ribbon
  const handleInsertDocParagraph = (
    type: 'h1' | 'h2' | 'p' | 'bullet' | 'callout' | 'table'
  ) => {
    const newId = `p-${Date.now()}`;
    const newP = {
      id: newId,
      type,
      content:
        type === 'h1'
          ? 'New Section'
          : type === 'h2'
          ? 'Sub-section'
          : type === 'bullet'
          ? 'Strategic item'
          : type === 'callout'
          ? 'Notice / Highlight'
          : type === 'table'
          ? 'Data Table'
          : 'Paragraph text...',
      align: 'left' as const,
      tableData:
        type === 'table'
          ? [
              ['Quarter', 'Revenue', 'Profit'],
              ['Q1', '$100,000', '$25,000'],
              ['Q2', '$140,000', '$38,000'],
            ]
          : undefined,
    };
    pushState({
      ...project,
      docs: {
        ...project.docs,
        paragraphs: [...project.docs.paragraphs, newP],
      },
    });
  };

  const handleAddSheetRow = () => {
    const activeSheet =
      project.sheets.sheets.find((s) => s.id === project.sheets.activeSheetId) ||
      project.sheets.sheets[0];
    const newSheets = project.sheets.sheets.map((s) =>
      s.id === activeSheet.id ? { ...s, rowCount: s.rowCount + 5 } : s
    );
    pushState({
      ...project,
      sheets: { ...project.sheets, sheets: newSheets },
    });
  };

  const handleAddSheetCol = () => {
    const activeSheet =
      project.sheets.sheets.find((s) => s.id === project.sheets.activeSheetId) ||
      project.sheets.sheets[0];
    const newSheets = project.sheets.sheets.map((s) =>
      s.id === activeSheet.id ? { ...s, colCount: s.colCount + 2 } : s
    );
    pushState({
      ...project,
      sheets: { ...project.sheets, sheets: newSheets },
    });
  };

  const handleAddSlide = (
    layout: 'title' | 'title-bullets' | 'two-column' | 'metric-highlight' = 'title-bullets'
  ) => {
    const newSlide = {
      id: `slide-${Date.now()}`,
      title: 'New Slide',
      layout,
      bgColor: '#0f172a',
      elements: [
        {
          id: `el-${Date.now()}-1`,
          type: 'heading' as const,
          content: 'New Presentation Slide',
          x: 10,
          y: 20,
          width: 80,
          style: { fontSize: 36, bold: true, color: '#ffffff', align: 'left' as const },
        },
        {
          id: `el-${Date.now()}-2`,
          type: 'bullet-list' as const,
          content: '• First strategic deliverable\n• Second milestone metric\n• Next steps',
          x: 10,
          y: 45,
          width: 80,
          style: { fontSize: 18, color: '#cbd5e1' },
        },
      ],
    };
    pushState({
      ...project,
      slides: {
        ...project.slides,
        slides: [...project.slides.slides, newSlide],
      },
    });
  };

  const handleDeleteSlide = () => {
    if (project.slides.slides.length <= 1) return;
    const newSlides = project.slides.slides.slice(0, -1);
    pushState({
      ...project,
      slides: { ...project.slides, slides: newSlides },
    });
  };

  return (
    <div
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      className={`flex flex-col h-screen w-screen overflow-hidden transition-colors relative ${
        darkMode ? 'bg-slate-950 text-slate-100' : 'bg-slate-100 text-slate-900'
      }`}
    >
      {/* 1. TOP FULL-WIDTH TITLE BAR (With Custom Window Minimize, Maximize, Close Buttons) */}
      <TitleBar
        appMode={appMode}
        docTitle={project.name}
        onTitleChange={(newTitle) => pushState({ ...project, name: newTitle })}
        autoSaveEnabled={autoSaveEnabled}
        setAutoSaveEnabled={setAutoSaveEnabled}
        onSave={handleSave}
        onUndo={handleUndo}
        onRedo={handleRedo}
        canUndo={historyIndex > 0}
        canRedo={historyIndex < history.length - 1}
        onExport={handleExport}
        darkMode={darkMode}
        setDarkMode={setDarkMode}
      />

      {/* 2. MAIN BODY AREA: AppRail on Left, Ribbon + Stage on Right */}
      <div className="flex-1 flex overflow-hidden">
        {/* Left Microsoft 365 Vertical App Rail */}
        <AppRail
          appMode={appMode}
          setAppMode={setAppMode}
          onOpenTemplates={() => setShowTemplates(true)}
          onOpenEvidence={() => setShowEvidence(true)}
          darkMode={darkMode}
        />

        {/* Main Workspace: Microsoft Fluent Ribbon + Canvas */}
        <div className="flex-1 flex flex-col overflow-hidden">
          <Ribbon
            appMode={appMode}
            onUndo={handleUndo}
            onRedo={handleRedo}
            canUndo={historyIndex > 0}
            canRedo={historyIndex < history.length - 1}
            onSave={handleSave}
            onOpenFile={handleOpenFile}
            onExport={handleExport}
            onOpenTemplates={() => setShowTemplates(true)}
            darkMode={darkMode}
            activeFormat={activeFormat}
            onFormatChange={(fmt) => setActiveFormat({ ...activeFormat, ...fmt })}
            onInsertDocParagraph={handleInsertDocParagraph}
            onAddSheetRow={handleAddSheetRow}
            onAddSheetCol={handleAddSheetCol}
            onAddSlide={handleAddSlide}
            onDeleteSlide={handleDeleteSlide}
            onStartSlideshow={() => setIsPresenting(true)}
          />

          {/* Main Canvas Stage */}
          <div className="flex-1 overflow-hidden relative">
          {appMode === 'docs' && (
            <DocsEditor
              data={project.docs}
              onChange={(updatedDocs: DocumentData) =>
                pushState({ ...project, docs: updatedDocs })
              }
              darkMode={darkMode}
            />
          )}

          {appMode === 'sheets' && (
            <SheetsEditor
              data={project.sheets}
              onChange={(updatedSheets: SpreadsheetData) =>
                pushState({ ...project, sheets: updatedSheets })
              }
              darkMode={darkMode}
              activeFormat={activeFormat}
            />
          )}

          {appMode === 'slides' && (
            <SlidesEditor
              data={project.slides}
              onChange={(updatedSlides: PresentationData) =>
                pushState({ ...project, slides: updatedSlides })
              }
              darkMode={darkMode}
              isPresenting={isPresenting}
              setIsPresenting={setIsPresenting}
            />
          )}

          {appMode === 'pdf' && (
            <PdfViewer
              pdfData={
                project.pdf || {
                  title: 'Anything Office Architecture & Technical Whitepaper.pdf',
                  totalPages: 3,
                  currentPage: 1,
                  zoom: 100,
                  rotation: 0,
                  extractedText:
                    'Anything Office Architecture: Unifying ECMA-376 OOXML processing, DAG formula evaluation, and DrawingML scene graphs into a sub-second desktop workspace.',
                }
              }
              onUpdatePdf={(updatedPdf: Partial<PdfDocumentData>) => {
                const current = project.pdf || {
                  title: 'Document.pdf',
                  totalPages: 1,
                  currentPage: 1,
                  zoom: 100,
                  rotation: 0,
                };
                pushState({ ...project, pdf: { ...current, ...updatedPdf } });
              }}
              onReflowToDocs={(reflowedDoc: DocumentData) => {
                pushState({ ...project, docs: reflowedDoc });
                setAppMode('docs');
                setSaveToast('PDF Reflowed to Editable Word Document!');
                setTimeout(() => setSaveToast(null), 3500);
              }}
              darkMode={darkMode}
            />
          )}
        </div>
      </div>
      </div>

      {/* Drag & Drop Overlay */}
      {isDragging && (
        <div className="absolute inset-0 z-50 bg-blue-950/80 backdrop-blur-md border-4 border-dashed border-blue-400 flex flex-col items-center justify-center text-white select-none pointer-events-none p-8 text-center animate-fadeIn">
          <UploadCloud className="w-16 h-16 text-blue-400 mb-4 animate-bounce" />
          <h2 className="text-2xl font-black mb-2">Drop Office Document to Open in Anything Office</h2>
          <p className="text-sm text-blue-200 max-w-lg mb-6">
            Opens Microsoft Office (.docx, .xlsx, .pptx), PDF (.pdf), LibreOffice (.odt, .ods, .odp), or CSV/Text files.
          </p>

          <div className="flex items-center space-x-2 text-xs font-mono">
            <span className="px-2.5 py-1 rounded bg-blue-600 text-white font-bold">Word / .DOCX</span>
            <span className="px-2.5 py-1 rounded bg-emerald-600 text-white font-bold">Excel / .XLSX</span>
            <span className="px-2.5 py-1 rounded bg-orange-600 text-white font-bold">PowerPoint / .PPTX</span>
            <span className="px-2.5 py-1 rounded bg-red-600 text-white font-bold">PDF / .PDF</span>
            <span className="px-2.5 py-1 rounded bg-purple-600 text-white font-bold">ODF / .ODT</span>
          </div>
        </div>
      )}

      {/* Auto-Save & Notification Toast */}
      {saveToast && (
        <div className="absolute bottom-6 right-6 z-50 bg-slate-900 text-white px-4 py-2.5 rounded-xl shadow-2xl text-xs font-semibold flex items-center space-x-2 animate-bounce border border-slate-700">
          <Check className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{saveToast}</span>
        </div>
      )}

      {/* Evidence Modal Dialog */}
      <EvidenceModal
        isOpen={showEvidence}
        onClose={() => setShowEvidence(false)}
        darkMode={darkMode}
      />

      {/* Template Gallery Modal Dialog */}
      <TemplateGalleryModal
        isOpen={showTemplates}
        onClose={() => setShowTemplates(false)}
        onSelectTemplate={(tpl) => pushState(tpl)}
        darkMode={darkMode}
      />
    </div>
  );
};

export default App;
