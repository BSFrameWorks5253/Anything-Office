import React, { useState, useEffect, useRef } from 'react';
import type {
  AppMode,
  AnythingProject,
  CellStyle,
  DocumentData,
  SpreadsheetData,
  PresentationData,
  PdfDocumentData,
  RecentDocumentItem,
} from './types/office';
import { DEFAULT_PROJECT } from './engine/sampleProjects';
import { DocumentImporter } from './engine/documentImporter';
import { DocumentExporter } from './engine/documentExporter';
import { TitleBar } from './components/TitleBar';
import { AppRail } from './components/AppRail';
import { HomeHub } from './components/HomeHub';
import { Ribbon } from './components/Ribbon';
import { DocsEditor } from './components/DocsEditor';
import { SheetsEditor } from './components/SheetsEditor';
import { SlidesEditor } from './components/SlidesEditor';
import { PdfViewer } from './components/PdfViewer';
import { EvidenceModal } from './components/EvidenceModal';
import { TemplateGalleryModal } from './components/TemplateGalleryModal';
import { Check, UploadCloud } from 'lucide-react';

const DEFAULT_RECENTS: RecentDocumentItem[] = [
  {
    id: 'sample-doc-1',
    name: 'Executive Product Strategy 2026.docx',
    type: 'docs',
    lastModified: Date.now() - 1000 * 60 * 24, // 24m ago
    size: '169 words · A4',
    pinned: true,
    source: 'sample',
    description: 'Unified document processing and OOXML architecture',
  },
  {
    id: 'sample-sheet-1',
    name: 'Q3 Global Financial Model & Cash Flow.xlsx',
    type: 'sheets',
    lastModified: Date.now() - 1000 * 60 * 120, // 2h ago
    size: '100 rows × 26 cols',
    pinned: true,
    source: 'sample',
    description: 'Dynamic reactive formula engine with Net Cash calculations',
  },
  {
    id: 'sample-slide-1',
    name: 'Anything Ecosystem Keynote 2026.pptx',
    type: 'slides',
    lastModified: Date.now() - 1000 * 60 * 60 * 5, // 5h ago
    size: '3 slides · 16:9 4K',
    pinned: false,
    source: 'sample',
    description: 'Corporate pitch deck with metric highlight layout',
  },
  {
    id: 'sample-pdf-1',
    name: 'Anything Office Architecture & Technical Whitepaper.pdf',
    type: 'pdf',
    lastModified: Date.now() - 1000 * 60 * 60 * 24, // 1d ago
    size: '3 pages · 2.4 MB',
    pinned: false,
    source: 'sample',
    description: 'Acrobat annotation suite and ECMA-376 specification',
  },
];

export const App: React.FC = () => {
  // Theme state
  const [darkMode, setDarkMode] = useState<boolean>(() => {
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  });

  // Active module mode: 'home' | 'docs' | 'sheets' | 'slides' | 'pdf'
  const [appMode, setAppMode] = useState<AppMode>('home');

  // Presentation fullscreen state
  const [isPresenting, setIsPresenting] = useState<boolean>(false);

  // Drag and drop overlay state
  const [isDragging, setIsDragging] = useState<boolean>(false);

  // AutoSave state
  const [autoSaveEnabled, setAutoSaveEnabled] = useState<boolean>(true);

  // Global hidden file input ref
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Recent Documents list state (persisted to localStorage)
  const [recentItems, setRecentItems] = useState<RecentDocumentItem[]>(() => {
    const saved = localStorage.getItem('anything_office_recents');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        console.error('Failed to parse recents', e);
      }
    }
    return DEFAULT_RECENTS;
  });

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

  // Format byte size to readable string
  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  // Record document into Recents and persist to localStorage
  const recordRecentDocument = (
    name: string,
    type: 'docs' | 'sheets' | 'slides' | 'pdf' | 'anything',
    size: string,
    projectData?: Partial<AnythingProject>
  ) => {
    setRecentItems((prev) => {
      const filtered = prev.filter((item) => item.name !== name);
      const newItem: RecentDocumentItem = {
        id: `rec-${Date.now()}`,
        name,
        type,
        lastModified: Date.now(),
        size,
        pinned: false,
        source: 'local',
        projectData,
      };
      const updated = [newItem, ...filtered].slice(0, 30);
      localStorage.setItem('anything_office_recents', JSON.stringify(updated));
      return updated;
    });
  };

  const handleTogglePinRecent = (id: string) => {
    setRecentItems((prev) => {
      const updated = prev.map((item) =>
        item.id === id ? { ...item, pinned: !item.pinned } : item
      );
      localStorage.setItem('anything_office_recents', JSON.stringify(updated));
      return updated;
    });
  };

  const handleDeleteRecent = (id: string) => {
    setRecentItems((prev) => {
      const updated = prev.filter((item) => item.id !== id);
      localStorage.setItem('anything_office_recents', JSON.stringify(updated));
      return updated;
    });
  };

  const handleSelectRecentItem = (item: RecentDocumentItem) => {
    if (item.projectData) {
      pushState({ ...project, ...item.projectData });
    }
    if (item.type === 'docs' || item.type === 'sheets' || item.type === 'slides' || item.type === 'pdf') {
      setAppMode(item.type);
    } else {
      setAppMode('docs');
    }
    setSaveToast(`Opened "${item.name}"`);
    setTimeout(() => setSaveToast(null), 2500);
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
      } else {
        setAppMode('docs');
      }

      // Add to Recents list
      const detectedType = result.mode === 'all' ? 'docs' : result.mode;
      recordRecentDocument(file.name, detectedType, formatFileSize(file.size));

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
  const handleExport = async (
    format: 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'csv' | 'markdown' | 'html' | 'bundle' | 'print'
  ) => {
    try {
      if (format === 'docx') {
        const blob = await DocumentExporter.exportDocx(project.docs);
        await DocumentExporter.downloadBlob(blob, `${project.docs.title || project.name}.docx`);
        setSaveToast('Exported authentic Microsoft Word document (.docx)');
      } else if (format === 'xlsx') {
        const blob = await DocumentExporter.exportXlsx(project.sheets);
        await DocumentExporter.downloadBlob(blob, `${project.name}.xlsx`);
        setSaveToast('Exported authentic Microsoft Excel workbook (.xlsx)');
      } else if (format === 'pptx') {
        const blob = await DocumentExporter.exportPptx(project.slides);
        await DocumentExporter.downloadBlob(blob, `${project.name}.pptx`);
        setSaveToast('Exported authentic Microsoft PowerPoint presentation (.pptx)');
      } else if (format === 'csv') {
        const activeSheet =
          project.sheets.sheets.find((s) => s.id === project.sheets.activeSheetId) ||
          project.sheets.sheets[0];
        const blob = DocumentExporter.exportCsv(activeSheet);
        await DocumentExporter.downloadBlob(blob, `${activeSheet.name}.csv`);
        setSaveToast('Exported CSV spreadsheet');
      } else if (format === 'markdown') {
        const blob = DocumentExporter.exportMarkdown(project.docs);
        await DocumentExporter.downloadBlob(blob, `${project.docs.title || project.name}.md`);
        setSaveToast('Exported Markdown document');
      } else if (format === 'html') {
        const blob = DocumentExporter.exportHtml(project);
        await DocumentExporter.downloadBlob(blob, `${project.name}.html`);
        setSaveToast('Exported HTML document');
      } else if (format === 'bundle') {
        const blob = DocumentExporter.exportBundle(project);
        await DocumentExporter.downloadBlob(blob, `${project.name}.anything`);
        setSaveToast('Exported Anything Unified Bundle (.anything)');
      } else if (format === 'print' || format === 'pdf') {
        window.print();
      }
      setTimeout(() => setSaveToast(null), 3500);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Export failed';
      setSaveToast(`Export error: ${msg}`);
      setTimeout(() => setSaveToast(null), 4000);
    }
  };

  // Specific Actions from Ribbon
  const handleInsertDocParagraph = (
    type: 'h1' | 'h2' | 'p' | 'bullet' | 'number' | 'callout' | 'table' | 'image' | 'hr'
  ) => {
    const newId = `p-${Date.now()}`;
    const newP = {
      id: newId,
      type,
      content:
        type === 'h1'
          ? 'New Section Heading'
          : type === 'h2'
          ? 'Sub-section Title'
          : type === 'bullet'
          ? 'Strategic item'
          : type === 'number'
          ? 'Numbered process step'
          : type === 'callout'
          ? 'Notice / Highlight Note'
          : type === 'table'
          ? 'Data Summary Table'
          : type === 'image'
          ? 'Figure 1: Project Overview'
          : type === 'hr'
          ? '---'
          : 'Paragraph text...',
      align: 'left' as const,
      imageUrl: type === 'image' ? 'https://images.unsplash.com/photo-1557804506-669a67965ba0?w=800&auto=format&fit=crop&q=80' : undefined,
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

  const handleInsertFormula = (formula: string) => {
    setAppMode('sheets');
    const activeSheet =
      project.sheets.sheets.find((s) => s.id === project.sheets.activeSheetId) ||
      project.sheets.sheets[0];
    const currentCells = { ...activeSheet.cells };
    currentCells['B10'] = { raw: formula };
    pushState({
      ...project,
      sheets: {
        ...project.sheets,
        sheets: project.sheets.sheets.map((s) =>
          s.id === activeSheet.id ? { ...s, cells: currentCells } : s
        ),
      },
    });
    setSaveToast(`Inserted formula: ${formula}`);
    setTimeout(() => setSaveToast(null), 2500);
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

        {/* Main Workspace: Microsoft Fluent Ribbon + Canvas OR Home Hub */}
        <div className="flex-1 flex flex-col overflow-hidden">
          {appMode === 'home' ? (
            <HomeHub
              onOpenAppMode={(mode) => setAppMode(mode)}
              onOpenProject={(p) => pushState(p)}
              onOpenFileClick={() => fileInputRef.current?.click()}
              onOpenEvidence={() => setShowEvidence(true)}
              onOpenTemplates={() => setShowTemplates(true)}
              recentItems={recentItems}
              onTogglePinRecent={handleTogglePinRecent}
              onDeleteRecent={handleDeleteRecent}
              onSelectRecentItem={handleSelectRecentItem}
              darkMode={darkMode}
            />
          ) : (
            <>
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
                onInsertFormula={handleInsertFormula}
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
            </>
          )}
        </div>
      </div>

      {/* Global Hidden File Input for Native Open Dialog */}
      <input
        ref={fileInputRef}
        type="file"
        className="hidden"
        accept=".docx,.xlsx,.pptx,.pdf,.odt,.ods,.odp,.csv,.txt,.json,.anything"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) {
            handleOpenFile(file);
          }
          e.target.value = '';
        }}
      />

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
