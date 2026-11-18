import React, { useMemo, useState } from 'react';
import type {
  DocumentData,
  DocParagraph,
  DocComment,
  TrackChange,
  DocCitation,
} from '../types/office';
import {
  AlignCenter,
  Plus,
  Trash2,
  Info,
  ZoomIn,
  ZoomOut,
  BookOpen,
  Layout,
  CheckCircle2,
  Search,
  X,
  List,
  ListOrdered,
  MessageSquare,
  History,
  Stamp,
  Bookmark,
  Users,
  ListTree,
  FileText,
} from 'lucide-react';

interface DocsEditorProps {
  data: DocumentData;
  onChange: (updated: DocumentData) => void;
  darkMode: boolean;
}

export const DocsEditor: React.FC<DocsEditorProps> = ({ data, onChange, darkMode }) => {
  const [zoomLevel, setZoomLevel] = useState<number>(100);
  const [showFindReplace, setShowFindReplace] = useState<boolean>(false);
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [replaceTerm, setReplaceTerm] = useState<string>('');
  const [selectedParagraphId, setSelectedParagraphId] = useState<string | null>(null);

  // Enterprise Feature Drawer States
  const [activeSideDrawer, setActiveSideDrawer] = useState<'comments' | 'trackChanges' | 'citations' | 'watermark' | 'mailMerge' | 'outline' | null>(null);
  const [newCommentText, setNewCommentText] = useState<string>('');

  // Document Outline (Headings Navigation)
  const headings = useMemo(() => {
    return data.paragraphs
      .filter((p) => p.type === 'h1' || p.type === 'h2' || p.type === 'h3')
      .map((p) => ({
        id: p.id,
        level: p.type === 'h1' ? 1 : p.type === 'h2' ? 2 : 3,
        text: p.content,
      }));
  }, [data.paragraphs]);

  const scrollToHeading = (id: string) => {
    setSelectedParagraphId(id);
    const el = document.getElementById(id);
    if (el) {
      el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  };

  const insertTableOfContents = () => {
    const tocId = `p-toc-${Date.now()}`;
    const tocContent =
      'TABLE OF CONTENTS\n\n' +
      headings
        .map(
          (h, i) =>
            `${h.level === 1 ? `${i + 1}. ` : h.level === 2 ? `   ${i + 1}. ` : `      • `}${h.text} .................................... Page ${Math.max(
              1,
              Math.ceil((i + 1) / 3)
            )}`
        )
        .join('\n');

    const tocParagraph: DocParagraph = {
      id: tocId,
      type: 'callout',
      content: tocContent,
      align: 'left',
    };
    const newParagraphs = [tocParagraph, ...data.paragraphs];
    onChange({ ...data, paragraphs: newParagraphs });
    setSelectedParagraphId(tocId);
  };

  // Citations Modal Input
  const [citationAuthor, setCitationAuthor] = useState<string>('Porter, M. E.');
  const [citationTitle, setCitationTitle] = useState<string>('Competitive Strategy: Techniques for Analyzing Industries');
  const [citationYear, setCitationYear] = useState<number>(2024);
  const [citationStyle, setCitationStyle] = useState<'APA' | 'MLA' | 'Chicago'>('APA');

  // Watermark text input
  const [watermarkInput, setWatermarkInput] = useState<string>(data.watermark?.text || 'CONFIDENTIAL');

  // Statistics calculation
  const stats = useMemo(() => {
    let words = 0;
    let characters = 0;

    for (const p of data.paragraphs) {
      if (p.content) {
        const text = p.content.trim();
        characters += text.length;
        if (text.length > 0) {
          words += text.split(/\s+/).length;
        }
      }
      if (p.tableData) {
        for (const row of p.tableData) {
          for (const cell of row) {
            characters += cell.length;
            if (cell.trim().length > 0) words += cell.trim().split(/\s+/).length;
          }
        }
      }
    }

    const readingMinutes = Math.max(1, Math.ceil(words / 200));
    return { words, characters, readingMinutes };
  }, [data]);

  // Find & Replace match count
  const matchCount = useMemo(() => {
    if (!searchTerm.trim()) return 0;
    let count = 0;
    const term = searchTerm.toLowerCase();
    for (const p of data.paragraphs) {
      if (p.content) {
        const matches = p.content.toLowerCase().split(term).length - 1;
        count += matches;
      }
    }
    return count;
  }, [data, searchTerm]);

  const handleReplace = () => {
    if (!searchTerm.trim()) return;
    let replaced = false;
    const term = searchTerm.toLowerCase();
    const newParagraphs = data.paragraphs.map((p) => {
      if (!replaced && p.content && p.content.toLowerCase().includes(term)) {
        replaced = true;
        const re = new RegExp(searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i');
        return { ...p, content: p.content.replace(re, replaceTerm) };
      }
      return p;
    });
    onChange({ ...data, paragraphs: newParagraphs });
  };

  const handleReplaceAll = () => {
    if (!searchTerm.trim()) return;
    const re = new RegExp(searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    const newParagraphs = data.paragraphs.map((p) => {
      if (p.content) {
        return { ...p, content: p.content.replace(re, replaceTerm) };
      }
      return p;
    });
    onChange({ ...data, paragraphs: newParagraphs });
  };

  const updateParagraph = (id: string, updates: Partial<DocParagraph>) => {
    // If track changes is enabled, record revision
    if (data.trackChangesEnabled && updates.content !== undefined) {
      const orig = data.paragraphs.find((p) => p.id === id);
      if (orig && orig.content !== updates.content) {
        const newChange: TrackChange = {
          id: `tc-${Date.now()}`,
          paragraphId: id,
          author: 'Editor (Author)',
          type: 'modify',
          oldText: orig.content,
          newText: updates.content,
          timestamp: Date.now(),
          status: 'pending',
        };
        const currentChanges = data.trackChanges || [];
        const newParagraphs = data.paragraphs.map((p) => (p.id === id ? { ...p, ...updates } : p));
        onChange({ ...data, paragraphs: newParagraphs, trackChanges: [newChange, ...currentChanges] });
        return;
      }
    }

    const newParagraphs = data.paragraphs.map((p) => (p.id === id ? { ...p, ...updates } : p));
    onChange({ ...data, paragraphs: newParagraphs });
  };

  const deleteParagraph = (id: string) => {
    if (data.paragraphs.length <= 1) return;
    const newParagraphs = data.paragraphs.filter((p) => p.id !== id);
    onChange({ ...data, paragraphs: newParagraphs });
  };

  const addParagraphAfter = (
    index: number,
    type: 'p' | 'h1' | 'h2' | 'h3' | 'bullet' | 'number' | 'callout' | 'table' | 'image' | 'hr' = 'p'
  ) => {
    const newId = `p-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`;
    const newP: DocParagraph = {
      id: newId,
      type,
      content:
        type === 'h1'
          ? 'New Heading 1'
          : type === 'h2'
          ? 'New Section Title'
          : type === 'h3'
          ? 'Subsection Heading'
          : type === 'bullet'
          ? 'Bullet list item'
          : type === 'number'
          ? 'Numbered step item'
          : type === 'callout'
          ? 'Important callout note'
          : type === 'table'
          ? 'Data Summary Table'
          : type === 'image'
          ? 'Figure 1: Project Architecture Overview'
          : type === 'hr'
          ? '---'
          : 'Type narrative paragraph text here...',
      align: 'left',
      imageUrl: type === 'image' ? 'https://images.unsplash.com/photo-1557804506-669a67965ba0?w=800&auto=format&fit=crop&q=80' : undefined,
      imageCaption: type === 'image' ? 'Figure 1. Strategic system workflow' : undefined,
      tableData:
        type === 'table'
          ? [
              ['Column 1', 'Column 2', 'Column 3'],
              ['Item A1', 'Value B1', 'Target C1'],
              ['Item A2', 'Value B2', 'Target C2'],
            ]
          : undefined,
    };

    const newParagraphs = [...data.paragraphs];
    newParagraphs.splice(index + 1, 0, newP);
    onChange({ ...data, paragraphs: newParagraphs });
    setSelectedParagraphId(newId);
  };

  const updateTableCell = (pId: string, rowIdx: number, colIdx: number, val: string) => {
    const target = data.paragraphs.find((p) => p.id === pId);
    if (!target || !target.tableData) return;

    const newTable = target.tableData.map((row, rI) =>
      row.map((cell, cI) => (rI === rowIdx && cI === colIdx ? val : cell))
    );
    updateParagraph(pId, { tableData: newTable });
  };

  const addTableRow = (pId: string) => {
    const target = data.paragraphs.find((p) => p.id === pId);
    if (!target || !target.tableData || target.tableData.length === 0) return;

    const colCount = target.tableData[0].length;
    const emptyRow = new Array(colCount).fill('New Value');
    updateParagraph(pId, { tableData: [...target.tableData, emptyRow] });
  };

  const addTableCol = (pId: string) => {
    const target = data.paragraphs.find((p) => p.id === pId);
    if (!target || !target.tableData || target.tableData.length === 0) return;

    const newTable = target.tableData.map((row, idx) => [
      ...row,
      idx === 0 ? `Col ${row.length + 1}` : 'Val',
    ]);
    updateParagraph(pId, { tableData: newTable });
  };

  const deleteTableRow = (pId: string, rowIdx: number) => {
    const target = data.paragraphs.find((p) => p.id === pId);
    if (!target || !target.tableData || target.tableData.length <= 1) return;

    const newTable = target.tableData.filter((_, idx) => idx !== rowIdx);
    updateParagraph(pId, { tableData: newTable });
  };

  // Comments Operations
  const addComment = () => {
    if (!selectedParagraphId || !newCommentText.trim()) return;
    const newComment: DocComment = {
      id: `c-${Date.now()}`,
      paragraphId: selectedParagraphId,
      author: 'Reviewer',
      text: newCommentText.trim(),
      timestamp: Date.now(),
      resolved: false,
    };
    const current = data.comments || [];
    onChange({ ...data, comments: [newComment, ...current] });
    setNewCommentText('');
  };

  const deleteComment = (commentId: string) => {
    const current = data.comments || [];
    onChange({ ...data, comments: current.filter((c) => c.id !== commentId) });
  };

  // Track Changes Operations
  const acceptChange = (changeId: string) => {
    const changes = data.trackChanges || [];
    const change = changes.find((c) => c.id === changeId);
    if (!change) return;

    const newChanges = changes.filter((c) => c.id !== changeId);
    onChange({ ...data, trackChanges: newChanges });
  };

  const rejectChange = (changeId: string) => {
    const changes = data.trackChanges || [];
    const change = changes.find((c) => c.id === changeId);
    if (!change) return;

    // Revert paragraph text
    const newParagraphs = data.paragraphs.map((p) =>
      p.id === change.paragraphId ? { ...p, content: change.oldText || '' } : p
    );
    const newChanges = changes.filter((c) => c.id !== changeId);
    onChange({ ...data, paragraphs: newParagraphs, trackChanges: newChanges });
  };

  // Citations Operations
  const addCitation = () => {
    if (!citationTitle.trim()) return;
    const newCit: DocCitation = {
      id: `cit-${Date.now()}`,
      style: citationStyle,
      author: citationAuthor,
      title: citationTitle,
      year: citationYear,
    };
    const cits = data.citations || [];
    onChange({ ...data, citations: [...cits, newCit] });
  };

  const insertBibliography = () => {
    const cits = data.citations || [];
    if (cits.length === 0) return;

    const bibTitle: DocParagraph = {
      id: `p-${Date.now()}-bibh`,
      type: 'h2',
      content: 'References & Bibliography',
      align: 'left',
    };

    const bibItems: DocParagraph[] = cits.map((c, i) => ({
      id: `p-${Date.now()}-bib${i}`,
      type: 'p',
      content: `${c.author} (${c.year}). ${c.title}. [${c.style} Citation Standard]`,
      align: 'left',
    }));

    onChange({
      ...data,
      paragraphs: [...data.paragraphs, bibTitle, ...bibItems],
    });
  };

  // Mail Merge Sample Records
  const sampleMailMergeRecords = [
    { First_Name: 'Alex', Company: 'Acme Enterprise', Date: 'November 2026', Amount: '$150,000' },
    { First_Name: 'Sarah', Company: 'Starlight Ventures', Date: 'December 2026', Amount: '$280,000' },
    { First_Name: 'Marcus', Company: 'Apex Dynamics', Date: 'January 2027', Amount: '$95,000' },
  ];

  const activeRecord =
    data.mailMerge?.records?.[data.mailMerge?.activeRecordIndex || 0] || sampleMailMergeRecords[0];

  const renderContentWithMailMerge = (content: string) => {
    if (!data.mailMerge?.enabled) return content;
    let res = content;
    for (const [key, val] of Object.entries(activeRecord)) {
      res = res.replace(new RegExp(`{{${key}}}`, 'g'), val);
    }
    return res;
  };

  const isLandscape = data.orientation === 'landscape';

  return (
    <div className="flex flex-col h-full w-full select-text overflow-hidden relative">
      {/* 1. TOP ADVANCED WORD TOOLBAR (Track Changes, Comments, Watermark, Citations, Mail Merge) */}
      <div
        className={`h-7 border-b select-none flex items-center justify-between px-3 text-xs shrink-0 transition-colors ${
          darkMode ? 'bg-slate-900 border-slate-800 text-slate-300' : 'bg-slate-100 border-slate-300 text-slate-700'
        }`}
      >
        <div className="flex items-center space-x-2">
          {/* Orientation Toggle */}
          <button
            onClick={() =>
              onChange({
                ...data,
                orientation: isLandscape ? 'portrait' : 'landscape',
              })
            }
            className="flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-medium hover:bg-slate-200 dark:hover:bg-slate-800"
            title="Toggle Page Orientation (Portrait / Landscape)"
          >
            <Layout className="w-3 h-3 text-blue-500" />
            <span>{isLandscape ? 'Landscape' : 'Portrait'}</span>
          </button>

          {/* Margins Preset Dropdown */}
          <select
            value={
              data.margins.left === 16 ? 'narrow' : data.margins.left === 48 ? 'wide' : 'normal'
            }
            onChange={(e) => {
              const val = e.target.value;
              const margins =
                val === 'narrow'
                  ? { top: 16, right: 16, bottom: 16, left: 16 }
                  : val === 'wide'
                  ? { top: 48, right: 48, bottom: 48, left: 48 }
                  : { top: 28, right: 28, bottom: 28, left: 28 };
              onChange({ ...data, margins });
            }}
            className="h-5 px-1 text-[11px] rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 focus:outline-none"
            title="Page Margins Preset"
          >
            <option value="normal">Normal (1")</option>
            <option value="narrow">Narrow (0.5")</option>
            <option value="wide">Wide (1.75")</option>
          </select>

          <div className="h-4 w-px bg-slate-300 dark:bg-slate-700" />

          {/* Track Changes Toggle */}
          <button
            onClick={() =>
              onChange({
                ...data,
                trackChangesEnabled: !data.trackChangesEnabled,
              })
            }
            className={`flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
              data.trackChangesEnabled
                ? 'bg-amber-600 text-white font-bold'
                : 'hover:bg-slate-200 dark:hover:bg-slate-800'
            }`}
            title="Toggle Track Changes / Redlining mode"
          >
            <History className="w-3 h-3" />
            <span>Track Changes {data.trackChangesEnabled ? 'ON' : 'OFF'}</span>
          </button>

          {/* Comments Drawer Button */}
          <button
            onClick={() => setActiveSideDrawer(activeSideDrawer === 'comments' ? null : 'comments')}
            className={`flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
              activeSideDrawer === 'comments'
                ? 'bg-blue-600 text-white'
                : 'hover:bg-slate-200 dark:hover:bg-slate-800'
            }`}
          >
            <MessageSquare className="w-3 h-3" />
            <span>Comments ({data.comments?.length || 0})</span>
          </button>

          {/* Watermark Drawer Button */}
          <button
            onClick={() => setActiveSideDrawer(activeSideDrawer === 'watermark' ? null : 'watermark')}
            className={`flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
              data.watermark?.enabled ? 'text-blue-600 dark:text-blue-400 font-bold' : ''
            }`}
          >
            <Stamp className="w-3 h-3" />
            <span>Watermark {data.watermark?.enabled ? `(${data.watermark.text})` : ''}</span>
          </button>

          {/* Citations Button */}
          <button
            onClick={() => setActiveSideDrawer(activeSideDrawer === 'citations' ? null : 'citations')}
            className="flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-medium hover:bg-slate-200 dark:hover:bg-slate-800"
          >
            <Bookmark className="w-3 h-3" />
            <span>Citations ({data.citations?.length || 0})</span>
          </button>

          {/* Mail Merge Button */}
          <button
            onClick={() => setActiveSideDrawer(activeSideDrawer === 'mailMerge' ? null : 'mailMerge')}
            className={`flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
              data.mailMerge?.enabled ? 'bg-indigo-600 text-white' : 'hover:bg-slate-200 dark:hover:bg-slate-800'
            }`}
          >
            <Users className="w-3 h-3" />
            <span>Mail Merge</span>
          </button>

          {/* Outline Drawer Button */}
          <button
            onClick={() => setActiveSideDrawer(activeSideDrawer === 'outline' ? null : 'outline')}
            className={`flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
              activeSideDrawer === 'outline'
                ? 'bg-blue-600 text-white'
                : 'hover:bg-slate-200 dark:hover:bg-slate-800'
            }`}
            title="Document Headings Navigation Outline"
          >
            <ListTree className="w-3 h-3" />
            <span>Outline ({headings.length})</span>
          </button>

          {/* Table of Contents Button */}
          <button
            onClick={insertTableOfContents}
            className="flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-medium hover:bg-slate-200 dark:hover:bg-slate-800"
            title="Insert Automated Table of Contents"
          >
            <FileText className="w-3 h-3 text-blue-500" />
            <span>+ TOC</span>
          </button>
        </div>

        {/* Find & Replace Trigger */}
        <button
          onClick={() => setShowFindReplace(!showFindReplace)}
          className={`flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
            showFindReplace ? 'bg-blue-600 text-white' : 'hover:bg-slate-200 dark:hover:bg-slate-800'
          }`}
        >
          <Search className="w-3 h-3" />
          <span>Find & Replace</span>
        </button>
      </div>

      {/* Floating Find & Replace Panel */}
      {showFindReplace && (
        <div className="absolute top-9 right-8 z-40 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 shadow-xl rounded-lg p-3 text-xs w-80 space-y-2 animate-fadeIn">
          <div className="flex items-center justify-between font-bold text-slate-700 dark:text-slate-200">
            <span className="flex items-center space-x-1">
              <Search className="w-3.5 h-3.5 text-blue-500" />
              <span>Find and Replace</span>
            </span>
            <button
              onClick={() => setShowFindReplace(false)}
              className="p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-400"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
          <div className="space-y-1.5">
            <input
              type="text"
              placeholder="Find text..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full px-2 py-1 text-xs rounded border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 focus:outline-none focus:ring-1 focus:ring-blue-500"
            />
            <input
              type="text"
              placeholder="Replace with..."
              value={replaceTerm}
              onChange={(e) => setReplaceTerm(e.target.value)}
              className="w-full px-2 py-1 text-xs rounded border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 focus:outline-none focus:ring-1 focus:ring-blue-500"
            />
          </div>
          <div className="flex items-center justify-between pt-1">
            <span className="text-[10px] text-slate-500 dark:text-slate-400">
              {matchCount} match{matchCount === 1 ? '' : 'es'} found
            </span>
            <div className="flex items-center space-x-1.5">
              <button
                onClick={handleReplace}
                className="px-2 py-0.5 rounded bg-slate-200 dark:bg-slate-800 hover:bg-slate-300 text-slate-700 dark:text-slate-200 text-[11px]"
              >
                Replace
              </button>
              <button
                onClick={handleReplaceAll}
                className="px-2 py-0.5 rounded bg-blue-600 hover:bg-blue-700 text-white text-[11px] font-semibold"
              >
                Replace All
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 2. MAIN CONTAINER WITH WORD CANVAS + OPTIONAL SIDEBAR DRAWER */}
      <div className="flex-1 flex overflow-hidden relative">
        {/* Main Canvas Scroll Area */}
        <div
          className={`flex-1 overflow-y-auto p-8 flex justify-center transition-colors relative ${
            darkMode ? 'bg-[#1e1e1e]' : 'bg-[#f3f2f1]'
          }`}
        >
          {/* Physical Sheet Container */}
          <div
            className={`rounded-xs p-12 transition-all relative overflow-hidden ${
              isLandscape ? 'w-[1100px] min-h-[800px]' : 'w-full max-w-4xl min-h-[1100px]'
            } ${
              darkMode
                ? 'bg-[#2b2b2b] text-slate-100 border border-neutral-700 shadow-2xl'
                : 'bg-white text-slate-900 border border-slate-300 shadow-[0_4px_24px_rgba(0,0,0,0.12)]'
            }`}
            style={{
              transform: `scale(${zoomLevel / 100})`,
              transformOrigin: 'top center',
              paddingTop: `${data.margins.top * 1.5}px`,
              paddingBottom: `${data.margins.bottom * 1.5}px`,
              paddingLeft: `${data.margins.left * 1.5}px`,
              paddingRight: `${data.margins.right * 1.5}px`,
              fontFamily: data.fontFamily || 'Calibri, sans-serif',
            }}
          >
            {/* Visual Background Watermark */}
            {data.watermark?.enabled && (
              <div
                className="absolute inset-0 flex items-center justify-center pointer-events-none select-none z-0"
                style={{ opacity: data.watermark.opacity ?? 0.08 }}
              >
                <span className="text-8xl font-black uppercase tracking-widest -rotate-45 text-slate-500 whitespace-nowrap">
                  {data.watermark.text || 'CONFIDENTIAL'}
                </span>
              </div>
            )}

            {/* Header Bar */}
            <div className="border-b border-slate-200 dark:border-neutral-700 pb-3 mb-8 flex justify-between items-center text-xs text-slate-400 select-none relative z-10">
              <span className="font-semibold tracking-wide uppercase text-[10px]">
                {data.title || 'Executive Document'}
              </span>
              <span className="text-[10px]">
                {data.pageSize || 'A4'} Format • {isLandscape ? 'Landscape' : 'Portrait'} • Normal Margins
              </span>
            </div>

            {/* Paragraph Blocks */}
            <div className="space-y-4 relative z-10">
              {data.paragraphs.map((p, idx) => {
                const paragraphComments = (data.comments || []).filter((c) => c.paragraphId === p.id);
                const paragraphTrackChanges = (data.trackChanges || []).filter(
                  (tc) => tc.paragraphId === p.id && tc.status === 'pending'
                );

                const isPageBreak = idx > 0 && idx % 5 === 0;

                return (
                  <React.Fragment key={p.id}>
                    {isPageBreak && (
                      <div className="my-8 py-2 border-y border-dashed border-slate-300 dark:border-neutral-700 flex items-center justify-between text-[10px] text-slate-400 uppercase tracking-widest select-none bg-slate-100/60 dark:bg-neutral-900/50 -mx-12 px-12">
                        <span>--- Page Break ---</span>
                        <span>Page {Math.floor(idx / 5) + 1} of {Math.max(1, Math.ceil(data.paragraphs.length / 5))}</span>
                      </div>
                    )}

                    <div
                      id={p.id}
                      onClick={() => setSelectedParagraphId(p.id)}
                      className={`group relative -mx-4 px-4 py-1.5 rounded transition-colors ${
                        selectedParagraphId === p.id
                          ? 'ring-1 ring-blue-500/50 bg-blue-50/20 dark:bg-blue-950/20'
                          : 'hover:bg-slate-50 dark:hover:bg-neutral-800/40'
                      }`}
                    >
                    {/* Floating Comment Badge indicator */}
                    {paragraphComments.length > 0 && (
                      <div
                        onClick={() => setActiveSideDrawer('comments')}
                        className="absolute -left-2 top-2 w-5 h-5 rounded-full bg-blue-600 text-white flex items-center justify-center text-[10px] font-bold shadow cursor-pointer z-20"
                        title={`${paragraphComments.length} comment(s)`}
                      >
                        {paragraphComments.length}
                      </div>
                    )}

                    {/* Pending Track Change Banner */}
                    {paragraphTrackChanges.length > 0 && (
                      <div className="mb-1 p-1 bg-amber-50 dark:bg-amber-950/40 border border-amber-300 dark:border-amber-700 rounded text-[10px] text-amber-800 dark:text-amber-200 flex items-center justify-between">
                        <span>Revision pending ({paragraphTrackChanges[0].author})</span>
                        <div className="space-x-1">
                          <button
                            onClick={() => acceptChange(paragraphTrackChanges[0].id)}
                            className="px-1.5 py-0.2 rounded bg-emerald-600 text-white font-bold"
                          >
                            Accept
                          </button>
                          <button
                            onClick={() => rejectChange(paragraphTrackChanges[0].id)}
                            className="px-1.5 py-0.2 rounded bg-rose-600 text-white font-bold"
                          >
                            Reject
                          </button>
                        </div>
                      </div>
                    )}

                    {/* Floating Block Controls on Hover */}
                    <div className="absolute right-2 -top-3 hidden group-hover:flex items-center space-x-1 bg-white dark:bg-neutral-800 border border-slate-300 dark:border-neutral-700 rounded-md shadow-md px-1.5 py-0.5 z-20 select-none text-[11px]">
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setSelectedParagraphId(p.id);
                          setActiveSideDrawer('comments');
                        }}
                        title="Add Comment on this block"
                        className="p-1 hover:text-blue-500 text-slate-500"
                      >
                        <MessageSquare className="w-3.5 h-3.5" />
                      </button>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          addParagraphAfter(idx, 'p');
                        }}
                        title="Add paragraph below"
                        className="p-1 hover:text-blue-500 text-slate-500"
                      >
                        <Plus className="w-3.5 h-3.5" />
                      </button>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          updateParagraph(p.id, {
                            align:
                              p.align === 'center'
                                ? 'right'
                                : p.align === 'right'
                                ? 'justify'
                                : p.align === 'justify'
                                ? 'left'
                                : 'center',
                          });
                        }}
                        title="Cycle alignment"
                        className="p-1 hover:text-blue-500 text-slate-500"
                      >
                        <AlignCenter className="w-3.5 h-3.5" />
                      </button>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          deleteParagraph(p.id);
                        }}
                        title="Delete block"
                        className="p-1 hover:text-rose-500 text-slate-500"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>

                    {/* Render Paragraph by Type */}
                    {p.type === 'h1' && (
                      <input
                        type="text"
                        value={renderContentWithMailMerge(p.content)}
                        onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                        className={`w-full text-3xl font-extrabold tracking-tight bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded px-1 text-[#185abd] dark:text-blue-400 ${
                          p.align === 'center' ? 'text-center' : p.align === 'right' ? 'text-right' : 'text-left'
                        }`}
                      />
                    )}

                    {p.type === 'h2' && (
                      <input
                        type="text"
                        value={renderContentWithMailMerge(p.content)}
                        onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                        className={`w-full text-xl font-bold tracking-tight text-slate-900 dark:text-white bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded px-1 mt-3 ${
                          p.align === 'center' ? 'text-center' : p.align === 'right' ? 'text-right' : 'text-left'
                        }`}
                      />
                    )}

                    {p.type === 'h3' && (
                      <input
                        type="text"
                        value={renderContentWithMailMerge(p.content)}
                        onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                        className={`w-full text-base font-semibold text-slate-800 dark:text-slate-100 bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded px-1 mt-2 ${
                          p.align === 'center' ? 'text-center' : p.align === 'right' ? 'text-right' : 'text-left'
                        }`}
                      />
                    )}

                    {p.type === 'p' && (
                      <textarea
                        rows={Math.max(2, Math.ceil(p.content.length / 85))}
                        value={renderContentWithMailMerge(p.content)}
                        onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                        className={`w-full leading-relaxed text-sm text-slate-900 dark:text-slate-100 bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded p-1 resize-none ${
                          p.align === 'center'
                            ? 'text-center'
                            : p.align === 'right'
                            ? 'text-right'
                            : p.align === 'justify'
                            ? 'text-justify'
                            : 'text-left'
                        }`}
                      />
                    )}

                    {p.type === 'bullet' && (
                      <div className="flex items-start space-x-2.5">
                        <span className="text-[#185abd] dark:text-blue-400 font-bold select-none text-base leading-relaxed">•</span>
                        <textarea
                          rows={Math.max(1, Math.ceil(p.content.length / 80))}
                          value={renderContentWithMailMerge(p.content)}
                          onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                          className="w-full text-sm leading-relaxed text-slate-900 dark:text-slate-100 bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded p-0.5 resize-none"
                        />
                      </div>
                    )}

                    {p.type === 'number' && (
                      <div className="flex items-start space-x-2.5">
                        <span className="text-[#185abd] dark:text-blue-400 font-bold select-none text-sm leading-relaxed">
                          {idx + 1}.
                        </span>
                        <textarea
                          rows={Math.max(1, Math.ceil(p.content.length / 80))}
                          value={renderContentWithMailMerge(p.content)}
                          onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                          className="w-full text-sm leading-relaxed text-slate-900 dark:text-slate-100 bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded p-0.5 resize-none"
                        />
                      </div>
                    )}

                    {p.type === 'callout' && (
                      <div className="p-3.5 rounded-md border-l-4 border-[#185abd] bg-blue-50/70 dark:bg-blue-900/40 text-blue-950 dark:text-blue-50 flex items-start space-x-3">
                        <Info className="w-4 h-4 text-[#185abd] dark:text-blue-300 shrink-0 mt-0.5" />
                        <textarea
                          rows={2}
                          value={renderContentWithMailMerge(p.content)}
                          onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                          className="w-full text-xs font-semibold leading-relaxed text-blue-950 dark:text-blue-50 bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 rounded resize-none"
                        />
                      </div>
                    )}

                    {p.type === 'image' && (
                      <div className="my-3 p-3 rounded-lg border border-slate-200 dark:border-neutral-700 bg-slate-50 dark:bg-neutral-900/50 flex flex-col items-center">
                        {p.imageUrl && (
                          <img
                            src={p.imageUrl}
                            alt="Document illustration"
                            className="max-h-72 object-cover rounded-md shadow-sm mb-2"
                          />
                        )}
                        <input
                          type="text"
                          value={p.content}
                          onChange={(e) => updateParagraph(p.id, { content: e.target.value })}
                          placeholder="Image caption or description..."
                          className="text-xs text-center text-slate-500 dark:text-slate-400 bg-transparent w-full focus:outline-none italic"
                        />
                      </div>
                    )}

                    {p.type === 'hr' && (
                      <div className="my-6 border-b-2 border-slate-200 dark:border-neutral-700" />
                    )}

                    {/* Word Data Table */}
                    {p.type === 'table' && p.tableData && (
                      <div className="my-3 overflow-x-auto rounded border border-slate-300 dark:border-neutral-700">
                        <table className="w-full text-xs border-collapse">
                          <thead>
                            <tr className="bg-slate-100 dark:bg-neutral-800 border-b border-slate-300 dark:border-neutral-700">
                              {p.tableData[0]?.map((colHeader, cIdx) => (
                                <th key={cIdx} className="p-2 text-left font-bold border-r border-inherit last:border-r-0">
                                  <input
                                    type="text"
                                    value={colHeader}
                                    onChange={(e) => updateTableCell(p.id, 0, cIdx, e.target.value)}
                                    className="w-full bg-transparent font-bold focus:outline-none"
                                  />
                                </th>
                              ))}
                              <th className="w-6 p-1 text-center font-normal text-[10px] text-slate-400">✕</th>
                            </tr>
                          </thead>
                          <tbody>
                            {p.tableData.slice(1).map((row, rIdx) => (
                              <tr
                                key={rIdx}
                                className="border-b border-slate-200 dark:border-neutral-800 last:border-b-0 hover:bg-slate-50 dark:hover:bg-neutral-800/40"
                              >
                                {row.map((cell, cIdx) => (
                                  <td
                                    key={cIdx}
                                    className="p-2 border-r border-slate-200 dark:border-neutral-700 last:border-r-0"
                                  >
                                    <input
                                      type="text"
                                      value={cell}
                                      onChange={(e) =>
                                        updateTableCell(p.id, rIdx + 1, cIdx, e.target.value)
                                      }
                                      className="w-full bg-transparent focus:outline-none"
                                    />
                                  </td>
                                ))}
                                <td className="w-6 p-1 text-center">
                                  <button
                                    onClick={() => deleteTableRow(p.id, rIdx + 1)}
                                    className="text-slate-400 hover:text-rose-500 text-xs"
                                    title="Delete row"
                                  >
                                    ×
                                  </button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                        <div className="p-1.5 bg-slate-50 dark:bg-neutral-900 border-t border-inherit flex items-center justify-between">
                          <button
                            onClick={() => deleteParagraph(p.id)}
                            className="text-[10px] text-rose-500 hover:underline px-2 py-0.5"
                          >
                            Delete Table
                          </button>
                          <div className="flex items-center space-x-2">
                            <button
                              onClick={() => addTableCol(p.id)}
                              className="text-[10px] text-[#185abd] dark:text-blue-400 hover:underline flex items-center space-x-1 px-2 py-0.5 font-medium"
                            >
                              <Plus className="w-3 h-3" />
                              <span>Insert Column</span>
                            </button>
                            <button
                              onClick={() => addTableRow(p.id)}
                              className="text-[10px] text-[#185abd] dark:text-blue-400 hover:underline flex items-center space-x-1 px-2 py-0.5 font-medium"
                            >
                              <Plus className="w-3 h-3" />
                              <span>Insert Row</span>
                            </button>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                </React.Fragment>
              );
            })}
          </div>

          {/* Running Document Footer */}
          <div className="border-t border-slate-200 dark:border-neutral-700 pt-3 mt-8 flex justify-between items-center text-[10px] text-slate-400 select-none relative z-10">
            <span>{data.footerText || 'Confidential • Proprietary Document'}</span>
            <span>Page {Math.max(1, Math.ceil(data.paragraphs.length / 5))} of {Math.max(1, Math.ceil(data.paragraphs.length / 5))}</span>
          </div>

            {/* Quick Insert Block at Bottom of Page */}
            <div className="mt-8 pt-4 border-t border-dashed border-slate-300 dark:border-neutral-700 flex flex-wrap gap-2 justify-center select-none">
              <button
                onClick={() => addParagraphAfter(data.paragraphs.length - 1, 'p')}
                className="px-2.5 py-1 text-xs rounded border border-slate-300 dark:border-neutral-700 hover:bg-slate-100 dark:hover:bg-neutral-800 flex items-center space-x-1 text-slate-600 dark:text-slate-300"
              >
                <Plus className="w-3 h-3" />
                <span>Paragraph</span>
              </button>
              <button
                onClick={() => addParagraphAfter(data.paragraphs.length - 1, 'h2')}
                className="px-2.5 py-1 text-xs rounded border border-slate-300 dark:border-neutral-700 hover:bg-slate-100 dark:hover:bg-neutral-800 flex items-center space-x-1 text-slate-600 dark:text-slate-300"
              >
                <Plus className="w-3 h-3" />
                <span>Heading</span>
              </button>
              <button
                onClick={() => addParagraphAfter(data.paragraphs.length - 1, 'bullet')}
                className="px-2.5 py-1 text-xs rounded border border-slate-300 dark:border-neutral-700 hover:bg-slate-100 dark:hover:bg-neutral-800 flex items-center space-x-1 text-slate-600 dark:text-slate-300"
              >
                <List className="w-3 h-3" />
                <span>Bullet List</span>
              </button>
              <button
                onClick={() => addParagraphAfter(data.paragraphs.length - 1, 'number')}
                className="px-2.5 py-1 text-xs rounded border border-slate-300 dark:border-neutral-700 hover:bg-slate-100 dark:hover:bg-neutral-800 flex items-center space-x-1 text-slate-600 dark:text-slate-300"
              >
                <ListOrdered className="w-3 h-3" />
                <span>Numbered List</span>
              </button>
              <button
                onClick={() => addParagraphAfter(data.paragraphs.length - 1, 'table')}
                className="px-2.5 py-1 text-xs rounded border border-slate-300 dark:border-neutral-700 hover:bg-slate-100 dark:hover:bg-neutral-800 flex items-center space-x-1 text-slate-600 dark:text-slate-300"
              >
                <Plus className="w-3 h-3" />
                <span>Table</span>
              </button>
            </div>
          </div>
        </div>

        {/* 3. RIGHT EXPANDABLE ENTERPRISE DRAWER (Comments, Watermark, Citations, Mail Merge) */}
        {activeSideDrawer && (
          <aside
            className={`w-72 border-l p-4 flex flex-col justify-between shrink-0 select-none text-xs z-30 shadow-xl ${
              darkMode ? 'bg-slate-900 border-slate-800 text-slate-100' : 'bg-white border-slate-200 text-slate-900'
            }`}
          >
            {/* Comments Drawer */}
            {activeSideDrawer === 'comments' && (
              <div className="flex-1 flex flex-col">
                <div className="flex items-center justify-between border-b pb-2 mb-3 dark:border-slate-800">
                  <div className="flex items-center space-x-1.5 font-bold">
                    <MessageSquare className="w-4 h-4 text-blue-500" />
                    <span>Document Comments</span>
                  </div>
                  <button onClick={() => setActiveSideDrawer(null)}>
                    <X className="w-4 h-4 text-slate-400" />
                  </button>
                </div>

                <div className="flex-1 overflow-y-auto space-y-2 mb-3">
                  {(data.comments || []).length === 0 ? (
                    <div className="text-center text-slate-400 text-xs py-8">
                      Select a paragraph and type a comment below.
                    </div>
                  ) : (
                    (data.comments || []).map((c) => (
                      <div key={c.id} className="p-2.5 rounded-lg border dark:border-slate-800 bg-slate-50 dark:bg-slate-800/60 text-xs">
                        <div className="flex justify-between items-center text-[10px] text-slate-400 mb-1">
                          <span className="font-semibold text-slate-700 dark:text-slate-300">{c.author}</span>
                          <button onClick={() => deleteComment(c.id)} className="text-slate-400 hover:text-rose-500">
                            <X className="w-3 h-3" />
                          </button>
                        </div>
                        <p className="text-slate-800 dark:text-slate-200">{c.text}</p>
                      </div>
                    ))
                  )}
                </div>

                <div className="border-t pt-2 dark:border-slate-800">
                  <textarea
                    rows={2}
                    value={newCommentText}
                    onChange={(e) => setNewCommentText(e.target.value)}
                    placeholder="Type review comment..."
                    className="w-full p-2 text-xs rounded border dark:bg-slate-800 dark:border-slate-700 focus:outline-none"
                  />
                  <button
                    onClick={addComment}
                    disabled={!selectedParagraphId || !newCommentText.trim()}
                    className="w-full mt-2 py-1.5 rounded bg-blue-600 hover:bg-blue-700 disabled:opacity-40 text-white font-medium text-xs"
                  >
                    Post Comment
                  </button>
                </div>
              </div>
            )}

            {/* Watermark Drawer */}
            {activeSideDrawer === 'watermark' && (
              <div className="flex-1 flex flex-col">
                <div className="flex items-center justify-between border-b pb-2 mb-3 dark:border-slate-800">
                  <div className="flex items-center space-x-1.5 font-bold">
                    <Stamp className="w-4 h-4 text-blue-500" />
                    <span>Watermark Settings</span>
                  </div>
                  <button onClick={() => setActiveSideDrawer(null)}>
                    <X className="w-4 h-4 text-slate-400" />
                  </button>
                </div>

                <div className="space-y-3">
                  <div>
                    <label className="text-[11px] font-semibold text-slate-500">Watermark Text:</label>
                    <input
                      type="text"
                      value={watermarkInput}
                      onChange={(e) => setWatermarkInput(e.target.value)}
                      className="w-full mt-1 p-2 rounded border dark:bg-slate-800 dark:border-slate-700 text-xs uppercase"
                    />
                  </div>

                  <div className="flex flex-wrap gap-1.5">
                    {['CONFIDENTIAL', 'DRAFT', 'URGENT', 'INTERNAL ONLY'].map((preset) => (
                      <button
                        key={preset}
                        onClick={() => setWatermarkInput(preset)}
                        className="px-2 py-1 rounded text-[10px] bg-slate-100 dark:bg-slate-800 hover:bg-slate-200"
                      >
                        {preset}
                      </button>
                    ))}
                  </div>

                  <button
                    onClick={() =>
                      onChange({
                        ...data,
                        watermark: {
                          text: watermarkInput,
                          enabled: !data.watermark?.enabled,
                        },
                      })
                    }
                    className={`w-full py-1.5 rounded font-medium text-xs ${
                      data.watermark?.enabled ? 'bg-rose-600 text-white' : 'bg-blue-600 text-white'
                    }`}
                  >
                    {data.watermark?.enabled ? 'Remove Watermark' : 'Apply Watermark'}
                  </button>
                </div>
              </div>
            )}

            {/* Citations & Bibliography Drawer */}
            {activeSideDrawer === 'citations' && (
              <div className="flex-1 flex flex-col">
                <div className="flex items-center justify-between border-b pb-2 mb-3 dark:border-slate-800">
                  <div className="flex items-center space-x-1.5 font-bold">
                    <Bookmark className="w-4 h-4 text-indigo-500" />
                    <span>Citations & References</span>
                  </div>
                  <button onClick={() => setActiveSideDrawer(null)}>
                    <X className="w-4 h-4 text-slate-400" />
                  </button>
                </div>

                <div className="space-y-2 mb-3 flex-1 overflow-y-auto">
                  <select
                    value={citationStyle}
                    onChange={(e) => setCitationStyle(e.target.value as any)}
                    className="w-full p-1.5 text-xs rounded border dark:bg-slate-800 dark:border-slate-700"
                  >
                    <option value="APA">APA Style (7th Edition)</option>
                    <option value="MLA">MLA Style (9th Edition)</option>
                    <option value="Chicago">Chicago Manual of Style</option>
                  </select>

                  <input
                    type="text"
                    value={citationAuthor}
                    onChange={(e) => setCitationAuthor(e.target.value)}
                    placeholder="Author (e.g. Porter, M. E.)"
                    className="w-full p-1.5 text-xs rounded border dark:bg-slate-800 dark:border-slate-700"
                  />
                  <input
                    type="text"
                    value={citationTitle}
                    onChange={(e) => setCitationTitle(e.target.value)}
                    placeholder="Title of Work..."
                    className="w-full p-1.5 text-xs rounded border dark:bg-slate-800 dark:border-slate-700"
                  />
                  <input
                    type="number"
                    value={citationYear}
                    onChange={(e) => setCitationYear(Number(e.target.value))}
                    className="w-full p-1.5 text-xs rounded border dark:bg-slate-800 dark:border-slate-700"
                  />

                  <button
                    onClick={addCitation}
                    className="w-full py-1.5 rounded bg-indigo-600 hover:bg-indigo-700 text-white font-medium text-xs"
                  >
                    + Add Citation
                  </button>

                  <div className="border-t pt-2 dark:border-slate-800">
                    <div className="text-[10px] font-bold text-slate-400 uppercase mb-1">
                      Saved Citations ({data.citations?.length || 0})
                    </div>
                    {(data.citations || []).map((c) => (
                      <div key={c.id} className="p-1.5 rounded bg-slate-100 dark:bg-slate-800 text-[11px] mb-1">
                        <span className="font-semibold">{c.author}</span> ({c.year}). {c.title}
                      </div>
                    ))}
                  </div>
                </div>

                <button
                  onClick={insertBibliography}
                  disabled={(data.citations || []).length === 0}
                  className="w-full py-1.5 rounded bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 text-white font-medium text-xs"
                >
                  Insert Bibliography Section
                </button>
              </div>
            )}

            {/* Mail Merge Drawer */}
            {activeSideDrawer === 'mailMerge' && (
              <div className="flex-1 flex flex-col">
                <div className="flex items-center justify-between border-b pb-2 mb-3 dark:border-slate-800">
                  <div className="flex items-center space-x-1.5 font-bold">
                    <Users className="w-4 h-4 text-purple-500" />
                    <span>Mail Merge Assistant</span>
                  </div>
                  <button onClick={() => setActiveSideDrawer(null)}>
                    <X className="w-4 h-4 text-slate-400" />
                  </button>
                </div>

                <div className="space-y-3">
                  <div className="text-xs text-slate-500">
                    Use placeholders in your text like <code className="bg-slate-200 dark:bg-slate-800 px-1 rounded font-mono">{'{{First_Name}}'}</code>, <code className="bg-slate-200 dark:bg-slate-800 px-1 rounded font-mono">{'{{Company}}'}</code>, <code className="bg-slate-200 dark:bg-slate-800 px-1 rounded font-mono">{'{{Amount}}'}</code>.
                  </div>

                  <button
                    onClick={() =>
                      onChange({
                        ...data,
                        mailMerge: {
                          enabled: !data.mailMerge?.enabled,
                          variables: ['First_Name', 'Company', 'Date', 'Amount'],
                          records: sampleMailMergeRecords,
                          activeRecordIndex: 0,
                        },
                      })
                    }
                    className={`w-full py-1.5 rounded font-medium text-xs ${
                      data.mailMerge?.enabled ? 'bg-purple-600 text-white' : 'bg-slate-200 dark:bg-slate-800 text-slate-800 dark:text-slate-200'
                    }`}
                  >
                    Preview Mail Merge {data.mailMerge?.enabled ? 'ON' : 'OFF'}
                  </button>

                  {data.mailMerge?.enabled && (
                    <div className="p-3 rounded-lg border dark:border-slate-800 bg-purple-50/50 dark:bg-purple-950/20 space-y-2">
                      <div className="text-[10px] font-bold uppercase text-purple-600 dark:text-purple-400">
                        Active Record ({((data.mailMerge?.activeRecordIndex || 0) + 1)} of {sampleMailMergeRecords.length})
                      </div>
                      <div className="font-mono text-[11px] space-y-0.5">
                        <div>Name: {activeRecord.First_Name}</div>
                        <div>Company: {activeRecord.Company}</div>
                        <div>Amount: {activeRecord.Amount}</div>
                      </div>

                      <div className="flex justify-between pt-1">
                        <button
                          onClick={() =>
                            onChange({
                              ...data,
                              mailMerge: {
                                ...data.mailMerge!,
                                activeRecordIndex: Math.max(0, (data.mailMerge?.activeRecordIndex || 0) - 1),
                              },
                            })
                          }
                          className="px-2 py-0.5 rounded border text-[10px]"
                        >
                          Previous
                        </button>
                        <button
                          onClick={() =>
                            onChange({
                              ...data,
                              mailMerge: {
                                ...data.mailMerge!,
                                activeRecordIndex: Math.min(
                                  sampleMailMergeRecords.length - 1,
                                  (data.mailMerge?.activeRecordIndex || 0) + 1
                                ),
                              },
                            })
                          }
                          className="px-2 py-0.5 rounded border text-[10px]"
                        >
                          Next Record
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* Document Outline / Headings Navigation Drawer */}
            {activeSideDrawer === 'outline' && (
              <div className="flex-1 flex flex-col">
                <div className="flex items-center justify-between border-b pb-2 mb-3 dark:border-slate-800">
                  <div className="flex items-center space-x-1.5 font-bold">
                    <ListTree className="w-4 h-4 text-blue-500" />
                    <span>Document Outline</span>
                  </div>
                  <button onClick={() => setActiveSideDrawer(null)}>
                    <X className="w-4 h-4 text-slate-400" />
                  </button>
                </div>

                <div className="text-xs text-slate-500 mb-2">
                  Headings hierarchy for rapid section navigation:
                </div>

                <div className="flex-1 overflow-y-auto space-y-1 mb-3">
                  {headings.length === 0 ? (
                    <div className="text-center text-slate-400 text-xs py-8">
                      No Headings found. Add Heading 1, 2, or 3 blocks to see outline tree.
                    </div>
                  ) : (
                    headings.map((h, i) => (
                      <div
                        key={h.id}
                        onClick={() => scrollToHeading(h.id)}
                        className={`p-1.5 rounded cursor-pointer transition-colors text-xs flex items-center space-x-1.5 ${
                          h.level === 1
                            ? 'font-bold text-blue-600 dark:text-blue-400 hover:bg-blue-50 dark:hover:bg-blue-950/40'
                            : h.level === 2
                            ? 'font-semibold text-slate-700 dark:text-slate-200 pl-4 hover:bg-slate-100 dark:hover:bg-slate-800'
                            : 'text-slate-500 dark:text-slate-400 pl-7 hover:bg-slate-100 dark:hover:bg-slate-800'
                        }`}
                      >
                        <span className="w-1.5 h-1.5 rounded-full bg-blue-500 shrink-0" />
                        <span className="truncate">{h.text || `Heading ${i + 1}`}</span>
                      </div>
                    ))
                  )}
                </div>

                <button
                  onClick={insertTableOfContents}
                  className="w-full py-1.5 rounded bg-blue-600 hover:bg-blue-700 text-white font-medium text-xs shadow-xs"
                >
                  Insert Table of Contents
                </button>
              </div>
            )}
          </aside>
        )}
      </div>

      {/* 4. AUTHENTIC MICROSOFT WORD STATUS BAR */}
      <footer
        className={`h-6 border-t select-none px-3 flex items-center justify-between text-[11px] shrink-0 transition-colors ${
          darkMode ? 'bg-slate-900 border-slate-800 text-slate-400' : 'bg-[#0078d4] text-white border-blue-700'
        }`}
      >
        <div className="flex items-center space-x-3">
          <span>Page 1 of {Math.max(1, Math.ceil(data.paragraphs.length / 5))}</span>
          <span>•</span>
          <span>{stats.words.toLocaleString()} words</span>
          <span>•</span>
          <span>{stats.characters.toLocaleString()} characters</span>
          <span>•</span>
          <span>~{stats.readingMinutes} min read</span>
          <span className="hidden md:inline flex items-center space-x-1">
            <CheckCircle2 className="w-3 h-3 text-emerald-400 inline mr-1" />
            <span>Accessibility: Good to go</span>
          </span>
        </div>

        <div className="flex items-center space-x-4">
          <div className="flex items-center space-x-1">
            <button title="Read Mode" className="p-0.5 hover:bg-white/10 rounded">
              <BookOpen className="w-3.5 h-3.5" />
            </button>
            <button title="Print Layout" className="p-0.5 bg-white/20 rounded">
              <Layout className="w-3.5 h-3.5" />
            </button>
          </div>

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
    </div>
  );
};
