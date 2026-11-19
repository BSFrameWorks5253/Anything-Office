import React, { useMemo, useState, useEffect } from 'react';
import type {
  DocumentData,
  DocParagraph,
  DocComment,
  TrackChange,
  DocCitation,
} from '../types/office';
import {
  Bold,
  Italic,
  Underline,
  Strikethrough,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  Plus,
  Trash2,
  Info,
  ZoomIn,
  ZoomOut,
  Search,
  X,
  MessageSquare,
  History,
  Stamp,
  Bookmark,
  Users,
  ListTree,
  Table as TableIcon,
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
  const [selectedParagraphId, setSelectedParagraphId] = useState<string | null>(data.paragraphs[0]?.id || null);

  // Active formatting state for current selection / paragraph
  const [activeFontFamily, setActiveFontFamily] = useState<string>(data.fontFamily || 'Calibri');
  const [activeFontSize, setActiveFontSize] = useState<number>(11);

  // Mini Floating Toolbar on Text Selection (Microsoft Word style)
  const [miniToolbarPos, setMiniToolbarPos] = useState<{ x: number; y: number } | null>(null);

  // Enterprise Feature Drawer States
  const [activeSideDrawer, setActiveSideDrawer] = useState<'comments' | 'trackChanges' | 'citations' | 'watermark' | 'mailMerge' | 'outline' | null>(null);
  const [newCommentText, setNewCommentText] = useState<string>('');

  // Citations Modal Input
  const [citationAuthor, setCitationAuthor] = useState<string>('Porter, M. E.');
  const [citationTitle, setCitationTitle] = useState<string>('Competitive Strategy: Techniques for Analyzing Industries');
  const [citationYear, setCitationYear] = useState<number>(2024);
  const [citationStyle, setCitationStyle] = useState<'APA' | 'MLA' | 'Chicago'>('APA');

  // Watermark text input
  const [watermarkInput, setWatermarkInput] = useState<string>(data.watermark?.text || 'CONFIDENTIAL');

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

  // Automated Table of Contents insertion
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

  // Real-time Statistics calculation
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

    const estimatedPages = Math.max(1, Math.ceil(data.paragraphs.length / 5));
    const readingMinutes = Math.max(1, Math.ceil(words / 200));
    return { words, characters, estimatedPages, readingMinutes };
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

  // Text selection listener for Mini Floating Toolbar
  useEffect(() => {
    const handleSelectionChange = () => {
      const selection = window.getSelection();
      if (!selection || selection.isCollapsed || selection.toString().trim().length === 0) {
        setMiniToolbarPos(null);
        return;
      }

      const range = selection.getRangeAt(0);
      const rect = range.getBoundingClientRect();
      if (rect.width > 0 && rect.height > 0) {
        setMiniToolbarPos({
          x: rect.left + rect.width / 2,
          y: Math.max(10, rect.top - 44),
        });
      }
    };

    document.addEventListener('selectionchange', handleSelectionChange);
    return () => document.removeEventListener('selectionchange', handleSelectionChange);
  }, []);

  // Update a single paragraph
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

  // Delete a paragraph
  const deleteParagraph = (id: string) => {
    if (data.paragraphs.length <= 1) return;
    const idx = data.paragraphs.findIndex((p) => p.id === id);
    const newParagraphs = data.paragraphs.filter((p) => p.id !== id);
    onChange({ ...data, paragraphs: newParagraphs });
    const nextIdx = Math.max(0, idx - 1);
    setSelectedParagraphId(newParagraphs[nextIdx]?.id || null);
  };

  // Add paragraph after index
  const addParagraphAfter = (
    index: number,
    type: DocParagraph['type'] = 'p',
    initialContent: string = ''
  ) => {
    const newId = `p-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`;
    const newP: DocParagraph = {
      id: newId,
      type,
      content: initialContent || (
        type === 'h1'
          ? 'Heading 1'
          : type === 'h2'
          ? 'Section Title'
          : type === 'h3'
          ? 'Subsection Heading'
          : type === 'bullet'
          ? 'List item'
          : type === 'number'
          ? 'Numbered step'
          : type === 'callout'
          ? 'Key strategic highlight or callout'
          : type === 'table'
          ? 'Summary Table'
          : type === 'image'
          ? 'Figure 1. Operational Overview'
          : type === 'hr'
          ? '---'
          : ''
      ),
      align: 'left',
      imageUrl: type === 'image' ? 'https://images.unsplash.com/photo-1557804506-669a67965ba0?w=800&auto=format&fit=crop&q=80' : undefined,
      imageCaption: type === 'image' ? 'Figure 1: Architectural diagram' : undefined,
      tableData:
        type === 'table'
          ? [
              ['Department', 'Q1 Budget', 'Q2 Forecast', 'Variance'],
              ['Engineering', '$450,000', '$465,000', '+3.3%'],
              ['Marketing', '$280,000', '$275,000', '-1.8%'],
              ['Operations', '$190,000', '$188,000', '-1.1%'],
            ]
          : undefined,
    };

    const newParagraphs = [...data.paragraphs];
    newParagraphs.splice(index + 1, 0, newP);
    onChange({ ...data, paragraphs: newParagraphs });
    setSelectedParagraphId(newId);

    setTimeout(() => {
      const el = document.getElementById(newId);
      if (el) {
        el.focus();
      }
    }, 50);
  };

  // Format active paragraph or selection
  const applyFormattingToSelected = (formatType: 'bold' | 'italic' | 'underline' | 'strike') => {
    if (!selectedParagraphId) return;
    const target = data.paragraphs.find((p) => p.id === selectedParagraphId);
    if (!target) return;

    const currentStyle = target.style || {};
    const updatedStyle = {
      ...currentStyle,
      [formatType]: !currentStyle[formatType],
    };
    updateParagraph(selectedParagraphId, { style: updatedStyle });
  };

  const applyAlignToSelected = (align: DocParagraph['align']) => {
    if (!selectedParagraphId) return;
    updateParagraph(selectedParagraphId, { align });
  };

  const applyTypeToSelected = (type: DocParagraph['type']) => {
    if (!selectedParagraphId) return;
    updateParagraph(selectedParagraphId, { type });
  };

  // Table manipulation methods
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
    const emptyRow = new Array(colCount).fill('Data');
    updateParagraph(pId, { tableData: [...target.tableData, emptyRow] });
  };

  const addTableCol = (pId: string) => {
    const target = data.paragraphs.find((p) => p.id === pId);
    if (!target || !target.tableData || target.tableData.length === 0) return;

    const newTable = target.tableData.map((row, idx) => [
      ...row,
      idx === 0 ? `Col ${row.length + 1}` : 'Value',
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
  const selectedParagraph = data.paragraphs.find((p) => p.id === selectedParagraphId);

  return (
    <div className="flex flex-col h-full w-full select-text overflow-hidden relative font-sans">
      {/* 1. TOP DEDICATED WORD DOCUMENT FORMATTING RIBBON */}
      <div
        className={`h-9 border-b select-none flex items-center justify-between px-3 text-xs shrink-0 transition-colors z-20 ${
          darkMode ? 'bg-slate-900 border-slate-800 text-slate-300' : 'bg-slate-100 border-slate-300 text-slate-700'
        }`}
      >
        <div className="flex items-center space-x-1.5 overflow-x-auto">
          {/* Paragraph Style Selector (Heading 1, 2, 3, Normal, Quote, Bullet, Number) */}
          <select
            value={selectedParagraph?.type || 'p'}
            onChange={(e) => applyTypeToSelected(e.target.value as any)}
            className="h-6 px-2 text-xs rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-medium focus:outline-none"
            title="Style Preset"
          >
            <option value="p">Normal Text</option>
            <option value="h1">Heading 1</option>
            <option value="h2">Heading 2</option>
            <option value="h3">Heading 3</option>
            <option value="bullet">• Bullet List</option>
            <option value="number">1. Numbered List</option>
            <option value="callout">💡 Callout Box</option>
            <option value="table">📊 Table</option>
          </select>

          {/* Font Family Picker */}
          <select
            value={activeFontFamily}
            onChange={(e) => {
              setActiveFontFamily(e.target.value);
              onChange({ ...data, fontFamily: e.target.value });
            }}
            className="h-6 px-2 text-xs rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 focus:outline-none font-sans"
            title="Font Family"
          >
            <option value="Calibri">Calibri</option>
            <option value="Arial">Arial</option>
            <option value="Times New Roman">Times New Roman</option>
            <option value="Segoe UI">Segoe UI</option>
            <option value="Georgia">Georgia</option>
            <option value="Courier New">Courier New</option>
          </select>

          {/* Font Size Picker */}
          <select
            value={activeFontSize}
            onChange={(e) => setActiveFontSize(Number(e.target.value))}
            className="h-6 px-1.5 text-xs rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 focus:outline-none font-mono"
            title="Font Size"
          >
            {[9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48].map((sz) => (
              <option key={sz} value={sz}>
                {sz}
              </option>
            ))}
          </select>

          <div className="h-4 w-px bg-slate-300 dark:bg-slate-700 mx-0.5" />

          {/* Bold, Italic, Underline, Strike */}
          <button
            onClick={() => applyFormattingToSelected('bold')}
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              selectedParagraph?.style?.bold ? 'bg-slate-200 dark:bg-slate-800 text-blue-600 font-bold' : ''
            }`}
            title="Bold (Ctrl+B)"
          >
            <Bold className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => applyFormattingToSelected('italic')}
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              selectedParagraph?.style?.italic ? 'bg-slate-200 dark:bg-slate-800 text-blue-600' : ''
            }`}
            title="Italic (Ctrl+I)"
          >
            <Italic className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => applyFormattingToSelected('underline')}
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              selectedParagraph?.style?.underline ? 'bg-slate-200 dark:bg-slate-800 text-blue-600' : ''
            }`}
            title="Underline (Ctrl+U)"
          >
            <Underline className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => applyFormattingToSelected('strike')}
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              selectedParagraph?.style?.strike ? 'bg-slate-200 dark:bg-slate-800 text-blue-600' : ''
            }`}
            title="Strikethrough"
          >
            <Strikethrough className="w-3.5 h-3.5" />
          </button>

          <div className="h-4 w-px bg-slate-300 dark:bg-slate-700 mx-0.5" />

          {/* Alignment */}
          <button
            onClick={() => applyAlignToSelected('left')}
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              selectedParagraph?.align === 'left' ? 'text-blue-600 font-bold' : ''
            }`}
            title="Align Left"
          >
            <AlignLeft className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => applyAlignToSelected('center')}
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              selectedParagraph?.align === 'center' ? 'text-blue-600 font-bold' : ''
            }`}
            title="Align Center"
          >
            <AlignCenter className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => applyAlignToSelected('right')}
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              selectedParagraph?.align === 'right' ? 'text-blue-600 font-bold' : ''
            }`}
            title="Align Right"
          >
            <AlignRight className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => applyAlignToSelected('justify')}
            className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 ${
              selectedParagraph?.align === 'justify' ? 'text-blue-600 font-bold' : ''
            }`}
            title="Justify"
          >
            <AlignJustify className="w-3.5 h-3.5" />
          </button>

          <div className="h-4 w-px bg-slate-300 dark:bg-slate-700 mx-0.5" />

          {/* Quick Insert Table */}
          <button
            onClick={() => {
              const currentIdx = data.paragraphs.findIndex((p) => p.id === selectedParagraphId);
              addParagraphAfter(currentIdx >= 0 ? currentIdx : data.paragraphs.length - 1, 'table');
            }}
            className="flex items-center space-x-1 px-2 py-0.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 font-medium text-[11px]"
            title="Insert Table"
          >
            <TableIcon className="w-3.5 h-3.5 text-blue-600 dark:text-blue-400" />
            <span>Table</span>
          </button>

          {/* Quick Insert Callout */}
          <button
            onClick={() => {
              const currentIdx = data.paragraphs.findIndex((p) => p.id === selectedParagraphId);
              addParagraphAfter(currentIdx >= 0 ? currentIdx : data.paragraphs.length - 1, 'callout');
            }}
            className="flex items-center space-x-1 px-2 py-0.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 font-medium text-[11px]"
            title="Insert Callout Box"
          >
            <Info className="w-3.5 h-3.5 text-amber-500" />
            <span>Callout</span>
          </button>
        </div>

        {/* Right tools (Orientation, Track Changes, Comments, Find & Replace) */}
        <div className="flex items-center space-x-2 shrink-0">
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
            <span>{isLandscape ? 'Landscape' : 'Portrait'}</span>
          </button>

          {/* Track Changes */}
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
            title="Track Changes / Redlining mode"
          >
            <History className="w-3 h-3" />
            <span>Track Changes</span>
          </button>

          {/* Comments */}
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
              data.watermark?.enabled ? 'text-blue-600 dark:text-blue-400 font-bold' : 'hover:bg-slate-200 dark:hover:bg-slate-800'
            }`}
            title="Configure Background Watermark"
          >
            <Stamp className="w-3 h-3" />
            <span>Watermark</span>
          </button>

          {/* Mail Merge Drawer Button */}
          <button
            onClick={() => setActiveSideDrawer(activeSideDrawer === 'mailMerge' ? null : 'mailMerge')}
            className={`flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
              data.mailMerge?.enabled ? 'bg-indigo-600 text-white font-bold' : 'hover:bg-slate-200 dark:hover:bg-slate-800'
            }`}
            title="Mail Merge Dynamic Fields"
          >
            <Users className="w-3 h-3" />
            <span>Mail Merge</span>
          </button>

          {/* Citations Drawer Button */}
          <button
            onClick={() => setActiveSideDrawer(activeSideDrawer === 'citations' ? null : 'citations')}
            className={`flex items-center space-x-1 px-2 py-0.5 rounded text-[11px] font-medium transition-colors ${
              activeSideDrawer === 'citations' ? 'bg-indigo-600 text-white' : 'hover:bg-slate-200 dark:hover:bg-slate-800'
            }`}
            title="Citations and Bibliography"
          >
            <Bookmark className="w-3 h-3" />
            <span>Citations ({data.citations?.length || 0})</span>
          </button>

          {/* Outline */}
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
            <span>Outline</span>
          </button>

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
      </div>

      {/* Floating Find & Replace Panel */}
      {showFindReplace && (
        <div className="absolute top-11 right-8 z-40 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 shadow-2xl rounded-lg p-3 text-xs w-80 space-y-2 animate-fadeIn">
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

      {/* Floating Mini Toolbar on Text Selection (Microsoft Word style) */}
      {miniToolbarPos && (
        <div
          className="fixed z-50 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 shadow-2xl rounded-md px-2 py-1 flex items-center space-x-1 text-xs animate-fadeIn -translate-x-1/2"
          style={{ left: miniToolbarPos.x, top: miniToolbarPos.y }}
        >
          <button
            onMouseDown={(e) => {
              e.preventDefault();
              applyFormattingToSelected('bold');
            }}
            className="p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 font-bold"
            title="Bold"
          >
            <Bold className="w-3.5 h-3.5" />
          </button>
          <button
            onMouseDown={(e) => {
              e.preventDefault();
              applyFormattingToSelected('italic');
            }}
            className="p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800"
            title="Italic"
          >
            <Italic className="w-3.5 h-3.5" />
          </button>
          <button
            onMouseDown={(e) => {
              e.preventDefault();
              applyFormattingToSelected('underline');
            }}
            className="p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800"
            title="Underline"
          >
            <Underline className="w-3.5 h-3.5" />
          </button>
          <div className="h-4 w-px bg-slate-200 dark:bg-slate-700 mx-0.5" />
          <button
            onMouseDown={(e) => {
              e.preventDefault();
              applyAlignToSelected('left');
            }}
            className="p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800"
          >
            <AlignLeft className="w-3.5 h-3.5" />
          </button>
          <button
            onMouseDown={(e) => {
              e.preventDefault();
              applyAlignToSelected('center');
            }}
            className="p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800"
          >
            <AlignCenter className="w-3.5 h-3.5" />
          </button>
          <button
            onMouseDown={(e) => {
              e.preventDefault();
              applyAlignToSelected('right');
            }}
            className="p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800"
          >
            <AlignRight className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* 2. AUTHENTIC PHYSICAL DOCUMENT WORKSPACE */}
      <div className="flex-1 flex overflow-hidden relative">
        {/* Main Canvas Scroll Area */}
        <div
          className={`flex-1 overflow-y-auto p-8 flex flex-col items-center transition-colors relative ${
            darkMode ? 'bg-[#18181b]' : 'bg-[#eef2f6]'
          }`}
        >
          {/* Authentic Physical Margin Ruler */}
          <div
            className={`mb-3 h-5 rounded-xs border flex items-center justify-between px-6 text-[9px] font-mono select-none ${
              isLandscape ? 'w-[1050px]' : 'w-full max-w-4xl'
            } ${
              darkMode ? 'bg-slate-900 border-slate-800 text-slate-500' : 'bg-white border-slate-300 text-slate-400'
            }`}
          >
            <div className="flex items-center space-x-8">
              <span>| 1"</span>
              <span>| 2"</span>
              <span>| 3"</span>
              <span>| 4"</span>
              <span>| 5"</span>
              <span>| 6"</span>
              <span>| 7"</span>
              {isLandscape && (
                <>
                  <span>| 8"</span>
                  <span>| 9"</span>
                  <span>| 10"</span>
                </>
              )}
            </div>
            <span className="font-semibold">{data.pageSize || 'A4'} Page Grid</span>
          </div>

          {/* Physical Sheet Container (Exact Microsoft Word Page) */}
          <div
            className={`rounded-xs p-16 transition-all relative overflow-hidden ${
              isLandscape ? 'w-[1050px] min-h-[750px]' : 'w-full max-w-4xl min-h-[1123px]'
            } ${
              darkMode
                ? 'bg-[#27272a] text-slate-100 border border-neutral-700 shadow-2xl'
                : 'bg-white text-slate-900 border border-slate-300 shadow-[0_8px_30px_rgb(0,0,0,0.12)]'
            }`}
            style={{
              transform: `scale(${zoomLevel / 100})`,
              transformOrigin: 'top center',
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

            {/* Document Header Zone */}
            <div className="border-b border-slate-200 dark:border-neutral-700 pb-2 mb-8 flex justify-between items-center text-xs text-slate-400 select-none relative z-10">
              <span className="font-bold tracking-wide uppercase text-[10px]">
                {data.title || 'Untitled Document'}
              </span>
              <span className="text-[10px]">
                {data.pageSize || 'A4'} • {isLandscape ? 'Landscape' : 'Portrait'} • Microsoft Word Parity
              </span>
            </div>

            {/* Paragraph Blocks Flow */}
            <div className="space-y-3 relative z-10">
              {data.paragraphs.map((p, idx) => {
                const paragraphComments = (data.comments || []).filter((c) => c.paragraphId === p.id);
                const paragraphTrackChanges = (data.trackChanges || []).filter(
                  (tc) => tc.paragraphId === p.id && tc.status === 'pending'
                );

                const isPageBreak = idx > 0 && idx % 6 === 0;

                return (
                  <React.Fragment key={p.id}>
                    {isPageBreak && (
                      <div className="my-10 py-2 border-y border-dashed border-slate-300 dark:border-neutral-700 flex items-center justify-between text-[10px] text-slate-400 uppercase tracking-widest select-none bg-slate-100/60 dark:bg-neutral-900/50 -mx-16 px-16">
                        <span>--- Page Break ---</span>
                        <span>Page {Math.floor(idx / 6) + 1} of {stats.estimatedPages}</span>
                      </div>
                    )}

                    <div
                      id={p.id}
                      onClick={() => setSelectedParagraphId(p.id)}
                      className={`group relative -mx-4 px-4 py-1.5 rounded transition-colors ${
                        selectedParagraphId === p.id
                          ? 'ring-1 ring-blue-500/50 bg-blue-50/15 dark:bg-blue-950/20'
                          : 'hover:bg-slate-50/80 dark:hover:bg-neutral-800/30'
                      }`}
                    >
                      {/* Floating Comment Badge indicator */}
                      {paragraphComments.length > 0 && (
                        <div
                          onClick={() => setActiveSideDrawer('comments')}
                          className="absolute -left-3 top-2 w-5 h-5 rounded-full bg-blue-600 text-white flex items-center justify-center text-[10px] font-bold shadow cursor-pointer z-20"
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
                          title="Add Comment"
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
                            deleteParagraph(p.id);
                          }}
                          title="Delete paragraph"
                          className="p-1 hover:text-rose-500 text-slate-500"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>

                      {/* Content-Editable Fluid Paragraph Rendering */}
                      {p.type === 'h1' && (
                        <div
                          contentEditable
                          suppressContentEditableWarning
                          onBlur={(e) => updateParagraph(p.id, { content: e.currentTarget.textContent || '' })}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                              addParagraphAfter(idx, 'p');
                            }
                          }}
                          className={`text-3xl font-extrabold tracking-tight text-[#185abd] dark:text-blue-400 focus:outline-none focus:ring-1 focus:ring-blue-500 rounded px-1 ${
                            p.align === 'center' ? 'text-center' : p.align === 'right' ? 'text-right' : 'text-left'
                          }`}
                          style={{
                            fontWeight: p.style?.bold ? 'bold' : undefined,
                            fontStyle: p.style?.italic ? 'italic' : undefined,
                            textDecoration: p.style?.underline ? 'underline' : undefined,
                          }}
                        >
                          {renderContentWithMailMerge(p.content)}
                        </div>
                      )}

                      {p.type === 'h2' && (
                        <div
                          contentEditable
                          suppressContentEditableWarning
                          onBlur={(e) => updateParagraph(p.id, { content: e.currentTarget.textContent || '' })}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                              addParagraphAfter(idx, 'p');
                            }
                          }}
                          className={`text-xl font-bold tracking-tight text-slate-900 dark:text-white focus:outline-none focus:ring-1 focus:ring-blue-500 rounded px-1 mt-3 ${
                            p.align === 'center' ? 'text-center' : p.align === 'right' ? 'text-right' : 'text-left'
                          }`}
                          style={{
                            fontWeight: p.style?.bold ? 'bold' : undefined,
                            fontStyle: p.style?.italic ? 'italic' : undefined,
                            textDecoration: p.style?.underline ? 'underline' : undefined,
                          }}
                        >
                          {renderContentWithMailMerge(p.content)}
                        </div>
                      )}

                      {p.type === 'h3' && (
                        <div
                          contentEditable
                          suppressContentEditableWarning
                          onBlur={(e) => updateParagraph(p.id, { content: e.currentTarget.textContent || '' })}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                              addParagraphAfter(idx, 'p');
                            }
                          }}
                          className={`text-base font-semibold text-slate-800 dark:text-slate-100 focus:outline-none focus:ring-1 focus:ring-blue-500 rounded px-1 mt-2 ${
                            p.align === 'center' ? 'text-center' : p.align === 'right' ? 'text-right' : 'text-left'
                          }`}
                        >
                          {renderContentWithMailMerge(p.content)}
                        </div>
                      )}

                      {p.type === 'p' && (
                        <div
                          contentEditable
                          suppressContentEditableWarning
                          onBlur={(e) => updateParagraph(p.id, { content: e.currentTarget.textContent || '' })}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter' && !e.shiftKey) {
                              e.preventDefault();
                              addParagraphAfter(idx, 'p');
                            } else if (e.key === 'Backspace' && e.currentTarget.textContent === '') {
                              e.preventDefault();
                              deleteParagraph(p.id);
                            }
                          }}
                          className={`text-sm leading-relaxed text-slate-800 dark:text-slate-200 focus:outline-none focus:ring-1 focus:ring-blue-500 rounded px-1 min-h-[1.5em] ${
                            p.align === 'center'
                              ? 'text-center'
                              : p.align === 'right'
                              ? 'text-right'
                              : p.align === 'justify'
                              ? 'text-justify'
                              : 'text-left'
                          }`}
                          style={{
                            fontWeight: p.style?.bold ? 'bold' : undefined,
                            fontStyle: p.style?.italic ? 'italic' : undefined,
                            textDecoration: p.style?.underline ? 'underline' : undefined,
                          }}
                        >
                          {renderContentWithMailMerge(p.content)}
                        </div>
                      )}

                      {p.type === 'bullet' && (
                        <div className="flex items-start space-x-2.5">
                          <span className="text-[#185abd] dark:text-blue-400 font-bold select-none text-base leading-relaxed">•</span>
                          <div
                            contentEditable
                            suppressContentEditableWarning
                            onBlur={(e) => updateParagraph(p.id, { content: e.currentTarget.textContent || '' })}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                e.preventDefault();
                                addParagraphAfter(idx, 'bullet');
                              } else if (e.key === 'Backspace' && e.currentTarget.textContent === '') {
                                e.preventDefault();
                                deleteParagraph(p.id);
                              }
                            }}
                            className="w-full text-sm leading-relaxed text-slate-800 dark:text-slate-200 focus:outline-none focus:ring-1 focus:ring-blue-500 rounded px-1"
                          >
                            {renderContentWithMailMerge(p.content)}
                          </div>
                        </div>
                      )}

                      {p.type === 'number' && (
                        <div className="flex items-start space-x-2.5">
                          <span className="text-[#185abd] dark:text-blue-400 font-bold select-none text-sm leading-relaxed">
                            {idx + 1}.
                          </span>
                          <div
                            contentEditable
                            suppressContentEditableWarning
                            onBlur={(e) => updateParagraph(p.id, { content: e.currentTarget.textContent || '' })}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                e.preventDefault();
                                addParagraphAfter(idx, 'number');
                              } else if (e.key === 'Backspace' && e.currentTarget.textContent === '') {
                                e.preventDefault();
                                deleteParagraph(p.id);
                              }
                            }}
                            className="w-full text-sm leading-relaxed text-slate-800 dark:text-slate-200 focus:outline-none focus:ring-1 focus:ring-blue-500 rounded px-1"
                          >
                            {renderContentWithMailMerge(p.content)}
                          </div>
                        </div>
                      )}

                      {p.type === 'callout' && (
                        <div className="p-4 rounded-md border-l-4 border-[#185abd] bg-blue-50/70 dark:bg-blue-900/40 text-blue-950 dark:text-blue-50 flex items-start space-x-3">
                          <Info className="w-4 h-4 text-[#185abd] dark:text-blue-300 shrink-0 mt-0.5" />
                          <div
                            contentEditable
                            suppressContentEditableWarning
                            onBlur={(e) => updateParagraph(p.id, { content: e.currentTarget.textContent || '' })}
                            className="w-full text-xs font-semibold leading-relaxed focus:outline-none"
                          >
                            {renderContentWithMailMerge(p.content)}
                          </div>
                        </div>
                      )}

                      {/* Word Data Table */}
                      {p.type === 'table' && p.tableData && (
                        <div className="my-4 overflow-x-auto rounded-sm border border-slate-300 dark:border-neutral-700">
                          <div className="bg-slate-100 dark:bg-neutral-800 px-3 py-1 border-b border-inherit flex items-center justify-between text-[11px] text-slate-600 dark:text-slate-300 font-medium select-none">
                            <span>Table: {p.content}</span>
                            <div className="flex items-center space-x-2">
                              <button
                                onClick={() => addTableRow(p.id)}
                                className="px-1.5 py-0.5 rounded bg-white dark:bg-neutral-700 hover:bg-slate-200 border border-slate-300 dark:border-neutral-600 text-[10px]"
                              >
                                + Row
                              </button>
                              <button
                                onClick={() => addTableCol(p.id)}
                                className="px-1.5 py-0.5 rounded bg-white dark:bg-neutral-700 hover:bg-slate-200 border border-slate-300 dark:border-neutral-600 text-[10px]"
                              >
                                + Column
                              </button>
                            </div>
                          </div>
                          <table className="w-full text-xs border-collapse">
                            <thead>
                              <tr className="bg-slate-50 dark:bg-neutral-800 border-b border-slate-300 dark:border-neutral-700">
                                {p.tableData[0]?.map((colHeader, cIdx) => (
                                  <th key={cIdx} className="p-2.5 text-left font-bold border-r border-inherit last:border-r-0">
                                    <input
                                      type="text"
                                      value={colHeader}
                                      onChange={(e) => updateTableCell(p.id, 0, cIdx, e.target.value)}
                                      className="w-full bg-transparent font-bold focus:outline-none"
                                    />
                                  </th>
                                ))}
                                <th className="w-8 p-1 text-center font-normal text-slate-400">✕</th>
                              </tr>
                            </thead>
                            <tbody>
                              {p.tableData.slice(1).map((row, rIdx) => (
                                <tr
                                  key={rIdx}
                                  className="border-b border-slate-200 dark:border-neutral-800 last:border-b-0 hover:bg-slate-50/60 dark:hover:bg-neutral-800/40"
                                >
                                  {row.map((cell, cIdx) => (
                                    <td
                                      key={cIdx}
                                      className="p-2 border-r border-slate-200 dark:border-neutral-700 last:border-r-0"
                                    >
                                      <input
                                        type="text"
                                        value={cell}
                                        onChange={(e) => updateTableCell(p.id, rIdx + 1, cIdx, e.target.value)}
                                        className="w-full bg-transparent focus:outline-none"
                                      />
                                    </td>
                                  ))}
                                  <td className="w-8 p-1 text-center">
                                    <button
                                      onClick={() => deleteTableRow(p.id, rIdx + 1)}
                                      className="text-slate-400 hover:text-rose-500 font-bold"
                                      title="Delete Row"
                                    >
                                      ✕
                                    </button>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      )}

                      {p.type === 'image' && (
                        <div className="my-4 p-3 rounded-lg border border-slate-200 dark:border-neutral-700 bg-slate-50 dark:bg-neutral-900/50 flex flex-col items-center">
                          {p.imageUrl && (
                            <img
                              src={p.imageUrl}
                              alt="Document visual"
                              className="max-h-80 object-cover rounded shadow-sm mb-2"
                            />
                          )}
                          <div
                            contentEditable
                            suppressContentEditableWarning
                            onBlur={(e) => updateParagraph(p.id, { content: e.currentTarget.textContent || '' })}
                            className="text-xs text-center text-slate-500 italic focus:outline-none w-full"
                          >
                            {p.content}
                          </div>
                        </div>
                      )}

                      {p.type === 'hr' && (
                        <div className="my-6 border-b-2 border-slate-200 dark:border-neutral-700" />
                      )}
                    </div>
                  </React.Fragment>
                );
              })}
            </div>

            {/* Document Footer Zone */}
            <div className="border-t border-slate-200 dark:border-neutral-700 pt-3 mt-12 flex justify-between items-center text-[10px] text-slate-400 select-none relative z-10">
              <span>Confidential & Proprietary • Anything Office</span>
              <span className="font-mono">Page 1 of {stats.estimatedPages}</span>
            </div>
          </div>
        </div>

        {/* 3. ENTERPRISE SIDEBAR DRAWER (Comments, Outline, Citations, Track Changes) */}
        {activeSideDrawer && (
          <div className="w-80 border-l border-slate-300 dark:border-slate-800 bg-white dark:bg-slate-900 flex flex-col shrink-0 shadow-2xl z-30 animate-fadeIn">
            <div className="p-3 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between font-bold text-xs">
              <span className="flex items-center space-x-1.5 capitalize">
                {activeSideDrawer === 'outline' && <ListTree className="w-4 h-4 text-blue-500" />}
                {activeSideDrawer === 'comments' && <MessageSquare className="w-4 h-4 text-blue-500" />}
                {activeSideDrawer === 'trackChanges' && <History className="w-4 h-4 text-amber-500" />}
                {activeSideDrawer === 'citations' && <Bookmark className="w-4 h-4 text-indigo-500" />}
                <span>{activeSideDrawer} Panel</span>
              </span>
              <button
                onClick={() => setActiveSideDrawer(null)}
                className="p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-400"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="flex-1 p-3 overflow-y-auto text-xs space-y-3">
              {/* Outline Tab */}
              {activeSideDrawer === 'outline' && (
                <div className="space-y-2">
                  <div className="flex items-center justify-between pb-1 border-b border-slate-200 dark:border-slate-800">
                    <span className="text-[10px] font-bold uppercase text-slate-400">Headings ({headings.length})</span>
                    <button
                      onClick={insertTableOfContents}
                      className="text-blue-600 dark:text-blue-400 hover:underline font-semibold text-[11px]"
                    >
                      + Insert TOC
                    </button>
                  </div>
                  {headings.map((h, i) => (
                    <div
                      key={h.id}
                      onClick={() => scrollToHeading(h.id)}
                      className={`cursor-pointer p-1.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-300 truncate ${
                        h.level === 1 ? 'font-bold pl-2 text-blue-600 dark:text-blue-400' : h.level === 2 ? 'pl-4 font-medium' : 'pl-6 text-slate-500'
                      }`}
                    >
                      {h.text || `Untitled Heading ${i + 1}`}
                    </div>
                  ))}
                </div>
              )}

              {/* Comments Tab */}
              {activeSideDrawer === 'comments' && (
                <div className="space-y-3">
                  <div className="space-y-1.5">
                    <textarea
                      rows={2}
                      placeholder="Type comment on selected paragraph..."
                      value={newCommentText}
                      onChange={(e) => setNewCommentText(e.target.value)}
                      className="w-full p-2 text-xs rounded border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 focus:outline-none focus:ring-1 focus:ring-blue-500"
                    />
                    <button
                      onClick={addComment}
                      className="w-full py-1 rounded bg-blue-600 hover:bg-blue-700 text-white font-semibold text-xs"
                    >
                      Post Comment
                    </button>
                  </div>

                  <div className="space-y-2 pt-2 border-t border-slate-200 dark:border-slate-800">
                    {(data.comments || []).map((c) => (
                      <div key={c.id} className="p-2.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/60">
                        <div className="flex items-center justify-between text-[10px] text-slate-400 mb-1">
                          <span className="font-bold text-slate-700 dark:text-slate-200">{c.author}</span>
                          <button
                            onClick={() => deleteComment(c.id)}
                            className="text-slate-400 hover:text-rose-500"
                          >
                            <Trash2 className="w-3 h-3" />
                          </button>
                        </div>
                        <p className="text-slate-800 dark:text-slate-200 text-xs">{c.text}</p>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Citations Drawer */}
              {activeSideDrawer === 'citations' && (
                <div className="space-y-3">
                  <div className="space-y-2">
                    <input
                      type="text"
                      placeholder="Author(s)..."
                      value={citationAuthor}
                      onChange={(e) => setCitationAuthor(e.target.value)}
                      className="w-full p-1.5 text-xs rounded border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800"
                    />
                    <input
                      type="text"
                      placeholder="Work Title..."
                      value={citationTitle}
                      onChange={(e) => setCitationTitle(e.target.value)}
                      className="w-full p-1.5 text-xs rounded border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800"
                    />
                    <div className="flex items-center space-x-2">
                      <input
                        type="number"
                        placeholder="Year"
                        value={citationYear}
                        onChange={(e) => setCitationYear(Number(e.target.value))}
                        className="w-24 p-1.5 text-xs rounded border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 font-mono"
                      />
                      <select
                        value={citationStyle}
                        onChange={(e) => setCitationStyle(e.target.value as any)}
                        className="flex-1 p-1.5 text-xs rounded border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800"
                      >
                        <option value="APA">APA 7th</option>
                        <option value="MLA">MLA 9th</option>
                        <option value="Chicago">Chicago</option>
                      </select>
                    </div>
                    <button
                      onClick={addCitation}
                      className="w-full py-1 rounded bg-indigo-600 hover:bg-indigo-700 text-white font-semibold text-xs"
                    >
                      Add Citation Entry
                    </button>
                    <button
                      onClick={insertBibliography}
                      className="w-full py-1 rounded bg-slate-200 dark:bg-slate-800 hover:bg-slate-300 text-slate-800 dark:text-slate-200 font-semibold text-xs"
                    >
                      Generate Bibliography
                    </button>
                  </div>
                </div>
              )}

              {/* Watermark Drawer */}
              {activeSideDrawer === 'watermark' && (
                <div className="space-y-3">
                  <div className="flex items-center justify-between pb-1 border-b border-slate-200 dark:border-slate-800">
                    <span className="font-bold text-slate-700 dark:text-slate-200">Watermark Text</span>
                    <button
                      onClick={() =>
                        onChange({
                          ...data,
                          watermark: {
                            text: watermarkInput,
                            enabled: !(data.watermark?.enabled),
                          },
                        })
                      }
                      className={`px-2 py-0.5 rounded text-xs font-bold ${
                        data.watermark?.enabled ? 'bg-blue-600 text-white' : 'bg-slate-200 dark:bg-slate-800 text-slate-700 dark:text-slate-300'
                      }`}
                    >
                      {data.watermark?.enabled ? 'Enabled' : 'Disabled'}
                    </button>
                  </div>
                  <input
                    type="text"
                    value={watermarkInput}
                    onChange={(e) => {
                      setWatermarkInput(e.target.value);
                      onChange({
                        ...data,
                        watermark: {
                          text: e.target.value,
                          enabled: true,
                        },
                      });
                    }}
                    placeholder="CONFIDENTIAL, DRAFT, INTERNAL..."
                    className="w-full p-2 text-xs rounded border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 font-bold uppercase tracking-wider text-slate-700 dark:text-slate-200"
                  />
                  <div className="flex items-center space-x-2 pt-2">
                    {['CONFIDENTIAL', 'DRAFT', 'URGENT', 'DO NOT COPY'].map((preset) => (
                      <button
                        key={preset}
                        onClick={() => {
                          setWatermarkInput(preset);
                          onChange({
                            ...data,
                            watermark: { text: preset, enabled: true },
                          });
                        }}
                        className="px-2 py-1 rounded bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 text-[10px] font-semibold"
                      >
                        {preset}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              {/* Mail Merge Drawer */}
              {activeSideDrawer === 'mailMerge' && (
                <div className="space-y-3">
                  <div className="flex items-center justify-between pb-1 border-b border-slate-200 dark:border-slate-800">
                    <span className="font-bold text-slate-700 dark:text-slate-200">Mail Merge Preview</span>
                    <button
                      onClick={() => {
                        const enabled = !data.mailMerge?.enabled;
                        onChange({
                          ...data,
                          mailMerge: {
                            enabled,
                            variables: ['First_Name', 'Company', 'Date', 'Amount'],
                            records: sampleMailMergeRecords,
                            activeRecordIndex: 0,
                          },
                        });
                      }}
                      className={`px-2 py-0.5 rounded text-xs font-bold ${
                        data.mailMerge?.enabled ? 'bg-indigo-600 text-white' : 'bg-slate-200 dark:bg-slate-800 text-slate-700 dark:text-slate-300'
                      }`}
                    >
                      {data.mailMerge?.enabled ? 'Active' : 'Enable'}
                    </button>
                  </div>
                  <p className="text-[11px] text-slate-500">
                    Use <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded font-mono">{"{{First_Name}}"}</code>, <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded font-mono">{"{{Company}}"}</code> in text.
                  </p>
                  <div className="space-y-1">
                    {sampleMailMergeRecords.map((rec, idx) => (
                      <div
                        key={idx}
                        onClick={() => {
                          if (data.mailMerge) {
                            onChange({
                              ...data,
                              mailMerge: { ...data.mailMerge, activeRecordIndex: idx, enabled: true },
                            });
                          }
                        }}
                        className={`p-2 rounded border cursor-pointer text-xs ${
                          (data.mailMerge?.activeRecordIndex || 0) === idx
                            ? 'border-indigo-500 bg-indigo-50 dark:bg-indigo-950/40 text-indigo-700 dark:text-indigo-200 font-bold'
                            : 'border-slate-200 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-800/40'
                        }`}
                      >
                        <div>{rec.First_Name} • {rec.Company}</div>
                        <div className="text-[10px] text-slate-400 font-normal">{rec.Date} — {rec.Amount}</div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      {/* 4. REAL-TIME WORD DOCUMENT STATUS BAR */}
      <footer
        className={`h-6 border-t px-4 flex items-center justify-between text-[11px] select-none shrink-0 transition-colors ${
          darkMode ? 'bg-slate-950 border-slate-800 text-slate-400' : 'bg-slate-200 border-slate-300 text-slate-600'
        }`}
      >
        <div className="flex items-center space-x-4">
          <span className="font-semibold text-blue-600 dark:text-blue-400">Microsoft Word Parity</span>
          <span>•</span>
          <span>Page 1 of {stats.estimatedPages}</span>
          <span>•</span>
          <span>{stats.words} Words</span>
          <span>•</span>
          <span>{stats.characters} Characters</span>
          <span>•</span>
          <span>~{stats.readingMinutes} min read</span>
        </div>

        <div className="flex items-center space-x-3">
          <button
            onClick={() => setZoomLevel((prev) => Math.max(50, prev - 10))}
            className="p-0.5 rounded hover:bg-slate-300 dark:hover:bg-slate-800"
            title="Zoom Out"
          >
            <ZoomOut className="w-3.5 h-3.5" />
          </button>
          <span className="font-mono text-[10px] w-9 text-center">{zoomLevel}%</span>
          <button
            onClick={() => setZoomLevel((prev) => Math.min(200, prev + 10))}
            className="p-0.5 rounded hover:bg-slate-300 dark:hover:bg-slate-800"
            title="Zoom In"
          >
            <ZoomIn className="w-3.5 h-3.5" />
          </button>
        </div>
      </footer>
    </div>
  );
};
