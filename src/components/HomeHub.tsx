import React, { useState } from 'react';
import type { AppMode, AnythingProject, RecentDocumentItem } from '../types/office';
import {
  FileText,
  Table as TableIcon,
  Presentation,
  FileDigit,
  Sparkles,
  Search,
  Pin,
  PinOff,
  Trash2,
  Clock,
  FolderOpen,
  UploadCloud,
  ChevronRight,
  LayoutGrid,
  List as ListIcon,
  BookOpen,
} from 'lucide-react';

interface HomeHubProps {
  onOpenAppMode: (mode: AppMode) => void;
  onOpenProject: (project: AnythingProject) => void;
  onOpenFileClick: () => void;
  onOpenEvidence: () => void;
  onOpenTemplates: () => void;
  recentItems: RecentDocumentItem[];
  onTogglePinRecent: (id: string) => void;
  onDeleteRecent: (id: string) => void;
  onSelectRecentItem: (item: RecentDocumentItem) => void;
  darkMode: boolean;
}

export const HomeHub: React.FC<HomeHubProps> = ({
  onOpenAppMode,
  onOpenFileClick,
  onOpenEvidence,
  onOpenTemplates,
  recentItems,
  onTogglePinRecent,
  onDeleteRecent,
  onSelectRecentItem,
  darkMode,
}) => {
  const [activeFilter, setActiveFilter] = useState<'all' | 'pinned' | 'docs' | 'sheets' | 'slides' | 'pdf'>('all');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [viewMode, setViewMode] = useState<'list' | 'grid'>('list');

  // Time-based greeting
  const getGreeting = () => {
    const hour = new Date().getHours();
    if (hour < 12) return 'Good morning';
    if (hour < 18) return 'Good afternoon';
    return 'Good evening';
  };

  // Filter recent documents
  const filteredRecents = recentItems.filter((item) => {
    // Search query match
    if (searchQuery.trim() && !item.name.toLowerCase().includes(searchQuery.toLowerCase())) {
      return false;
    }
    // Tab filters
    if (activeFilter === 'pinned') return item.pinned;
    if (activeFilter === 'docs') return item.type === 'docs';
    if (activeFilter === 'sheets') return item.type === 'sheets';
    if (activeFilter === 'slides') return item.type === 'slides';
    if (activeFilter === 'pdf') return item.type === 'pdf';
    return true;
  });

  const formatRelativeTime = (timestamp: number): string => {
    const diff = Math.floor((Date.now() - timestamp) / 1000);
    if (diff < 60) return 'Just now';
    if (diff < 3600) return `${Math.floor(diff / 60)}m ago`;
    if (diff < 86400) return `${Math.floor(diff / 3600)}h ago`;
    if (diff < 604800) return `${Math.floor(diff / 86400)}d ago`;
    return new Date(timestamp).toLocaleDateString();
  };

  const getModuleBadge = (type: RecentDocumentItem['type']) => {
    switch (type) {
      case 'docs':
        return {
          label: 'Word',
          ext: '.docx',
          bg: 'bg-blue-600',
          lightBg: 'bg-blue-50 dark:bg-blue-950/40 text-blue-600 dark:text-blue-400',
          icon: FileText,
        };
      case 'sheets':
        return {
          label: 'Excel',
          ext: '.xlsx',
          bg: 'bg-emerald-600',
          lightBg: 'bg-emerald-50 dark:bg-emerald-950/40 text-emerald-600 dark:text-emerald-400',
          icon: TableIcon,
        };
      case 'slides':
        return {
          label: 'PowerPoint',
          ext: '.pptx',
          bg: 'bg-orange-600',
          lightBg: 'bg-orange-50 dark:bg-orange-950/40 text-orange-600 dark:text-orange-400',
          icon: Presentation,
        };
      case 'pdf':
        return {
          label: 'Acrobat PDF',
          ext: '.pdf',
          bg: 'bg-red-600',
          lightBg: 'bg-red-50 dark:bg-red-950/40 text-red-600 dark:text-red-400',
          icon: FileDigit,
        };
      case 'anything':
      default:
        return {
          label: 'Unified',
          ext: '.anything',
          bg: 'bg-purple-600',
          lightBg: 'bg-purple-50 dark:bg-purple-950/40 text-purple-600 dark:text-purple-400',
          icon: Sparkles,
        };
    }
  };

  return (
    <div
      className={`flex-1 overflow-y-auto select-none transition-colors ${
        darkMode ? 'bg-slate-950 text-slate-100' : 'bg-[#faf9f8] text-slate-900'
      }`}
    >
      <div className="max-w-6xl mx-auto px-8 py-8 space-y-9">
        {/* 1. HERO GREETING & QUICK OPEN */}
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-2 border-b border-slate-200 dark:border-slate-800">
          <div>
            <h1 className="text-2xl font-black tracking-tight text-slate-900 dark:text-white flex items-center gap-2.5">
              <span>{getGreeting()}</span>
              <span className="text-xs px-2.5 py-0.5 rounded-full font-bold bg-violet-600/10 text-violet-600 dark:text-violet-400 border border-violet-500/20">
                Anything Office Hub
              </span>
            </h1>
            <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
              Create, calculate, present, and annotate across Microsoft 365 and Adobe formats.
            </p>
          </div>

          <div className="flex items-center gap-2.5">
            <button
              onClick={onOpenFileClick}
              className="flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-semibold bg-white dark:bg-slate-850 border border-slate-300 dark:border-slate-700 hover:border-slate-400 shadow-xs transition-all active:scale-95"
            >
              <FolderOpen className="w-4 h-4 text-blue-500" />
              <span>Open Document...</span>
            </button>

            <button
              onClick={onOpenEvidence}
              className="flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-semibold bg-indigo-50 dark:bg-indigo-950/40 text-indigo-700 dark:text-indigo-300 border border-indigo-200 dark:border-indigo-800 hover:bg-indigo-100 dark:hover:bg-indigo-900/60 shadow-xs transition-all active:scale-95"
            >
              <BookOpen className="w-4 h-4 text-indigo-500" />
              <span>Architecture Evidence</span>
            </button>
          </div>
        </div>

        {/* 2. CREATE NEW / CREATE DOCUMENT BAR (Microsoft 365 + Adobe Acrobat) */}
        <div>
          <div className="flex items-center justify-between mb-3">
            <h2 className="text-xs font-bold uppercase tracking-wider text-slate-400">
              Create New
            </h2>
            <button
              onClick={onOpenTemplates}
              className="text-xs font-semibold text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-0.5"
            >
              <span>Explore All Templates</span>
              <ChevronRight className="w-3.5 h-3.5" />
            </button>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-3.5">
            {/* Blank Word Document */}
            <button
              onClick={() => onOpenAppMode('docs')}
              className={`p-4 rounded-2xl border text-left flex flex-col justify-between h-36 transition-all duration-200 group hover:-translate-y-1 hover:shadow-xl ${
                darkMode
                  ? 'bg-slate-900 border-slate-800 hover:border-blue-500/60'
                  : 'bg-white border-slate-200/90 hover:border-blue-400 shadow-xs'
              }`}
            >
              <div className="flex items-center justify-between">
                <div className="w-10 h-10 rounded-xl bg-blue-600 text-white flex items-center justify-center font-bold shadow-md shadow-blue-500/20 group-hover:scale-105 transition-transform">
                  <FileText className="w-5 h-5" />
                </div>
                <span className="text-[10px] font-mono font-bold text-blue-600 dark:text-blue-400 px-1.5 py-0.5 rounded bg-blue-50 dark:bg-blue-950/50">
                  .docx
                </span>
              </div>
              <div>
                <h3 className="text-xs font-bold text-slate-900 dark:text-white group-hover:text-blue-600 dark:group-hover:text-blue-400 transition-colors">
                  Blank Document
                </h3>
                <p className="text-[11px] text-slate-400 mt-0.5">Word & Writer Flow</p>
              </div>
            </button>

            {/* Blank Excel Workbook */}
            <button
              onClick={() => onOpenAppMode('sheets')}
              className={`p-4 rounded-2xl border text-left flex flex-col justify-between h-36 transition-all duration-200 group hover:-translate-y-1 hover:shadow-xl ${
                darkMode
                  ? 'bg-slate-900 border-slate-800 hover:border-emerald-500/60'
                  : 'bg-white border-slate-200/90 hover:border-emerald-400 shadow-xs'
              }`}
            >
              <div className="flex items-center justify-between">
                <div className="w-10 h-10 rounded-xl bg-emerald-600 text-white flex items-center justify-center font-bold shadow-md shadow-emerald-500/20 group-hover:scale-105 transition-transform">
                  <TableIcon className="w-5 h-5" />
                </div>
                <span className="text-[10px] font-mono font-bold text-emerald-600 dark:text-emerald-400 px-1.5 py-0.5 rounded bg-emerald-50 dark:bg-emerald-950/50">
                  .xlsx
                </span>
              </div>
              <div>
                <h3 className="text-xs font-bold text-slate-900 dark:text-white group-hover:text-emerald-600 dark:group-hover:text-emerald-400 transition-colors">
                  Blank Workbook
                </h3>
                <p className="text-[11px] text-slate-400 mt-0.5">Excel & Calc Grid</p>
              </div>
            </button>

            {/* Blank PowerPoint Presentation */}
            <button
              onClick={() => onOpenAppMode('slides')}
              className={`p-4 rounded-2xl border text-left flex flex-col justify-between h-36 transition-all duration-200 group hover:-translate-y-1 hover:shadow-xl ${
                darkMode
                  ? 'bg-slate-900 border-slate-800 hover:border-orange-500/60'
                  : 'bg-white border-slate-200/90 hover:border-orange-400 shadow-xs'
              }`}
            >
              <div className="flex items-center justify-between">
                <div className="w-10 h-10 rounded-xl bg-orange-600 text-white flex items-center justify-center font-bold shadow-md shadow-orange-500/20 group-hover:scale-105 transition-transform">
                  <Presentation className="w-5 h-5" />
                </div>
                <span className="text-[10px] font-mono font-bold text-orange-600 dark:text-orange-400 px-1.5 py-0.5 rounded bg-orange-50 dark:bg-orange-950/50">
                  .pptx
                </span>
              </div>
              <div>
                <h3 className="text-xs font-bold text-slate-900 dark:text-white group-hover:text-orange-600 dark:group-hover:text-orange-400 transition-colors">
                  New Presentation
                </h3>
                <p className="text-[11px] text-slate-400 mt-0.5">PowerPoint 16:9 Deck</p>
              </div>
            </button>

            {/* Adobe Acrobat PDF Tool Suite */}
            <button
              onClick={() => onOpenAppMode('pdf')}
              className={`p-4 rounded-2xl border text-left flex flex-col justify-between h-36 transition-all duration-200 group hover:-translate-y-1 hover:shadow-xl ${
                darkMode
                  ? 'bg-slate-900 border-slate-800 hover:border-red-500/60'
                  : 'bg-white border-slate-200/90 hover:border-red-400 shadow-xs'
              }`}
            >
              <div className="flex items-center justify-between">
                <div className="w-10 h-10 rounded-xl bg-red-600 text-white flex items-center justify-center font-bold shadow-md shadow-red-500/20 group-hover:scale-105 transition-transform">
                  <FileDigit className="w-5 h-5" />
                </div>
                <span className="text-[10px] font-mono font-bold text-red-600 dark:text-red-400 px-1.5 py-0.5 rounded bg-red-50 dark:bg-red-950/50">
                  .pdf
                </span>
              </div>
              <div>
                <h3 className="text-xs font-bold text-slate-900 dark:text-white group-hover:text-red-600 dark:group-hover:text-red-400 transition-colors">
                  PDF Acrobat Suite
                </h3>
                <p className="text-[11px] text-slate-400 mt-0.5">View, Markup & Sign</p>
              </div>
            </button>

            {/* Unified Workspace (.anything) */}
            <button
              onClick={onOpenTemplates}
              className={`p-4 rounded-2xl border text-left flex flex-col justify-between h-36 transition-all duration-200 group hover:-translate-y-1 hover:shadow-xl ${
                darkMode
                  ? 'bg-slate-900 border-slate-800 hover:border-purple-500/60'
                  : 'bg-white border-slate-200/90 hover:border-purple-400 shadow-xs'
              }`}
            >
              <div className="flex items-center justify-between">
                <div className="w-10 h-10 rounded-xl bg-purple-600 text-white flex items-center justify-center font-bold shadow-md shadow-purple-500/20 group-hover:scale-105 transition-transform">
                  <Sparkles className="w-5 h-5" />
                </div>
                <span className="text-[10px] font-mono font-bold text-purple-600 dark:text-purple-400 px-1.5 py-0.5 rounded bg-purple-50 dark:bg-purple-950/50">
                  Template
                </span>
              </div>
              <div>
                <h3 className="text-xs font-bold text-slate-900 dark:text-white group-hover:text-purple-600 dark:group-hover:text-purple-400 transition-colors">
                  Template Gallery
                </h3>
                <p className="text-[11px] text-slate-400 mt-0.5">Pro Business Packs</p>
              </div>
            </button>
          </div>
        </div>

        {/* 3. RECENT DOCUMENTS HUB (The Core Experience) */}
        <div className="space-y-4">
          {/* Header & Filter Controls */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            {/* Filter Pills */}
            <div className="flex items-center gap-1 overflow-x-auto pb-1 sm:pb-0">
              {[
                { id: 'all', label: 'Recent' },
                { id: 'pinned', label: 'Pinned' },
                { id: 'docs', label: 'Word' },
                { id: 'sheets', label: 'Excel' },
                { id: 'slides', label: 'PowerPoint' },
                { id: 'pdf', label: 'PDF' },
              ].map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => setActiveFilter(tab.id as any)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                    activeFilter === tab.id
                      ? 'bg-slate-900 text-white dark:bg-white dark:text-slate-900 shadow-xs'
                      : 'hover:bg-slate-200/60 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-400'
                  }`}
                >
                  {tab.label}
                </button>
              ))}
            </div>

            {/* Search & View Toggle */}
            <div className="flex items-center gap-2">
              <div className="relative">
                <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="text"
                  placeholder="Filter recents..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="pl-8 pr-3 py-1.5 rounded-lg text-xs border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 focus:outline-none focus:ring-1 focus:ring-blue-500 w-44 sm:w-52"
                />
              </div>

              <div className="flex items-center border border-slate-200 dark:border-slate-800 rounded-lg p-0.5 bg-white dark:bg-slate-900">
                <button
                  onClick={() => setViewMode('list')}
                  className={`p-1 rounded ${
                    viewMode === 'list'
                      ? 'bg-slate-100 dark:bg-slate-800 text-slate-900 dark:text-white'
                      : 'text-slate-400'
                  }`}
                  title="List View"
                >
                  <ListIcon className="w-3.5 h-3.5" />
                </button>
                <button
                  onClick={() => setViewMode('grid')}
                  className={`p-1 rounded ${
                    viewMode === 'grid'
                      ? 'bg-slate-100 dark:bg-slate-800 text-slate-900 dark:text-white'
                      : 'text-slate-400'
                  }`}
                  title="Grid View"
                >
                  <LayoutGrid className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>
          </div>

          {/* List or Grid Display */}
          {filteredRecents.length === 0 ? (
            <div className="py-16 text-center border-2 border-dashed border-slate-200 dark:border-slate-800 rounded-2xl flex flex-col items-center">
              <Clock className="w-8 h-8 text-slate-300 dark:text-slate-600 mb-2" />
              <p className="text-xs font-semibold text-slate-500 dark:text-slate-400">
                No matching documents found
              </p>
              <p className="text-[11px] text-slate-400 mt-1">
                Create a new document or open any Office/PDF file to start.
              </p>
            </div>
          ) : viewMode === 'list' ? (
            /* LIST VIEW (Table style) */
            <div className="rounded-2xl border border-slate-200 dark:border-slate-800 overflow-hidden bg-white dark:bg-slate-900 shadow-xs">
              <div className="grid grid-cols-12 px-4 py-2.5 text-[11px] font-bold text-slate-400 uppercase tracking-wider border-b border-inherit bg-slate-50/50 dark:bg-slate-950/40">
                <div className="col-span-6 sm:col-span-5">Name</div>
                <div className="col-span-3 sm:col-span-3">Type & Format</div>
                <div className="col-span-3 sm:col-span-3">Last Modified</div>
                <div className="hidden sm:block sm:col-span-1 text-right">Action</div>
              </div>

              <div className="divide-y divide-slate-100 dark:divide-slate-800/80">
                {filteredRecents.map((item) => {
                  const badge = getModuleBadge(item.type);
                  const Icon = badge.icon;
                  return (
                    <div
                      key={item.id}
                      onClick={() => onSelectRecentItem(item)}
                      className="grid grid-cols-12 px-4 py-3 items-center hover:bg-slate-50 dark:hover:bg-slate-800/60 cursor-pointer transition-colors group text-xs"
                    >
                      {/* Name & Icon */}
                      <div className="col-span-6 sm:col-span-5 flex items-center space-x-3 pr-2 truncate">
                        <div
                          className={`w-8 h-8 rounded-lg flex items-center justify-center shrink-0 ${badge.bg} text-white shadow-xs`}
                        >
                          <Icon className="w-4 h-4" />
                        </div>
                        <div className="truncate">
                          <h4 className="font-semibold text-slate-900 dark:text-white truncate group-hover:text-blue-600 dark:group-hover:text-blue-400 transition-colors">
                            {item.name}
                          </h4>
                          <span className="text-[10px] text-slate-400">
                            {item.source === 'template' ? 'Built-in Template' : item.size}
                          </span>
                        </div>
                      </div>

                      {/* Type Pill */}
                      <div className="col-span-3 sm:col-span-3 flex items-center gap-1.5">
                        <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${badge.lightBg}`}>
                          {badge.label}
                        </span>
                        <span className="font-mono text-[10px] text-slate-400 hidden md:inline">
                          {badge.ext}
                        </span>
                      </div>

                      {/* Last Modified */}
                      <div className="col-span-3 sm:col-span-3 text-[11px] text-slate-500 dark:text-slate-400 flex items-center gap-1">
                        <Clock className="w-3 h-3 text-slate-400" />
                        <span>{formatRelativeTime(item.lastModified)}</span>
                      </div>

                      {/* Actions (Pin, Delete) */}
                      <div className="hidden sm:flex sm:col-span-1 items-center justify-end space-x-1">
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            onTogglePinRecent(item.id);
                          }}
                          className={`p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-700 transition-colors ${
                            item.pinned ? 'text-amber-500 font-bold' : 'text-slate-400 opacity-0 group-hover:opacity-100'
                          }`}
                          title={item.pinned ? 'Unpin Document' : 'Pin to Top'}
                        >
                          {item.pinned ? <Pin className="w-3.5 h-3.5 fill-current" /> : <PinOff className="w-3.5 h-3.5" />}
                        </button>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            onDeleteRecent(item.id);
                          }}
                          className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-400 hover:text-red-500 transition-colors opacity-0 group-hover:opacity-100"
                          title="Remove from Recents"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          ) : (
            /* GRID VIEW (Cards) */
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
              {filteredRecents.map((item) => {
                const badge = getModuleBadge(item.type);
                const Icon = badge.icon;
                return (
                  <div
                    key={item.id}
                    onClick={() => onSelectRecentItem(item)}
                    className={`p-4 rounded-2xl border text-left flex flex-col justify-between h-40 transition-all duration-200 group cursor-pointer hover:-translate-y-1 hover:shadow-lg ${
                      darkMode
                        ? 'bg-slate-900 border-slate-800 hover:border-slate-700'
                        : 'bg-white border-slate-200 hover:border-slate-300 shadow-xs'
                    }`}
                  >
                    <div className="flex items-start justify-between">
                      <div
                        className={`w-9 h-9 rounded-xl flex items-center justify-center ${badge.bg} text-white shadow-xs`}
                      >
                        <Icon className="w-4 h-4" />
                      </div>
                      <div className="flex items-center space-x-1">
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            onTogglePinRecent(item.id);
                          }}
                          className={`p-1 rounded ${
                            item.pinned ? 'text-amber-500' : 'text-slate-400 opacity-0 group-hover:opacity-100'
                          }`}
                          title={item.pinned ? 'Unpin' : 'Pin'}
                        >
                          <Pin className={`w-3.5 h-3.5 ${item.pinned ? 'fill-current' : ''}`} />
                        </button>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            onDeleteRecent(item.id);
                          }}
                          className="p-1 rounded text-slate-400 hover:text-red-500 opacity-0 group-hover:opacity-100"
                          title="Remove"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>

                    <div>
                      <h4 className="text-xs font-bold text-slate-900 dark:text-white truncate group-hover:text-blue-600 dark:group-hover:text-blue-400 transition-colors">
                        {item.name}
                      </h4>
                      <div className="flex items-center justify-between text-[11px] text-slate-400 mt-2">
                        <span className={`px-1.5 py-0.2 rounded text-[10px] font-bold ${badge.lightBg}`}>
                          {badge.label}
                        </span>
                        <span>{formatRelativeTime(item.lastModified)}</span>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* 4. DRAG & DROP QUICK INGESTION ZONE */}
        <div
          onClick={onOpenFileClick}
          className="border-2 border-dashed border-slate-300 dark:border-slate-800 hover:border-blue-500 dark:hover:border-blue-500 rounded-2xl p-6 flex flex-col items-center justify-center text-center cursor-pointer transition-colors bg-white/40 dark:bg-slate-900/40 group"
        >
          <UploadCloud className="w-8 h-8 text-blue-500 mb-2 group-hover:scale-110 transition-transform" />
          <h3 className="text-xs font-bold text-slate-800 dark:text-slate-200">
            Drop any document to open in Anything Office
          </h3>
          <p className="text-[11px] text-slate-400 mt-0.5">
            Supports Microsoft Office (.docx, .xlsx, .pptx), PDF (.pdf), LibreOffice (.odt, .ods, .odp), or CSV
          </p>
        </div>
      </div>
    </div>
  );
};
