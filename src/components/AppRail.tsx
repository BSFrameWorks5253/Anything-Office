import React from 'react';
import type { AppMode } from '../types/office';
import {
  FileText,
  Table as TableIcon,
  Presentation,
  FileDigit,
  LayoutTemplate,
  BookOpen,
} from 'lucide-react';

interface AppRailProps {
  appMode: AppMode;
  setAppMode: (mode: AppMode) => void;
  onOpenTemplates: () => void;
  onOpenEvidence: () => void;
  darkMode: boolean;
}

export const AppRail: React.FC<AppRailProps> = ({
  appMode,
  setAppMode,
  onOpenTemplates,
  onOpenEvidence,
  darkMode,
}) => {
  return (
    <aside
      className={`w-12 shrink-0 border-r flex flex-col items-center justify-between py-2 select-none transition-colors z-20 ${
        darkMode ? 'bg-slate-900 border-slate-800 text-slate-300' : 'bg-slate-50 border-slate-200 text-slate-600'
      }`}
    >
      {/* Microsoft 365 Apps Navigation: Word, Excel, PowerPoint, PDF */}
      <nav className="flex flex-col items-center space-y-2 w-full">
        {/* Word (Docs) */}
          <button
            onClick={() => setAppMode('docs')}
            title="Microsoft Word (Docs)"
            className={`relative w-9 h-9 rounded-lg flex items-center justify-center transition-all ${
              appMode === 'docs'
                ? 'bg-blue-600/10 text-blue-600 dark:text-blue-400 font-bold'
                : 'hover:bg-slate-200/60 dark:hover:bg-slate-800 text-slate-500'
            }`}
          >
            {appMode === 'docs' && (
              <span className="absolute left-0 top-2 bottom-2 w-1 rounded-r bg-[#185abd]" />
            )}
            <FileText className="w-4 h-4" />
          </button>

          {/* Excel (Sheets) */}
          <button
            onClick={() => setAppMode('sheets')}
            title="Microsoft Excel (Sheets)"
            className={`relative w-9 h-9 rounded-lg flex items-center justify-center transition-all ${
              appMode === 'sheets'
                ? 'bg-emerald-600/10 text-emerald-600 dark:text-emerald-400 font-bold'
                : 'hover:bg-slate-200/60 dark:hover:bg-slate-800 text-slate-500'
            }`}
          >
            {appMode === 'sheets' && (
              <span className="absolute left-0 top-2 bottom-2 w-1 rounded-r bg-[#107c41]" />
            )}
            <TableIcon className="w-4 h-4" />
          </button>

          {/* PowerPoint (Slides) */}
          <button
            onClick={() => setAppMode('slides')}
            title="Microsoft PowerPoint (Slides)"
            className={`relative w-9 h-9 rounded-lg flex items-center justify-center transition-all ${
              appMode === 'slides'
                ? 'bg-orange-600/10 text-orange-600 dark:text-orange-400 font-bold'
                : 'hover:bg-slate-200/60 dark:hover:bg-slate-800 text-slate-500'
            }`}
          >
            {appMode === 'slides' && (
              <span className="absolute left-0 top-2 bottom-2 w-1 rounded-r bg-[#c43e1c]" />
            )}
            <Presentation className="w-4 h-4" />
          </button>

          {/* Adobe / Microsoft PDF */}
          <button
            onClick={() => setAppMode('pdf')}
            title="PDF Document Viewer & Word Reflow"
            className={`relative w-9 h-9 rounded-lg flex items-center justify-center transition-all ${
              appMode === 'pdf'
                ? 'bg-red-600/10 text-red-600 dark:text-red-400 font-bold'
                : 'hover:bg-slate-200/60 dark:hover:bg-slate-800 text-slate-500'
            }`}
          >
            {appMode === 'pdf' && (
              <span className="absolute left-0 top-2 bottom-2 w-1 rounded-r bg-[#e11d48]" />
            )}
            <FileDigit className="w-4 h-4" />
          </button>
        </nav>

      {/* Bottom Actions: Templates & Architecture Evidence */}
      <div className="flex flex-col items-center space-y-2">
        <button
          onClick={onOpenTemplates}
          title="Template Gallery"
          className="w-9 h-9 rounded-lg flex items-center justify-center text-slate-500 hover:text-slate-800 dark:hover:text-slate-200 hover:bg-slate-200/60 dark:hover:bg-slate-800 transition-colors"
        >
          <LayoutTemplate className="w-4 h-4" />
        </button>

        <button
          onClick={onOpenEvidence}
          title="Architecture & Standards Evidence"
          className="w-9 h-9 rounded-lg flex items-center justify-center text-slate-500 hover:text-indigo-600 dark:hover:text-indigo-400 hover:bg-slate-200/60 dark:hover:bg-slate-800 transition-colors"
        >
          <BookOpen className="w-4 h-4" />
        </button>
      </div>
    </aside>
  );
};
