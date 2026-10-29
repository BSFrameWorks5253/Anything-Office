import React from 'react';
import type { AppMode } from '../types/office';
import { AdobeLogo } from './AdobeLogo';
import {
  Save,
  RotateCcw,
  RotateCw,
  Search,
  Share2,
  Sun,
  Moon,
  Minus,
  Square,
  X,
} from 'lucide-react';

interface TitleBarProps {
  appMode: AppMode;
  docTitle: string;
  onTitleChange: (newTitle: string) => void;
  autoSaveEnabled: boolean;
  setAutoSaveEnabled: (val: boolean) => void;
  onSave: () => void;
  onUndo: () => void;
  onRedo: () => void;
  canUndo: boolean;
  canRedo: boolean;
  onExport: (format: 'bundle' | 'html' | 'csv' | 'markdown' | 'print') => void;
  darkMode: boolean;
  setDarkMode: (val: boolean) => void;
}

export const TitleBar: React.FC<TitleBarProps> = ({
  appMode,
  docTitle,
  onTitleChange,
  autoSaveEnabled,
  setAutoSaveEnabled,
  onSave,
  onUndo,
  onRedo,
  canUndo,
  canRedo,
  onExport,
  darkMode,
  setDarkMode,
}) => {
  const getAppPill = () => {
    switch (appMode) {
      case 'home':
        return { name: 'Home Hub', color: 'bg-violet-600' };
      case 'docs':
        return { name: 'Word', color: 'bg-blue-600' };
      case 'sheets':
        return { name: 'Excel', color: 'bg-emerald-600' };
      case 'slides':
        return { name: 'PowerPoint', color: 'bg-orange-600' };
      case 'pdf':
        return { name: 'Acrobat PDF', color: 'bg-red-600' };
      default:
        return { name: 'Workspace', color: 'bg-blue-600' };
    }
  };

  const currentPill = getAppPill();

  return (
    <header
      className={`h-9 w-full flex items-center justify-between px-2 text-xs border-b select-none transition-colors z-40 shrink-0 ${
        darkMode ? 'bg-slate-900 border-slate-800 text-slate-200' : 'bg-slate-100 border-slate-300 text-slate-800'
      }`}
      style={{ WebkitAppRegion: 'drag' } as any}
    >
      {/* 1. LEFT: App Brand & Quick Access Toolbar */}
      <div
        className="flex items-center space-x-2"
        style={{ WebkitAppRegion: 'no-drag' } as any}
      >
        {/* Monogram & Title */}
        <div className="flex items-center space-x-2 mr-1">
          <AdobeLogo size="sm" monogram="Ao" theme="anything" />
          <span className="font-bold text-xs tracking-tight hidden sm:inline">
            Anything Office
          </span>
          <span className={`text-[10px] font-bold text-white px-1.5 py-0.2 rounded ${currentPill.color}`}>
            {currentPill.name}
          </span>
        </div>

        <div className="h-4 w-px bg-slate-300 dark:bg-slate-700" />

        {/* AutoSave Toggle Switch */}
        <div className="flex items-center space-x-1.5 pl-1">
          <button
            onClick={() => setAutoSaveEnabled(!autoSaveEnabled)}
            className={`w-7 h-4 flex items-center rounded-full p-0.5 transition-colors ${
              autoSaveEnabled ? 'bg-blue-600' : 'bg-slate-400'
            }`}
            title="AutoSave: Automatically saves changes to local storage"
          >
            <div
              className={`bg-white w-3 h-3 rounded-full shadow-sm transform transition-transform ${
                autoSaveEnabled ? 'translate-x-3' : 'translate-x-0'
              }`}
            />
          </button>
          <span className="text-[11px] font-medium text-slate-500 dark:text-slate-400 hidden md:inline">
            AutoSave
          </span>
        </div>

        {/* Quick Save */}
        <button
          onClick={onSave}
          title="Save Document (Ctrl+S)"
          className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-300 transition-colors"
        >
          <Save className="w-3.5 h-3.5" />
        </button>

        {/* Quick Undo / Redo */}
        <div className="flex items-center space-x-0.5">
          <button
            onClick={onUndo}
            disabled={!canUndo}
            title="Undo (Ctrl+Z)"
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 disabled:opacity-30 text-slate-600 dark:text-slate-300 transition-colors"
          >
            <RotateCcw className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={onRedo}
            disabled={!canRedo}
            title="Redo (Ctrl+Y)"
            className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 disabled:opacity-30 text-slate-600 dark:text-slate-300 transition-colors"
          >
            <RotateCw className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* 2. CENTER: Document Title & Microsoft Search Bar */}
      <div
        className="flex items-center space-x-3 max-w-md w-full justify-center px-4"
        style={{ WebkitAppRegion: 'no-drag' } as any}
      >
        <div className="flex items-center space-x-1.5 max-w-[200px] sm:max-w-xs">
          <input
            type="text"
            value={docTitle}
            onChange={(e) => onTitleChange(e.target.value)}
            className="bg-transparent text-center hover:bg-slate-200/60 dark:hover:bg-slate-800 rounded px-2 py-0.5 focus:outline-none focus:ring-1 focus:ring-blue-500 font-semibold text-xs truncate"
          />
          <span className="text-[10px] text-slate-400 font-normal hidden lg:inline">
            - Saved
          </span>
        </div>

        {/* Microsoft Search Box (Alt+Q) */}
        <div className="hidden lg:flex items-center bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-md px-2.5 py-0.5 w-56 shadow-xs text-slate-400 focus-within:ring-1 focus-within:ring-blue-500">
          <Search className="w-3 h-3 text-slate-400 mr-2 shrink-0" />
          <input
            type="text"
            placeholder="Search (Alt+Q)"
            className="w-full bg-transparent text-xs text-slate-700 dark:text-slate-200 focus:outline-none"
          />
        </div>
      </div>

      {/* 3. RIGHT: Share, Theme, Windows Native Control Buttons */}
      <div
        className="flex items-center space-x-1"
        style={{ WebkitAppRegion: 'no-drag' } as any}
      >
        {/* Share Button */}
        <button
          onClick={() => onExport('bundle')}
          className="flex items-center space-x-1.5 px-2.5 py-1 rounded bg-blue-600 hover:bg-blue-700 text-white font-medium text-xs shadow-xs mr-1"
        >
          <Share2 className="w-3 h-3" />
          <span className="hidden sm:inline">Share</span>
        </button>

        {/* Theme Toggle */}
        <button
          onClick={() => setDarkMode(!darkMode)}
          className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-400"
          title="Toggle Dark / Light Mode"
        >
          {darkMode ? <Sun className="w-3.5 h-3.5 text-amber-400" /> : <Moon className="w-3.5 h-3.5" />}
        </button>

        {/* Native Windows Desktop Window Controls */}
        <div className="flex items-center pl-1 border-l border-slate-300 dark:border-slate-700 h-9 -mr-2">
          {/* Minimize */}
          <button
            onClick={() => window.electronAPI?.minimize()}
            className="h-full px-3 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-600 dark:text-slate-300 flex items-center justify-center transition-colors"
            title="Minimize"
          >
            <Minus className="w-3.5 h-3.5" />
          </button>

          {/* Maximize / Restore */}
          <button
            onClick={() => window.electronAPI?.maximize()}
            className="h-full px-3 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-600 dark:text-slate-300 flex items-center justify-center transition-colors"
            title="Maximize"
          >
            <Square className="w-3 h-3" />
          </button>

          {/* Close */}
          <button
            onClick={() => window.electronAPI?.close()}
            className="h-full px-3 hover:bg-red-600 hover:text-white text-slate-600 dark:text-slate-300 flex items-center justify-center transition-colors"
            title="Close"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>
    </header>
  );
};
