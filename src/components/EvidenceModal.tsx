import React, { useState } from 'react';
import { OFFICE_EVIDENCE } from '../engine/evidenceData';
import type { EvidenceSection } from '../engine/evidenceData';
import { AdobeLogo } from './AdobeLogo';
import {
  X,
  FileCode,
  Layers,
  Cpu,
  CheckCircle2,
  FileText,
  Table as TableIcon,
  Presentation,
  ShieldCheck,
  Code2,
  Terminal,
  Sparkles,
} from 'lucide-react';

interface EvidenceModalProps {
  isOpen: boolean;
  onClose: () => void;
  darkMode: boolean;
}

export const EvidenceModal: React.FC<EvidenceModalProps> = ({ isOpen, onClose, darkMode }) => {
  const [activeTab, setActiveTab] = useState<'word' | 'excel' | 'powerpoint' | 'unified' | 'brand'>('word');

  if (!isOpen) return null;

  const currentSection: EvidenceSection =
    activeTab !== 'brand' ? OFFICE_EVIDENCE[activeTab] : OFFICE_EVIDENCE.unified;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm select-none animate-fadeIn">
      <div
        className={`w-full max-w-5xl h-[88vh] rounded-2xl shadow-2xl border flex flex-col overflow-hidden transition-all ${
          darkMode ? 'bg-slate-900 border-slate-700 text-slate-100' : 'bg-white border-slate-200 text-slate-900'
        }`}
      >
        {/* Header */}
        <div className="px-6 py-4 border-b border-inherit flex items-center justify-between bg-slate-50/50 dark:bg-slate-800/40">
          <div className="flex items-center space-x-3">
            <div className="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center text-white">
              <ShieldCheck className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <h2 className="text-base font-bold">Office Architecture & Technical Evidence</h2>
                <span className="text-[11px] font-mono px-2 py-0.5 rounded-full bg-indigo-100 dark:bg-indigo-950 text-indigo-700 dark:text-indigo-300 border border-indigo-300 dark:border-indigo-800">
                  ECMA-376 & ISO/IEC 29500 Verified
                </span>
              </div>
              <p className="text-xs text-slate-500 dark:text-slate-400">
                Reverse-engineering analysis of Microsoft Word, Excel, PowerPoint & the Anything Unified Engine
              </p>
            </div>
          </div>

          <button
            onClick={onClose}
            className="p-1.5 rounded-lg hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-500 hover:text-slate-800 dark:hover:text-slate-200"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Tab Navigation */}
        <div className="flex border-b border-inherit bg-slate-100/60 dark:bg-slate-900 px-6 pt-2 space-x-2 text-xs">
          <button
            onClick={() => setActiveTab('word')}
            className={`flex items-center space-x-2 px-4 py-2.5 rounded-t-lg font-medium border-t-2 transition-all ${
              activeTab === 'word'
                ? 'border-blue-600 bg-white dark:bg-slate-800 text-blue-600 dark:text-blue-400 font-bold shadow-sm'
                : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-300'
            }`}
          >
            <FileText className="w-4 h-4 text-blue-500" />
            <span>Microsoft Word (.docx)</span>
          </button>

          <button
            onClick={() => setActiveTab('excel')}
            className={`flex items-center space-x-2 px-4 py-2.5 rounded-t-lg font-medium border-t-2 transition-all ${
              activeTab === 'excel'
                ? 'border-emerald-600 bg-white dark:bg-slate-800 text-emerald-600 dark:text-emerald-400 font-bold shadow-sm'
                : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-300'
            }`}
          >
            <TableIcon className="w-4 h-4 text-emerald-500" />
            <span>Microsoft Excel (.xlsx)</span>
          </button>

          <button
            onClick={() => setActiveTab('powerpoint')}
            className={`flex items-center space-x-2 px-4 py-2.5 rounded-t-lg font-medium border-t-2 transition-all ${
              activeTab === 'powerpoint'
                ? 'border-amber-600 bg-white dark:bg-slate-800 text-amber-600 dark:text-amber-400 font-bold shadow-sm'
                : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-300'
            }`}
          >
            <Presentation className="w-4 h-4 text-amber-500" />
            <span>Microsoft PowerPoint (.pptx)</span>
          </button>

          <button
            onClick={() => setActiveTab('unified')}
            className={`flex items-center space-x-2 px-4 py-2.5 rounded-t-lg font-medium border-t-2 transition-all ${
              activeTab === 'unified'
                ? 'border-purple-600 bg-white dark:bg-slate-800 text-purple-600 dark:text-purple-400 font-bold shadow-sm'
                : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-300'
            }`}
          >
            <Layers className="w-4 h-4 text-purple-500" />
            <span>Anything Unified Suite</span>
          </button>

          <button
            onClick={() => setActiveTab('brand')}
            className={`flex items-center space-x-2 px-4 py-2.5 rounded-t-lg font-medium border-t-2 transition-all ${
              activeTab === 'brand'
                ? 'border-cyan-500 bg-white dark:bg-slate-800 text-cyan-600 dark:text-cyan-400 font-bold shadow-sm'
                : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-300'
            }`}
          >
            <Sparkles className="w-4 h-4 text-cyan-500" />
            <span>App Identity & Adobe Logo</span>
          </button>
        </div>

        {/* Content Body */}
        <div className="flex-1 overflow-y-auto p-6 space-y-6">
          {activeTab === 'brand' ? (
            <div className="space-y-6">
              {/* Brand Summary Banner */}
              <div className="p-5 rounded-2xl border border-cyan-500/30 bg-gradient-to-r from-slate-900 to-slate-950 text-white flex items-center justify-between shadow-xl">
                <div className="flex items-center space-x-5">
                  <AdobeLogo size="xl" showLabel={false} monogram="Om" />
                  <div>
                    <div className="flex items-center space-x-2">
                      <h3 className="text-xl font-black tracking-tight">OmniSuite</h3>
                      <span className="text-xs px-2 py-0.5 rounded-full bg-cyan-500/20 text-cyan-300 font-mono">
                        by Anything
                      </span>
                    </div>
                    <p className="text-xs text-slate-300 max-w-lg mt-1 leading-relaxed">
                      Named <strong>OmniSuite</strong> (from Latin <em>Omnis</em>, meaning "all" or "universal") to signify a single unified app for everything across documents, spreadsheets, presentations, and cross-format compatibility.
                    </p>
                  </div>
                </div>

                <div className="flex flex-col items-end text-xs font-mono text-cyan-400 space-y-1">
                  <span>Adobe Creative Cloud Inspired</span>
                  <span className="text-slate-400">Two-Letter Squircle Paradigm</span>
                </div>
              </div>

              {/* Logo Sizing & Variations */}
              <div>
                <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-3 flex items-center space-x-1.5">
                  <Sparkles className="w-4 h-4 text-cyan-500" />
                  <span>Adobe Design Language System & Scales</span>
                </h4>
                <div className="grid grid-cols-4 gap-4">
                  <div className="p-4 rounded-xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-800/80 flex flex-col items-center justify-center text-center space-y-3">
                    <AdobeLogo size="xl" monogram="Om" />
                    <div>
                      <div className="font-bold text-xs">Extra Large (80px)</div>
                      <div className="text-[10px] text-slate-400">Desktop Icon / Splash Screen</div>
                    </div>
                  </div>
                  <div className="p-4 rounded-xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-800/80 flex flex-col items-center justify-center text-center space-y-3">
                    <AdobeLogo size="lg" monogram="Om" />
                    <div>
                      <div className="font-bold text-xs">Large (44px)</div>
                      <div className="text-[10px] text-slate-400">App Switcher / Modal Head</div>
                    </div>
                  </div>
                  <div className="p-4 rounded-xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-800/80 flex flex-col items-center justify-center text-center space-y-3">
                    <AdobeLogo size="md" monogram="Om" />
                    <div>
                      <div className="font-bold text-xs">Medium (32px)</div>
                      <div className="text-[10px] text-slate-400">Ribbon Toolbar Brand Mark</div>
                    </div>
                  </div>
                  <div className="p-4 rounded-xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-800/80 flex flex-col items-center justify-center text-center space-y-3">
                    <AdobeLogo size="sm" monogram="Om" />
                    <div>
                      <div className="font-bold text-xs">Small (24px)</div>
                      <div className="text-[10px] text-slate-400">Window Titlebar / Tab Icon</div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Anatomy Breakdown */}
              <div>
                <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-3 flex items-center space-x-1.5">
                  <Cpu className="w-4 h-4 text-emerald-500" />
                  <span>Adobe App Grid Comparison (Ps, Ai, Id vs Om)</span>
                </h4>
                <div className="grid grid-cols-4 gap-3 text-xs">
                  <div className="p-3.5 rounded-xl border border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-900 flex flex-col items-center text-center space-y-2">
                    <AdobeLogo size="lg" monogram="Ps" theme="word" />
                    <div>
                      <div className="font-bold text-xs">Photoshop</div>
                      <div className="text-[10px] text-slate-400">Deep Blue + Cyan</div>
                    </div>
                  </div>
                  <div className="p-3.5 rounded-xl border border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-900 flex flex-col items-center text-center space-y-2">
                    <AdobeLogo size="lg" monogram="Ai" theme="powerpoint" />
                    <div>
                      <div className="font-bold text-xs">Illustrator</div>
                      <div className="text-[10px] text-slate-400">Deep Umber + Amber</div>
                    </div>
                  </div>
                  <div className="p-3.5 rounded-xl border border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-900 flex flex-col items-center text-center space-y-2">
                    <AdobeLogo size="lg" monogram="Id" theme="excel" />
                    <div>
                      <div className="font-bold text-xs">InDesign</div>
                      <div className="text-[10px] text-slate-400">Deep Green + Accent</div>
                    </div>
                  </div>
                  <div className="p-3.5 rounded-xl border-2 border-cyan-500/60 bg-cyan-950/20 flex flex-col items-center text-center space-y-2 shadow-md">
                    <AdobeLogo size="lg" monogram="Ao" theme="anything" />
                    <div>
                      <div className="font-bold text-xs text-cyan-400">Anything Office</div>
                      <div className="text-[10px] text-cyan-300/80">Deep Obsidian + Cyan</div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          ) : (
            <>
              {/* Section Summary */}
              <div className="p-4 rounded-xl border border-indigo-200 dark:border-indigo-900 bg-indigo-50/50 dark:bg-indigo-950/30">
                <div className="flex items-center justify-between mb-2">
                  <h3 className="text-sm font-bold text-indigo-950 dark:text-indigo-200">
                    {currentSection.title}
                  </h3>
                  <span className="text-[11px] font-mono px-2 py-0.5 rounded bg-indigo-200 dark:bg-indigo-900 text-indigo-800 dark:text-indigo-200">
                    {currentSection.badge}
                  </span>
                </div>
                <p className="text-xs leading-relaxed text-slate-700 dark:text-slate-300">
                  {currentSection.summary}
                </p>
              </div>

              {/* Specifications Table */}
              <div>
                <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-2 flex items-center space-x-1.5">
                  <FileCode className="w-4 h-4 text-blue-500" />
                  <span>Standard Formal Specifications</span>
                </h4>
                <div className="grid grid-cols-2 gap-3 text-xs">
                  <div className="p-3 rounded-lg border border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-900">
                    <span className="text-slate-400 block text-[10px]">INDUSTRY STANDARD</span>
                    <span className="font-semibold text-slate-800 dark:text-slate-200">
                      {currentSection.specs.standard}
                    </span>
                  </div>
                  <div className="p-3 rounded-lg border border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-900">
                    <span className="text-slate-400 block text-[10px]">ISO RATIFIED STANDARD</span>
                    <span className="font-semibold text-slate-800 dark:text-slate-200">
                      {currentSection.specs.isoStandard}
                    </span>
                  </div>
                  <div className="p-3 rounded-lg border border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-900">
                    <span className="text-slate-400 block text-[10px]">PHYSICAL PACKAGING</span>
                    <span className="font-semibold text-slate-800 dark:text-slate-200">
                      {currentSection.specs.containerFormat}
                    </span>
                  </div>
                  <div className="p-3 rounded-lg border border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-900">
                    <span className="text-slate-400 block text-[10px]">XML SCHEMA NAMESPACE</span>
                    <span className="font-mono text-[11px] text-indigo-600 dark:text-indigo-400 truncate block">
                      {currentSection.specs.coreSchemaNamespace}
                    </span>
                  </div>
                </div>
              </div>

              {/* Core Engine Mechanics */}
              <div>
                <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-3 flex items-center space-x-1.5">
                  <Cpu className="w-4 h-4 text-emerald-500" />
                  <span>Internal Runtime Engines & Algorithms</span>
                </h4>
                <div className="grid grid-cols-2 gap-4">
                  {currentSection.engineDetails.map((eng: any, idx: number) => (
                    <div
                      key={idx}
                      className="p-4 rounded-xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-800/80 shadow-sm"
                    >
                      <div className="flex items-center justify-between mb-1">
                        <span className="font-bold text-xs text-slate-800 dark:text-slate-100">
                          {eng.title}
                        </span>
                        <span className="text-[10px] font-mono px-2 py-0.5 rounded bg-emerald-100 dark:bg-emerald-950/60 text-emerald-700 dark:text-emerald-300">
                          {eng.algorithmicComplexity}
                        </span>
                      </div>
                      <div className="text-[11px] font-medium text-indigo-600 dark:text-indigo-400 mb-2">
                        Mechanism: {eng.mechanism}
                      </div>
                      <p className="text-xs leading-relaxed text-slate-600 dark:text-slate-300">
                        {eng.description}
                      </p>
                    </div>
                  ))}
                </div>
              </div>

              {/* Actual XML Structure & Physical Evidence */}
              <div>
                <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-3 flex items-center space-x-1.5">
                  <Code2 className="w-4 h-4 text-purple-500" />
                  <span>OPC Zip Container File Layout & Schemas</span>
                </h4>
                <div className="space-y-3">
                  {currentSection.fileStructure.map((file: any, idx: number) => (
                    <div
                      key={idx}
                      className="rounded-xl border border-slate-300 dark:border-slate-800 overflow-hidden bg-slate-900 text-slate-100 text-xs"
                    >
                      <div className="px-4 py-2 bg-slate-800/90 border-b border-slate-700 flex items-center justify-between">
                        <div className="flex items-center space-x-2 font-mono text-[11px] text-amber-300">
                          <Terminal className="w-3.5 h-3.5" />
                          <span>{file.name}</span>
                        </div>
                        <span className="text-[10px] text-slate-400">{file.role}</span>
                      </div>
                      <pre className="p-4 font-mono text-[11px] text-slate-300 overflow-x-auto leading-relaxed">
                        <code>{file.contentPreview}</code>
                      </pre>
                    </div>
                  ))}
                </div>
              </div>

              {/* Architectural Comparison Matrix */}
              <div>
                <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-3 flex items-center space-x-1.5">
                  <CheckCircle2 className="w-4 h-4 text-emerald-500" />
                  <span>Comparative Architecture: Legacy Microsoft Office vs Anything Office</span>
                </h4>
                <div className="border border-slate-200 dark:border-slate-800 rounded-xl overflow-hidden">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="bg-slate-100 dark:bg-slate-800 border-b border-inherit text-left">
                        <th className="p-3 font-bold w-1/4">Capability / Layer</th>
                        <th className="p-3 font-bold w-3/8 text-slate-600 dark:text-slate-400">
                          Microsoft Office (Siloed Architecture)
                        </th>
                        <th className="p-3 font-bold w-3/8 text-indigo-600 dark:text-indigo-400">
                          Anything Office (Unified Engine)
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-inherit">
                      {currentSection.comparisons.map((c: any, idx: number) => (
                        <tr key={idx} className="hover:bg-slate-50 dark:hover:bg-slate-850">
                          <td className="p-3 font-semibold text-slate-700 dark:text-slate-300">
                            {c.feature}
                          </td>
                          <td className="p-3 text-slate-500 dark:text-slate-400">{c.msOffice}</td>
                          <td className="p-3 text-indigo-700 dark:text-indigo-300 font-medium">
                            {c.anythingApp}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-3 border-t border-inherit bg-slate-50/50 dark:bg-slate-900 flex justify-between items-center text-xs text-slate-500">
          <span>Anything Office Architecture Engine • 2026</span>
          <button
            onClick={onClose}
            className="px-4 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white font-medium"
          >
            Close Evidence Viewer
          </button>
        </div>
      </div>
    </div>
  );
};
