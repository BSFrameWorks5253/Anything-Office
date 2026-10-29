import React from 'react';
import type { AnythingProject } from '../types/office';
import { DEFAULT_PROJECT } from '../engine/sampleProjects';
import { X, FileText, Table as TableIcon, Presentation, Sparkles, FolderPlus } from 'lucide-react';

interface TemplateGalleryModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSelectTemplate: (template: AnythingProject) => void;
  darkMode: boolean;
}

export const TemplateGalleryModal: React.FC<TemplateGalleryModalProps> = ({
  isOpen,
  onClose,
  onSelectTemplate,
  darkMode,
}) => {
  if (!isOpen) return null;

  const templates: {
    title: string;
    description: string;
    type: 'all' | 'docs' | 'sheets' | 'slides';
    badge: string;
    action: () => AnythingProject;
  }[] = [
    {
      title: 'Complete Corporate Strategy Suite',
      description:
        'All-in-one suite loaded with Strategic Brief (Docs), FY26 Financial Model (Sheets), and Milestones Deck (Slides).',
      type: 'all',
      badge: 'All-in-One Suite',
      action: () => ({ ...DEFAULT_PROJECT, id: `proj-${Date.now()}` }),
    },
    {
      title: 'Quarterly Financial & KPI Model',
      description:
        'Spreadsheet model pre-configured with revenue channels, OPEX deductions, and SUM/AVERAGE/IF formulas.',
      type: 'sheets',
      badge: 'Sheets Focused',
      action: () => ({
        ...DEFAULT_PROJECT,
        id: `proj-${Date.now()}`,
        name: 'Financial & KPI Model',
      }),
    },
    {
      title: 'Investor Pitch & Product Deck',
      description:
        'High-impact 16:9 presentation slides with metric callouts, feature cards, and execution roadmap.',
      type: 'slides',
      badge: 'Slides Focused',
      action: () => ({
        ...DEFAULT_PROJECT,
        id: `proj-${Date.now()}`,
        name: 'Investor Pitch Deck',
      }),
    },
    {
      title: 'Blank Document & Clean Slate',
      description: 'Start with empty clean canvases for Docs, Sheets, and Slides.',
      type: 'all',
      badge: 'Clean Workspace',
      action: () => ({
        id: `proj-${Date.now()}`,
        name: 'Untitled Workspace',
        version: '1.0.0',
        lastModified: Date.now(),
        docs: {
          title: 'Untitled Document',
          pageSize: 'A4',
          margins: { top: 20, right: 24, bottom: 20, left: 24 },
          paragraphs: [
            { id: 'p-1', type: 'h1', content: 'Untitled Document', align: 'left' },
            { id: 'p-2', type: 'p', content: 'Start typing your document text here...', align: 'left' },
          ],
        },
        sheets: {
          activeSheetId: 'sheet-1',
          sheets: [
            {
              id: 'sheet-1',
              name: 'Sheet 1',
              rowCount: 20,
              colCount: 10,
              cells: {},
            },
          ],
        },
        slides: {
          theme: 'modern-dark',
          aspectRatio: '16:9',
          slides: [
            {
              id: 'slide-1',
              title: 'Welcome Slide',
              layout: 'title',
              bgColor: '#0f172a',
              elements: [
                {
                  id: 'el-1',
                  type: 'heading',
                  content: 'New Presentation',
                  subcontent: 'Click here to edit subtitle',
                  x: 10,
                  y: 35,
                  width: 80,
                  style: { fontSize: 44, bold: true, color: '#ffffff', align: 'center' },
                },
              ],
            },
          ],
        },
      }),
    },
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm select-none animate-fadeIn">
      <div
        className={`w-full max-w-3xl rounded-2xl shadow-2xl border flex flex-col overflow-hidden transition-all ${
          darkMode ? 'bg-slate-900 border-slate-700 text-slate-100' : 'bg-white border-slate-200 text-slate-900'
        }`}
      >
        <div className="px-6 py-4 border-b border-inherit flex items-center justify-between bg-slate-50/50 dark:bg-slate-800/40">
          <div className="flex items-center space-x-2.5">
            <div className="p-2 rounded-lg bg-indigo-600 text-white">
              <Sparkles className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-base font-bold">Template Gallery</h2>
              <p className="text-xs text-slate-500 dark:text-slate-400">
                Choose a pre-built template to accelerate your document, sheet, or slide workflow
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

        <div className="p-6 grid grid-cols-2 gap-4">
          {templates.map((tpl, idx) => (
            <div
              key={idx}
              onClick={() => {
                onSelectTemplate(tpl.action());
                onClose();
              }}
              className={`p-4 rounded-xl border cursor-pointer hover:scale-[1.02] transition-all flex flex-col justify-between ${
                darkMode
                  ? 'border-slate-700 bg-slate-800/60 hover:border-indigo-500 hover:bg-slate-800'
                  : 'border-slate-200 bg-slate-50 hover:border-indigo-500 hover:bg-white'
              }`}
            >
              <div>
                <div className="flex items-center justify-between mb-2">
                  <span className="text-[10px] font-mono px-2 py-0.5 rounded bg-indigo-100 dark:bg-indigo-950 text-indigo-700 dark:text-indigo-300 font-bold">
                    {tpl.badge}
                  </span>
                  {tpl.type === 'docs' && <FileText className="w-4 h-4 text-blue-500" />}
                  {tpl.type === 'sheets' && <TableIcon className="w-4 h-4 text-emerald-500" />}
                  {tpl.type === 'slides' && <Presentation className="w-4 h-4 text-amber-500" />}
                  {tpl.type === 'all' && <FolderPlus className="w-4 h-4 text-purple-500" />}
                </div>
                <h3 className="font-bold text-sm mb-1">{tpl.title}</h3>
                <p className="text-xs text-slate-500 dark:text-slate-400 leading-relaxed">
                  {tpl.description}
                </p>
              </div>

              <div className="mt-4 pt-3 border-t border-inherit flex justify-end">
                <span className="text-xs font-semibold text-indigo-600 dark:text-indigo-400 hover:underline">
                  Load Template →
                </span>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};
