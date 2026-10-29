import React, { useState, useEffect } from 'react';
import type { PresentationData, SlideData, SlideElement } from '../types/office';
import {
  Plus,
  Play,
  Trash2,
  Copy,
  ChevronLeft,
  ChevronRight,
  Minimize2,
  Clock,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import confetti from 'canvas-confetti';

interface SlidesEditorProps {
  data: PresentationData;
  onChange: (updated: PresentationData) => void;
  darkMode: boolean;
  isPresenting: boolean;
  setIsPresenting: (val: boolean) => void;
}

export const SlidesEditor: React.FC<SlidesEditorProps> = ({
  data,
  onChange,
  darkMode,
  isPresenting,
  setIsPresenting,
}) => {
  const [activeSlideIndex, setActiveSlideIndex] = useState<number>(0);
  const [presenterTime, setPresenterTime] = useState<number>(0);
  const [selectedElementId, setSelectedElementId] = useState<string | null>(null);
  const [zoomLevel, setZoomLevel] = useState<number>(100);

  const activeSlide = data.slides[activeSlideIndex] || data.slides[0];

  // Timer when presenting
  useEffect(() => {
    let interval: any;
    if (isPresenting) {
      interval = setInterval(() => {
        setPresenterTime((prev) => prev + 1);
      }, 1000);
    } else {
      setPresenterTime(0);
    }
    return () => clearInterval(interval);
  }, [isPresenting]);

  // Keyboard navigation for presentation mode
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (isPresenting) {
        if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
          e.preventDefault();
          if (activeSlideIndex < data.slides.length - 1) {
            setActiveSlideIndex((prev) => prev + 1);
          } else {
            confetti({ particleCount: 80, spread: 70, origin: { y: 0.6 } });
          }
        } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
          e.preventDefault();
          if (activeSlideIndex > 0) {
            setActiveSlideIndex((prev) => prev - 1);
          }
        } else if (e.key === 'Escape') {
          setIsPresenting(false);
        }
      } else {
        if (e.key === 'F5') {
          e.preventDefault();
          setIsPresenting(true);
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isPresenting, activeSlideIndex, data.slides.length, setIsPresenting]);

  const updateActiveSlide = (updates: Partial<SlideData>) => {
    const newSlides = data.slides.map((s, idx) =>
      idx === activeSlideIndex ? { ...s, ...updates } : s
    );
    onChange({ ...data, slides: newSlides });
  };

  const updateElement = (elId: string, updates: Partial<SlideElement>) => {
    const newElements = activeSlide.elements.map((el) =>
      el.id === elId ? { ...el, ...updates } : el
    );
    updateActiveSlide({ elements: newElements });
  };

  const duplicateSlide = (idx: number) => {
    const target = data.slides[idx];
    const clone: SlideData = {
      ...target,
      id: `slide-${Date.now()}`,
      title: `${target.title} (Copy)`,
      elements: target.elements.map((el) => ({ ...el, id: `el-${Date.now()}-${Math.random()}` })),
    };
    const newSlides = [...data.slides];
    newSlides.splice(idx + 1, 0, clone);
    onChange({ ...data, slides: newSlides });
    setActiveSlideIndex(idx + 1);
  };

  const deleteSlide = (idx: number) => {
    if (data.slides.length <= 1) return;
    const newSlides = data.slides.filter((_, i) => i !== idx);
    onChange({ ...data, slides: newSlides });
    setActiveSlideIndex(Math.max(0, idx - 1));
  };

  const addElement = (type: 'heading' | 'shape' | 'metric' | 'bullet-list') => {
    const newEl: SlideElement = {
      id: `el-${Date.now()}`,
      type,
      content:
        type === 'heading'
          ? 'New Slide Title'
          : type === 'metric'
          ? '99.9%'
          : type === 'bullet-list'
          ? '• First insight point\n• Second milestone metric\n• Third priority deliverable'
          : 'Interactive shape card text description',
      subcontent: type === 'metric' ? 'Metric Description' : undefined,
      x: 10,
      y: 40,
      width: type === 'metric' ? 30 : 60,
      style: {
        fontSize: type === 'heading' ? 32 : 16,
        bold: true,
        bgColor: type === 'shape' ? '#1e293b' : type === 'metric' ? '#312e81' : undefined,
        color: '#ffffff',
        align: 'left',
      },
    };
    updateActiveSlide({ elements: [...activeSlide.elements, newEl] });
  };

  const formatTime = (secs: number) => {
    const m = Math.floor(secs / 60);
    const s = secs % 60;
    return `${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
  };

  // FULLSCREEN PRESENTER MODE OVERLAY
  if (isPresenting) {
    return (
      <div className="fixed inset-0 z-50 bg-black flex flex-col items-center justify-center text-white select-none">
        {/* Top Floating Presenter Bar */}
        <div className="absolute top-4 left-6 right-6 flex items-center justify-between text-xs text-slate-400 bg-slate-900/80 backdrop-blur-md px-4 py-2 rounded-full border border-slate-700/60">
          <div className="flex items-center space-x-2">
            <span className="w-2.5 h-2.5 rounded-full bg-orange-500 animate-pulse" />
            <span className="font-semibold text-slate-200">PowerPoint Presenter Mode</span>
            <span>•</span>
            <span>
              Slide {activeSlideIndex + 1} of {data.slides.length}
            </span>
          </div>

          <div className="flex items-center space-x-4">
            <div className="flex items-center space-x-1.5 font-mono text-amber-400">
              <Clock className="w-3.5 h-3.5" />
              <span>{formatTime(presenterTime)}</span>
            </div>
            <button
              onClick={() => setIsPresenting(false)}
              className="flex items-center space-x-1 px-3 py-1 rounded bg-slate-800 hover:bg-slate-700 text-white font-medium"
            >
              <Minimize2 className="w-3.5 h-3.5" />
              <span>End Show (Esc)</span>
            </button>
          </div>
        </div>

        {/* 16:9 Presentation Stage */}
        <div
          className="relative w-full max-w-6xl aspect-[16/9] shadow-2xl rounded-xl overflow-hidden border border-slate-800 p-12 flex flex-col justify-between"
          style={{ backgroundColor: activeSlide.bgColor || '#0f172a' }}
        >
          {/* Elements */}
          <div className="relative w-full h-full">
            {activeSlide.elements.map((el) => (
              <div
                key={el.id}
                className="absolute"
                style={{
                  left: `${el.x}%`,
                  top: `${el.y}%`,
                  width: `${el.width}%`,
                }}
              >
                {el.type === 'heading' && (
                  <div>
                    <h1
                      className="font-extrabold tracking-tight leading-tight"
                      style={{
                        fontSize: `${el.style?.fontSize || 36}px`,
                        color: el.style?.color || '#ffffff',
                        textAlign: el.style?.align || 'left',
                      }}
                    >
                      {el.content}
                    </h1>
                    {el.subcontent && (
                      <p className="text-xl text-slate-400 mt-2 font-normal">{el.subcontent}</p>
                    )}
                  </div>
                )}

                {el.type === 'metric' && (
                  <div
                    className="p-6 rounded-2xl border border-orange-500/40 shadow-lg text-center backdrop-blur-md"
                    style={{
                      backgroundColor: el.style?.bgColor || '#312e81',
                      color: el.style?.color || '#ffffff',
                    }}
                  >
                    <div className="text-4xl font-black tracking-tight">{el.content}</div>
                    {el.subcontent && (
                      <div className="text-xs uppercase tracking-wider text-slate-300 mt-2 font-medium">
                        {el.subcontent}
                      </div>
                    )}
                  </div>
                )}

                {el.type === 'shape' && (
                  <div
                    className="p-5 rounded-xl border border-slate-700/60 whitespace-pre-line text-sm leading-relaxed"
                    style={{
                      backgroundColor: el.style?.bgColor || '#1e293b',
                      color: el.style?.color || '#f1f5f9',
                    }}
                  >
                    {el.content}
                  </div>
                )}

                {el.type === 'bullet-list' && (
                  <div
                    className="whitespace-pre-line text-base leading-loose font-medium"
                    style={{ color: el.style?.color || '#cbd5e1' }}
                  >
                    {el.content}
                  </div>
                )}
              </div>
            ))}
          </div>

          <div className="flex justify-between items-center text-xs text-slate-500">
            <span>Anything Office Presentation Deck</span>
            <span>Slide {activeSlideIndex + 1}</span>
          </div>
        </div>

        {/* Bottom Floating Navigation Arrows */}
        <div className="absolute bottom-6 flex items-center space-x-3 bg-slate-900/90 backdrop-blur-md px-4 py-2 rounded-full border border-slate-700/60">
          <button
            onClick={() => setActiveSlideIndex((prev) => Math.max(0, prev - 1))}
            disabled={activeSlideIndex === 0}
            className="p-2 rounded-full hover:bg-slate-800 disabled:opacity-30"
          >
            <ChevronLeft className="w-5 h-5" />
          </button>
          <span className="text-xs text-slate-300 font-mono">
            {activeSlideIndex + 1} / {data.slides.length}
          </span>
          <button
            onClick={() =>
              setActiveSlideIndex((prev) => Math.min(data.slides.length - 1, prev + 1))
            }
            disabled={activeSlideIndex === data.slides.length - 1}
            className="p-2 rounded-full hover:bg-slate-800 disabled:opacity-30"
          >
            <ChevronRight className="w-5 h-5" />
          </button>
        </div>
      </div>
    );
  }

  // STANDARD POWERPOINT EDITOR VIEW
  return (
    <div className="flex flex-col h-full w-full select-none text-xs overflow-hidden">
      <div className="flex flex-1 overflow-hidden">
        {/* 1. LEFT POWERPOINT SLIDES THUMBNAILS PANEL */}
        <div
          className={`w-52 border-r flex flex-col justify-between shrink-0 transition-colors ${
            darkMode ? 'bg-slate-900 border-slate-800 text-slate-300' : 'bg-slate-50 border-slate-300 text-slate-700'
          }`}
        >
          <div className="p-2.5 overflow-y-auto space-y-2">
            <div className="flex items-center justify-between text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-1">
              <span>Slides ({data.slides.length})</span>
              <button
                onClick={() => addElement('heading')}
                className="text-orange-600 dark:text-orange-400 hover:underline flex items-center space-x-0.5"
              >
                <Plus className="w-3 h-3" />
                <span>New</span>
              </button>
            </div>

            {/* Slide Thumbnails List */}
            <div className="space-y-2">
              {data.slides.map((slide, idx) => {
                const isSelected = idx === activeSlideIndex;
                return (
                  <div
                    key={slide.id}
                    onClick={() => setActiveSlideIndex(idx)}
                    className="flex items-start space-x-1.5 cursor-pointer group"
                  >
                    {/* Slide Number */}
                    <span className="text-[10px] font-mono text-slate-400 pt-1 w-3 text-right">
                      {idx + 1}
                    </span>

                    {/* Thumbnail Card with Signature PowerPoint Orange Border when active */}
                    <div
                      className={`flex-1 rounded-sm p-1 transition-all border relative group/thumb ${
                        isSelected
                          ? 'border-[#c43e1c] ring-2 ring-[#c43e1c]/40 bg-white dark:bg-slate-800 shadow-sm'
                          : 'border-slate-300 dark:border-slate-700 hover:border-slate-400 bg-white dark:bg-slate-900'
                      }`}
                    >
                      <div
                        className="w-full aspect-[16/9] rounded-xs border border-slate-700/40 p-1.5 flex flex-col justify-center items-center text-center overflow-hidden"
                        style={{ backgroundColor: slide.bgColor || '#0f172a' }}
                      >
                        <span className="text-[9px] font-bold text-white line-clamp-2">
                          {slide.elements[0]?.content || slide.title}
                        </span>
                      </div>

                      {/* Quick Thumbnail Hover Actions */}
                      <div className="absolute top-1 right-1 hidden group-hover/thumb:flex items-center space-x-0.5 bg-black/80 rounded px-1 py-0.5">
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            duplicateSlide(idx);
                          }}
                          title="Duplicate Slide"
                          className="p-0.5 text-white hover:text-amber-400"
                        >
                          <Copy className="w-2.5 h-2.5" />
                        </button>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            deleteSlide(idx);
                          }}
                          title="Delete Slide"
                          className="p-0.5 text-white hover:text-rose-400"
                        >
                          <Trash2 className="w-2.5 h-2.5" />
                        </button>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>

        {/* 2. MAIN POWERPOINT CANVAS STAGE */}
        <div
          className={`flex-1 flex flex-col items-center justify-between p-6 overflow-y-auto transition-colors ${
            darkMode ? 'bg-[#1e1e1e]' : 'bg-[#e9e9e9]'
          }`}
        >
          {/* 16:9 Slide Canvas */}
          <div
            className="relative w-full max-w-4xl aspect-[16/9] shadow-2xl rounded-sm border border-slate-300 dark:border-neutral-700 p-8 flex flex-col justify-between overflow-hidden"
            style={{
              backgroundColor: activeSlide.bgColor || '#0f172a',
              transform: `scale(${zoomLevel / 100})`,
              transformOrigin: 'center center',
            }}
          >
            {/* Interactive Elements with 8-Point PowerPoint Selection Handles */}
            <div className="relative w-full h-full">
              {activeSlide.elements.map((el) => {
                const isSelected = selectedElementId === el.id;
                return (
                  <div
                    key={el.id}
                    onClick={(e) => {
                      e.stopPropagation();
                      setSelectedElementId(el.id);
                    }}
                    className={`absolute rounded-xs group transition-all ${
                      isSelected
                        ? 'ring-1 ring-white/80 border border-dashed border-white'
                        : 'hover:ring-1 hover:ring-white/40'
                    }`}
                    style={{
                      left: `${el.x}%`,
                      top: `${el.y}%`,
                      width: `${el.width}%`,
                    }}
                  >
                    {/* Real 8-Point PowerPoint Resize Handles */}
                    {isSelected && (
                      <>
                        <div className="absolute -top-1 -left-1 w-2 h-2 bg-white border border-slate-900 cursor-nwse-resize" />
                        <div className="absolute -top-1 left-1/2 -translate-x-1/2 w-2 h-2 bg-white border border-slate-900 cursor-ns-resize" />
                        <div className="absolute -top-1 -right-1 w-2 h-2 bg-white border border-slate-900 cursor-nesw-resize" />
                        <div className="absolute top-1/2 -left-1 -translate-y-1/2 w-2 h-2 bg-white border border-slate-900 cursor-ew-resize" />
                        <div className="absolute top-1/2 -right-1 -translate-y-1/2 w-2 h-2 bg-white border border-slate-900 cursor-ew-resize" />
                        <div className="absolute -bottom-1 -left-1 w-2 h-2 bg-white border border-slate-900 cursor-nesw-resize" />
                        <div className="absolute -bottom-1 left-1/2 -translate-x-1/2 w-2 h-2 bg-white border border-slate-900 cursor-ns-resize" />
                        <div className="absolute -bottom-1 -right-1 w-2 h-2 bg-white border border-slate-900 cursor-nwse-resize" />
                      </>
                    )}

                    {el.type === 'heading' && (
                      <div className="p-1">
                        <input
                          type="text"
                          value={el.content}
                          onChange={(e) => updateElement(el.id, { content: e.target.value })}
                          className="w-full font-extrabold tracking-tight bg-transparent focus:outline-none"
                          style={{
                            fontSize: `${el.style?.fontSize || 32}px`,
                            color: el.style?.color || '#ffffff',
                            textAlign: el.style?.align || 'left',
                          }}
                        />
                        {el.subcontent !== undefined && (
                          <input
                            type="text"
                            value={el.subcontent}
                            onChange={(e) => updateElement(el.id, { subcontent: e.target.value })}
                            className="w-full text-base text-slate-400 bg-transparent focus:outline-none mt-1"
                          />
                        )}
                      </div>
                    )}

                    {el.type === 'metric' && (
                      <div
                        className="p-4 rounded-xl border border-orange-500/30 text-center shadow"
                        style={{
                          backgroundColor: el.style?.bgColor || '#312e81',
                          color: el.style?.color || '#ffffff',
                        }}
                      >
                        <input
                          type="text"
                          value={el.content}
                          onChange={(e) => updateElement(el.id, { content: e.target.value })}
                          className="w-full text-2xl font-black text-center bg-transparent focus:outline-none"
                        />
                        <input
                          type="text"
                          value={el.subcontent || ''}
                          onChange={(e) => updateElement(el.id, { subcontent: e.target.value })}
                          className="w-full text-[11px] uppercase tracking-wider text-center text-slate-300 bg-transparent focus:outline-none mt-1"
                        />
                      </div>
                    )}

                    {el.type === 'shape' && (
                      <div
                        className="p-4 rounded-md border border-slate-700/60"
                        style={{
                          backgroundColor: el.style?.bgColor || '#1e293b',
                          color: el.style?.color || '#f1f5f9',
                        }}
                      >
                        <textarea
                          rows={3}
                          value={el.content}
                          onChange={(e) => updateElement(el.id, { content: e.target.value })}
                          className="w-full text-xs leading-relaxed bg-transparent focus:outline-none resize-none"
                        />
                      </div>
                    )}

                    {el.type === 'bullet-list' && (
                      <div className="p-2">
                        <textarea
                          rows={4}
                          value={el.content}
                          onChange={(e) => updateElement(el.id, { content: e.target.value })}
                          className="w-full text-sm leading-relaxed bg-transparent focus:outline-none resize-none"
                          style={{ color: el.style?.color || '#cbd5e1' }}
                        />
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            {/* Slide Footer */}
            <div className="flex justify-between items-center text-[10px] text-slate-500 border-t border-slate-800 pt-2 select-none">
              <span>Slide {activeSlideIndex + 1} of {data.slides.length}</span>
              <span>16:9 Widescreen</span>
            </div>
          </div>

          {/* Speaker Notes Drawer */}
          <div className="w-full max-w-4xl mt-3 p-2 rounded-sm border border-slate-300 dark:border-neutral-700 bg-white dark:bg-slate-900">
            <div className="text-[9px] font-bold uppercase tracking-wider text-slate-400 mb-0.5">
              Click to add speaker notes
            </div>
            <textarea
              rows={1}
              value={activeSlide.notes || ''}
              onChange={(e) => updateActiveSlide({ notes: e.target.value })}
              placeholder="Private presenter talking points..."
              className="w-full text-xs bg-transparent focus:outline-none text-slate-600 dark:text-slate-300 resize-none"
            />
          </div>
        </div>
      </div>

      {/* 3. AUTHENTIC MICROSOFT POWERPOINT STATUS BAR */}
      <footer
        className={`h-6 border-t select-none px-3 flex items-center justify-between text-[11px] shrink-0 transition-colors ${
          darkMode ? 'bg-slate-900 border-slate-800 text-slate-400' : 'bg-[#c43e1c] text-white border-orange-800'
        }`}
      >
        <div className="flex items-center space-x-3">
          <span>Slide {activeSlideIndex + 1} of {data.slides.length}</span>
          <span>•</span>
          <span>Office Theme</span>
          <span>•</span>
          <button className="hover:underline">Notes</button>
        </div>

        <div className="flex items-center space-x-4">
          <button
            onClick={() => setIsPresenting(true)}
            className="flex items-center space-x-1 hover:bg-white/20 px-1.5 py-0.5 rounded"
            title="Slide Show (F5)"
          >
            <Play className="w-3 h-3 fill-current" />
            <span>Slide Show</span>
          </button>

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
