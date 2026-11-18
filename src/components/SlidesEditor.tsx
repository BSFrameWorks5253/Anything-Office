import React, { useState, useEffect, useRef } from 'react';
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
  ArrowUp,
  ArrowDown,
  Type,
  Square,
  TrendingUp,
  Image as ImageIcon,
  FileText,
  AlignLeft,
  AlignCenter,
  AlignRight,
  Palette,
  Sparkles,
  Workflow,
  Sliders,
  Check,
  X,
  Move,
  ArrowUpToLine,
  ArrowDownToLine,
} from 'lucide-react';
import confetti from 'canvas-confetti';

export const SLIDE_THEMES: {
  id: PresentationData['theme'];
  name: string;
  bgColor: string;
  cardBg: string;
  accent: string;
  textColor: string;
}[] = [
  { id: 'modern-dark', name: 'Modern Dark', bgColor: '#0f172a', cardBg: '#1e293b', accent: '#38bdf8', textColor: '#ffffff' },
  { id: 'corporate-blue', name: 'Corporate Blue', bgColor: '#03254c', cardBg: '#0f3c68', accent: '#60a5fa', textColor: '#ffffff' },
  { id: 'emerald', name: 'Emerald Growth', bgColor: '#064e3b', cardBg: '#065f46', accent: '#34d399', textColor: '#ffffff' },
  { id: 'sunset', name: 'Sunset Warm', bgColor: '#450a0a', cardBg: '#7f1d1d', accent: '#fb923c', textColor: '#ffffff' },
  { id: 'slate', name: 'Minimal Light', bgColor: '#f8fafc', cardBg: '#ffffff', accent: '#2563eb', textColor: '#0f172a' },
  { id: 'cyber', name: 'Cyber Neon', bgColor: '#18181b', cardBg: '#27272a', accent: '#c084fc', textColor: '#ffffff' },
];

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
  const [showNotesInPresentation, setShowNotesInPresentation] = useState<boolean>(false);
  const [laserPointerActive, setLaserPointerActive] = useState<boolean>(false);
  const [laserPos, setLaserPos] = useState<{ x: number; y: number }>({ x: 0, y: 0 });
  const [showSmartArtModal, setShowSmartArtModal] = useState<boolean>(false);
  const [showThemeModal, setShowThemeModal] = useState<boolean>(false);

  // Mouse Drag-to-Move and Drag-to-Resize State
  const canvasRef = useRef<HTMLDivElement>(null);
  const [draggingElementId, setDraggingElementId] = useState<string | null>(null);
  const [resizingElementId, setResizingElementId] = useState<string | null>(null);
  const [dragStartPos, setDragStartPos] = useState<{ mouseX: number; mouseY: number; elX: number; elY: number }>({
    mouseX: 0,
    mouseY: 0,
    elX: 0,
    elY: 0,
  });
  const [resizeStartPos, setResizeStartPos] = useState<{ mouseX: number; mouseY: number; startWidth: number }>({
    mouseX: 0,
    mouseY: 0,
    startWidth: 0,
  });

  const activeSlide = data.slides[activeSlideIndex] || data.slides[0];

  // Z-Ordering functions
  const bringToFront = (elId: string) => {
    const idx = activeSlide.elements.findIndex((e) => e.id === elId);
    if (idx === -1 || idx === activeSlide.elements.length - 1) return;
    const newElements = [...activeSlide.elements];
    const [item] = newElements.splice(idx, 1);
    newElements.push(item);
    updateActiveSlide({ elements: newElements });
  };

  const sendToBack = (elId: string) => {
    const idx = activeSlide.elements.findIndex((e) => e.id === elId);
    if (idx <= 0) return;
    const newElements = [...activeSlide.elements];
    const [item] = newElements.splice(idx, 1);
    newElements.unshift(item);
    updateActiveSlide({ elements: newElements });
  };

  const bringForward = (elId: string) => {
    const idx = activeSlide.elements.findIndex((e) => e.id === elId);
    if (idx === -1 || idx === activeSlide.elements.length - 1) return;
    const newElements = [...activeSlide.elements];
    const temp = newElements[idx];
    newElements[idx] = newElements[idx + 1];
    newElements[idx + 1] = temp;
    updateActiveSlide({ elements: newElements });
  };

  const sendBackward = (elId: string) => {
    const idx = activeSlide.elements.findIndex((e) => e.id === elId);
    if (idx <= 0) return;
    const newElements = [...activeSlide.elements];
    const temp = newElements[idx];
    newElements[idx] = newElements[idx - 1];
    newElements[idx - 1] = temp;
    updateActiveSlide({ elements: newElements });
  };

  // Drag handlers
  const startDrag = (elId: string, e: React.MouseEvent) => {
    e.stopPropagation();
    setSelectedElementId(elId);
    const target = activeSlide.elements.find((el) => el.id === elId);
    if (!target) return;
    setDraggingElementId(elId);
    setDragStartPos({
      mouseX: e.clientX,
      mouseY: e.clientY,
      elX: target.x,
      elY: target.y,
    });
  };

  const startResize = (elId: string, e: React.MouseEvent) => {
    e.stopPropagation();
    const target = activeSlide.elements.find((el) => el.id === elId);
    if (!target) return;
    setResizingElementId(elId);
    setResizeStartPos({
      mouseX: e.clientX,
      mouseY: e.clientY,
      startWidth: target.width,
    });
  };

  // Global mousemove and mouseup listeners for drag/resize
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (!canvasRef.current) return;
      const rect = canvasRef.current.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) return;

      if (draggingElementId) {
        const deltaX = ((e.clientX - dragStartPos.mouseX) / rect.width) * 100;
        const deltaY = ((e.clientY - dragStartPos.mouseY) / rect.height) * 100;
        const newX = Math.max(0, Math.min(90, Math.round(dragStartPos.elX + deltaX)));
        const newY = Math.max(0, Math.min(90, Math.round(dragStartPos.elY + deltaY)));
        updateElement(draggingElementId, { x: newX, y: newY });
      } else if (resizingElementId) {
        const deltaW = ((e.clientX - resizeStartPos.mouseX) / rect.width) * 100;
        const newWidth = Math.max(10, Math.min(95, Math.round(resizeStartPos.startWidth + deltaW)));
        updateElement(resizingElementId, { width: newWidth });
      }
    };

    const handleMouseUp = () => {
      setDraggingElementId(null);
      setResizingElementId(null);
    };

    if (draggingElementId || resizingElementId) {
      window.addEventListener('mousemove', handleMouseMove);
      window.addEventListener('mouseup', handleMouseUp);
    }
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [draggingElementId, resizingElementId, dragStartPos, resizeStartPos]);

  const applyTheme = (themeId: PresentationData['theme']) => {
    const themeObj = SLIDE_THEMES.find((t) => t.id === themeId);
    if (!themeObj) return;

    const newSlides = data.slides.map((s) => ({
      ...s,
      bgColor: themeObj.bgColor,
    }));

    onChange({
      ...data,
      theme: themeId,
      slides: newSlides,
    });
    setShowThemeModal(false);
  };

  const setSlideTransition = (transition: SlideData['transition'], applyAll: boolean = false) => {
    if (applyAll) {
      const newSlides = data.slides.map((s) => ({ ...s, transition }));
      onChange({ ...data, slides: newSlides });
    } else {
      updateActiveSlide({ transition });
    }
  };

  const toggleAspectRatio = () => {
    const nextRatio = data.aspectRatio === '4:3' ? '16:9' : '4:3';
    onChange({ ...data, aspectRatio: nextRatio });
  };

  const insertSmartArt = (type: 'process' | 'hierarchy' | 'cycle' | 'matrix' | 'funnel') => {
    const timestamp = Date.now();
    const currentTheme = SLIDE_THEMES.find((t) => t.id === data.theme) || SLIDE_THEMES[0];
    let newElements: SlideElement[] = [];

    if (type === 'process') {
      newElements = [
        {
          id: `el-${timestamp}-h`,
          type: 'heading',
          content: 'Strategic Delivery Process Flow',
          subcontent: 'Sequential execution milestones & phase gate delivery',
          x: 10,
          y: 12,
          width: 80,
          style: { fontSize: 28, bold: true, color: currentTheme.textColor, align: 'left' },
        },
        {
          id: `el-${timestamp}-p1`,
          type: 'shape',
          content: '01. Inception\n• Requirements mapping\n• System specifications\n• Risk mitigation',
          x: 6,
          y: 36,
          width: 20,
          style: { bgColor: currentTheme.cardBg, color: currentTheme.textColor, borderColor: currentTheme.accent, borderWidth: 2 },
        },
        {
          id: `el-${timestamp}-p2`,
          type: 'shape',
          content: '02. Architecture\n• ECMA-376 OOXML Core\n• Piece Table Buffer\n• Reactive DAG Engine',
          x: 29,
          y: 36,
          width: 20,
          style: { bgColor: currentTheme.cardBg, color: currentTheme.textColor, borderColor: currentTheme.accent, borderWidth: 2 },
        },
        {
          id: `el-${timestamp}-p3`,
          type: 'shape',
          content: '03. Acceleration\n• Sub-second native boot\n• Zero-flash render\n• Desktop distribution',
          x: 52,
          y: 36,
          width: 20,
          style: { bgColor: currentTheme.cardBg, color: currentTheme.textColor, borderColor: currentTheme.accent, borderWidth: 2 },
        },
        {
          id: `el-${timestamp}-p4`,
          type: 'shape',
          content: '04. Deployment\n• Enterprise rollout\n• Sovereign offline mode\n• 100% parity verified',
          x: 75,
          y: 36,
          width: 20,
          style: { bgColor: currentTheme.cardBg, color: currentTheme.textColor, borderColor: currentTheme.accent, borderWidth: 2 },
        },
      ];
    } else if (type === 'hierarchy') {
      newElements = [
        {
          id: `el-${timestamp}-h`,
          type: 'heading',
          content: 'Executive Organization Hierarchy',
          subcontent: 'Cross-functional engineering & product leadership structure',
          x: 10,
          y: 10,
          width: 80,
          style: { fontSize: 28, bold: true, color: currentTheme.textColor, align: 'left' },
        },
        {
          id: `el-${timestamp}-lead`,
          type: 'shape',
          content: 'Chief Executive Officer / Chief Architect\nExecutive Engineering Direction',
          x: 35,
          y: 26,
          width: 30,
          style: { bgColor: currentTheme.accent, color: '#ffffff', bold: true, align: 'center' },
        },
        {
          id: `el-${timestamp}-sub1`,
          type: 'shape',
          content: 'Lead Systems Engineer\nCore OOXML Engine',
          x: 8,
          y: 56,
          width: 26,
          style: { bgColor: currentTheme.cardBg, color: currentTheme.textColor, align: 'center' },
        },
        {
          id: `el-${timestamp}-sub2`,
          type: 'shape',
          content: 'Head of Product Design\nFluent Ribbon UX & Themes',
          x: 37,
          y: 56,
          width: 26,
          style: { bgColor: currentTheme.cardBg, color: currentTheme.textColor, align: 'center' },
        },
        {
          id: `el-${timestamp}-sub3`,
          type: 'shape',
          content: 'QA & Compliance Lead\nECMA-376 Parity Auditing',
          x: 66,
          y: 56,
          width: 26,
          style: { bgColor: currentTheme.cardBg, color: currentTheme.textColor, align: 'center' },
        },
      ];
    } else if (type === 'matrix') {
      newElements = [
        {
          id: `el-${timestamp}-h`,
          type: 'heading',
          content: 'SWOT Strategic Matrix',
          subcontent: 'Internal capabilities versus market environment',
          x: 10,
          y: 10,
          width: 80,
          style: { fontSize: 28, bold: true, color: currentTheme.textColor, align: 'left' },
        },
        {
          id: `el-${timestamp}-s`,
          type: 'shape',
          content: 'STRENGTHS (Internal)\n• Sub-second startup with zero telemetry\n• Native .docx, .xlsx, .pptx OOXML\n• Complete offline sovereign security',
          x: 8,
          y: 28,
          width: 40,
          style: { bgColor: currentTheme.cardBg, color: '#34d399', borderColor: '#059669', borderWidth: 2 },
        },
        {
          id: `el-${timestamp}-w`,
          type: 'shape',
          content: 'WEAKNESSES (Internal)\n• Initial desktop brand awareness\n• Third-party macro ecosystem migration\n• Cloud team real-time co-authoring rollout',
          x: 52,
          y: 28,
          width: 40,
          style: { bgColor: currentTheme.cardBg, color: '#fb923c', borderColor: '#ea580c', borderWidth: 2 },
        },
        {
          id: `el-${timestamp}-o`,
          type: 'shape',
          content: 'OPPORTUNITIES (External)\n• Displace expensive SaaS subscriptions\n• Edge/air-gapped defense compliance\n• Lightweight alternative to 4GB+ suites',
          x: 8,
          y: 62,
          width: 40,
          style: { bgColor: currentTheme.cardBg, color: '#38bdf8', borderColor: '#0284c7', borderWidth: 2 },
        },
        {
          id: `el-${timestamp}-t`,
          type: 'shape',
          content: 'THREATS (External)\n• Legacy vendor licensing lock-in\n• Proprietary font and template dependencies\n• Cloud-only ecosystem barriers',
          x: 52,
          y: 62,
          width: 40,
          style: { bgColor: currentTheme.cardBg, color: '#f87171', borderColor: '#dc2626', borderWidth: 2 },
        },
      ];
    } else if (type === 'cycle') {
      newElements = [
        {
          id: `el-${timestamp}-h`,
          type: 'heading',
          content: 'Continuous Lifecycle Loop',
          subcontent: 'Iterative feedback, measurement & optimization cycle',
          x: 10,
          y: 10,
          width: 80,
          style: { fontSize: 28, bold: true, color: currentTheme.textColor, align: 'left' },
        },
        {
          id: `el-${timestamp}-c1`,
          type: 'metric',
          content: '01. Plan',
          subcontent: 'Requirements & Scope Analysis',
          x: 12,
          y: 30,
          width: 34,
          style: { bgColor: currentTheme.cardBg, color: currentTheme.accent },
        },
        {
          id: `el-${timestamp}-c2`,
          type: 'metric',
          content: '02. Build',
          subcontent: 'Rapid Prototyping & Core Engine',
          x: 54,
          y: 30,
          width: 34,
          style: { bgColor: currentTheme.cardBg, color: currentTheme.accent },
        },
        {
          id: `el-${timestamp}-c3`,
          type: 'metric',
          content: '03. Validate',
          subcontent: 'Parity Testing & Benchmarking',
          x: 54,
          y: 62,
          width: 34,
          style: { bgColor: currentTheme.cardBg, color: currentTheme.accent },
        },
        {
          id: `el-${timestamp}-c4`,
          type: 'metric',
          content: '04. Deploy',
          subcontent: 'Production Release & Telemetry',
          x: 12,
          y: 62,
          width: 34,
          style: { bgColor: currentTheme.cardBg, color: currentTheme.accent },
        },
      ];
    } else {
      // Funnel
      newElements = [
        {
          id: `el-${timestamp}-h`,
          type: 'heading',
          content: 'Enterprise Conversion Pipeline',
          subcontent: 'Top-of-funnel discovery to closed production contracts',
          x: 10,
          y: 10,
          width: 80,
          style: { fontSize: 28, bold: true, color: currentTheme.textColor, align: 'left' },
        },
        {
          id: `el-${timestamp}-f1`,
          type: 'metric',
          content: '1. Brand Discovery (500K)',
          subcontent: 'Top-of-funnel organic & direct interest',
          x: 8,
          y: 28,
          width: 84,
          style: { bgColor: currentTheme.cardBg, color: '#60a5fa' },
        },
        {
          id: `el-${timestamp}-f2`,
          type: 'metric',
          content: '2. Active Evaluation (120K)',
          subcontent: 'Self-hosted and desktop client installations',
          x: 16,
          y: 46,
          width: 68,
          style: { bgColor: currentTheme.cardBg, color: '#38bdf8' },
        },
        {
          id: `el-${timestamp}-f3`,
          type: 'metric',
          content: '3. Technical Proof of Concept (24K)',
          subcontent: 'Pilot deployments and compliance evaluations',
          x: 24,
          y: 64,
          width: 52,
          style: { bgColor: currentTheme.cardBg, color: '#34d399' },
        },
        {
          id: `el-${timestamp}-f4`,
          type: 'metric',
          content: '4. Enterprise Contract (6.4K)',
          subcontent: 'Multi-seat licensed production organizations',
          x: 32,
          y: 82,
          width: 36,
          style: { bgColor: currentTheme.accent, color: '#ffffff' },
        },
      ];
    }

    updateActiveSlide({ elements: [...activeSlide.elements, ...newElements] });
    setShowSmartArtModal(false);
  };

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
      elements: target.elements.map((el) => ({ ...el, id: `el-${Date.now()}-${Math.random().toString(36).substr(2, 4)}` })),
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

  const moveSlide = (idx: number, direction: 'up' | 'down') => {
    const targetIdx = direction === 'up' ? idx - 1 : idx + 1;
    if (targetIdx < 0 || targetIdx >= data.slides.length) return;

    const newSlides = [...data.slides];
    const [moved] = newSlides.splice(idx, 1);
    newSlides.splice(targetIdx, 0, moved);
    onChange({ ...data, slides: newSlides });
    setActiveSlideIndex(targetIdx);
  };

  const addSlideWithLayout = (layout: SlideData['layout']) => {
    const newId = `slide-${Date.now()}`;
    let elements: SlideElement[] = [];

    if (layout === 'title') {
      elements = [
        {
          id: `el-${Date.now()}-1`,
          type: 'heading',
          content: 'Presentation Title',
          subcontent: 'Executive Subtitle & Presenter Name',
          x: 10,
          y: 35,
          width: 80,
          style: { fontSize: 44, bold: true, color: '#ffffff', align: 'center' },
        },
      ];
    } else if (layout === 'two-column') {
      elements = [
        {
          id: `el-${Date.now()}-1`,
          type: 'heading',
          content: 'Two-Column Strategy Comparison',
          x: 10,
          y: 15,
          width: 80,
          style: { fontSize: 32, bold: true, color: '#ffffff', align: 'left' },
        },
        {
          id: `el-${Date.now()}-2`,
          type: 'shape',
          content: 'Option A: Cloud Platform\n• Accelerated time-to-market\n• Native elastic scalability\n• Global edge delivery',
          x: 10,
          y: 35,
          width: 38,
          style: { bgColor: '#1e293b', color: '#f8fafc', fontSize: 14 },
        },
        {
          id: `el-${Date.now()}-3`,
          type: 'shape',
          content: 'Option B: Hybrid Core\n• Dedicated on-premise governance\n• Regulatory data residency\n• Predictable operational cost',
          x: 52,
          y: 35,
          width: 38,
          style: { bgColor: '#1e293b', color: '#f8fafc', fontSize: 14 },
        },
      ];
    } else if (layout === 'metric-highlight') {
      elements = [
        {
          id: `el-${Date.now()}-1`,
          type: 'heading',
          content: 'Key Performance Indicators',
          x: 10,
          y: 15,
          width: 80,
          style: { fontSize: 32, bold: true, color: '#ffffff', align: 'left' },
        },
        {
          id: `el-${Date.now()}-2`,
          type: 'metric',
          content: '$12.4M',
          subcontent: 'Annual Recurring Revenue (+38% YoY)',
          x: 10,
          y: 38,
          width: 25,
          style: { bgColor: '#064e3b', color: '#ffffff', fontSize: 36, bold: true },
        },
        {
          id: `el-${Date.now()}-3`,
          type: 'metric',
          content: '99.98%',
          subcontent: 'Global Uptime SLA Commitment',
          x: 37,
          y: 38,
          width: 25,
          style: { bgColor: '#1e3a8a', color: '#ffffff', fontSize: 36, bold: true },
        },
        {
          id: `el-${Date.now()}-4`,
          type: 'metric',
          content: '142K+',
          subcontent: 'Active Enterprise Daily Users',
          x: 64,
          y: 38,
          width: 25,
          style: { bgColor: '#701a75', color: '#ffffff', fontSize: 36, bold: true },
        },
      ];
    } else {
      // Default title + bullets
      elements = [
        {
          id: `el-${Date.now()}-1`,
          type: 'heading',
          content: 'Strategic Objectives & Roadmap',
          x: 10,
          y: 15,
          width: 80,
          style: { fontSize: 32, bold: true, color: '#ffffff', align: 'left' },
        },
        {
          id: `el-${Date.now()}-2`,
          type: 'bullet-list',
          content: '• Phase 1: Core application architecture unification\n• Phase 2: OpenXML bidirectional document engine\n• Phase 3: Desktop production distribution',
          x: 10,
          y: 38,
          width: 80,
          style: { fontSize: 18, color: '#cbd5e1' },
        },
      ];
    }

    const newSlide: SlideData = {
      id: newId,
      title: 'New Slide',
      layout,
      bgColor: '#0f172a',
      elements,
    };

    const newSlides = [...data.slides];
    newSlides.splice(activeSlideIndex + 1, 0, newSlide);
    onChange({ ...data, slides: newSlides });
    setActiveSlideIndex(activeSlideIndex + 1);
  };

  const addElement = (type: SlideElement['type']) => {
    const newId = `el-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`;
    const newEl: SlideElement = {
      id: newId,
      type,
      content:
        type === 'heading'
          ? 'New Slide Title'
          : type === 'metric'
          ? '99.9%'
          : type === 'bullet-list'
          ? '• First strategic point\n• Second milestone metric\n• Next deliverables'
          : type === 'image'
          ? 'Architecture Schema'
          : 'Interactive shape card text description',
      subcontent: type === 'metric' ? 'Metric Description' : undefined,
      imageUrl: type === 'image' ? 'https://images.unsplash.com/photo-1557804506-669a67965ba0?w=800&auto=format&fit=crop&q=80' : undefined,
      x: 15,
      y: 40,
      width: type === 'metric' ? 26 : type === 'image' ? 40 : 60,
      style: {
        fontSize: type === 'heading' ? 32 : 16,
        bold: true,
        bgColor: type === 'shape' ? '#1e293b' : type === 'metric' ? '#312e81' : undefined,
        color: '#ffffff',
        align: 'left',
      },
    };
    updateActiveSlide({ elements: [...activeSlide.elements, newEl] });
    setSelectedElementId(newId);
  };

  const deleteElement = (elId: string) => {
    const newElements = activeSlide.elements.filter((el) => el.id !== elId);
    updateActiveSlide({ elements: newElements });
    setSelectedElementId(null);
  };

  const duplicateElement = (elId: string) => {
    const target = activeSlide.elements.find((el) => el.id === elId);
    if (!target) return;
    const clone: SlideElement = {
      ...target,
      id: `el-${Date.now()}`,
      x: Math.min(80, target.x + 5),
      y: Math.min(80, target.y + 5),
    };
    updateActiveSlide({ elements: [...activeSlide.elements, clone] });
    setSelectedElementId(clone.id);
  };

  const formatTime = (secs: number) => {
    const m = Math.floor(secs / 60);
    const s = secs % 60;
    return `${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
  };

  const selectedElement = activeSlide.elements.find((el) => el.id === selectedElementId);

  // 1. FULLSCREEN PRESENTER MODE OVERLAY
  if (isPresenting) {
    return (
      <div
        className="fixed inset-0 z-50 bg-black flex flex-col items-center justify-center text-white select-none relative overflow-hidden"
        onMouseMove={(e) => {
          if (laserPointerActive) {
            setLaserPos({ x: e.clientX, y: e.clientY });
          }
        }}
      >
        {/* Laser pointer dot */}
        {laserPointerActive && (
          <div
            className="fixed w-4 h-4 rounded-full bg-red-500 shadow-[0_0_15px_#ff0000] pointer-events-none z-50 -translate-x-1/2 -translate-y-1/2"
            style={{ left: laserPos.x, top: laserPos.y }}
          />
        )}

        {/* Top Floating Presenter Bar */}
        <div className="absolute top-4 left-6 right-6 flex items-center justify-between text-xs text-slate-400 bg-slate-900/90 backdrop-blur-md px-5 py-2.5 rounded-full border border-slate-700/60 z-40">
          <div className="flex items-center space-x-2.5">
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
              onClick={() => setLaserPointerActive(!laserPointerActive)}
              className={`px-3 py-1 rounded text-xs font-medium transition-colors ${
                laserPointerActive ? 'bg-red-600 text-white' : 'bg-slate-800 hover:bg-slate-700 text-slate-300'
              }`}
            >
              Laser Pointer {laserPointerActive ? 'ON' : 'OFF'}
            </button>

            <button
              onClick={() => setShowNotesInPresentation(!showNotesInPresentation)}
              className={`px-3 py-1 rounded text-xs font-medium transition-colors ${
                showNotesInPresentation ? 'bg-blue-600 text-white' : 'bg-slate-800 hover:bg-slate-700 text-slate-300'
              }`}
            >
              Speaker Notes
            </button>

            <button
              onClick={() => setIsPresenting(false)}
              className="flex items-center space-x-1 px-3 py-1 rounded bg-slate-800 hover:bg-slate-700 text-white font-medium"
            >
              <Minimize2 className="w-3.5 h-3.5" />
              <span>End Show (Esc)</span>
            </button>
          </div>
        </div>

        {/* 16:9 / 4:3 Presentation Stage with Active Transition Keyframes */}
        <div
          key={`presenter-slide-${activeSlideIndex}`}
          className={`relative w-full max-w-6xl ${
            data.aspectRatio === '4:3' ? 'aspect-[4/3]' : 'aspect-[16/9]'
          } shadow-2xl rounded-xl overflow-hidden border border-slate-800 p-12 flex flex-col justify-between ${
            activeSlide.transition && activeSlide.transition !== 'none'
              ? `transition-slide-${activeSlide.transition}`
              : 'transition-slide-fade'
          }`}
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

                {el.type === 'image' && el.imageUrl && (
                  <img
                    src={el.imageUrl}
                    alt={el.content}
                    className="w-full h-auto rounded-lg shadow-md object-cover"
                  />
                )}
              </div>
            ))}
          </div>

          <div className="flex justify-between items-center text-xs text-slate-500">
            <span>Anything Office Presentation Deck</span>
            <span>Slide {activeSlideIndex + 1}</span>
          </div>
        </div>

        {/* Floating Speaker Notes Panel in Slideshow */}
        {showNotesInPresentation && (
          <div className="absolute bottom-20 left-12 w-96 p-4 rounded-xl bg-slate-900/95 border border-slate-700 shadow-2xl text-xs z-40">
            <div className="font-bold text-amber-400 uppercase text-[10px] mb-1">
              Presenter Notes for Slide {activeSlideIndex + 1}
            </div>
            <div className="text-slate-200 leading-relaxed max-h-40 overflow-y-auto">
              {activeSlide.notes || 'No notes added for this slide.'}
            </div>
          </div>
        )}

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

  // 2. STANDARD POWERPOINT EDITOR VIEW
  return (
    <div className="flex flex-col h-full w-full select-none text-xs overflow-hidden">
      {/* Element Properties Toolbar if an element is selected */}
      {selectedElement && (
        <div
          className={`h-8 px-4 border-b flex items-center justify-between text-xs shrink-0 select-none ${
            darkMode ? 'bg-slate-950 border-slate-800 text-slate-300' : 'bg-slate-100 border-slate-300 text-slate-700'
          }`}
        >
          <div className="flex items-center space-x-2">
            <span className="font-semibold text-[11px] uppercase text-orange-500">
              Element: {selectedElement.type}
            </span>

            <div className="h-4 w-px bg-slate-300 dark:bg-slate-700" />

            {/* Font Size */}
            <div className="flex items-center space-x-1 font-mono text-xs">
              <button
                onClick={() =>
                  updateElement(selectedElement.id, {
                    style: {
                      ...selectedElement.style,
                      fontSize: Math.max(10, (selectedElement.style?.fontSize || 20) - 2),
                    },
                  })
                }
                className="px-1.5 py-0.5 rounded border dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800"
              >
                -
              </button>
              <span>{selectedElement.style?.fontSize || 20}pt</span>
              <button
                onClick={() =>
                  updateElement(selectedElement.id, {
                    style: {
                      ...selectedElement.style,
                      fontSize: Math.min(72, (selectedElement.style?.fontSize || 20) + 2),
                    },
                  })
                }
                className="px-1.5 py-0.5 rounded border dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800"
              >
                +
              </button>
            </div>

            <div className="h-4 w-px bg-slate-300 dark:bg-slate-700" />

            {/* Alignment */}
            <button
              onClick={() =>
                updateElement(selectedElement.id, {
                  style: { ...selectedElement.style, align: 'left' },
                })
              }
              className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
            >
              <AlignLeft className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() =>
                updateElement(selectedElement.id, {
                  style: { ...selectedElement.style, align: 'center' },
                })
              }
              className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
            >
              <AlignCenter className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() =>
                updateElement(selectedElement.id, {
                  style: { ...selectedElement.style, align: 'right' },
                })
              }
              className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
            >
              <AlignRight className="w-3.5 h-3.5" />
            </button>

            <div className="h-4 w-px bg-slate-300 dark:bg-slate-700" />

            {/* Z-Ordering Layer Controls */}
            <div className="flex items-center space-x-1">
              <span className="text-[10px] text-slate-400 font-semibold uppercase">Layer:</span>
              <button
                onClick={() => bringForward(selectedElement.id)}
                className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
                title="Bring Forward"
              >
                <ArrowUp className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => sendBackward(selectedElement.id)}
                className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
                title="Send Backward"
              >
                <ArrowDown className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => bringToFront(selectedElement.id)}
                className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
                title="Bring to Front"
              >
                <ArrowUpToLine className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => sendToBack(selectedElement.id)}
                className="p-1 rounded hover:bg-slate-200 dark:hover:bg-slate-800"
                title="Send to Back"
              >
                <ArrowDownToLine className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>

          <div className="flex items-center space-x-1.5">
            <button
              onClick={() => duplicateElement(selectedElement.id)}
              className="px-2 py-0.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 flex items-center space-x-1"
            >
              <Copy className="w-3 h-3" />
              <span>Duplicate</span>
            </button>
            <button
              onClick={() => deleteElement(selectedElement.id)}
              className="px-2 py-0.5 rounded hover:bg-slate-200 dark:hover:bg-slate-800 text-rose-500 flex items-center space-x-1"
            >
              <Trash2 className="w-3 h-3" />
              <span>Delete</span>
            </button>
          </div>
        </div>
      )}

      <div className="flex flex-1 overflow-hidden">
        {/* Left Thumbnails Panel */}
        <div
          className={`w-56 border-r flex flex-col justify-between shrink-0 transition-colors ${
            darkMode ? 'bg-slate-900 border-slate-800 text-slate-300' : 'bg-slate-50 border-slate-300 text-slate-700'
          }`}
        >
          <div className="p-2.5 overflow-y-auto space-y-2">
            <div className="flex items-center justify-between text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-1">
              <span>Slides ({data.slides.length})</span>
              <button
                onClick={() => addSlideWithLayout('title-bullets')}
                className="text-orange-600 dark:text-orange-400 hover:underline flex items-center space-x-0.5 font-bold"
              >
                <Plus className="w-3 h-3" />
                <span>New Slide</span>
              </button>
            </div>

            {/* Quick Layout Buttons */}
            <div className="grid grid-cols-2 gap-1 pb-2 border-b border-slate-200 dark:border-slate-800">
              <button
                onClick={() => addSlideWithLayout('title')}
                className="px-1.5 py-1 text-[10px] rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800 text-center truncate"
              >
                Title
              </button>
              <button
                onClick={() => addSlideWithLayout('two-column')}
                className="px-1.5 py-1 text-[10px] rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800 text-center truncate"
              >
                Two-Col
              </button>
              <button
                onClick={() => addSlideWithLayout('metric-highlight')}
                className="px-1.5 py-1 text-[10px] rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800 text-center truncate"
              >
                Metrics
              </button>
              <button
                onClick={() => addSlideWithLayout('blank')}
                className="px-1.5 py-1 text-[10px] rounded border border-slate-300 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-800 text-center truncate"
              >
                Blank
              </button>
            </div>

            {/* Slide Thumbnails List */}
            <div className="space-y-2">
              {data.slides.map((slide, idx) => {
                const isSelected = idx === activeSlideIndex;
                return (
                  <div
                    key={slide.id}
                    onClick={() => {
                      setActiveSlideIndex(idx);
                      setSelectedElementId(null);
                    }}
                    className="flex items-start space-x-1.5 cursor-pointer group"
                  >
                    <span className="text-[10px] font-mono text-slate-400 pt-1 w-3 text-right">
                      {idx + 1}
                    </span>

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

                      {/* Hover Actions: Move Up, Move Down, Duplicate, Delete */}
                      <div className="absolute top-1 right-1 hidden group-hover/thumb:flex items-center space-x-0.5 bg-black/85 rounded px-1 py-0.5 z-20">
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            moveSlide(idx, 'up');
                          }}
                          disabled={idx === 0}
                          title="Move Slide Up"
                          className="p-0.5 text-white hover:text-amber-400 disabled:opacity-20"
                        >
                          <ArrowUp className="w-2.5 h-2.5" />
                        </button>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            moveSlide(idx, 'down');
                          }}
                          disabled={idx === data.slides.length - 1}
                          title="Move Slide Down"
                          className="p-0.5 text-white hover:text-amber-400 disabled:opacity-20"
                        >
                          <ArrowDown className="w-2.5 h-2.5" />
                        </button>
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

        {/* Main Canvas Stage */}
        <div
          onClick={() => setSelectedElementId(null)}
          className={`flex-1 flex flex-col items-center justify-between p-6 overflow-y-auto transition-colors ${
            darkMode ? 'bg-[#1e1e1e]' : 'bg-[#e9e9e9]'
          }`}
        >
          {/* Canvas Floating Toolbar for adding elements, transitions, themes, and SmartArt */}
          <div className="mb-3 flex items-center space-x-1.5 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 px-3 py-1.5 rounded-full shadow-md select-none text-xs">
            <span className="text-[10px] font-bold uppercase text-slate-400 mr-1">Insert:</span>
            <button
              onClick={() => addElement('heading')}
              className="flex items-center space-x-1 px-2 py-0.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
            >
              <Type className="w-3.5 h-3.5 text-orange-500" />
              <span>Text</span>
            </button>
            <button
              onClick={() => addElement('bullet-list')}
              className="flex items-center space-x-1 px-2 py-0.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
            >
              <FileText className="w-3.5 h-3.5 text-blue-500" />
              <span>Bullets</span>
            </button>
            <button
              onClick={() => addElement('metric')}
              className="flex items-center space-x-1 px-2 py-0.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
            >
              <TrendingUp className="w-3.5 h-3.5 text-emerald-500" />
              <span>Metric</span>
            </button>
            <button
              onClick={() => addElement('shape')}
              className="flex items-center space-x-1 px-2 py-0.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
            >
              <Square className="w-3.5 h-3.5 text-purple-500" />
              <span>Shape Card</span>
            </button>
            <button
              onClick={() => addElement('image')}
              className="flex items-center space-x-1 px-2 py-0.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
            >
              <ImageIcon className="w-3.5 h-3.5 text-teal-500" />
              <span>Image</span>
            </button>
            <button
              onClick={() => setShowSmartArtModal(true)}
              className="flex items-center space-x-1 px-2.5 py-0.5 rounded bg-gradient-to-r from-orange-500 to-amber-500 text-white font-semibold shadow-xs hover:brightness-110"
              title="Insert Microsoft SmartArt Diagrams"
            >
              <Workflow className="w-3.5 h-3.5" />
              <span>SmartArt</span>
            </button>

            <div className="h-4 w-px bg-slate-300 dark:bg-slate-700 mx-1" />

            {/* Slide Transition Selector */}
            <div className="flex items-center space-x-1">
              <Sliders className="w-3.5 h-3.5 text-slate-400" />
              <select
                value={activeSlide.transition || 'fade'}
                onChange={(e) => setSlideTransition(e.target.value as any)}
                className="h-6 px-1.5 text-[11px] rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 focus:outline-none"
                title="Slide Transition Effect"
              >
                <option value="none">No Transition</option>
                <option value="fade">Fade Transition</option>
                <option value="slide-left">Slide Left</option>
                <option value="slide-up">Slide Up</option>
                <option value="zoom">Zoom Transition</option>
              </select>
              <button
                onClick={() => setSlideTransition(activeSlide.transition || 'fade', true)}
                title="Apply Current Transition to All Slides"
                className="px-1.5 py-0.5 rounded text-[10px] text-slate-500 hover:text-orange-500 border border-slate-200 dark:border-slate-700"
              >
                All
              </button>
            </div>

            <div className="h-4 w-px bg-slate-300 dark:bg-slate-700 mx-1" />

            {/* Themes Selector Button */}
            <button
              onClick={() => setShowThemeModal(true)}
              className="flex items-center space-x-1 px-2 py-0.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200"
              title="Change Presentation Theme"
            >
              <Palette className="w-3.5 h-3.5 text-pink-500" />
              <span>Themes</span>
            </button>

            {/* Aspect Ratio Toggle */}
            <button
              onClick={toggleAspectRatio}
              className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-700 font-mono text-[10px] hover:bg-slate-100 dark:hover:bg-slate-800"
              title="Toggle Aspect Ratio (16:9 / 4:3)"
            >
              {data.aspectRatio || '16:9'}
            </button>
          </div>

          {/* Slide Canvas */}
          <div
            ref={canvasRef}
            className={`relative w-full max-w-4xl ${
              data.aspectRatio === '4:3' ? 'aspect-[4/3]' : 'aspect-[16/9]'
            } shadow-2xl rounded-sm border border-slate-300 dark:border-neutral-700 p-8 flex flex-col justify-between overflow-hidden`}
            style={{
              backgroundColor: activeSlide.bgColor || '#0f172a',
              transform: `scale(${zoomLevel / 100})`,
              transformOrigin: 'center center',
            }}
          >
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
                    className={`absolute rounded-xs group transition-shadow ${
                      isSelected
                        ? 'ring-1 ring-orange-500 border border-dashed border-orange-400 z-30'
                        : 'hover:ring-1 hover:ring-white/40'
                    }`}
                    style={{
                      left: `${el.x}%`,
                      top: `${el.y}%`,
                      width: `${el.width}%`,
                    }}
                  >
                    {/* Draggable Header Badge when selected */}
                    {isSelected && (
                      <div
                        onMouseDown={(e) => startDrag(el.id, e)}
                        className="absolute -top-6 left-0 right-0 h-5 bg-orange-600 text-white rounded-t flex items-center justify-between px-2 cursor-move text-[10px] font-bold select-none z-40 shadow-md"
                        title="Click and drag to move element across canvas"
                      >
                        <span className="flex items-center space-x-1">
                          <Move className="w-3 h-3" />
                          <span>Drag Element</span>
                        </span>
                        <span className="font-mono text-[9px]">X: {el.x}% · Y: {el.y}% · W: {el.width}%</span>
                      </div>
                    )}

                    {/* 8-Point Selection Handles */}
                    {isSelected && (
                      <>
                        <div className="absolute -top-1 -left-1 w-2 h-2 bg-white border border-slate-900 cursor-nwse-resize" />
                        <div className="absolute -top-1 left-1/2 -translate-x-1/2 w-2 h-2 bg-white border border-slate-900 cursor-ns-resize" />
                        <div className="absolute -top-1 -right-1 w-2 h-2 bg-white border border-slate-900 cursor-nesw-resize" />
                        <div className="absolute top-1/2 -left-1 -translate-y-1/2 w-2 h-2 bg-white border border-slate-900 cursor-ew-resize" />
                        <div className="absolute top-1/2 -right-1 -translate-y-1/2 w-2 h-2 bg-white border border-slate-900 cursor-ew-resize" />
                        <div className="absolute -bottom-1 -left-1 w-2 h-2 bg-white border border-slate-900 cursor-nesw-resize" />
                        <div className="absolute -bottom-1 left-1/2 -translate-x-1/2 w-2 h-2 bg-white border border-slate-900 cursor-ns-resize" />
                        {/* Interactive Resize Corner Handle */}
                        <div
                          onMouseDown={(e) => startResize(el.id, e)}
                          className="absolute -bottom-1.5 -right-1.5 w-3.5 h-3.5 bg-orange-500 border-2 border-white rounded-full cursor-nwse-resize z-50 shadow hover:scale-125 transition-transform"
                          title="Click and drag to resize width"
                        />
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

                    {el.type === 'image' && (
                      <div className="p-1">
                        {el.imageUrl && (
                          <img
                            src={el.imageUrl}
                            alt="Slide graphic"
                            className="w-full rounded-md shadow-md object-cover"
                          />
                        )}
                        <input
                          type="text"
                          value={el.content}
                          onChange={(e) => updateElement(el.id, { content: e.target.value })}
                          placeholder="Image label..."
                          className="w-full text-[10px] text-center text-slate-400 bg-transparent focus:outline-none mt-1"
                        />
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            <div className="flex justify-between items-center text-[10px] text-slate-500 border-t border-slate-800 pt-2 select-none">
              <span>Slide {activeSlideIndex + 1} of {data.slides.length}</span>
              <span>16:9 Widescreen</span>
            </div>
          </div>

          {/* Speaker Notes Drawer */}
          <div className="w-full max-w-4xl mt-3 p-2.5 rounded-sm border border-slate-300 dark:border-neutral-700 bg-white dark:bg-slate-900">
            <div className="text-[9px] font-bold uppercase tracking-wider text-slate-400 mb-0.5">
              Speaker Notes (Visible to Presenter)
            </div>
            <textarea
              rows={2}
              value={activeSlide.notes || ''}
              onChange={(e) => updateActiveSlide({ notes: e.target.value })}
              placeholder="Type private presenter talking points, keynote cues, or speech reminders..."
              className="w-full text-xs bg-transparent focus:outline-none text-slate-600 dark:text-slate-300 resize-none"
            />
          </div>
        </div>
      </div>

      {/* Status Bar */}
      <footer
        className={`h-6 border-t select-none px-3 flex items-center justify-between text-[11px] shrink-0 transition-colors ${
          darkMode ? 'bg-slate-900 border-slate-800 text-slate-400' : 'bg-[#c43e1c] text-white border-orange-800'
        }`}
      >
        <div className="flex items-center space-x-3">
          <span>Slide {activeSlideIndex + 1} of {data.slides.length}</span>
          <span>•</span>
          <span className="font-mono">{data.aspectRatio || '16:9'} HD</span>
          <span>•</span>
          <span className="capitalize">Theme: {data.theme || 'modern-dark'}</span>
          <span>•</span>
          <span className="capitalize">Transition: {activeSlide.transition || 'fade'}</span>
        </div>

        <div className="flex items-center space-x-4">
          <button
            onClick={() => setIsPresenting(true)}
            className="flex items-center space-x-1 hover:bg-white/20 px-2 py-0.5 rounded font-semibold"
            title="Start Fullscreen Slide Show (F5)"
          >
            <Play className="w-3 h-3 fill-current" />
            <span>Start Show (F5)</span>
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

      {/* SMARTART MODAL */}
      {showSmartArtModal && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-xl max-w-2xl w-full p-6 shadow-2xl animate-fadeIn">
            <div className="flex justify-between items-center mb-4 pb-2 border-b border-slate-200 dark:border-slate-800">
              <div className="flex items-center space-x-2">
                <Workflow className="w-5 h-5 text-orange-500" />
                <h3 className="font-bold text-sm text-slate-900 dark:text-white">
                  Insert Microsoft SmartArt Graphic
                </h3>
              </div>
              <button
                onClick={() => setShowSmartArtModal(false)}
                className="p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-400"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <p className="text-xs text-slate-500 dark:text-slate-400 mb-4">
              SmartArt graphics visually communicate information. Choose a diagram layout below to insert editable graphic nodes directly onto your slide:
            </p>

            <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
              {[
                { id: 'process', title: 'Process Flow', desc: 'Sequential horizontal chevron milestones' },
                { id: 'hierarchy', title: 'Org Hierarchy', desc: 'Executive leadership tree branching' },
                { id: 'matrix', title: '2x2 SWOT Matrix', desc: 'Four quadrant strategic evaluation' },
                { id: 'cycle', title: 'Circular Cycle', desc: 'Continuous loop of 4 phases' },
                { id: 'funnel', title: 'Pipeline Funnel', desc: 'Multi-stage conversion pipeline' },
              ].map((item) => (
                <div
                  key={item.id}
                  onClick={() => insertSmartArt(item.id as any)}
                  className="p-3.5 rounded-lg border border-slate-200 dark:border-slate-800 hover:border-orange-500 dark:hover:border-orange-500 hover:bg-orange-50/20 dark:hover:bg-orange-950/20 cursor-pointer transition-all group"
                >
                  <div className="flex items-center space-x-1.5 font-bold text-xs text-slate-900 dark:text-white group-hover:text-orange-500">
                    <Sparkles className="w-3.5 h-3.5 text-orange-500" />
                    <span>{item.title}</span>
                  </div>
                  <p className="text-[11px] text-slate-500 dark:text-slate-400 mt-1 leading-normal">
                    {item.desc}
                  </p>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* THEMES MODAL */}
      {showThemeModal && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-xl max-w-xl w-full p-6 shadow-2xl animate-fadeIn">
            <div className="flex justify-between items-center mb-4 pb-2 border-b border-slate-200 dark:border-slate-800">
              <div className="flex items-center space-x-2">
                <Palette className="w-5 h-5 text-pink-500" />
                <h3 className="font-bold text-sm text-slate-900 dark:text-white">
                  Presentation Themes & Color Palettes
                </h3>
              </div>
              <button
                onClick={() => setShowThemeModal(false)}
                className="p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-400"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <p className="text-xs text-slate-500 dark:text-slate-400 mb-4">
              Select a presentation theme to instantly restyle all slides with coordinated backgrounds, accents, and contrast typography:
            </p>

            <div className="grid grid-cols-2 gap-3">
              {SLIDE_THEMES.map((th) => {
                const isActive = data.theme === th.id;
                return (
                  <div
                    key={th.id}
                    onClick={() => applyTheme(th.id)}
                    className={`p-3.5 rounded-lg border cursor-pointer transition-all flex items-center justify-between ${
                      isActive
                        ? 'border-orange-500 ring-2 ring-orange-500/30'
                        : 'border-slate-200 dark:border-slate-800 hover:border-slate-400'
                    }`}
                    style={{ backgroundColor: th.bgColor }}
                  >
                    <div className="space-y-1">
                      <span className="font-bold text-xs" style={{ color: th.textColor }}>
                        {th.name}
                      </span>
                      <div className="flex items-center space-x-1.5">
                        <span className="w-3 h-3 rounded-full" style={{ backgroundColor: th.accent }} />
                        <span className="w-3 h-3 rounded-full" style={{ backgroundColor: th.cardBg }} />
                      </div>
                    </div>

                    {isActive && <Check className="w-4 h-4 text-orange-400" />}
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
