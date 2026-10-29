import type { PdfAnnotation } from './annotations';

export type AppMode = 'home' | 'docs' | 'sheets' | 'slides' | 'pdf' | 'evidence';

export interface RecentDocumentItem {
  id: string;
  name: string;
  type: 'docs' | 'sheets' | 'slides' | 'pdf' | 'anything';
  lastModified: number;
  size: string;
  pinned: boolean;
  source: 'local' | 'template' | 'sample';
  description?: string;
  projectData?: Partial<AnythingProject>;
}

export interface PdfDocumentData {
  title: string;
  url?: string;
  totalPages: number;
  currentPage: number;
  zoom: number;
  rotation: number;
  pageThumbnails?: string[];
  extractedText?: string;
  annotations?: PdfAnnotation[];
}

export interface DocParagraph {
  id: string;
  type: 'p' | 'h1' | 'h2' | 'h3' | 'blockquote' | 'bullet' | 'number' | 'callout' | 'table';
  content: string;
  align?: 'left' | 'center' | 'right' | 'justify';
  tableData?: string[][]; // for tables
}

export interface DocumentData {
  title: string;
  paragraphs: DocParagraph[];
  pageSize: 'A4' | 'Letter';
  margins: { top: number; right: number; bottom: number; left: number };
  headerText?: string;
  footerText?: string;
}

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  align?: 'left' | 'center' | 'right';
  format?: 'text' | 'number' | 'currency' | 'percent';
  bgColor?: string;
  textColor?: string;
}

export interface CellData {
  raw: string; // formula "=SUM(A1:B2)" or literal text "120"
  computed?: string | number;
  error?: string;
  style?: CellStyle;
}

export interface SheetData {
  id: string;
  name: string;
  cells: Record<string, CellData>; // key: "A1", "B2"
  rowCount: number;
  colCount: number;
}

export interface SpreadsheetData {
  sheets: SheetData[];
  activeSheetId: string;
}

export interface SlideElement {
  id: string;
  type: 'text' | 'heading' | 'shape' | 'metric' | 'bullet-list';
  content: string;
  subcontent?: string;
  x: number; // percentage 0-100
  y: number; // percentage 0-100
  width: number; // percentage 0-100
  height?: number;
  style?: {
    fontSize?: number;
    color?: string;
    bgColor?: string;
    borderColor?: string;
    align?: 'left' | 'center' | 'right';
    bold?: boolean;
    shapeType?: 'box' | 'badge' | 'card' | 'circle';
  };
}

export interface SlideData {
  id: string;
  title: string;
  layout: 'title' | 'title-bullets' | 'two-column' | 'metric-highlight' | 'blank';
  elements: SlideElement[];
  bgColor?: string;
  notes?: string;
}

export interface PresentationData {
  slides: SlideData[];
  theme: 'modern-dark' | 'corporate-blue' | 'emerald' | 'minimal-light';
  aspectRatio: '16:9' | '4:3';
}

export interface AnythingProject {
  id: string;
  name: string;
  version: string;
  lastModified: number;
  docs: DocumentData;
  sheets: SpreadsheetData;
  slides: PresentationData;
  pdf?: PdfDocumentData;
}
