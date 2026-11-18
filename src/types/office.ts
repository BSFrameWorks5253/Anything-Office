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

export interface PdfFormField {
  id: string;
  type: 'text' | 'checkbox' | 'date' | 'signature';
  name: string;
  value: string;
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
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
  formFields?: PdfFormField[];
}

export interface DocParagraphStyle {
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  bgColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  lineHeight?: number;
}

export interface DocComment {
  id: string;
  paragraphId: string;
  author: string;
  text: string;
  timestamp: number;
  resolved?: boolean;
}

export interface TrackChange {
  id: string;
  paragraphId: string;
  author: string;
  type: 'insert' | 'delete' | 'modify';
  oldText?: string;
  newText?: string;
  timestamp: number;
  status: 'pending' | 'accepted' | 'rejected';
}

export interface DocCitation {
  id: string;
  style: 'APA' | 'MLA' | 'Chicago';
  author: string;
  title: string;
  year: number;
  publisher?: string;
  url?: string;
}

export interface DocWatermark {
  text: string;
  enabled: boolean;
  opacity?: number;
  color?: string;
}

export interface DocMailMerge {
  enabled: boolean;
  variables: string[];
  records: Record<string, string>[];
  activeRecordIndex: number;
}

export interface DocParagraph {
  id: string;
  type: 'p' | 'h1' | 'h2' | 'h3' | 'blockquote' | 'bullet' | 'number' | 'callout' | 'table' | 'image' | 'hr';
  content: string;
  align?: 'left' | 'center' | 'right' | 'justify';
  tableData?: string[][]; // for tables
  imageUrl?: string;
  imageCaption?: string;
  style?: DocParagraphStyle;
}

export interface DocumentData {
  title: string;
  paragraphs: DocParagraph[];
  pageSize: 'A4' | 'Letter' | 'Legal';
  orientation?: 'portrait' | 'landscape';
  margins: { top: number; right: number; bottom: number; left: number };
  fontFamily?: string;
  headerText?: string;
  footerText?: string;
  comments?: DocComment[];
  trackChanges?: TrackChange[];
  trackChangesEnabled?: boolean;
  watermark?: DocWatermark;
  citations?: DocCitation[];
  mailMerge?: DocMailMerge;
}

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  align?: 'left' | 'center' | 'right';
  format?: 'text' | 'number' | 'currency' | 'percent' | 'date';
  decimals?: number;
  bgColor?: string;
  textColor?: string;
  border?: 'none' | 'all' | 'box' | 'bottom';
}

export interface CellData {
  raw: string; // formula "=SUM(A1:B2)" or literal text "120"
  computed?: string | number;
  error?: string;
  style?: CellStyle;
}

export interface ConditionalFormatRule {
  id: string;
  range: string;
  rule: 'greaterThan' | 'lessThan' | 'equals' | 'contains';
  value: string | number;
  bgColor: string;
  textColor: string;
}

export interface DataValidationRule {
  range: string;
  type: 'list';
  options: string[];
}

export interface PivotTableConfig {
  rowField: string;
  valueField: string;
  aggregation: 'SUM' | 'COUNT' | 'AVERAGE' | 'MIN' | 'MAX';
}

export interface SheetData {
  id: string;
  name: string;
  cells: Record<string, CellData>; // key: "A1", "B2"
  rowCount: number;
  colCount: number;
  freezeHeader?: boolean;
  conditionalFormats?: ConditionalFormatRule[];
  validations?: DataValidationRule[];
  pivotTable?: PivotTableConfig;
}

export interface SpreadsheetData {
  sheets: SheetData[];
  activeSheetId: string;
}

export interface SlideElementStyle {
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  bgColor?: string;
  borderColor?: string;
  borderWidth?: number;
  align?: 'left' | 'center' | 'right';
  bold?: boolean;
  italic?: boolean;
  shapeType?: 'box' | 'badge' | 'card' | 'circle' | 'arrow' | 'star';
  borderRadius?: number;
  opacity?: number;
}

export interface SlideElement {
  id: string;
  type: 'text' | 'heading' | 'shape' | 'metric' | 'bullet-list' | 'image';
  content: string;
  subcontent?: string;
  x: number; // percentage 0-100
  y: number; // percentage 0-100
  width: number; // percentage 0-100
  height?: number;
  imageUrl?: string;
  style?: SlideElementStyle;
}

export interface SmartArtItem {
  title: string;
  desc: string;
}

export interface SlideData {
  id: string;
  title: string;
  layout: 'title' | 'title-bullets' | 'two-column' | 'metric-highlight' | 'blank';
  elements: SlideElement[];
  bgColor?: string;
  notes?: string;
  transition?: 'fade' | 'slide-left' | 'slide-up' | 'zoom' | 'none';
  smartArt?: {
    type: 'process' | 'hierarchy' | 'cycle' | 'grid';
    items: SmartArtItem[];
  };
}

export interface PresentationData {
  slides: SlideData[];
  theme: 'modern-dark' | 'corporate-blue' | 'emerald' | 'sunset' | 'slate' | 'cyber';
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
