export type AnnotationTool =
  | 'select'
  | 'hand'
  | 'text'
  | 'highlight'
  | 'underline'
  | 'strikethrough'
  | 'sticky'
  | 'ink'
  | 'signature';

export interface CommentReply {
  id: string;
  author: string;
  text: string;
  timestamp: number;
}

export interface Point {
  x: number;
  y: number;
}

export interface InkStroke {
  points: Point[];
  color: string;
  width: number;
}

export interface PdfAnnotation {
  id: string;
  type: 'text' | 'highlight' | 'underline' | 'strikethrough' | 'sticky' | 'ink' | 'signature';
  pageNumber: number;
  // Normalized to base A4 dimensions (794 x 1123) for seamless zoom invariance
  x: number;
  y: number;
  width: number;
  height: number;
  content?: string;
  color: string;
  bgColor?: string;
  fontSize?: number;
  strokeWidth?: number;
  strokes?: InkStroke[]; // for freehand drawings
  author?: string;
  timestamp: number;
  replies?: CommentReply[];
  signatureDataUrl?: string;
  signerName?: string;
  opacity?: number;
}

export interface AnnotationStyleConfig {
  textColor: string;
  bgColor: string;
  fontSize: number;
  highlightColor: string;
  inkColor: string;
  inkWidth: number;
  stickyColor: string;
}
