import React, { useState } from 'react';
import type { PdfAnnotation } from '../../types/annotations';
import {
  X,
  MessageSquare,
  Trash2,
  Send,
  Clock,
  Filter,
} from 'lucide-react';

interface CommentsSidebarProps {
  isOpen: boolean;
  onClose: () => void;
  annotations: PdfAnnotation[];
  currentPage: number;
  onSelectAnnotation: (id: string) => void;
  selectedAnnotationId: string | null;
  onAddReply: (annotationId: string, replyText: string) => void;
  onDeleteAnnotation: (id: string) => void;
  darkMode: boolean;
}

export const CommentsSidebar: React.FC<CommentsSidebarProps> = ({
  isOpen,
  onClose,
  annotations,
  currentPage,
  onSelectAnnotation,
  selectedAnnotationId,
  onAddReply,
  onDeleteAnnotation,
  darkMode,
}) => {
  const [filterCurrentOnly, setFilterCurrentOnly] = useState<boolean>(false);
  const [replyTextMap, setReplyTextMap] = useState<Record<string, string>>({});

  if (!isOpen) return null;

  const commentAnnotations = annotations.filter(
    (a) => a.type === 'sticky' || a.type === 'text' || (a.replies && a.replies.length > 0)
  );

  const displayedComments = filterCurrentOnly
    ? commentAnnotations.filter((a) => a.pageNumber === currentPage)
    : commentAnnotations;

  const handleSendReply = (annotationId: string) => {
    const text = replyTextMap[annotationId]?.trim();
    if (text) {
      onAddReply(annotationId, text);
      setReplyTextMap({ ...replyTextMap, [annotationId]: '' });
    }
  };

  const formatTimestamp = (ts: number): string => {
    const diff = Math.floor((Date.now() - ts) / 1000);
    if (diff < 60) return 'Just now';
    if (diff < 3600) return `${Math.floor(diff / 60)}m ago`;
    if (diff < 86400) return `${Math.floor(diff / 3600)}h ago`;
    return new Date(ts).toLocaleDateString();
  };

  return (
    <aside
      className={`w-80 shrink-0 border-l flex flex-col h-full z-20 select-none transition-all ${
        darkMode ? 'bg-slate-900 border-slate-800 text-slate-100' : 'bg-white border-slate-200 text-slate-800 shadow-lg'
      }`}
    >
      {/* Sidebar Header */}
      <div className="p-4 border-b border-inherit flex items-center justify-between">
        <div className="flex items-center space-x-2">
          <MessageSquare className="w-4 h-4 text-amber-500" />
          <h3 className="font-bold text-sm">Comments & Threads</h3>
          <span className="text-[11px] px-2 py-0.5 rounded-full bg-amber-500/15 text-amber-600 dark:text-amber-400 font-bold">
            {commentAnnotations.length}
          </span>
        </div>
        <button
          onClick={onClose}
          className="p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-400 hover:text-slate-600"
        >
          <X className="w-4 h-4" />
        </button>
      </div>

      {/* Filter Bar */}
      <div className="px-4 py-2 border-b border-inherit flex items-center justify-between text-xs text-slate-500 bg-slate-50/50 dark:bg-slate-950/30">
        <div className="flex items-center gap-1.5">
          <Filter className="w-3.5 h-3.5" />
          <span>Filter:</span>
        </div>
        <button
          onClick={() => setFilterCurrentOnly(!filterCurrentOnly)}
          className={`px-2 py-0.5 rounded font-medium transition-colors ${
            filterCurrentOnly
              ? 'bg-amber-500 text-white font-bold'
              : 'hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-400'
          }`}
        >
          {filterCurrentOnly ? `Page ${currentPage} Only` : 'All Pages'}
        </button>
      </div>

      {/* Comments List */}
      <div className="flex-1 overflow-y-auto p-4 space-y-3.5">
        {displayedComments.length === 0 ? (
          <div className="text-center py-12 text-slate-400 text-xs flex flex-col items-center">
            <MessageSquare className="w-8 h-8 opacity-30 mb-2" />
            <p className="font-medium">No comments yet</p>
            <p className="text-[11px] opacity-75 mt-1 max-w-[200px]">
              Use the Sticky Note (C) or Text (T) tool on the toolbar to add annotations.
            </p>
          </div>
        ) : (
          displayedComments.map((ann) => {
            const isSelected = ann.id === selectedAnnotationId;
            return (
              <div
                key={ann.id}
                onClick={() => onSelectAnnotation(ann.id)}
                className={`p-3 rounded-xl border transition-all text-xs cursor-pointer ${
                  isSelected
                    ? 'border-amber-500 bg-amber-500/10 shadow-sm ring-1 ring-amber-500'
                    : 'border-slate-200 dark:border-slate-800 hover:border-slate-300 dark:hover:border-slate-700 bg-white dark:bg-slate-800/50'
                }`}
              >
                {/* Meta Header */}
                <div className="flex items-center justify-between mb-1.5">
                  <div className="flex items-center space-x-1.5">
                    <div className="w-5 h-5 rounded-full bg-amber-500/20 text-amber-600 dark:text-amber-400 flex items-center justify-center font-bold text-[10px]">
                      {ann.author ? ann.author[0].toUpperCase() : 'U'}
                    </div>
                    <span className="font-semibold text-slate-800 dark:text-slate-200">
                      {ann.author || 'Reviewer'}
                    </span>
                  </div>

                  <div className="flex items-center space-x-2">
                    <span className="px-1.5 py-0.2 rounded text-[10px] font-mono bg-slate-100 dark:bg-slate-800 text-slate-500 font-medium">
                      p.{ann.pageNumber}
                    </span>
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        onDeleteAnnotation(ann.id);
                      }}
                      className="text-slate-400 hover:text-red-500 p-0.5 rounded transition-colors"
                      title="Delete Comment"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </div>

                {/* Comment Body */}
                <p className="text-slate-700 dark:text-slate-300 leading-relaxed font-normal my-1">
                  {ann.content || (ann.type === 'sticky' ? 'Sticky Note' : 'Text Annotation')}
                </p>

                {/* Timestamp */}
                <div className="flex items-center space-x-1 text-[10px] text-slate-400 mt-1">
                  <Clock className="w-3 h-3" />
                  <span>{formatTimestamp(ann.timestamp)}</span>
                </div>

                {/* Replies Thread */}
                {ann.replies && ann.replies.length > 0 && (
                  <div className="mt-2.5 pt-2 border-t border-slate-200 dark:border-slate-800 space-y-2">
                    {ann.replies.map((reply) => (
                      <div key={reply.id} className="pl-3 border-l-2 border-amber-400 text-[11px] space-y-0.5">
                        <div className="flex justify-between items-center text-slate-500">
                          <span className="font-semibold">{reply.author}</span>
                          <span className="text-[9px]">{formatTimestamp(reply.timestamp)}</span>
                        </div>
                        <p className="text-slate-700 dark:text-slate-300">{reply.text}</p>
                      </div>
                    ))}
                  </div>
                )}

                {/* Quick Reply Input */}
                <div
                  className="mt-2.5 pt-2 border-t border-slate-200 dark:border-slate-800 flex items-center space-x-1"
                  onClick={(e) => e.stopPropagation()}
                >
                  <input
                    type="text"
                    placeholder="Reply to thread..."
                    value={replyTextMap[ann.id] || ''}
                    onChange={(e) =>
                      setReplyTextMap({ ...replyTextMap, [ann.id]: e.target.value })
                    }
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') handleSendReply(ann.id);
                    }}
                    className="flex-1 bg-slate-100 dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-lg px-2 py-1 text-xs focus:outline-none focus:ring-1 focus:ring-amber-500"
                  />
                  <button
                    onClick={() => handleSendReply(ann.id)}
                    className="p-1 rounded-lg bg-amber-500 text-white hover:bg-amber-600 transition-colors"
                    title="Send Reply"
                  >
                    <Send className="w-3 h-3" />
                  </button>
                </div>
              </div>
            );
          })
        )}
      </div>
    </aside>
  );
};
