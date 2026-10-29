import React, { useState, useRef, useEffect } from 'react';
import { PenTool, Type, Upload, X, Check, RotateCcw, ShieldCheck } from 'lucide-react';

interface SignatureModalProps {
  isOpen: boolean;
  onClose: () => void;
  onApplySignature: (signatureData: {
    dataUrl: string;
    signerName: string;
    includeDate: boolean;
  }) => void;
  darkMode: boolean;
}

const CALLIGRAPHY_FONTS = [
  { id: 'font-1', name: 'Elegant Script', font: 'Dancing Script, "Brush Script MT", cursive' },
  { id: 'font-2', name: 'Executive Cursive', font: 'Caveat, "Segoe Script", cursive' },
  { id: 'font-3', name: 'Modern Signature', font: '"Great Vibes", "Lucida Handwriting", cursive' },
  { id: 'font-4', name: 'Formal Calligraphy', font: '"Edwardian Script ITC", cursive' },
];

export const SignatureModal: React.FC<SignatureModalProps> = ({
  isOpen,
  onClose,
  onApplySignature,
  darkMode,
}) => {
  const [tab, setTab] = useState<'draw' | 'type' | 'upload'>('draw');
  const [signerName, setSignerName] = useState<string>('Alex Johnson');
  const [includeDate, setIncludeDate] = useState<boolean>(true);
  const [signatureColor, setSignatureColor] = useState<string>('#1e3a8a'); // classic blue ink
  const [selectedFont, setSelectedFont] = useState<string>(CALLIGRAPHY_FONTS[0].font);
  const [uploadedImage, setUploadedImage] = useState<string | null>(null);

  // Drawing canvas state
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [isDrawing, setIsDrawing] = useState<boolean>(false);
  const [hasDrawn, setHasDrawn] = useState<boolean>(false);

  useEffect(() => {
    if (isOpen && tab === 'draw') {
      const canvas = canvasRef.current;
      if (canvas) {
        const ctx = canvas.getContext('2d');
        if (ctx) {
          ctx.strokeStyle = signatureColor;
          ctx.lineWidth = 2.5;
          ctx.lineCap = 'round';
          ctx.lineJoin = 'round';
        }
      }
    }
  }, [isOpen, tab, signatureColor]);

  if (!isOpen) return null;

  const startDrawing = (e: React.MouseEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const rect = canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    ctx.beginPath();
    ctx.moveTo(x, y);
    setIsDrawing(true);
    setHasDrawn(true);
  };

  const draw = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!isDrawing) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const rect = canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    ctx.lineTo(x, y);
    ctx.stroke();
  };

  const stopDrawing = () => {
    setIsDrawing(false);
  };

  const clearCanvas = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    setHasDrawn(false);
  };

  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (event) => {
        setUploadedImage(event.target?.result as string);
      };
      reader.readAsDataURL(file);
    }
  };

  const generateTypedSignatureDataUrl = (): string => {
    const canvas = document.createElement('canvas');
    canvas.width = 500;
    canvas.height = 160;
    const ctx = canvas.getContext('2d');
    if (!ctx) return '';

    ctx.fillStyle = signatureColor;
    ctx.font = `64px ${selectedFont}`;
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';
    ctx.fillText(signerName || 'Signature', 250, 80);
    return canvas.toDataURL('image/png');
  };

  const handleApply = () => {
    let finalDataUrl = '';

    if (tab === 'draw') {
      const canvas = canvasRef.current;
      if (canvas && hasDrawn) {
        finalDataUrl = canvas.toDataURL('image/png');
      }
    } else if (tab === 'type') {
      finalDataUrl = generateTypedSignatureDataUrl();
    } else if (tab === 'upload') {
      finalDataUrl = uploadedImage || '';
    }

    if (!finalDataUrl && tab === 'draw') {
      // Fallback typed signature if user didn't draw
      finalDataUrl = generateTypedSignatureDataUrl();
    }

    onApplySignature({
      dataUrl: finalDataUrl,
      signerName: signerName || 'Authorized Signatory',
      includeDate,
    });
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-md animate-fadeIn">
      <div
        className={`w-full max-w-xl rounded-2xl shadow-2xl border overflow-hidden transition-all flex flex-col ${
          darkMode ? 'bg-slate-900 border-slate-700 text-slate-100' : 'bg-white border-slate-200 text-slate-800'
        }`}
      >
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-inherit">
          <div className="flex items-center space-x-2.5">
            <div className="w-8 h-8 rounded-lg bg-blue-600/10 text-blue-600 flex items-center justify-center font-bold">
              <ShieldCheck className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-base font-bold">Create Digital Signature</h2>
              <p className="text-xs text-slate-400">Adobe Acrobat Certified Signature Stamp</p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-600 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Tab Selection */}
        <div className="flex border-b border-inherit px-6 pt-3 gap-6 text-sm font-semibold">
          <button
            onClick={() => setTab('draw')}
            className={`pb-3 flex items-center gap-2 border-b-2 transition-all ${
              tab === 'draw'
                ? 'border-blue-600 text-blue-600 font-bold'
                : 'border-transparent text-slate-400 hover:text-slate-600'
            }`}
          >
            <PenTool className="w-4 h-4" />
            <span>Draw</span>
          </button>

          <button
            onClick={() => setTab('type')}
            className={`pb-3 flex items-center gap-2 border-b-2 transition-all ${
              tab === 'type'
                ? 'border-blue-600 text-blue-600 font-bold'
                : 'border-transparent text-slate-400 hover:text-slate-600'
            }`}
          >
            <Type className="w-4 h-4" />
            <span>Type</span>
          </button>

          <button
            onClick={() => setTab('upload')}
            className={`pb-3 flex items-center gap-2 border-b-2 transition-all ${
              tab === 'upload'
                ? 'border-blue-600 text-blue-600 font-bold'
                : 'border-transparent text-slate-400 hover:text-slate-600'
            }`}
          >
            <Upload className="w-4 h-4" />
            <span>Upload Image</span>
          </button>
        </div>

        {/* Tab Contents */}
        <div className="p-6 space-y-4">
          {/* TAB 1: DRAW */}
          {tab === 'draw' && (
            <div className="space-y-3">
              <div className="flex justify-between items-center text-xs">
                <span className="text-slate-500 font-medium">Draw your signature in the box below:</span>
                <div className="flex items-center space-x-2">
                  {/* Ink Colors */}
                  <div className="flex items-center space-x-1.5">
                    {['#1e3a8a', '#0f172a', '#dc2626'].map((color) => (
                      <button
                        key={color}
                        onClick={() => setSignatureColor(color)}
                        className={`w-5 h-5 rounded-full border transition-transform ${
                          signatureColor === color ? 'scale-125 ring-2 ring-blue-500 border-white' : 'border-black/20'
                        }`}
                        style={{ backgroundColor: color }}
                      />
                    ))}
                  </div>

                  <button
                    onClick={clearCanvas}
                    className="flex items-center gap-1 text-slate-400 hover:text-red-500 transition-colors ml-2"
                    title="Clear Signature"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    <span>Clear</span>
                  </button>
                </div>
              </div>

              <div className="relative border-2 border-dashed border-slate-300 dark:border-slate-700 rounded-xl bg-slate-50 dark:bg-slate-950/60 overflow-hidden h-44 flex items-center justify-center">
                <canvas
                  ref={canvasRef}
                  width={480}
                  height={170}
                  onMouseDown={startDrawing}
                  onMouseMove={draw}
                  onMouseUp={stopDrawing}
                  onMouseLeave={stopDrawing}
                  className="cursor-crosshair w-full h-full"
                />
                {!hasDrawn && (
                  <span className="absolute pointer-events-none text-slate-400 text-xs font-medium">
                    Sign using mouse or stylus here
                  </span>
                )}
              </div>
            </div>
          )}

          {/* TAB 2: TYPE */}
          {tab === 'type' && (
            <div className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1.5">
                  Your Full Name
                </label>
                <input
                  type="text"
                  value={signerName}
                  onChange={(e) => setSignerName(e.target.value)}
                  placeholder="Enter your name..."
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 font-medium"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-500 uppercase tracking-wider mb-2">
                  Calligraphy Style
                </label>
                <div className="grid grid-cols-2 gap-2.5">
                  {CALLIGRAPHY_FONTS.map((item) => (
                    <button
                      key={item.id}
                      onClick={() => setSelectedFont(item.font)}
                      className={`p-3 rounded-xl border text-left transition-all ${
                        selectedFont === item.font
                          ? 'border-blue-600 bg-blue-50/50 dark:bg-blue-900/20 ring-1 ring-blue-600'
                          : 'border-slate-200 dark:border-slate-800 hover:border-slate-400'
                      }`}
                    >
                      <div className="text-[10px] text-slate-400 font-mono mb-1">{item.name}</div>
                      <div
                        className="text-2xl truncate text-slate-900 dark:text-slate-100"
                        style={{ fontFamily: item.font, color: signatureColor }}
                      >
                        {signerName || 'Signature'}
                      </div>
                    </button>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* TAB 3: UPLOAD */}
          {tab === 'upload' && (
            <div className="space-y-3">
              <label className="block text-xs font-medium text-slate-500">
                Upload scanned signature image (PNG, JPG, SVG):
              </label>

              <div className="border-2 border-dashed border-slate-300 dark:border-slate-700 rounded-xl p-6 flex flex-col items-center justify-center text-center bg-slate-50 dark:bg-slate-950/60 h-44">
                {uploadedImage ? (
                  <div className="flex flex-col items-center gap-2">
                    <img src={uploadedImage} alt="Uploaded Signature" className="max-h-24 object-contain" />
                    <button
                      onClick={() => setUploadedImage(null)}
                      className="text-xs text-red-500 hover:underline"
                    >
                      Remove image
                    </button>
                  </div>
                ) : (
                  <>
                    <Upload className="w-8 h-8 text-slate-400 mb-2" />
                    <p className="text-xs text-slate-500 mb-3">Drag and drop signature image or browse</p>
                    <input
                      type="file"
                      accept="image/*"
                      onChange={handleImageUpload}
                      className="text-xs text-slate-400 file:mr-2 file:py-1 file:px-3 file:rounded-md file:border-0 file:text-xs file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer"
                    />
                  </>
                )}
              </div>
            </div>
          )}

          {/* Options: Date Stamp & Signer Name */}
          <div className="pt-2 border-t border-inherit flex items-center justify-between text-xs text-slate-500">
            <label className="flex items-center space-x-2 cursor-pointer">
              <input
                type="checkbox"
                checked={includeDate}
                onChange={(e) => setIncludeDate(e.target.checked)}
                className="rounded border-slate-300 text-blue-600 focus:ring-blue-500"
              />
              <span>Include automated timestamp ({new Date().toLocaleDateString()})</span>
            </label>
          </div>
        </div>

        {/* Footer Actions */}
        <div className="flex items-center justify-end px-6 py-4 border-t border-inherit gap-3 bg-slate-50/50 dark:bg-slate-950/30">
          <button
            onClick={onClose}
            className="px-4 py-2 rounded-xl text-xs font-semibold text-slate-600 dark:text-slate-300 hover:bg-slate-200/60 dark:hover:bg-slate-800 transition-colors"
          >
            Cancel
          </button>
          <button
            onClick={handleApply}
            className="flex items-center gap-1.5 px-5 py-2 rounded-xl text-xs font-bold bg-blue-600 hover:bg-blue-700 text-white shadow-md shadow-blue-500/20 transition-all active:scale-95"
          >
            <Check className="w-4 h-4" />
            <span>Apply Signature Stamp</span>
          </button>
        </div>
      </div>
    </div>
  );
};
