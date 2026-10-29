import React from 'react';

interface AdobeLogoProps {
  size?: 'sm' | 'md' | 'lg' | 'xl' | '2xl';
  showLabel?: boolean;
  monogram?: string; // Default: 'Ao' (Anything Office)
  className?: string;
  theme?: 'anything' | 'word' | 'excel' | 'powerpoint';
}

export const AdobeLogo: React.FC<AdobeLogoProps> = ({
  size = 'md',
  showLabel = false,
  monogram = 'Ao',
  className = '',
  theme = 'anything',
}) => {
  const pxMap = {
    sm: 24,
    md: 32,
    lg: 40,
    xl: 56,
    '2xl': 88,
  };

  const dim = pxMap[size];

  // Authentic Adobe / Microsoft Fluent design palettes
  const themes = {
    anything: {
      bgTop: '#0a101d',
      bgBottom: '#030712',
      border: '#0284c7',
      text: '#38bdf8',
      glow: 'rgba(56, 189, 248, 0.25)',
    },
    word: {
      bgTop: '#002050',
      bgBottom: '#001433',
      border: '#185abd',
      text: '#2b78e4',
      glow: 'rgba(43, 120, 228, 0.25)',
    },
    excel: {
      bgTop: '#003318',
      bgBottom: '#00200f',
      border: '#107c41',
      text: '#107c41',
      glow: 'rgba(16, 124, 65, 0.25)',
    },
    powerpoint: {
      bgTop: '#3b1207',
      bgBottom: '#260a04',
      border: '#c43e1c',
      text: '#d83b01',
      glow: 'rgba(216, 59, 1, 0.25)',
    },
  };

  const currentTheme = themes[theme];

  return (
    <div className={`flex items-center space-x-2.5 ${className}`}>
      {/* Authentic Adobe Squircle Monogram */}
      <svg
        width={dim}
        height={dim}
        viewBox="0 0 512 512"
        className="select-none shrink-0 transition-transform duration-200 hover:scale-105"
        style={{ filter: `drop-shadow(0 3px 8px ${currentTheme.glow})` }}
      >
        <defs>
          <linearGradient id={`adobeBg-${size}-${theme}`} x1="0%" y1="0%" x2="0%" y2="100%">
            <stop offset="0%" stopColor={currentTheme.bgTop} />
            <stop offset="100%" stopColor={currentTheme.bgBottom} />
          </linearGradient>
        </defs>

        {/* 22% Adobe Squircle Curvature */}
        <rect
          x="28"
          y="28"
          width="456"
          height="456"
          rx="108"
          ry="108"
          fill={`url(#adobeBg-${size}-${theme})`}
          stroke={currentTheme.border}
          strokeWidth="22"
        />

        {/* Centered Adobe-style 2-letter Typography */}
        <text
          x="256"
          y="272"
          textAnchor="middle"
          dominantBaseline="central"
          fontFamily="-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif"
          fontSize="224"
          fontWeight="900"
          letterSpacing="-6"
          fill={currentTheme.text}
        >
          {monogram}
        </text>
      </svg>

      {/* Brand Typography */}
      {showLabel && (
        <div className="flex flex-col text-left select-none">
          <div className="flex items-center space-x-1.5 leading-none">
            <span className="font-bold tracking-tight text-slate-800 dark:text-slate-100 text-sm">
              Anything Office
            </span>
          </div>
          <span className="text-[10px] text-slate-400 font-medium tracking-wide mt-0.5">
            The Unified Workspace
          </span>
        </div>
      )}
    </div>
  );
};
