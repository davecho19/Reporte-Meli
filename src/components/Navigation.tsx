import React from 'react';
import { FileDown, Loader2 } from 'lucide-react';

export type NavView = 'dashboard' | 'presentation1' | 'presentation2';

interface Props {
  currentView: NavView;
  onSelectView: (view: NavView) => void;
  onDownloadCombinedPPTX?: () => void;
  isExportingPPTX?: boolean;
}

export const Navigation: React.FC<Props> = ({
  currentView,
  onSelectView,
  onDownloadCombinedPPTX,
  isExportingPPTX = false,
}) => {
  return (
    <header className="no-print sticky top-0 z-40 w-full bg-[#09101d]/95 backdrop-blur border-b border-slate-800/80 px-4 sm:px-6 py-3">
      <div className="max-w-[1440px] mx-auto flex items-center justify-between gap-4">
        {/* Zone 1: Single text element wordmark */}
        <div className="flex items-center gap-2">
          <a
            href="#"
            onClick={(e) => {
              e.preventDefault();
              onSelectView('dashboard');
            }}
            className="text-lg font-black tracking-tight text-white hover:text-sky-400 transition-colors whitespace-nowrap"
          >
            Upconnect <span className="text-sky-400 font-extrabold">/</span> Connect
          </a>
        </div>

        {/* Zone 2: Clean single-line navigation controls */}
        <nav className="hidden md:flex items-center gap-1 p-1 bg-slate-900/90 border border-slate-800 rounded-xl">
          <button
            onClick={() => onSelectView('dashboard')}
            className={`px-3.5 py-1.5 text-xs font-semibold rounded-lg transition-colors whitespace-nowrap ${
              currentView === 'dashboard'
                ? 'bg-blue-600 text-white shadow-sm'
                : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
          >
            Dashboard
          </button>

          <button
            onClick={() => onSelectView('presentation1')}
            className={`px-3.5 py-1.5 text-xs font-semibold rounded-lg transition-colors whitespace-nowrap ${
              currentView === 'presentation1'
                ? 'bg-blue-600 text-white shadow-sm'
                : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
          >
            Reporte 1: Distribuidor Upconnect / Distribuidor Connect
          </button>

          <button
            onClick={() => onSelectView('presentation2')}
            className={`px-3.5 py-1.5 text-xs font-semibold rounded-lg transition-colors whitespace-nowrap ${
              currentView === 'presentation2'
                ? 'bg-blue-600 text-white shadow-sm'
                : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
          >
            Reporte 2: Auditoría & Socios
          </button>

          {onDownloadCombinedPPTX && (
            <button
              onClick={onDownloadCombinedPPTX}
              disabled={isExportingPPTX}
              className="flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-bold rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white transition-all shadow-md shadow-emerald-950/40 ml-1.5 whitespace-nowrap disabled:opacity-50"
              title="Descargar presentación completa con los 2 reportes consolidados (.pptx)"
            >
              {isExportingPPTX ? (
                <Loader2 className="w-3.5 h-3.5 animate-spin" />
              ) : (
                <FileDown className="w-3.5 h-3.5" />
              )}
              <span>Descargar PPTX Unificado</span>
            </button>
          )}
        </nav>
      </div>

      {/* Mobile nav bar */}
      <div className="md:hidden flex items-center justify-between gap-1 pt-2.5 mt-2 border-t border-slate-800 text-xs font-semibold">
        <button
          onClick={() => onSelectView('dashboard')}
          className={`flex-1 py-1.5 text-center rounded ${
            currentView === 'dashboard' ? 'bg-blue-600 text-white' : 'text-slate-400'
          }`}
        >
          Dashboard
        </button>
        <button
          onClick={() => onSelectView('presentation1')}
          className={`flex-1 py-1.5 text-center rounded ${
            currentView === 'presentation1' ? 'bg-blue-600 text-white' : 'text-slate-400'
          }`}
        >
          Distribuidor Upconnect / Connect
        </button>
        <button
          onClick={() => onSelectView('presentation2')}
          className={`flex-1 py-1.5 text-center rounded ${
            currentView === 'presentation2' ? 'bg-blue-600 text-white' : 'text-slate-400'
          }`}
        >
          Reporte 2
        </button>
        {onDownloadCombinedPPTX && (
          <button
            onClick={onDownloadCombinedPPTX}
            disabled={isExportingPPTX}
            className="py-1.5 px-2.5 text-center rounded bg-emerald-600 text-white flex items-center justify-center gap-1 font-bold whitespace-nowrap"
          >
            {isExportingPPTX ? <Loader2 className="w-3 h-3 animate-spin" /> : <FileDown className="w-3 h-3" />}
            <span>PPTX Unificado</span>
          </button>
        )}
      </div>
    </header>
  );
};
