import React, { useState } from 'react';
import { X, Layers, ArrowUp, ArrowDown, Check, RotateCcw, MoveRight } from 'lucide-react';

export interface SlideItem {
  id: string;
  title: string;
  subtitle?: string;
}

interface Props {
  isOpen: boolean;
  onClose: () => void;
  slides: SlideItem[];
  reportTitle: string;
  onSaveOrder: (newOrder: string[]) => void;
  onResetDefault: () => void;
}

export const ReorderSlidesModal: React.FC<Props> = ({
  isOpen,
  onClose,
  slides,
  reportTitle,
  onSaveOrder,
  onResetDefault,
}) => {
  const [orderedSlides, setOrderedSlides] = useState<SlideItem[]>(slides);
  const [fromPos, setFromPos] = useState<number>(1);
  const [toPos, setToPos] = useState<number>(2);

  // Sync with props whenever modal opens or slides change
  React.useEffect(() => {
    setOrderedSlides(slides);
    setFromPos(Math.min(8, slides.length));
    setToPos(2);
  }, [slides, isOpen]);

  if (!isOpen) return null;

  const total = orderedSlides.length;

  const handleMoveUp = (index: number) => {
    if (index <= 0) return;
    const next = [...orderedSlides];
    const temp = next[index - 1];
    next[index - 1] = next[index];
    next[index] = temp;
    setOrderedSlides(next);
  };

  const handleMoveDown = (index: number) => {
    if (index >= orderedSlides.length - 1) return;
    const next = [...orderedSlides];
    const temp = next[index + 1];
    next[index + 1] = next[index];
    next[index] = temp;
    setOrderedSlides(next);
  };

  // Exact move requested by user: "ejemplo si quiero que la 8 vaya a la posicion 2 y asi"
  const handleExecuteExactMove = () => {
    const fromIndex = fromPos - 1;
    const toIndex = Math.max(0, Math.min(total - 1, toPos - 1));
    if (fromIndex === toIndex || fromIndex < 0 || fromIndex >= total) return;

    const next = [...orderedSlides];
    const [moved] = next.splice(fromIndex, 1);
    next.splice(toIndex, 0, moved);
    setOrderedSlides(next);
  };

  const handleDirectPositionChange = (currentIndex: number, targetPosition: number) => {
    const toIndex = Math.max(0, Math.min(total - 1, targetPosition - 1));
    if (currentIndex === toIndex) return;

    const next = [...orderedSlides];
    const [moved] = next.splice(currentIndex, 1);
    next.splice(toIndex, 0, moved);
    setOrderedSlides(next);
  };

  const handleConfirm = () => {
    onSaveOrder(orderedSlides.map((s) => s.id));
    onClose();
  };

  const handleReset = () => {
    onResetDefault();
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm no-print animate-in fade-in duration-200">
      <div className="w-full max-w-2xl bg-[#0c1322] border border-slate-700/80 rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-800 bg-[#09101d]">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-sky-500/20 border border-sky-500/30 flex items-center justify-center text-sky-400">
              <Layers className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-base font-bold text-white tracking-tight">
                Mover y Reorganizar Diapositivas
              </h3>
              <p className="text-xs text-slate-400">
                {reportTitle} · Total: {total} diapositivas
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="text-slate-400 hover:text-white p-1.5 rounded-lg hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Quick Position Move Box (e.g., "mover la 8 a la posición 2") */}
        <div className="px-6 py-3 bg-sky-950/40 border-b border-sky-800/40 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center gap-2 text-slate-200">
            <span className="font-semibold text-sky-300">Movimiento directo:</span>
            <span>Mover diapositiva</span>
            <select
              value={fromPos}
              onChange={(e) => setFromPos(Number(e.target.value))}
              className="bg-slate-900 border border-slate-700 text-white font-bold rounded px-2 py-1 focus:outline-none focus:border-sky-500"
            >
              {orderedSlides.map((s, idx) => (
                <option key={s.id} value={idx + 1}>
                  #{idx + 1}: {s.title.substring(0, 24)}...
                </option>
              ))}
            </select>
            <MoveRight className="w-4 h-4 text-sky-400" />
            <span>a la posición</span>
            <select
              value={toPos}
              onChange={(e) => setToPos(Number(e.target.value))}
              className="bg-slate-900 border border-slate-700 text-white font-bold rounded px-2 py-1 focus:outline-none focus:border-sky-500"
            >
              {orderedSlides.map((_, idx) => (
                <option key={idx + 1} value={idx + 1}>
                  Posición #{idx + 1}
                </option>
              ))}
            </select>
          </div>
          <button
            onClick={handleExecuteExactMove}
            className="px-3 py-1 bg-sky-600 hover:bg-sky-500 text-white font-bold rounded-lg transition-colors shadow-sm"
          >
            Aplicar Mover
          </button>
        </div>

        {/* Slide List */}
        <div className="p-6 overflow-y-auto space-y-2 flex-1">
          <p className="text-xs text-slate-400 mb-3">
            Ajuste el orden con los botones de flechas o cambiando directamente el número de posición:
          </p>

          <div className="space-y-2">
            {orderedSlides.map((slide, index) => {
              const pos = index + 1;
              const isFirst = index === 0;
              const isLast = index === total - 1;

              return (
                <div
                  key={slide.id}
                  className="flex items-center justify-between p-3 rounded-xl bg-slate-900/80 border border-slate-800 hover:border-slate-700 transition-all gap-3"
                >
                  <div className="flex items-center gap-3 flex-1 min-w-0">
                    <span className="w-8 h-8 rounded-lg bg-slate-800 border border-slate-700 text-sky-400 font-black text-xs flex items-center justify-center shrink-0">
                      {pos}
                    </span>
                    <div className="truncate">
                      <div className="text-sm font-bold text-white truncate">
                        {slide.title}
                      </div>
                      {slide.subtitle && (
                        <div className="text-xs text-slate-400 truncate">
                          {slide.subtitle}
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="flex items-center gap-1.5 shrink-0">
                    {/* Position selector */}
                    <div className="flex items-center gap-1 text-xs text-slate-400 mr-2">
                      <span className="hidden sm:inline">Pos:</span>
                      <select
                        value={pos}
                        onChange={(e) => handleDirectPositionChange(index, Number(e.target.value))}
                        className="bg-slate-950 border border-slate-700 text-white font-bold rounded px-1.5 py-0.5 text-xs"
                      >
                        {orderedSlides.map((_, i) => (
                          <option key={i + 1} value={i + 1}>
                            {i + 1}
                          </option>
                        ))}
                      </select>
                    </div>

                    <button
                      onClick={() => handleMoveUp(index)}
                      disabled={isFirst}
                      className="p-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:hover:bg-slate-800 text-slate-300 transition-colors"
                      title="Subir una posición"
                    >
                      <ArrowUp className="w-4 h-4" />
                    </button>
                    <button
                      onClick={() => handleMoveDown(index)}
                      disabled={isLast}
                      className="p-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:hover:bg-slate-800 text-slate-300 transition-colors"
                      title="Bajar una posición"
                    >
                      <ArrowDown className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Footer */}
        <div className="flex items-center justify-between px-6 py-4 border-t border-slate-800 bg-[#09101d]">
          <button
            onClick={handleReset}
            className="flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>Restablecer Orden Original</span>
          </button>

          <div className="flex items-center gap-2">
            <button
              onClick={onClose}
              className="px-4 py-2 text-xs font-medium rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors"
            >
              Cancelar
            </button>
            <button
              onClick={handleConfirm}
              className="flex items-center gap-1.5 px-4 py-2 text-xs font-bold rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white transition-colors shadow-lg shadow-emerald-950/40"
            >
              <Check className="w-4 h-4" />
              <span>Guardar Nuevo Orden</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
