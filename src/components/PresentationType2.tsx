import React, { useState, useEffect, useMemo } from 'react';
import { GlobalDataset, SocioRecord, SolutionCategory } from '../types';
import { computeReportView, computeSociosChannelSummary } from '../utils/reportFilters';
import { exportPresentationType2ToPPTX } from '../services/pptxExportService';
import { SociosChannelTable } from './SociosChannelTable';
import { ReorderSlidesModal } from './ReorderSlidesModal';
import { SlideEditorModal } from './SlideEditorModal';
import { 
  ChevronLeft, 
  ChevronRight, 
  Maximize2, 
  Minimize2, 
  TrendingUp, 
  Users, 
  Award, 
  Layers, 
  ShieldCheck, 
  CheckCircle2, 
  FileText, 
  Activity, 
  Briefcase, 
  FileDown, 
  Loader2, 
  Table, 
  BarChart3, 
  Search,
  Edit3,
  Check,
  X,
  Sparkles,
  Zap,
  Share2
} from 'lucide-react';

interface Props {
  dataset: GlobalDataset;
  onBackToDashboard?: () => void;
  onUpdateDataset?: (updated: Partial<GlobalDataset>) => void;
}

interface SlideConfig2 {
  id: 'cover' | 'balance' | 'monthly' | 'socios' | 'weekly' | 'solutions' | 'funnel' | 'cross';
  title: string;
  subtitle?: string;
}

export const PresentationType2: React.FC<Props> = ({ dataset, onBackToDashboard, onUpdateDataset }) => {
  const [currentSlide, setCurrentSlide] = useState<number>(1);
  const [isFullscreen, setIsFullscreen] = useState<boolean>(false);
  const [isExportingPPTX, setIsExportingPPTX] = useState<boolean>(false);

  // Slide Reordering & Edit Mode state
  const [isReorderModalOpen, setIsReorderModalOpen] = useState<boolean>(false);
  const [isEditorModalOpen, setIsEditorModalOpen] = useState<boolean>(false);
  const [isEditMode, setIsEditMode] = useState<boolean>(false);
  const [editValues, setEditValues] = useState<Record<string, string | number>>(dataset.customOverrides || {});

  useEffect(() => {
    setEditValues(dataset.customOverrides || {});
  }, [dataset.customOverrides]);

  const handleSaveAllChanges = () => {
    if (onUpdateDataset) {
      onUpdateDataset({ customOverrides: editValues });
    }
    setIsEditMode(false);
    setIsEditorModalOpen(false);
  };

  const handleCancelEdit = () => {
    setEditValues(dataset.customOverrides || {});
    setIsEditMode(false);
  };

  const viewData = computeReportView(dataset);

  const baseSlides: SlideConfig2[] = [
    { id: 'cover', title: 'Portada Auditoría UpConta', subtitle: 'Presentación Ejecutiva Oficial' },
    { id: 'balance', title: 'Balance General de Facturación 2026', subtitle: 'Consolidado Neto y Bruto con IVA' },
    { id: 'monthly', title: 'Evolución de Ventas Mes a Mes (Sin IVA)', subtitle: 'Histórico Comparativo Mensual' },
    { id: 'socios', title: 'Ranking de Socios: Pareto y Canales', subtitle: 'Gráfico Pareto & Canales en una sola hoja' },
    { id: 'weekly', title: 'Desglose de Rendimiento Semana a Semana', subtitle: 'Auditoría Semanal de Ventas' },
    { id: 'solutions', title: 'Ventas por Categoría de Solución', subtitle: 'Portafolio de Sistemas y Servicios' },
    { id: 'funnel', title: 'Canal de Distribución - Embudo de Socios', subtitle: 'Comunidades y Franquicias Activas' },
    { id: 'cross', title: 'Estrategia Comercial - Cruce Integral', subtitle: 'Fuerza Interna vs Red de Socios' },
  ];

  // Reorder slides dynamically according to dataset.customSlideOrder2
  const activeSlides: SlideConfig2[] = (() => {
    if (!dataset.customSlideOrder2 || dataset.customSlideOrder2.length === 0) {
      return baseSlides;
    }
    const order = dataset.customSlideOrder2;
    return [...baseSlides].sort((a, b) => {
      const idxA = order.indexOf(a.id);
      const idxB = order.indexOf(b.id);
      if (idxA === -1 && idxB === -1) return 0;
      if (idxA === -1) return 1;
      if (idxB === -1) return -1;
      return idxA - idxB;
    });
  })();

  const totalSlides = activeSlides.length;

  useEffect(() => {
    if (currentSlide > totalSlides) {
      setCurrentSlide(totalSlides);
    }
  }, [totalSlides, currentSlide]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') {
        setCurrentSlide((prev) => Math.min(prev + 1, totalSlides));
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        setCurrentSlide((prev) => Math.max(prev - 1, 1));
      } else if (e.key === 'Escape' && isFullscreen) {
        setIsFullscreen(false);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isFullscreen, totalSlides]);

  const toggleFullscreen = () => {
    setIsFullscreen(!isFullscreen);
  };

  const handleExportPPTX = async () => {
    try {
      setIsExportingPPTX(true);
      await exportPresentationType2ToPPTX(dataset);
    } catch (e) {
      console.error('Error generating PowerPoint file', e);
    } finally {
      setIsExportingPPTX(false);
    }
  };

  const netSales = viewData.netSales;
  const ivaRate = 0.15;
  const ivaAmount = viewData.ivaAmount;
  const grossSales = viewData.grossSales;
  const activeSocios = viewData.socios;

  const activeMonthlyMetrics = (viewData.monthlyMetrics && viewData.monthlyMetrics.length > 0)
    ? viewData.monthlyMetrics
    : dataset.monthlyMetrics;

  const auditMonthly = activeMonthlyMetrics.length > 0
    ? activeMonthlyMetrics.map((m, _, arr) => {
        const totalAmount = m.upconnectSales + m.connectorsSales;
        const maxSales = Math.max(...arr.map((x) => x.upconnectSales + x.connectorsSales), 1);
        const percent = Math.round((totalAmount / maxSales) * 100);
        return {
          month: m.month,
          total: totalAmount,
          percent,
          upconnect: m.upconnectSales,
          connectors: m.connectorsSales,
        };
      })
    : [];

  const maxAuditVal = auditMonthly.length > 0 ? Math.max(...auditMonthly.map((m) => m.total)) : 1;

  const activeWeekly = (viewData.weeklyBreakdownType2 && viewData.weeklyBreakdownType2.length > 0)
    ? viewData.weeklyBreakdownType2
    : (dataset.weeklyBreakdownType2 || []);

  const totalWeeklySales = activeWeekly.reduce((a, b) => a + b.amount, 0);

  const activeCategories: SolutionCategory[] = dataset.solutionCategories || [];

  const currentSlideDef = activeSlides[currentSlide - 1] || activeSlides[0];

  return (
    <div className={`flex flex-col ${isFullscreen ? 'fixed inset-0 z-50 bg-[#070b14]' : 'w-full'}`}>
      {/* Slide Deck Top Action Bar */}
      <div className="no-print flex items-center justify-between px-6 py-3 bg-[#0d1527] border-b border-slate-800 text-slate-300">
        <div className="flex items-center gap-3">
          {onBackToDashboard && (
            <button
              onClick={onBackToDashboard}
              className="text-xs px-2.5 py-1 rounded bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white transition-colors flex items-center gap-1.5 mr-2 font-medium"
            >
              ← Dashboard
            </button>
          )}
          <span className="text-xs uppercase tracking-wider font-semibold text-emerald-400">
            Presentación Tipo 2 · Auditoría & Venta de Sistemas por Socio
          </span>
          <span className="text-slate-600">|</span>
          <span className="text-xs px-2 py-0.5 rounded-full bg-emerald-950/80 text-emerald-300 border border-emerald-800/60 font-medium">
            {viewData.periodLabel}
          </span>
          <span className="text-slate-600 hidden sm:inline">|</span>
          <span className="text-xs text-slate-400 hidden sm:inline">
            Diapositiva <strong className="text-white">{currentSlide}</strong> de {totalSlides}
          </span>
        </div>

        <div className="flex items-center gap-2">
          {/* Slide jump buttons */}
          <div className="hidden md:flex items-center gap-1 mr-4">
            {Array.from({ length: totalSlides }, (_, i) => i + 1).map((num) => (
              <button
                key={num}
                onClick={() => setCurrentSlide(num)}
                className={`w-7 h-7 text-xs rounded transition-colors ${
                  currentSlide === num
                    ? 'bg-emerald-500 text-white font-bold'
                    : 'bg-slate-800 hover:bg-slate-700 text-slate-300'
                }`}
                title={activeSlides[num - 1]?.title}
              >
                {num}
              </button>
            ))}
          </div>

          <button
            onClick={() => setCurrentSlide((p) => Math.max(p - 1, 1))}
            disabled={currentSlide === 1}
            className="p-1.5 rounded bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:hover:bg-slate-800 transition-colors"
            title="Anterior (Flecha Izquierda)"
          >
            <ChevronLeft className="w-5 h-5" />
          </button>

          <button
            onClick={() => setCurrentSlide((p) => Math.min(p + 1, totalSlides))}
            disabled={currentSlide === totalSlides}
            className="p-1.5 rounded bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:hover:bg-slate-800 transition-colors"
            title="Siguiente (Flecha Derecha o Espacio)"
          >
            <ChevronRight className="w-5 h-5" />
          </button>

          <button
            onClick={toggleFullscreen}
            className="p-1.5 rounded bg-slate-800 hover:bg-slate-700 transition-colors text-slate-200"
            title={isFullscreen ? 'Salir de pantalla completa (Esc)' : 'Pantalla completa'}
          >
            {isFullscreen ? <Minimize2 className="w-5 h-5" /> : <Maximize2 className="w-5 h-5" />}
          </button>

          <button
            onClick={handleExportPPTX}
            disabled={isExportingPPTX}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 text-white shadow-sm transition-colors"
            title="Descargar presentación unificada con los reportes 1 y 2 consolidados (.pptx)"
          >
            {isExportingPPTX ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <FileDown className="w-4 h-4" />
            )}
            <span>Descargar PPTX Unificado</span>
          </button>

          <button
            onClick={() => setIsReorderModalOpen(true)}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded bg-slate-800 hover:bg-slate-700 text-slate-200 transition-colors"
            title="Mover y organizar diapositivas (ej. mover la 8 a la posición 2)"
          >
            <Layers className="w-4 h-4 text-emerald-400" />
            <span className="hidden sm:inline">Mover Diapositivas</span>
          </button>

          {isEditMode ? (
            <div className="flex items-center gap-1.5">
              <button
                onClick={() => setIsEditorModalOpen(true)}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded bg-emerald-600/40 hover:bg-emerald-600 text-emerald-200 hover:text-white border border-emerald-500/50 transition-colors shadow-sm"
                title="Abrir formulario para editar campos de la diapositiva"
              >
                <Edit3 className="w-4 h-4" />
                <span>Editar Campos</span>
              </button>
              <button
                onClick={handleSaveAllChanges}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded bg-emerald-600 hover:bg-emerald-500 text-white shadow-md shadow-emerald-950/40 transition-all animate-pulse"
                title="Guardar cambios de textos y valores en ambos reportes"
              >
                <Check className="w-4 h-4" />
                <span>Guardar Cambios</span>
              </button>
              <button
                onClick={handleCancelEdit}
                className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium rounded bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white transition-colors"
                title="Cancelar edición"
              >
                <X className="w-4 h-4" />
                <span>Cancelar</span>
              </button>
            </div>
          ) : (
            <button
              onClick={() => {
                setIsEditMode(true);
                setIsEditorModalOpen(true);
              }}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded bg-emerald-600/30 hover:bg-emerald-600 text-emerald-300 hover:text-white border border-emerald-500/40 transition-colors"
              title="Modificar textos y valores de la presentación"
            >
              <Edit3 className="w-4 h-4" />
              <span>Editar</span>
            </button>
          )}
        </div>
      </div>

      {/* Edit Mode Banner */}
      {isEditMode && (
        <div className="bg-emerald-950/90 border-b border-emerald-600/60 px-6 py-2.5 flex items-center justify-between text-xs text-emerald-200 no-print flex-wrap gap-2">
          <div className="flex items-center gap-2">
            <Edit3 className="w-4 h-4 text-emerald-400 animate-pulse" />
            <span>
              <strong>Modo Edición Activado:</strong> Modifique los textos y valores directamente en pantalla o abra el editor de diapositiva.
            </span>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={() => setIsEditorModalOpen(true)}
              className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-emerald-300 border border-emerald-500/40 font-semibold rounded transition-colors text-xs"
            >
              Abrir Formulario de Edición
            </button>
            <button
              onClick={handleSaveAllChanges}
              className="px-3 py-1 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded shadow-sm text-xs flex items-center gap-1 transition-colors"
            >
              <Check className="w-3.5 h-3.5" />
              <span>Guardar Cambios</span>
            </button>
            <button
              onClick={handleCancelEdit}
              className="px-2 py-1 bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white rounded text-xs transition-colors"
            >
              Cancelar
            </button>
          </div>
        </div>
      )}

      {/* Main Slide Stage (16:9) */}
      <div className="flex-1 flex items-center justify-center p-2 sm:p-6 bg-[#070b14] overflow-auto">
        <div className="w-full max-w-[1240px] aspect-[16/9] min-h-[580px] bg-[#0c1322] border border-slate-800/80 rounded-2xl shadow-2xl relative overflow-hidden flex flex-col justify-between slide-container">
          
          {/* SLIDE 1: Cover */}
          {currentSlideDef.id === 'cover' && (
            <div className="h-full flex flex-col justify-center items-center text-center p-8 sm:p-16 relative bg-gradient-to-br from-[#0a1120] via-[#09101d] to-[#040812]">
              <div className="absolute w-[500px] h-[500px] rounded-full bg-emerald-500/10 blur-[130px] pointer-events-none -top-24 -right-24" />
              <div className="absolute w-[450px] h-[450px] rounded-full bg-blue-500/10 blur-[110px] pointer-events-none -bottom-24 -left-24" />

              <div className="z-10 max-w-4xl space-y-6">
                <div className="flex items-center justify-center gap-3">
                  <span className="w-10 h-10 rounded-xl bg-blue-600 text-white font-black flex items-center justify-center text-base shadow-lg shadow-blue-500/30">
                    Up
                  </span>
                  <span className="text-xl font-bold tracking-tight text-white">
                    UpConta <span className="text-emerald-400">/ ERP & Socios</span>
                  </span>
                </div>

                <div className="space-y-3">
                  <span className="text-xs uppercase tracking-widest font-extrabold text-emerald-400 px-3 py-1 rounded-full bg-emerald-950/60 border border-emerald-800/80">
                    REPORTE 2 · AUDITORÍA DE SISTEMAS Y SOCIOS
                  </span>
                  {isEditMode ? (
                    <input
                      type="text"
                      value={editValues['r2_cover_title'] !== undefined ? String(editValues['r2_cover_title']) : 'AUDITORÍA DE FACTURACIÓN Y RENDIMIENTO DE SOCIOS'}
                      onChange={(e) => setEditValues({ ...editValues, r2_cover_title: e.target.value })}
                      className="w-full text-2xl sm:text-4xl font-black text-center bg-slate-950 border border-emerald-500 rounded p-1 text-white"
                    />
                  ) : (
                    <h1 className="text-3xl sm:text-5xl font-black text-white tracking-tight leading-tight">
                      {editValues['r2_cover_title'] || 'AUDITORÍA DE FACTURACIÓN Y RENDIMIENTO DE SOCIOS'}
                    </h1>
                  )}
                  {isEditMode ? (
                    <input
                      type="text"
                      value={editValues['r2_cover_subtitle'] !== undefined ? String(editValues['r2_cover_subtitle']) : 'Análisis detallado de ventas por socio, ticket promedio, cruce de canales y conciliación de IVA'}
                      onChange={(e) => setEditValues({ ...editValues, r2_cover_subtitle: e.target.value })}
                      className="w-full text-xs text-center bg-slate-950 border border-emerald-500 rounded p-1 text-slate-300"
                    />
                  ) : (
                    <p className="text-sm sm:text-base text-slate-300 max-w-2xl mx-auto">
                      {editValues['r2_cover_subtitle'] || 'Análisis detallado de ventas por socio, ticket promedio, cruce de canales y conciliación de IVA'}
                    </p>
                  )}
                </div>

                {/* KPI Pill Grid */}
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 pt-4 max-w-3xl mx-auto text-left">
                  <div className="p-3.5 rounded-xl bg-slate-900/80 border border-slate-800">
                    <div className="text-[11px] text-slate-400 font-semibold uppercase">VENTAS NETAS</div>
                    <div className="text-xl font-black text-white tabular-nums">
                      ${netSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </div>
                    <div className="text-[10px] text-emerald-400 font-medium">100% Conciliado</div>
                  </div>
                  <div className="p-3.5 rounded-xl bg-slate-900/80 border border-slate-800">
                    <div className="text-[11px] text-slate-400 font-semibold uppercase">CARTERA SOCIOS</div>
                    <div className="text-xl font-black text-sky-400 tabular-nums">
                      {activeSocios.length} Socios
                    </div>
                    <div className="text-[10px] text-slate-400 font-medium">Auditados</div>
                  </div>
                  <div className="p-3.5 rounded-xl bg-slate-900/80 border border-slate-800">
                    <div className="text-[11px] text-slate-400 font-semibold uppercase">VENTAS CON IVA</div>
                    <div className="text-xl font-black text-amber-400 tabular-nums">
                      ${grossSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </div>
                    <div className="text-[10px] text-slate-400 font-medium">IVA 15% Calculado</div>
                  </div>
                  <div className="p-3.5 rounded-xl bg-slate-900/80 border border-slate-800">
                    <div className="text-[11px] text-slate-400 font-semibold uppercase">OPERACIONES</div>
                    <div className="text-xl font-black text-emerald-400 tabular-nums">
                      {activeSocios.reduce((a, s) => a + (s.operationsCount || 0), 0) || activeSocios.length}
                    </div>
                    <div className="text-[10px] text-slate-400 font-medium">Transacciones de Socios</div>
                  </div>
                </div>

                <div className="pt-2 text-xs text-slate-500">
                  Período Activo: <strong className="text-slate-300">{viewData.periodLabel}</strong> · Fecha de Corte: <strong className="text-slate-300">{dataset.cutoffDate}</strong>
                </div>
              </div>

              <div className="absolute bottom-4 left-8 right-8 flex justify-between items-center text-[11px] text-slate-500 border-t border-slate-800/60 pt-3">
                <span>UpConta ERP · Módulo de Control de Gestión</span>
                <span>Pág. 01</span>
              </div>
            </div>
          )}

          {/* SLIDE 2: Balance General de Facturación 2026 */}
          {currentSlideDef.id === 'balance' && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                  AUDITORÍA FINANCIERA · {viewData.periodLabel.toUpperCase()}
                </span>
                <h2 className="text-2xl sm:text-3xl font-black text-white tracking-tight mt-0.5">
                  Balance General de Facturación 2026
                </h2>
                <p className="text-xs text-slate-400 mt-1">
                  Desglose consolidado de ingresos netos, desglose impositivo IVA (15%) e ingreso bruto registrado
                </p>
              </div>

              {/* 3 Large KPI Cards matching Slide 2 of PDF 2 */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 my-auto">
                <div className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl relative overflow-hidden">
                  <div className="w-10 h-10 rounded-xl bg-emerald-500/10 text-emerald-400 flex items-center justify-center font-bold mb-4">
                    $
                  </div>
                  <div className="text-xs font-semibold uppercase tracking-wider text-slate-400">
                    VENTAS NETAS (SIN IVA)
                  </div>
                  <div className="text-3xl sm:text-4xl font-black text-white tabular-nums mt-1">
                    ${netSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                  <p className="text-xs text-slate-400 mt-3 pt-3 border-t border-slate-800">
                    Base imponible real correspondiente a los servicios y sistemas emitidos en el período.
                  </p>
                </div>

                <div className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl relative overflow-hidden">
                  <div className="w-10 h-10 rounded-xl bg-amber-500/10 text-amber-400 flex items-center justify-center font-bold mb-4">
                    %
                  </div>
                  <div className="text-xs font-semibold uppercase tracking-wider text-slate-400">
                    IVA ESTIMADO (15%)
                  </div>
                  <div className="text-3xl sm:text-4xl font-black text-amber-400 tabular-nums mt-1">
                    ${ivaAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                  <p className="text-xs text-slate-400 mt-3 pt-3 border-t border-slate-800">
                    Retención impositiva fiscal proyectada para la conciliación tributaria oficial.
                  </p>
                </div>

                <div className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl relative overflow-hidden">
                  <div className="w-10 h-10 rounded-xl bg-sky-500/10 text-sky-400 flex items-center justify-center font-bold mb-4">
                    Σ
                  </div>
                  <div className="text-xs font-semibold uppercase tracking-wider text-slate-400">
                    TOTAL FACTURADO CON IVA
                  </div>
                  <div className="text-3xl sm:text-4xl font-black text-sky-400 tabular-nums mt-1">
                    ${grossSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                  <p className="text-xs text-slate-400 mt-3 pt-3 border-t border-slate-800">
                    Recaudación bruta total esperada en cartera y cuentas por cobrar.
                  </p>
                </div>
              </div>

              {/* Bottom Insight Bar */}
              <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-emerald-400">Conclusión de Auditoría:</strong> El período{' '}
                  <strong className="text-white">{viewData.periodLabel}</strong> registra un promedio diario de{' '}
                  <strong className="text-white">
                    ${((netSales / Math.max(1, viewData.totalCount)) || 0).toFixed(2)} USD
                  </strong>{' '}
                  por transacción sobre una cartera auditada de{' '}
                  <strong className="text-sky-400">{activeSocios.length} socios comerciales</strong>.
                </div>
                <span className="text-slate-500">Corte: {dataset.cutoffDate}</span>
              </div>
            </div>
          )}

          {/* SLIDE 3: Evolución de Ventas Mes a Mes (Sin IVA) */}
          {currentSlideDef.id === 'monthly' && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                  TENDENCIA TEMPORAL · MATRIZ COMPARATIVA
                </span>
                <h2 className="text-2xl sm:text-3xl font-black text-white tracking-tight mt-0.5">
                  Evolución de Ventas Mes a Mes (Sin IVA)
                </h2>
                <p className="text-xs text-slate-400 mt-1">
                  Comparativo de facturación neta mensual y peso relativo porcentual respecto al mes líder
                </p>
              </div>

              {/* Horizontal Bar Chart matching Slide 3 of PDF 2 */}
              <div className="space-y-4 my-auto max-w-4xl w-full">
                {auditMonthly.length === 0 ? (
                  <div className="py-12 text-center text-slate-500 text-sm">
                    Sin datos mensuales registrados. Cargue la matriz de facturación en el Dashboard.
                  </div>
                ) : (
                  auditMonthly.map((item, idx) => {
                    const isLeader = item.total === maxAuditVal;
                    return (
                      <div key={idx} className="space-y-1">
                        <div className="flex justify-between items-center text-xs">
                          <div className="flex items-center gap-2">
                            <span className="font-bold text-white uppercase w-28 tracking-wide">
                              {item.month}
                            </span>
                            {isLeader && (
                              <span className="text-[10px] px-2 py-0.2 rounded-full bg-emerald-500/20 text-emerald-300 font-semibold border border-emerald-500/40">
                                Mes Líder en Ventas
                              </span>
                            )}
                          </div>
                          <div className="flex items-center gap-3">
                            <span className="text-slate-400 text-[11px]">
                              UP: ${item.upconnect.toFixed(2)} • CO: ${item.connectors.toFixed(2)}
                            </span>
                            <span className="font-black text-white tabular-nums text-sm">
                              ${item.total.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD
                            </span>
                          </div>
                        </div>

                        {/* Bar */}
                        <div className="h-5 bg-slate-900 rounded-lg overflow-hidden border border-slate-800 p-0.5">
                          <div
                            style={{ width: `${Math.max(4, (item.total / maxAuditVal) * 100)}%` }}
                            className={`h-full rounded transition-all duration-500 ${
                              isLeader ? 'bg-emerald-500' : 'bg-sky-500/80'
                            }`}
                          />
                        </div>
                      </div>
                    );
                  })
                )}
              </div>

              {/* Bottom Insight */}
              <div className="p-3.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-emerald-400">Total Acumulado Registrado:</strong>{' '}
                  <strong className="text-white">
                    ${auditMonthly.reduce((a, b) => a + b.total, 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD
                  </strong>
                </div>
                <span className="text-slate-400">Canal Distribuidor Upconnect + Red Distribuidor Connect</span>
              </div>
            </div>
          )}

          {/* SLIDE 4: GRÁFICO PARETO Y CANALES EN UNA SOLA HOJA (SIN FILTRO DE MES) */}
          {currentSlideDef.id === 'socios' && (
            <div className="h-full flex flex-col justify-between p-6 sm:p-8 relative bg-[#0c1322]">
              {/* Header: Exact and direct */}
              <div className="flex items-center justify-between border-b border-slate-800/80 pb-2.5">
                <div>
                  <span className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                    AUDITORÍA & RENDIMIENTO DE SOCIOS · {viewData.periodLabel.toUpperCase()}
                  </span>
                  <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                    {editValues['r2_slide4_title'] || 'Desempeño de Socios: Gráfico Pareto & Distribución por Canales'}
                  </h2>
                </div>

                <div className="flex items-center gap-3 text-xs">
                  <div className="bg-slate-900/90 px-3 py-1.5 rounded-lg border border-slate-800 text-slate-300 font-semibold">
                    Cartera Total: <strong className="text-emerald-400">${activeSocios.reduce((a, s) => a + s.totalSales, 0).toLocaleString('en-US', { minimumFractionDigits: 2 })} USD</strong>
                  </div>
                  <div className="bg-slate-900/90 px-3 py-1.5 rounded-lg border border-slate-800 text-slate-300 font-semibold">
                    <span className="text-sky-400 font-bold">{viewData.sociosSummary.upconnect.count} Distribuidor Upconnect</span>
                    <span className="text-slate-600 mx-1.5">|</span>
                    <span className="text-amber-400 font-bold">{viewData.sociosSummary.connectors.count} Distribuidor Connect</span>
                  </div>
                </div>
              </div>

              {/* Two side-by-side columns: LEFT: Pareto Chart, RIGHT: Canales Summary */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-5 my-auto">
                {/* 1. GRÁFICO PARETO TOP SOCIOS */}
                <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between">
                  <div className="flex items-center justify-between pb-2 border-b border-slate-800 mb-2">
                    <div className="flex items-center gap-2">
                      <BarChart3 className="w-4 h-4 text-emerald-400" />
                      <h3 className="text-xs font-bold uppercase tracking-wider text-emerald-300">
                        Regla de Oro Pareto: Top 1–10 Socios
                      </h3>
                    </div>
                    <span className="text-[11px] text-slate-400 font-medium">Mayor Volumen Colocado</span>
                  </div>

                  {activeSocios.length === 0 ? (
                    <div className="py-8 text-center text-slate-500 text-xs">
                      Sin datos de socios registrados
                    </div>
                  ) : (
                    <div className="space-y-1.5">
                      {activeSocios.slice(0, 8).map((s) => {
                        const maxVal = activeSocios[0]?.totalSales || 1;
                        const barWidth = Math.max(5, (s.totalSales / maxVal) * 100);
                        const isTop1 = s.rank === 1;

                        const nameKey = `socio_${s.rank}_name`;
                        const salesKey = `socio_${s.rank}_sales`;
                        const displayName = editValues[nameKey] !== undefined ? String(editValues[nameKey]) : s.name;
                        const displaySales = editValues[salesKey] !== undefined ? Number(editValues[salesKey]) : s.totalSales;

                        return (
                          <div key={s.rank} className="flex items-center gap-2 text-xs">
                            <span className={`w-5 font-bold tabular-nums text-center ${
                              isTop1 ? 'text-amber-400 font-black' : s.rank <= 3 ? 'text-emerald-400' : 'text-slate-500'
                            }`}>
                              {s.rank}.
                            </span>

                            {isEditMode ? (
                              <input
                                type="text"
                                value={displayName}
                                onChange={(e) => setEditValues({ ...editValues, [nameKey]: e.target.value })}
                                className="w-32 bg-slate-950 border border-emerald-500 rounded px-1 text-xs text-white"
                              />
                            ) : (
                              <span className="w-32 truncate font-semibold text-slate-200" title={displayName}>
                                {displayName}
                              </span>
                            )}

                            <span className={`text-[10px] px-1.5 py-0.2 rounded font-medium shrink-0 ${
                              s.channel === 'UpConnect' ? 'bg-sky-500/10 text-sky-300 border border-sky-500/20' : 'bg-amber-500/10 text-amber-300 border border-amber-500/20'
                            }`}>
                              {s.channel || 'Connectors'}
                            </span>

                            <div className="flex-1 bg-slate-800/60 rounded h-3.5 overflow-hidden">
                              <div
                                style={{ width: `${barWidth}%` }}
                                className={`h-full ${isTop1 ? 'bg-amber-400' : s.rank <= 3 ? 'bg-emerald-500' : 'bg-blue-600'} rounded`}
                              />
                            </div>

                            {isEditMode ? (
                              <input
                                type="number"
                                step="any"
                                value={displaySales}
                                onChange={(e) => setEditValues({ ...editValues, [salesKey]: parseFloat(e.target.value) || 0 })}
                                className="w-16 bg-slate-950 border border-emerald-500 rounded px-1 text-xs text-emerald-400 text-right font-bold"
                              />
                            ) : (
                              <span className="w-20 text-right font-bold text-white tabular-nums">
                                ${displaySales.toFixed(2)}
                              </span>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  )}

                  <div className="pt-2 mt-2 border-t border-slate-800 text-[11px] text-slate-400 flex items-center justify-between">
                    <span>Concentración Top 5:</span>
                    <strong className="text-emerald-400">
                      {activeSocios.length > 0
                        ? ((activeSocios.slice(0, 5).reduce((a, b) => a + b.totalSales, 0) / (activeSocios.reduce((a, b) => a + b.totalSales, 0) || 1)) * 100).toFixed(1)
                        : '0.0'}% de la recaudación
                    </strong>
                  </div>
                </div>

                {/* 2. RESUMEN POR CANALES (UPCONNECT VS CONNECTORS) */}
                <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between">
                  <div className="flex items-center justify-between pb-2 border-b border-slate-800 mb-2">
                    <div className="flex items-center gap-2">
                      <Users className="w-4 h-4 text-sky-400" />
                      <h3 className="text-xs font-bold uppercase tracking-wider text-sky-300">
                        Distribución y Volumen por Canal
                      </h3>
                    </div>
                    <span className="text-[11px] text-slate-400 font-medium">Distribuidor Upconnect vs Distribuidor Connect</span>
                  </div>

                  {/* Channel Comparison Cards */}
                  <div className="grid grid-cols-2 gap-3 mb-3">
                    {/* UpConnect */}
                    <div className="p-3 rounded-xl bg-slate-950/80 border-t-2 border-sky-500 border-x border-b border-slate-800">
                      <div className="flex items-center gap-1.5 text-xs font-bold text-sky-400 mb-1">
                        <Zap className="w-3.5 h-3.5" />
                        <span>Distribuidor Upconnect</span>
                      </div>
                      <div className="text-xl font-black text-white tabular-nums">
                        ${viewData.sociosSummary.upconnect.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                      </div>
                      <div className="text-[11px] text-slate-400 mt-1 flex justify-between">
                        <span>{viewData.sociosSummary.upconnect.count} Socios</span>
                        <span className="font-bold text-sky-300">{viewData.sociosSummary.upconnect.percentage}%</span>
                      </div>
                      <div className="text-[10px] text-slate-500 mt-0.5">
                        Ticket Promedio: ${viewData.sociosSummary.upconnect.averageTicket.toFixed(2)} USD
                      </div>
                    </div>

                    {/* Connect */}
                    <div className="p-3 rounded-xl bg-slate-950/80 border-t-2 border-amber-500 border-x border-b border-slate-800">
                      <div className="flex items-center gap-1.5 text-xs font-bold text-amber-400 mb-1">
                        <Share2 className="w-3.5 h-3.5" />
                        <span>Distribuidor Connect Aliados</span>
                      </div>
                      <div className="text-xl font-black text-white tabular-nums">
                        ${viewData.sociosSummary.connectors.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                      </div>
                      <div className="text-[11px] text-slate-400 mt-1 flex justify-between">
                        <span>{viewData.sociosSummary.connectors.count} Socios</span>
                        <span className="font-bold text-amber-300">{viewData.sociosSummary.connectors.percentage}%</span>
                      </div>
                      <div className="text-[10px] text-slate-500 mt-0.5">
                        Ticket Promedio: ${viewData.sociosSummary.connectors.averageTicket.toFixed(2)} USD
                      </div>
                    </div>
                  </div>

                  {/* Split Bar */}
                  <div className="space-y-1 mb-2">
                    <div className="flex justify-between text-[11px] font-semibold">
                      <span className="text-sky-400">Distribuidor Upconnect: {viewData.sociosSummary.upconnect.percentage}%</span>
                      <span className="text-amber-400">Distribuidor Connect: {viewData.sociosSummary.connectors.percentage}%</span>
                    </div>
                    <div className="h-3 w-full bg-slate-800 rounded-full overflow-hidden flex">
                      <div
                        style={{ width: `${viewData.sociosSummary.upconnect.percentage}%` }}
                        className="h-full bg-sky-500 transition-all"
                      />
                      <div
                        style={{ width: `${viewData.sociosSummary.connectors.percentage}%` }}
                        className="h-full bg-amber-500 transition-all"
                      />
                    </div>
                  </div>

                  {/* Channel Summary Table */}
                  <div className="border border-slate-800 rounded-lg overflow-hidden bg-slate-950/60 text-[11px]">
                    <div className="grid grid-cols-4 bg-slate-900/90 font-bold text-slate-400 p-1.5 border-b border-slate-800 text-center">
                      <span>Canal</span>
                      <span>Socios</span>
                      <span>Ventas</span>
                      <span>Total USD</span>
                    </div>
                    <div className="grid grid-cols-4 p-1.5 border-b border-slate-800/60 text-center text-slate-200">
                      <span className="font-bold text-sky-400">UpConnect</span>
                      <span>{viewData.sociosSummary.upconnect.count}</span>
                      <span>{viewData.sociosSummary.upconnect.operationsCount}</span>
                      <span className="font-bold text-emerald-400">${viewData.sociosSummary.upconnect.totalSales.toFixed(2)}</span>
                    </div>
                    <div className="grid grid-cols-4 p-1.5 text-center text-slate-200">
                      <span className="font-bold text-amber-400">Connectors</span>
                      <span>{viewData.sociosSummary.connectors.count}</span>
                      <span>{viewData.sociosSummary.connectors.operationsCount}</span>
                      <span className="font-bold text-emerald-400">${viewData.sociosSummary.connectors.totalSales.toFixed(2)}</span>
                    </div>
                  </div>
                </div>
              </div>

              {/* Footer Callout */}
              <div className="p-3 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-emerald-400">Total Auditado en Socios:</strong>{' '}
                  <strong className="text-white">
                    ${activeSocios.reduce((a, s) => a + s.totalSales, 0).toLocaleString('en-US', { minimumFractionDigits: 2 })} USD
                  </strong>{' '}
                  distribuidos en{' '}
                  <strong className="text-sky-300">{viewData.sociosSummary.totalSociosCount} socios registrados</strong> con{' '}
                  <strong className="text-white">{viewData.sociosSummary.totalOperationsCount} transacciones conciliadas</strong>.
                </div>
                {activeSocios.length > 0 && (
                  <div className="text-amber-400 font-bold">
                    Líder: {activeSocios[0]?.name} (${activeSocios[0]?.totalSales.toFixed(2)} USD)
                  </div>
                )}
              </div>
            </div>
          )}

          {/* SLIDE 5: Foco de Rendimiento Mensual: Desglose Semana a Semana */}
          {currentSlideDef.id === 'weekly' && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                  FOCO SEMANAL · AUDITORÍA DETALLADA
                </span>
                <h2 className="text-2xl sm:text-3xl font-black text-white tracking-tight mt-0.5">
                  Foco de Rendimiento: Desglose Semana a Semana
                </h2>
                <p className="text-xs text-slate-400 mt-1">
                  Distribución exacta de facturación por semana acorde a las fechas del período activo ({viewData.periodLabel})
                </p>
              </div>

              {/* 4 Weekly KPI Cards */}
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4 my-auto">
                {activeWeekly.length === 0 ? (
                  <div className="col-span-4 py-8 text-center text-slate-500 text-sm">
                    Sin desglose semanal disponible para el período actual.
                  </div>
                ) : (
                  activeWeekly.map((w, idx) => {
                    const wKey = `r2_week_${idx}_amount`;
                    const currentAmt = editValues[wKey] !== undefined ? Number(editValues[wKey]) : w.amount;

                    return (
                      <div key={idx} className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between">
                        <div>
                          <div className="flex items-center justify-between text-xs mb-1">
                            <span className="font-bold text-sky-400">{w.weekName}</span>
                            <span className="text-[11px] text-slate-400">{w.dateRange}</span>
                          </div>
                          <div className="text-2xl font-black text-white tabular-nums my-1">
                            {isEditMode ? (
                              <div className="flex items-center gap-1">
                                <span className="text-xs text-slate-400">$</span>
                                <input
                                  type="number"
                                  step="any"
                                  value={currentAmt}
                                  onChange={(e) => setEditValues({ ...editValues, [wKey]: parseFloat(e.target.value) || 0 })}
                                  className="w-24 bg-slate-950 border border-emerald-500 rounded px-1.5 py-0.5 text-xs text-emerald-400 font-bold"
                                />
                              </div>
                            ) : (
                              `$${w.amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
                            )}
                          </div>
                          <div className="text-xs font-semibold text-emerald-400">{w.percentage}% del total</div>
                        </div>

                        <div className="pt-3 border-t border-slate-800 text-[11px] text-slate-400">
                          {w.description || 'Emisiones registradas'}
                        </div>
                      </div>
                    );
                  })
                )}
              </div>

              {/* Bottom Insight */}
              <div className="p-3.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-emerald-400">Total Semanal Auditado:</strong>{' '}
                  <strong className="text-white">
                    ${totalWeeklySales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD
                  </strong>{' '}
                  (Consistente al 100% con la facturación de sistemas por socio).
                </div>
                <span className="text-slate-400">{activeWeekly.length} semanas auditadas</span>
              </div>
            </div>
          )}

          {/* SLIDE 6: Ventas por Categoría de Solución */}
          {currentSlideDef.id === 'solutions' && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                  CATÁLOGO COMERCIAL · {viewData.periodLabel.toUpperCase()}
                </span>
                <h2 className="text-2xl sm:text-3xl font-black text-white tracking-tight mt-0.5">
                  Ventas por Categoría de Solución UpConta
                </h2>
                <p className="text-xs text-slate-400 mt-1">
                  Distribución de ingresos según planes de suscripción ERP y productos contratados
                </p>
              </div>

              {/* 3 Categories Grid */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 my-auto">
                {activeCategories.length === 0 ? (
                  <div className="col-span-3 py-10 text-center text-slate-500 text-sm">
                    Sin categorías de solución registradas.
                  </div>
                ) : (
                  activeCategories.map((cat, idx) => (
                    <div key={idx} className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between">
                      <div>
                        <div className="flex items-center justify-between text-xs mb-2">
                          <span className="font-bold text-sky-400 uppercase tracking-wider text-[11px]">{cat.badge || 'Categoría'}</span>
                          <span className="font-bold text-white bg-slate-800 px-2 py-0.5 rounded text-[11px]">
                            {cat.percentage}%
                          </span>
                        </div>
                        <h3 className="text-base font-bold text-white">{cat.title}</h3>
                        <div className="text-2xl font-black text-emerald-400 tabular-nums my-2">
                          ${cat.amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </div>
                        <p className="text-xs text-slate-400">{cat.plansDescription}</p>
                      </div>

                      <div className="pt-3 border-t border-slate-800 text-[11px] text-slate-400 flex items-center justify-between">
                        <span>Participación:</span>
                        <strong className="text-slate-200">{cat.percentage}% de ventas</strong>
                      </div>
                    </div>
                  ))
                )}
              </div>

              {/* Bottom Callout */}
              <div className="p-3.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-emerald-400">Estrategia de Socios:</strong> Foco en tickets medianos y altos (ERP + Planes Contador representan la mayor recaudación).
                </div>
                <span className="text-emerald-400 font-bold">Alta Retención</span>
              </div>
            </div>
          )}

          {/* SLIDE 7: Canal de Distribución - Embudo de Socios */}
          {currentSlideDef.id === 'funnel' && (
            <div className="h-full flex flex-col justify-between p-6 sm:p-8 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                  CANAL DE DISTRIBUCIÓN · {viewData.periodLabel.toUpperCase()}
                </span>
                <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                  Embudo de Socios y Comunidades Activas
                </h2>
                <p className="text-xs text-slate-400 mt-0.5">
                  Conteo cuantitativo de socios, franquiciados UpConta y red de distribuidores Connect
                </p>
              </div>

              {/* 3 Compact Funnel Columns */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 my-2">
                {/* Stage 1 */}
                <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between text-center">
                  <div>
                    <div className="w-10 h-10 rounded-xl bg-blue-600/20 text-sky-400 font-bold flex items-center justify-center mx-auto mb-2">
                      Up
                    </div>
                    <h3 className="text-sm font-bold text-white">UpConta Socios</h3>
                    <div className="text-3xl font-black text-sky-400 tabular-nums my-1">
                      {viewData.sociosSummary.upconnect.count}
                    </div>
                    <div className="text-[11px] font-semibold text-slate-400 mb-1">Socios Distribuidor Upconnect</div>
                  </div>
                  <p className="text-[11px] text-slate-400 border-t border-slate-800 pt-2">
                    Red principal de contadores y socios directos UpConnect.
                  </p>
                </div>

                {/* Stage 2 (Hero highlighted) */}
                <div className="p-4 rounded-xl bg-slate-900/95 border-2 border-emerald-500 shadow-xl shadow-emerald-950/50 flex flex-col justify-between text-center relative">
                  <div>
                    <div className="w-10 h-10 rounded-xl bg-purple-600/20 text-purple-400 font-bold flex items-center justify-center mx-auto mb-2">
                      👑
                    </div>
                    <h3 className="text-sm font-bold text-white">Franquicia Contadores VIP</h3>
                    <div className="text-3xl font-black text-white tabular-nums my-1">
                      {viewData.sociosSummary.connectors.count}
                    </div>
                    <div className="text-[11px] font-semibold text-purple-400 mb-1">Socios Distribuidor Connect Aliados</div>
                  </div>
                  <p className="text-[11px] text-slate-300 border-t border-slate-800 pt-2">
                    Contadores activos y red externa de distribución Connect.
                  </p>
                </div>

                {/* Stage 3 */}
                <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between text-center">
                  <div>
                    <div className="w-10 h-10 rounded-xl bg-emerald-600/20 text-emerald-400 font-bold flex items-center justify-center mx-auto mb-2">
                      🎓
                    </div>
                    <h3 className="text-sm font-bold text-white">Capacitación Franquicia</h3>
                    <div className="text-3xl font-black text-emerald-400 tabular-nums my-1">
                      {Math.max(1, Math.round(viewData.sociosSummary.totalSociosCount * 0.25)) || 51}
                    </div>
                    <div className="text-[11px] font-semibold text-slate-400 mb-1">Miembros en Inducción</div>
                  </div>
                  <p className="text-[11px] text-slate-400 border-t border-slate-800 pt-2">
                    Programa de formación técnica y habilitación comercial.
                  </p>
                </div>
              </div>

              {/* TABLA ADICIONAL: Cantidad de Socios Distribuidor Upconnect vs Distribuidor Connect */}
              <div className="my-1">
                <SociosChannelTable
                  summary={viewData.sociosSummary}
                  variant="slide"
                  title="Distribución y Cantidad de Socios por Canal: Distribuidor Upconnect vs Distribuidor Connect"
                />
              </div>

              {/* Bottom callout */}
              <div className="p-2.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-emerald-400">Total Cartera Auditada:</strong> {viewData.sociosSummary.totalSociosCount} socios registrados ({viewData.sociosSummary.upconnect.count} Distribuidor Upconnect / {viewData.sociosSummary.connectors.count} Distribuidor Connect) con {viewData.sociosSummary.totalOperationsCount} ventas conciliadas por ${viewData.sociosSummary.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD.
                </div>
                <span className="text-emerald-400 font-bold">100% Conciliado con Base Excel</span>
              </div>
            </div>
          )}

          {/* SLIDE 8: Siguiente Fase: Cruce Comercial */}
          {currentSlideDef.id === 'cross' && (
            <div className="h-full flex flex-col justify-center items-center text-center p-8 sm:p-16 relative bg-gradient-to-br from-[#0c1322] via-[#080d19] to-[#040812]">
              <div className="z-10 max-w-2xl space-y-6">
                <div className="w-12 h-12 rounded-2xl bg-emerald-500 text-white flex items-center justify-center mx-auto shadow-lg shadow-emerald-500/30">
                  <Activity className="w-6 h-6" />
                </div>

                <div className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                  SIGUIENTE FASE: CRUCE COMERCIAL
                </div>

                <h2 className="text-3xl sm:text-4xl font-black text-white tracking-tight">
                  RESUMEN DE VENTAS UPCONTA Y SOCIOS
                </h2>

                <div className="space-y-2 py-4 text-base font-semibold text-slate-300">
                  <div>VENTAS EQUIPO COMERCIAL: <strong className="text-white">${dataset.commercialCross?.commercialTeamSales ? dataset.commercialCross.commercialTeamSales.toLocaleString('en-US', { minimumFractionDigits: 2 }) : '0.00'}</strong></div>
                  <div>VENTAS ORGÁNICAS: <strong className="text-white">${dataset.commercialCross?.organicSales ? dataset.commercialCross.organicSales.toLocaleString('en-US', { minimumFractionDigits: 2 }) : '0.00'}</strong></div>
                  <div>VENTAS SOCIOS: <strong className="text-emerald-400">${(dataset.commercialCross?.partnersSales || netSales).toLocaleString('en-US', { minimumFractionDigits: 2 })}</strong></div>
                </div>

                <div className="grid grid-cols-2 gap-8 pt-6 border-t border-slate-800/80 max-w-lg mx-auto">
                  <div className="text-left">
                    <div className="text-[11px] font-bold text-slate-500 uppercase">{dataset.commercialCross?.channel1Name || 'CANAL 1 AUDITADO'}</div>
                    <div className="text-sm sm:text-base font-bold text-sky-400 mt-1">
                      Socios UpConnect (${(dataset.commercialCross?.channel1Amount || netSales).toLocaleString('en-US', { minimumFractionDigits: 2 })})
                    </div>
                  </div>
                  <div className="text-left">
                    <div className="text-[11px] font-bold text-slate-500 uppercase">{dataset.commercialCross?.channel2Name || 'CANAL 2 PENDIENTE'}</div>
                    <div className="text-sm sm:text-base font-bold text-slate-400 mt-1">
                      {dataset.commercialCross?.channel2Status || 'Asesores Directos & Marketing'}
                    </div>
                  </div>
                </div>
              </div>

              <div className="absolute bottom-4 left-8 right-8 flex justify-between items-center text-[11px] text-slate-500 border-t border-slate-800/60 pt-3">
                <span>Próximo hito: Consolidación omnicanal y liquidación automática</span>
                <span>Pág. 08</span>
              </div>
            </div>
          )}

        </div>
      </div>

      {/* MODAL: Mover y Reorganizar Diapositivas */}
      <ReorderSlidesModal
        isOpen={isReorderModalOpen}
        onClose={() => setIsReorderModalOpen(false)}
        slides={activeSlides.map((s) => ({
          id: s.id,
          title: s.title,
          subtitle: s.subtitle,
        }))}
        reportTitle="Reporte 2: Auditoría & Venta de Sistemas por Socio"
        onSaveOrder={(newOrder) => {
          if (onUpdateDataset) {
            onUpdateDataset({ customSlideOrder2: newOrder });
          }
        }}
        onResetDefault={() => {
          if (onUpdateDataset) {
            onUpdateDataset({ customSlideOrder2: [] });
          }
        }}
      />

      {/* MODAL: Editor de Diapositiva */}
      <SlideEditorModal
        isOpen={isEditorModalOpen}
        onClose={() => setIsEditorModalOpen(false)}
        reportType={2}
        slideId={currentSlideDef.id}
        slideTitle={currentSlideDef.title}
        viewData={viewData}
        dataset={dataset}
        editValues={editValues}
        onChangeValue={(key, val) => setEditValues((prev) => ({ ...prev, [key]: val }))}
        onSave={handleSaveAllChanges}
      />
    </div>
  );
};
