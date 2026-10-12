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
  Share2,
  RotateCcw,
  DollarSign,
  MonitorPlay
} from 'lucide-react';

interface Props {
  dataset: GlobalDataset;
  onBackToDashboard?: () => void;
  onUpdateDataset?: (updated: Partial<GlobalDataset>) => void;
  onOpenOnlineReport?: () => void;
}

interface SlideConfig2 {
  id: 'cover' | 'balance' | 'monthly' | 'socios' | 'weekly' | 'solutions' | 'funnel' | 'cross';
  title: string;
  subtitle?: string;
}

export const PresentationType2: React.FC<Props> = ({ dataset, onBackToDashboard, onUpdateDataset, onOpenOnlineReport }) => {
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
      await exportPresentationType2ToPPTX(dataset, customCommunities);
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

  const activeMonthlyMetrics = useMemo(() => {
    const raw = (viewData.sistemasMonthlyMetrics && viewData.sistemasMonthlyMetrics.length > 0)
      ? viewData.sistemasMonthlyMetrics
      : (dataset.sistemasMonthlyMetrics || []);
    if (dataset.reportFilter?.type === 'month' && dataset.reportFilter.month) {
      const targetMonth = dataset.reportFilter.month.toUpperCase();
      const filtered = raw.filter((m) => m.month.toUpperCase() === targetMonth);
      if (filtered.length > 0) return filtered;
      return [{
        month: targetMonth,
        upconnectSales: viewData.sociosSummary.upconnect.totalSales,
        connectorsSales: viewData.sociosSummary.connectors.totalSales,
        upconnectCount: viewData.sociosSummary.upconnect.operationsCount,
        connectorsCount: viewData.sociosSummary.connectors.operationsCount,
      }];
    }
    return raw;
  }, [viewData.sistemasMonthlyMetrics, dataset.sistemasMonthlyMetrics, dataset.reportFilter, viewData.sociosSummary]);

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

  const activeCategories: SolutionCategory[] = (viewData.solutionCategoriesReport2 && viewData.solutionCategoriesReport2.length > 0)
    ? viewData.solutionCategoriesReport2
    : (dataset.solutionCategories || []);

  // Weekly Line Chart State (Pestaña 5)
  const [weeklyMetric, setWeeklyMetric] = useState<'sales' | 'count'>('sales');
  const [hoveredWeekIndex, setHoveredWeekIndex] = useState<number | null>(null);

  // Communities / Embudo State & Calculations (Pestaña 7)
  const autoUpSocios = viewData.sociosSummary.upconnect.count || 118;
  const autoCoSocios = viewData.sociosSummary.connectors.count || 129;
  const autoFormacion = Math.max(1, Math.round((autoUpSocios + autoCoSocios) * 0.25)) || 51;

  const [customCommunities, setCustomCommunities] = useState<{
    upcontaSocios?: number;
    franquiciaVIP?: number;
    formacionComercial?: number;
  }>(() => {
    const ov = dataset.customOverrides || {};
    return {
      upcontaSocios: ov['r2_upcontaSocios'] !== undefined ? Number(ov['r2_upcontaSocios']) : (ov['upcontaSocios'] !== undefined ? Number(ov['upcontaSocios']) : undefined),
      franquiciaVIP: ov['r2_franquiciaVIP'] !== undefined ? Number(ov['r2_franquiciaVIP']) : (ov['franquiciaVIP'] !== undefined ? Number(ov['franquiciaVIP']) : undefined),
      formacionComercial: ov['r2_formacionComercial'] !== undefined ? Number(ov['r2_formacionComercial']) : (ov['formacionComercial'] !== undefined ? Number(ov['formacionComercial']) : undefined),
    };
  });

  useEffect(() => {
    const ov = dataset.customOverrides || {};
    if (ov['r2_upcontaSocios'] !== undefined || ov['r2_franquiciaVIP'] !== undefined || ov['r2_formacionComercial'] !== undefined ||
        ov['upcontaSocios'] !== undefined || ov['franquiciaVIP'] !== undefined || ov['formacionComercial'] !== undefined) {
      setCustomCommunities((prev) => ({
        ...prev,
        upcontaSocios: ov['r2_upcontaSocios'] !== undefined ? Number(ov['r2_upcontaSocios']) : (ov['upcontaSocios'] !== undefined ? Number(ov['upcontaSocios']) : prev.upcontaSocios),
        franquiciaVIP: ov['r2_franquiciaVIP'] !== undefined ? Number(ov['r2_franquiciaVIP']) : (ov['franquiciaVIP'] !== undefined ? Number(ov['franquiciaVIP']) : prev.franquiciaVIP),
        formacionComercial: ov['r2_formacionComercial'] !== undefined ? Number(ov['r2_formacionComercial']) : (ov['formacionComercial'] !== undefined ? Number(ov['formacionComercial']) : prev.formacionComercial),
      }));
    }
  }, [dataset.customOverrides]);

  const [isEditingCommunities, setIsEditingCommunities] = useState<boolean>(false);
  const [tempUpconta, setTempUpconta] = useState<number>(0);
  const [tempFranquicia, setTempFranquicia] = useState<number>(0);
  const [tempFormacion, setTempFormacion] = useState<number>(0);

  const upcontaMembers = editValues['r2_upcontaSocios'] !== undefined
    ? Number(editValues['r2_upcontaSocios'])
    : (customCommunities.upcontaSocios ?? (autoUpSocios > 0 ? autoUpSocios : 118));

  const franquiciaMembers = editValues['r2_franquiciaVIP'] !== undefined
    ? Number(editValues['r2_franquiciaVIP'])
    : (customCommunities.franquiciaVIP ?? (autoCoSocios > 0 ? autoCoSocios : 129));

  const formacionMembers = editValues['r2_formacionComercial'] !== undefined
    ? Number(editValues['r2_formacionComercial'])
    : (customCommunities.formacionComercial ?? (autoCoSocios > 0 ? autoFormacion : 51));

  const isManuallyModified =
    customCommunities.upcontaSocios !== undefined ||
    customCommunities.franquiciaVIP !== undefined ||
    customCommunities.formacionComercial !== undefined ||
    editValues['r2_upcontaSocios'] !== undefined ||
    editValues['r2_franquiciaVIP'] !== undefined ||
    editValues['r2_formacionComercial'] !== undefined ||
    editValues['upcontaSocios'] !== undefined ||
    editValues['franquiciaVIP'] !== undefined ||
    editValues['formacionComercial'] !== undefined;

  const handleSaveCommunitiesModal = (up: number, vip: number, form: number) => {
    const updated = {
      r2_upcontaSocios: up,
      r2_franquiciaVIP: vip,
      r2_formacionComercial: form,
      upcontaSocios: up,
      franquiciaVIP: vip,
      formacionComercial: form,
    };
    setCustomCommunities(updated);
    setEditValues((prev) => ({
      ...prev,
      ...updated,
    }));
    if (onUpdateDataset) {
      onUpdateDataset({
        customOverrides: {
          ...(dataset.customOverrides || {}),
          ...updated,
        },
      });
    }
    setIsEditingCommunities(false);
  };

  const handleResetCommunities = () => {
    setCustomCommunities({});
    const newOverrides = { ...(dataset.customOverrides || {}) };
    delete newOverrides.r2_upcontaSocios;
    delete newOverrides.r2_franquiciaVIP;
    delete newOverrides.r2_formacionComercial;
    delete newOverrides.upcontaSocios;
    delete newOverrides.franquiciaVIP;
    delete newOverrides.formacionComercial;
    setEditValues(newOverrides);
    if (onUpdateDataset) {
      onUpdateDataset({ customOverrides: newOverrides });
    }
  };

  // Dynamic Communities Channel Table Summary for Slide 7:
  // Updates immediately when upcontaMembers or franquiciaMembers change
  const communitiesTableSummary = useMemo(() => {
    const upCount = upcontaMembers;
    const coCount = franquiciaMembers;
    const totalMembers = upCount + coCount;
    const upPct = totalMembers > 0 ? parseFloat(((upCount / totalMembers) * 100).toFixed(1)) : 0;
    const coPct = totalMembers > 0 ? parseFloat(((coCount / totalMembers) * 100).toFixed(1)) : 0;

    const baseSummary = viewData.sociosSummary;
    const upOps = baseSummary.upconnect.operationsCount;
    const coOps = baseSummary.connectors.operationsCount;
    const totalOps = baseSummary.totalOperationsCount;

    const upNetSales = baseSummary.upconnect.totalSales;
    const coNetSales = baseSummary.connectors.totalSales;
    const allSales = baseSummary.totalSales;

    const upAvgTicket = upOps > 0 ? parseFloat((upNetSales / upOps).toFixed(2)) : (upCount > 0 ? parseFloat((upNetSales / upCount).toFixed(2)) : 0);
    const coAvgTicket = coOps > 0 ? parseFloat((coNetSales / coOps).toFixed(2)) : (coCount > 0 ? parseFloat((coNetSales / coCount).toFixed(2)) : 0);
    const avgTicket = totalOps > 0 ? parseFloat((allSales / totalOps).toFixed(2)) : 0;

    return {
      upconnect: {
        channel: 'UpConnect' as const,
        channelLabel: 'Distribuidor Upconnect',
        count: upCount,
        percentage: upPct,
        operationsCount: upOps,
        totalSales: upNetSales,
        averageTicket: upAvgTicket,
        socios: activeSocios.filter((s) => s.channel === 'UpConnect'),
      },
      connectors: {
        channel: 'Connectors' as const,
        channelLabel: 'Distribuidor Connect',
        count: coCount,
        percentage: coPct,
        operationsCount: coOps,
        totalSales: coNetSales,
        averageTicket: coAvgTicket,
        socios: activeSocios.filter((s) => s.channel === 'Connectors'),
      },
      totalSociosCount: totalMembers,
      totalOperationsCount: totalOps,
      totalSales: allSales,
      averageTicket: avgTicket,
    };
  }, [upcontaMembers, franquiciaMembers, viewData.sociosSummary, activeSocios]);

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

          {onOpenOnlineReport && (
            <button
              onClick={onOpenOnlineReport}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded bg-sky-600 hover:bg-sky-500 text-white shadow-sm transition-colors"
              title="Ver presentación unificada interactiva en línea con ambos reportes"
            >
              <MonitorPlay className="w-4 h-4 text-sky-200" />
              <span>Reporte en Línea</span>
            </button>
          )}

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
                  Período Activo: <strong className="text-slate-300">{viewData.periodLabel}</strong> · Fecha de Corte: <strong className="text-slate-300">{viewData.cutoffDateReport2}</strong>
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
                  <strong className="text-white">{viewData.periodLabel}</strong> registra un promedio de{' '}
                  <strong className="text-white">
                    ${((netSales / Math.max(1, activeSocios.reduce((a, s) => a + (s.operationsCount || 1), 0))) || 0).toFixed(2)} USD
                  </strong>{' '}
                  por transacción sobre una cartera auditada de{' '}
                  <strong className="text-sky-400">{activeSocios.length} socios comerciales</strong>.
                </div>
                <span className="text-slate-500">Corte: {viewData.cutoffDateReport2}</span>
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
                  <strong className="text-emerald-400">Total {viewData.isFiltered ? `del Período (${viewData.periodLabel})` : 'Acumulado Registrado'}:</strong>{' '}
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
                        <span>Distribuidor Connect</span>
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
                      <span className="font-bold text-amber-400">Connect</span>
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

          {/* SLIDE 5: Foco de Rendimiento Mensual: Desglose Semana a Semana con Gráficos Estadísticos Lineales */}
          {currentSlideDef.id === 'weekly' && (() => {
            const diffs = (viewData.weeklyDifferentiators && viewData.weeklyDifferentiators.length > 0)
              ? viewData.weeklyDifferentiators
              : [
                  { weekId: 'w1', weekName: 'Semana 1', dateRange: '01 al 06', totalSales: 0, totalCount: 0, facturacionSales: 0, facturacionCount: 0, erpSales: 0, erpCount: 0, contadoresSales: 0, contadoresCount: 0 },
                  { weekId: 'w2', weekName: 'Semana 2', dateRange: '07 al 13', totalSales: 0, totalCount: 0, facturacionSales: 0, facturacionCount: 0, erpSales: 0, erpCount: 0, contadoresSales: 0, contadoresCount: 0 },
                  { weekId: 'w3', weekName: 'Semana 3', dateRange: '14 al 20', totalSales: 0, totalCount: 0, facturacionSales: 0, facturacionCount: 0, erpSales: 0, erpCount: 0, contadoresSales: 0, contadoresCount: 0 },
                  { weekId: 'w4', weekName: 'Semana 4', dateRange: '21 al 31', totalSales: 0, totalCount: 0, facturacionSales: 0, facturacionCount: 0, erpSales: 0, erpCount: 0, contadoresSales: 0, contadoresCount: 0 },
                ];

            const totFact = diffs.reduce((a, b) => a + (weeklyMetric === 'sales' ? b.facturacionSales : b.facturacionCount), 0);
            const totErp = diffs.reduce((a, b) => a + (weeklyMetric === 'sales' ? b.erpSales : b.erpCount), 0);
            const totCont = diffs.reduce((a, b) => a + (weeklyMetric === 'sales' ? b.contadoresSales : b.contadoresCount), 0);
            const totCombined = totFact + totErp + totCont || 1;

            // Coordinate calculations for responsive SVG
            const svgW = 760;
            const svgH = 175;
            const padL = 65;
            const padR = 40;
            const padT = 22;
            const padB = 35;
            const plotW = svgW - padL - padR;
            const plotH = svgH - padT - padB;

            const xPoints = diffs.map((_, i) => padL + (i / Math.max(1, diffs.length - 1)) * plotW);

            const allVals = diffs.flatMap((d) => 
              weeklyMetric === 'sales' 
                ? [d.facturacionSales, d.erpSales, d.contadoresSales]
                : [d.facturacionCount, d.erpCount, d.contadoresCount]
            );
            const maxVal = Math.max(...allVals, 10);
            const scaleMax = maxVal * 1.18;

            const getY = (val: number) => padT + plotH - (val / scaleMax) * plotH;

            const factPoints = diffs.map((d, i) => ({
              x: xPoints[i],
              y: getY(weeklyMetric === 'sales' ? d.facturacionSales : d.facturacionCount),
              val: weeklyMetric === 'sales' ? d.facturacionSales : d.facturacionCount,
            }));

            const erpPoints = diffs.map((d, i) => ({
              x: xPoints[i],
              y: getY(weeklyMetric === 'sales' ? d.erpSales : d.erpCount),
              val: weeklyMetric === 'sales' ? d.erpSales : d.erpCount,
            }));

            const contPoints = diffs.map((d, i) => ({
              x: xPoints[i],
              y: getY(weeklyMetric === 'sales' ? d.contadoresSales : d.contadoresCount),
              val: weeklyMetric === 'sales' ? d.contadoresSales : d.contadoresCount,
            }));

            const makePolyline = (pts: { x: number; y: number }[]) => pts.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
            const makeArea = (pts: { x: number; y: number }[]) => {
              const base = `${pts[0].x.toFixed(1)},${(padT + plotH).toFixed(1)}`;
              const line = pts.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
              const end = `${pts[pts.length - 1].x.toFixed(1)},${(padT + plotH).toFixed(1)}`;
              return `${base} ${line} ${end}`;
            };

            const gridSteps = [0, 0.33, 0.66, 1];

            return (
              <div className="h-full flex flex-col justify-between p-6 sm:p-8 relative bg-[#0c1322]">
                {/* Header */}
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-800/80 pb-2.5">
                  <div>
                    <span className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                      COMPARATIVA SEMANA A SEMANA · AUDITORÍA POR DIFERENCIADOR
                    </span>
                    <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                      Foco de Rendimiento: Desglose y Gráficos Lineales
                    </h2>
                    <p className="text-xs text-slate-400 mt-0.5">
                      Comparativa estadística semanal según diferenciadores: Planes de Facturación, ERP y Contadores ({viewData.periodLabel})
                    </p>
                  </div>

                  {/* Metric Switcher */}
                  <div className="flex items-center gap-2 bg-slate-900/90 p-1 rounded-xl border border-slate-800 text-xs">
                    <button
                      onClick={() => setWeeklyMetric('sales')}
                      className={`px-3 py-1 rounded-lg font-semibold flex items-center gap-1.5 transition-all ${
                        weeklyMetric === 'sales'
                          ? 'bg-emerald-500 text-white shadow-md shadow-emerald-500/30'
                          : 'text-slate-400 hover:text-white'
                      }`}
                    >
                      <DollarSign className="w-3.5 h-3.5" />
                      <span>Facturación ($ USD)</span>
                    </button>
                    <button
                      onClick={() => setWeeklyMetric('count')}
                      className={`px-3 py-1 rounded-lg font-semibold flex items-center gap-1.5 transition-all ${
                        weeklyMetric === 'count'
                          ? 'bg-sky-500 text-white shadow-md shadow-sky-500/30'
                          : 'text-slate-400 hover:text-white'
                      }`}
                    >
                      <Activity className="w-3.5 h-3.5" />
                      <span>Volumen (Operaciones)</span>
                    </button>
                  </div>
                </div>

                {/* Linear Chart Section */}
                <div className="bg-slate-900/90 rounded-2xl border border-slate-800 p-4 shadow-xl my-2 flex flex-col justify-between">
                  {/* Chart Legend & KPI Summary */}
                  <div className="flex flex-wrap items-center justify-between gap-3 pb-2 border-b border-slate-800/60 text-xs">
                    <div className="flex items-center gap-4">
                      {/* Facturación */}
                      <div className="flex items-center gap-1.5">
                        <span className="w-3 h-1 rounded bg-sky-400"></span>
                        <span className="text-slate-300 font-medium">Planes de Facturación:</span>
                        <strong className="text-sky-300 font-bold tabular-nums">
                          {weeklyMetric === 'sales' ? `$${totFact.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : `${totFact} ops`}
                        </strong>
                        <span className="text-[10px] text-slate-500 font-semibold">({((totFact / totCombined) * 100).toFixed(1)}%)</span>
                      </div>

                      {/* ERP */}
                      <div className="flex items-center gap-1.5">
                        <span className="w-3 h-1 rounded bg-emerald-400"></span>
                        <span className="text-slate-300 font-medium">Sistemas ERP:</span>
                        <strong className="text-emerald-300 font-bold tabular-nums">
                          {weeklyMetric === 'sales' ? `$${totErp.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : `${totErp} ops`}
                        </strong>
                        <span className="text-[10px] text-slate-500 font-semibold">({((totErp / totCombined) * 100).toFixed(1)}%)</span>
                      </div>

                      {/* Contadores */}
                      <div className="flex items-center gap-1.5">
                        <span className="w-3 h-1 rounded bg-purple-400"></span>
                        <span className="text-slate-300 font-medium">Planes Contadores:</span>
                        <strong className="text-purple-300 font-bold tabular-nums">
                          {weeklyMetric === 'sales' ? `$${totCont.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : `${totCont} ops`}
                        </strong>
                        <span className="text-[10px] text-slate-500 font-semibold">({((totCont / totCombined) * 100).toFixed(1)}%)</span>
                      </div>
                    </div>

                    <div className="text-[11px] text-slate-400">
                      Pasa el cursor sobre cada semana para ver la comparativa detallada
                    </div>
                  </div>

                  {/* SVG Multi-Line Statistical Chart */}
                  <div className="relative w-full h-[175px] mt-2">
                    <svg viewBox={`0 0 ${svgW} ${svgH}`} className="w-full h-full overflow-visible">
                      <defs>
                        <linearGradient id="gradFact" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="#38bdf8" stopOpacity="0.25" />
                          <stop offset="100%" stopColor="#38bdf8" stopOpacity="0.0" />
                        </linearGradient>
                        <linearGradient id="gradErp" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="#10b981" stopOpacity="0.25" />
                          <stop offset="100%" stopColor="#10b981" stopOpacity="0.0" />
                        </linearGradient>
                        <linearGradient id="gradCont" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="#a855f7" stopOpacity="0.25" />
                          <stop offset="100%" stopColor="#a855f7" stopOpacity="0.0" />
                        </linearGradient>
                      </defs>

                      {/* Horizontal Grid Lines & Y Axis Labels */}
                      {gridSteps.map((step, sIdx) => {
                        const val = scaleMax * step;
                        const y = padT + plotH - step * plotH;
                        const label = weeklyMetric === 'sales'
                          ? `$${Math.round(val).toLocaleString('en-US')}`
                          : `${Math.round(val)}`;
                        return (
                          <g key={sIdx}>
                            <line
                              x1={padL}
                              x2={svgW - padR}
                              y1={y}
                              y2={y}
                              stroke="#1e293b"
                              strokeDasharray="3 3"
                              strokeWidth="1"
                            />
                            <text
                              x={padL - 10}
                              y={y + 3.5}
                              textAnchor="end"
                              fill="#64748b"
                              fontSize="9.5"
                              fontFamily="monospace"
                            >
                              {label}
                            </text>
                          </g>
                        );
                      })}

                      {/* Gradient Areas */}
                      <polygon points={makeArea(factPoints)} fill="url(#gradFact)" />
                      <polygon points={makeArea(erpPoints)} fill="url(#gradErp)" />
                      <polygon points={makeArea(contPoints)} fill="url(#gradCont)" />

                      {/* Polylines */}
                      <polyline
                        points={makePolyline(factPoints)}
                        fill="none"
                        stroke="#38bdf8"
                        strokeWidth="2.5"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      />
                      <polyline
                        points={makePolyline(erpPoints)}
                        fill="none"
                        stroke="#10b981"
                        strokeWidth="2.5"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      />
                      <polyline
                        points={makePolyline(contPoints)}
                        fill="none"
                        stroke="#a855f7"
                        strokeWidth="2.5"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      />

                      {/* Hover guideline */}
                      {hoveredWeekIndex !== null && (
                        <line
                          x1={xPoints[hoveredWeekIndex]}
                          x2={xPoints[hoveredWeekIndex]}
                          y1={padT}
                          y2={padT + plotH}
                          stroke="#94a3b8"
                          strokeDasharray="2 2"
                          strokeWidth="1.5"
                          opacity="0.8"
                        />
                      )}

                      {/* Dots & X-axis Labels */}
                      {diffs.map((d, idx) => {
                        const isHovered = hoveredWeekIndex === idx;
                        const fp = factPoints[idx];
                        const ep = erpPoints[idx];
                        const cp = contPoints[idx];

                        return (
                          <g key={idx}>
                            {/* Hover transparent hit column */}
                            <rect
                              x={xPoints[idx] - 40}
                              y={padT - 10}
                              width="80"
                              height={plotH + 40}
                              fill="transparent"
                              className="cursor-pointer"
                              onMouseEnter={() => setHoveredWeekIndex(idx)}
                              onMouseLeave={() => setHoveredWeekIndex(null)}
                            />

                            {/* Facturación Dot */}
                            <circle
                              cx={fp.x}
                              cy={fp.y}
                              r={isHovered ? 6 : 4}
                              fill="#0f172a"
                              stroke="#38bdf8"
                              strokeWidth={isHovered ? 3 : 2}
                              className="transition-all"
                            />

                            {/* ERP Dot */}
                            <circle
                              cx={ep.x}
                              cy={ep.y}
                              r={isHovered ? 6 : 4}
                              fill="#0f172a"
                              stroke="#10b981"
                              strokeWidth={isHovered ? 3 : 2}
                              className="transition-all"
                            />

                            {/* Contadores Dot */}
                            <circle
                              cx={cp.x}
                              cy={cp.y}
                              r={isHovered ? 6 : 4}
                              fill="#0f172a"
                              stroke="#a855f7"
                              strokeWidth={isHovered ? 3 : 2}
                              className="transition-all"
                            />

                            {/* X Axis Labels */}
                            <text
                              x={xPoints[idx]}
                              y={padT + plotH + 15}
                              textAnchor="middle"
                              fill={isHovered ? '#ffffff' : '#cbd5e1'}
                              fontSize="11"
                              fontWeight={isHovered ? 'bold' : '600'}
                            >
                              {d.weekName}
                            </text>
                            <text
                              x={xPoints[idx]}
                              y={padT + plotH + 28}
                              textAnchor="middle"
                              fill="#64748b"
                              fontSize="9.5"
                            >
                              {d.dateRange}
                            </text>
                          </g>
                        );
                      })}
                    </svg>

                    {/* Floating Tooltip when hovering */}
                    {hoveredWeekIndex !== null && diffs[hoveredWeekIndex] && (() => {
                      const hw = diffs[hoveredWeekIndex];
                      const leftPos = Math.min(Math.max(xPoints[hoveredWeekIndex] - 80, 20), svgW - 200);
                      return (
                        <div
                          style={{ left: `${(leftPos / svgW) * 100}%`, top: '-8px' }}
                          className="absolute z-20 pointer-events-none bg-slate-950/95 border border-slate-700 rounded-xl p-2.5 shadow-2xl backdrop-blur-md text-[11px] space-y-1 w-48 transition-all"
                        >
                          <div className="font-bold text-white border-b border-slate-800 pb-1 flex justify-between">
                            <span>{hw.weekName}</span>
                            <span className="text-slate-400 font-normal">{hw.dateRange}</span>
                          </div>
                          <div className="flex justify-between items-center text-sky-400">
                            <span>Facturación:</span>
                            <strong className="font-mono">
                              {weeklyMetric === 'sales' ? `$${hw.facturacionSales.toFixed(2)}` : `${hw.facturacionCount} ops`}
                            </strong>
                          </div>
                          <div className="flex justify-between items-center text-emerald-400">
                            <span>Sistemas ERP:</span>
                            <strong className="font-mono">
                              {weeklyMetric === 'sales' ? `$${hw.erpSales.toFixed(2)}` : `${hw.erpCount} ops`}
                            </strong>
                          </div>
                          <div className="flex justify-between items-center text-purple-400">
                            <span>Contadores:</span>
                            <strong className="font-mono">
                              {weeklyMetric === 'sales' ? `$${hw.contadoresSales.toFixed(2)}` : `${hw.contadoresCount} ops`}
                            </strong>
                          </div>
                          <div className="border-t border-slate-800 pt-1 flex justify-between items-center font-bold text-white">
                            <span>Total Semana:</span>
                            <span className="text-emerald-400 font-mono">
                              {weeklyMetric === 'sales' ? `$${hw.totalSales.toFixed(2)} USD` : `${hw.totalCount} ops`}
                            </span>
                          </div>
                        </div>
                      );
                    })()}
                  </div>
                </div>

                {/* 4 Weekly KPI Cards */}
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 my-1">
                  {diffs.map((w, idx) => {
                    const wKey = `r2_week_${idx}_amount`;
                    const currentAmt = editValues[wKey] !== undefined ? Number(editValues[wKey]) : w.totalSales;
                    const weekPct = totalWeeklySales > 0 ? ((w.totalSales / totalWeeklySales) * 100).toFixed(1) : '25.0';

                    return (
                      <div
                        key={idx}
                        className={`p-3 rounded-xl border transition-all ${
                          hoveredWeekIndex === idx
                            ? 'bg-slate-900 border-sky-500/60 shadow-lg'
                            : 'bg-slate-900/80 border-slate-800'
                        }`}
                        onMouseEnter={() => setHoveredWeekIndex(idx)}
                        onMouseLeave={() => setHoveredWeekIndex(null)}
                      >
                        <div className="flex items-center justify-between text-xs mb-1">
                          <span className="font-bold text-white">{w.weekName}</span>
                          <span className="text-[10px] text-slate-400">{w.dateRange}</span>
                        </div>

                        <div className="text-xl font-black text-white tabular-nums my-1">
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
                            `$${w.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
                          )}
                        </div>

                        <div className="text-[10px] font-semibold text-emerald-400 mb-2">
                          {weekPct}% del total del período
                        </div>

                        {/* Breakdown pills */}
                        <div className="space-y-1 text-[10px] border-t border-slate-800 pt-1.5 text-slate-300">
                          <div className="flex justify-between items-center">
                            <span className="text-sky-400">Facturación:</span>
                            <span className="font-mono text-slate-300">${w.facturacionSales.toFixed(2)}</span>
                          </div>
                          <div className="flex justify-between items-center">
                            <span className="text-emerald-400">ERP:</span>
                            <span className="font-mono text-slate-300">${w.erpSales.toFixed(2)}</span>
                          </div>
                          <div className="flex justify-between items-center">
                            <span className="text-purple-400">Contadores:</span>
                            <span className="font-mono text-slate-300">${w.contadoresSales.toFixed(2)}</span>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* Bottom Insight */}
                <div className="p-3 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                  <div>
                    <strong className="text-emerald-400">Total Semanal Auditado:</strong>{' '}
                    <strong className="text-white">
                      ${totalWeeklySales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD
                    </strong>{' '}
                    ({activeWeekly.length} semanas auditadas). Consistente al 100% con la facturación de sistemas por socio.
                  </div>
                  <span className="text-slate-400">{viewData.periodLabel}</span>
                </div>
              </div>
            );
          })()}

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
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
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

                <div className="flex items-center gap-2">
                  {isManuallyModified && (
                    <button
                      onClick={handleResetCommunities}
                      className="text-xs px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-amber-400 border border-amber-500/30 flex items-center gap-1.5 transition-colors font-medium"
                      title="Restablecer a valores automáticos del Excel"
                    >
                      <RotateCcw className="w-3.5 h-3.5" />
                      <span>Restablecer Automático</span>
                    </button>
                  )}
                  <button
                    onClick={() => {
                      setTempUpconta(upcontaMembers);
                      setTempFranquicia(franquiciaMembers);
                      setTempFormacion(formacionMembers);
                      setIsEditingCommunities(true);
                    }}
                    className="text-xs px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white font-semibold flex items-center gap-1.5 shadow-md shadow-emerald-600/30 transition-colors"
                  >
                    <Edit3 className="w-3.5 h-3.5" />
                    <span>Modificar Miembros</span>
                  </button>
                </div>
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
                    {isEditMode ? (
                      <div className="my-1">
                        <label className="text-[10px] text-sky-300 font-medium block">Cantidad Socios:</label>
                        <input
                          type="number"
                          min="0"
                          value={upcontaMembers}
                          onChange={(e) => {
                            const val = parseInt(e.target.value, 10) || 0;
                            setCustomCommunities((prev) => ({ ...prev, upcontaSocios: val }));
                            setEditValues((prev) => ({ ...prev, r2_upcontaSocios: val, upcontaSocios: val }));
                          }}
                          className="w-full text-2xl font-black text-sky-400 bg-slate-950 border border-sky-500 rounded px-2 py-0.5 tabular-nums focus:outline-none text-center"
                        />
                      </div>
                    ) : (
                      <div className="text-3xl font-black text-sky-400 tabular-nums my-1">
                        {upcontaMembers}
                      </div>
                    )}
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
                    {isEditMode ? (
                      <div className="my-1">
                        <label className="text-[10px] text-purple-300 font-medium block">Cantidad Franquiciados:</label>
                        <input
                          type="number"
                          min="0"
                          value={franquiciaMembers}
                          onChange={(e) => {
                            const val = parseInt(e.target.value, 10) || 0;
                            setCustomCommunities((prev) => ({ ...prev, franquiciaVIP: val }));
                            setEditValues((prev) => ({ ...prev, r2_franquiciaVIP: val, franquiciaVIP: val }));
                          }}
                          className="w-full text-2xl font-black text-white bg-slate-950 border border-emerald-500 rounded px-2 py-0.5 tabular-nums focus:outline-none text-center"
                        />
                      </div>
                    ) : (
                      <div className="text-3xl font-black text-white tabular-nums my-1">
                        {franquiciaMembers}
                      </div>
                    )}
                    <div className="text-[11px] font-semibold text-purple-400 mb-1">Socios Distribuidor Connect</div>
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
                    {isEditMode ? (
                      <div className="my-1">
                        <label className="text-[10px] text-emerald-300 font-medium block">Miembros en Formación:</label>
                        <input
                          type="number"
                          min="0"
                          value={formacionMembers}
                          onChange={(e) => {
                            const val = parseInt(e.target.value, 10) || 0;
                            setCustomCommunities((prev) => ({ ...prev, formacionComercial: val }));
                            setEditValues((prev) => ({ ...prev, r2_formacionComercial: val, formacionComercial: val }));
                          }}
                          className="w-full text-2xl font-black text-emerald-400 bg-slate-950 border border-emerald-500 rounded px-2 py-0.5 tabular-nums focus:outline-none text-center"
                        />
                      </div>
                    ) : (
                      <div className="text-3xl font-black text-emerald-400 tabular-nums my-1">
                        {formacionMembers}
                      </div>
                    )}
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
                  summary={communitiesTableSummary}
                  variant="slide"
                  title="Distribución y Cantidad de Socios por Canal: Distribuidor Upconnect vs Distribuidor Connect"
                />
              </div>

              {/* Bottom callout */}
              <div className="p-2.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-emerald-400">Total Cartera Auditada:</strong> {communitiesTableSummary.totalSociosCount} socios registrados ({communitiesTableSummary.upconnect.count} Distribuidor Upconnect / {communitiesTableSummary.connectors.count} Distribuidor Connect) con {communitiesTableSummary.totalOperationsCount} ventas conciliadas por ${communitiesTableSummary.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD.
                  {isManuallyModified && (
                    <span className="ml-2 text-amber-400 font-medium text-[11px]">(Valores personalizados manualmente)</span>
                  )}
                </div>
                <span className="text-emerald-400 font-bold">100% Conciliado con Base Excel</span>
              </div>
            </div>
          )}

          {/* SLIDE 8: Siguiente Fase: Cruce Comercial */}
          {currentSlideDef.id === 'cross' && (() => {
            const activeCross = viewData.commercialCrossReport2 || dataset.commercialCross;
            return (
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
                    <div>VENTAS EQUIPO COMERCIAL: <strong className="text-white">${activeCross?.commercialTeamSales ? activeCross.commercialTeamSales.toLocaleString('en-US', { minimumFractionDigits: 2 }) : '0.00'}</strong></div>
                    <div>VENTAS ORGÁNICAS: <strong className="text-white">${activeCross?.organicSales ? activeCross.organicSales.toLocaleString('en-US', { minimumFractionDigits: 2 }) : '0.00'}</strong></div>
                    <div>VENTAS SOCIOS: <strong className="text-emerald-400">${(activeCross?.partnersSales || netSales).toLocaleString('en-US', { minimumFractionDigits: 2 })}</strong></div>
                  </div>

                  <div className="grid grid-cols-2 gap-8 pt-6 border-t border-slate-800/80 max-w-lg mx-auto">
                    <div className="text-left">
                      <div className="text-[11px] font-bold text-slate-500 uppercase">{activeCross?.channel1Name || 'CANAL 1 AUDITADO'}</div>
                      <div className="text-sm sm:text-base font-bold text-sky-400 mt-1">
                        Socios UpConnect (${(activeCross?.channel1Amount || netSales).toLocaleString('en-US', { minimumFractionDigits: 2 })})
                      </div>
                    </div>
                    <div className="text-left">
                      <div className="text-[11px] font-bold text-slate-500 uppercase">{activeCross?.channel2Name || 'CANAL 2 PENDIENTE'}</div>
                      <div className="text-sm sm:text-base font-bold text-slate-400 mt-1">
                        {activeCross?.channel2Status || 'Asesores Directos & Marketing'}
                      </div>
                    </div>
                  </div>
                </div>

                <div className="absolute bottom-4 left-8 right-8 flex justify-between items-center text-[11px] text-slate-500 border-t border-slate-800/60 pt-3">
                  <span>Próximo hito: Consolidación omnicanal y liquidación automática</span>
                  <span>Pág. 08</span>
                </div>
              </div>
            );
          })()}

        </div>
      </div>

      {/* MODAL: Modificar Miembros de Comunidades / Embudo */}
      {isEditingCommunities && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="w-full max-w-md bg-slate-950 border border-slate-800 rounded-2xl p-6 shadow-2xl space-y-5">
            <div className="flex items-center justify-between pb-3 border-b border-slate-800">
              <div className="flex items-center gap-2">
                <Users className="w-5 h-5 text-emerald-400" />
                <h3 className="text-base font-bold text-white">Modificar Conteo de Miembros</h3>
              </div>
              <button
                onClick={() => setIsEditingCommunities(false)}
                className="p-1 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <p className="text-xs text-slate-400">
              Ajusta manualmente los miembros del embudo. Los cambios afectarán de inmediato la cantidad y porcentajes de la tabla de distribución por canal.
            </p>

            <div className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-slate-300 mb-1">
                  UpConta Socios (Distribuidor Upconnect)
                </label>
                <div className="text-[11px] text-slate-500 mb-1.5">
                  Socios calculados en Excel: <strong className="text-sky-400">{autoUpSocios}</strong>
                </div>
                <input
                  type="number"
                  min="0"
                  value={tempUpconta}
                  onChange={(e) => setTempUpconta(parseInt(e.target.value, 10) || 0)}
                  className="w-full px-3 py-2 text-sm bg-slate-900 border border-slate-700 rounded-xl text-white focus:outline-none focus:border-emerald-500"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-300 mb-1">
                  Franquicia Contadores VIP (Distribuidor Connect)
                </label>
                <div className="text-[11px] text-slate-500 mb-1.5">
                  Socios calculados en Excel: <strong className="text-purple-400">{autoCoSocios}</strong>
                </div>
                <input
                  type="number"
                  min="0"
                  value={tempFranquicia}
                  onChange={(e) => setTempFranquicia(parseInt(e.target.value, 10) || 0)}
                  className="w-full px-3 py-2 text-sm bg-slate-900 border border-slate-700 rounded-xl text-white focus:outline-none focus:border-emerald-500"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-300 mb-1">
                  Capacitación Franquicia (Formación Comercial)
                </label>
                <div className="text-[11px] text-slate-500 mb-1.5">
                  Estimado automático: <strong className="text-emerald-400">{autoFormacion}</strong>
                </div>
                <input
                  type="number"
                  min="0"
                  value={tempFormacion}
                  onChange={(e) => setTempFormacion(parseInt(e.target.value, 10) || 0)}
                  className="w-full px-3 py-2 text-sm bg-slate-900 border border-slate-700 rounded-xl text-white focus:outline-none focus:border-emerald-500"
                />
              </div>
            </div>

            <div className="flex items-center justify-between pt-3 border-t border-slate-800 gap-2">
              <button
                onClick={handleResetCommunities}
                className="text-xs px-3 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors flex items-center gap-1.5 font-medium"
              >
                <RotateCcw className="w-3.5 h-3.5" />
                <span>Restablecer</span>
              </button>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => setIsEditingCommunities(false)}
                  className="text-xs px-3 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white transition-colors"
                >
                  Cancelar
                </button>
                <button
                  onClick={() => handleSaveCommunitiesModal(tempUpconta, tempFranquicia, tempFormacion)}
                  className="text-xs px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white font-bold transition-colors shadow-lg shadow-emerald-600/30 flex items-center gap-1.5"
                >
                  <Check className="w-3.5 h-3.5" />
                  <span>Guardar Cambios</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

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
