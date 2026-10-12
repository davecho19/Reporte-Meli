import React, { useState, useEffect, useMemo } from 'react';
import { GlobalDataset, SocioRecord, SolutionCategory } from '../types';
import { computeReportView, computeSociosChannelSummary } from '../utils/reportFilters';
import { exportCombinedPresentationsToPPTX } from '../services/pptxExportService';
import { SociosChannelTable } from './SociosChannelTable';
import { SlideEditorModal } from './SlideEditorModal';
import { 
  ChevronLeft, 
  ChevronRight, 
  Maximize2, 
  Minimize2, 
  FileDown, 
  Loader2, 
  Edit3, 
  Check, 
  X, 
  Sparkles, 
  RotateCcw,
  MonitorPlay,
  Layers,
  Users,
  Award,
  Zap,
  Share2,
  BarChart3,
  DollarSign,
  Activity,
  ArrowRight,
  TrendingUp,
  LayoutGrid
} from 'lucide-react';

interface Props {
  dataset: GlobalDataset;
  onBackToDashboard: () => void;
  onUpdateDataset: (updated: Partial<GlobalDataset>) => void;
}

export interface UnifiedSlideDef {
  id: string;
  reportPart: 1 | 2;
  partBadge: string;
  slideNumberInPart: number;
  globalIndex: number;
  title: string;
  subtitle: string;
  internalId: string;
}

export const UnifiedOnlinePresentation: React.FC<Props> = ({
  dataset,
  onBackToDashboard,
  onUpdateDataset,
}) => {
  const [currentSlide, setCurrentSlide] = useState<number>(1);
  const [isFullscreen, setIsFullscreen] = useState<boolean>(false);
  const [isExportingPPTX, setIsExportingPPTX] = useState<boolean>(false);
  const [showSlideDrawer, setShowSlideDrawer] = useState<boolean>(false);

  // Edit Mode state
  const [isEditorModalOpen, setIsEditorModalOpen] = useState<boolean>(false);
  const [isEditMode, setIsEditMode] = useState<boolean>(false);
  const [editValues, setEditValues] = useState<Record<string, string | number>>(dataset.customOverrides || {});

  // Weekly Line Chart State (Report 2 Slide 5 / Global Slide 12)
  const [weeklyMetric, setWeeklyMetric] = useState<'sales' | 'count'>('sales');
  const [hoveredWeekIndex, setHoveredWeekIndex] = useState<number | null>(null);

  useEffect(() => {
    setEditValues(dataset.customOverrides || {});
  }, [dataset.customOverrides]);

  const handleSaveAllChanges = () => {
    onUpdateDataset({ customOverrides: editValues });
    setIsEditMode(false);
    setIsEditorModalOpen(false);
  };

  const handleCancelEdit = () => {
    setEditValues(dataset.customOverrides || {});
    setIsEditMode(false);
  };

  const viewData = computeReportView(dataset);

  // Unified 15-Slide Master Deck Definition
  const unifiedSlides: UnifiedSlideDef[] = [
    // PARTE I: REPORTE 1 (Firmas y Emisiones)
    { id: 'r1_cover', reportPart: 1, partBadge: 'PARTE I · REPORTE 1', slideNumberInPart: 1, globalIndex: 1, title: 'Portada Oficial: Distribuidor Upconnect / Connect', subtitle: 'Emisiones & Rendimiento de Canales', internalId: 'cover' },
    { id: 'r1_balance', reportPart: 1, partBadge: 'PARTE I · REPORTE 1', slideNumberInPart: 2, globalIndex: 2, title: 'Balance General de Facturación 2026', subtitle: 'Consolidado Oficial de Emisiones', internalId: 'balance' },
    { id: 'r1_channels', reportPart: 1, partBadge: 'PARTE I · REPORTE 1', slideNumberInPart: 3, globalIndex: 3, title: 'Desglose de Ventas por Canal y Top Productos', subtitle: 'Distribuidor Upconnect vs Distribuidor Connect', internalId: 'channels' },
    { id: 'r1_weekly', reportPart: 1, partBadge: 'PARTE I · REPORTE 1', slideNumberInPart: 4, globalIndex: 4, title: 'Foco de Rendimiento Semanal (Firmas)', subtitle: 'Auditoría Semana a Semana', internalId: 'weekly' },
    { id: 'r1_durations', reportPart: 1, partBadge: 'PARTE I · REPORTE 1', slideNumberInPart: 5, globalIndex: 5, title: 'Duración de la Cartera (1 a 5 Años)', subtitle: 'Portafolio de Vigencias de Firma', internalId: 'durations' },
    { id: 'r1_communities', reportPart: 1, partBadge: 'PARTE I · REPORTE 1', slideNumberInPart: 6, globalIndex: 6, title: 'Comunidades y Red de Franquiciados', subtitle: 'Ecosistema de Socios y Franquicias', internalId: 'communities' },
    { id: 'r1_cross', reportPart: 1, partBadge: 'PARTE I · REPORTE 1', slideNumberInPart: 7, globalIndex: 7, title: 'Siguiente Fase: Cruce Comercial (Reporte 1)', subtitle: 'Resumen de Ventas y Canales', internalId: 'cross' },

    // PARTE II: REPORTE 2 (Auditoría & Socios UpConta)
    { id: 'r2_cover', reportPart: 2, partBadge: 'PARTE II · REPORTE 2', slideNumberInPart: 1, globalIndex: 8, title: 'Portada Auditoría UpConta: Sistemas & Socios', subtitle: 'Presentación Ejecutiva Oficial UpConta', internalId: 'cover' },
    { id: 'r2_balance', reportPart: 2, partBadge: 'PARTE II · REPORTE 2', slideNumberInPart: 2, globalIndex: 9, title: 'Balance General de Facturación 2026 (Neto / IVA)', subtitle: 'Consolidado Neto y Bruto con IVA (15%)', internalId: 'balance' },
    { id: 'r2_monthly', reportPart: 2, partBadge: 'PARTE II · REPORTE 2', slideNumberInPart: 3, globalIndex: 10, title: 'Evolución de Ventas Mes a Mes (Sin IVA)', subtitle: 'Histórico Comparativo y Mes Líder', internalId: 'monthly' },
    { id: 'r2_pareto', reportPart: 2, partBadge: 'PARTE II · REPORTE 2', slideNumberInPart: 4, globalIndex: 11, title: 'Ranking de Socios: Pareto Top 1-10 & Canales', subtitle: 'Regla de Oro Pareto & Distribución', internalId: 'socios' },
    { id: 'r2_weekly', reportPart: 2, partBadge: 'PARTE II · REPORTE 2', slideNumberInPart: 5, globalIndex: 12, title: 'Desglose y Gráficos Lineales Semana a Semana', subtitle: 'Comparativa: Facturación, ERP y Contadores', internalId: 'weekly' },
    { id: 'r2_solutions', reportPart: 2, partBadge: 'PARTE II · REPORTE 2', slideNumberInPart: 6, globalIndex: 13, title: 'Ventas por Categoría de Solución UpConta', subtitle: 'Portafolio de Sistemas, ERP y Servicios', internalId: 'solutions' },
    { id: 'r2_funnel', reportPart: 2, partBadge: 'PARTE II · REPORTE 2', slideNumberInPart: 7, globalIndex: 14, title: 'Canal de Distribución - Embudo de Socios', subtitle: 'Conteo Cuantitativo y Franquiciados', internalId: 'funnel' },
    { id: 'r2_cross', reportPart: 2, partBadge: 'PARTE II · REPORTE 2', slideNumberInPart: 8, globalIndex: 15, title: 'Estrategia Comercial - Cruce Integral UpConta', subtitle: 'Fuerza Comercial vs Red de Socios', internalId: 'cross' },
  ];

  const totalSlides = unifiedSlides.length;
  const currentSlideDef = unifiedSlides[currentSlide - 1] || unifiedSlides[0];

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') {
        setCurrentSlide((prev) => Math.min(prev + 1, totalSlides));
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        setCurrentSlide((prev) => Math.max(prev - 1, 1));
      } else if (e.key === 'Escape') {
        if (isFullscreen) setIsFullscreen(false);
        if (showSlideDrawer) setShowSlideDrawer(false);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isFullscreen, totalSlides, showSlideDrawer]);

  const toggleFullscreen = () => {
    setIsFullscreen(!isFullscreen);
  };

  // Communities / Embudo State & Management
  // Report 1 Communities
  const autoUpSociosR1 = viewData.firmaSociosSummary.upconnect.count || 118;
  const autoCoSociosR1 = viewData.firmaSociosSummary.connectors.count || 129;
  const autoFormacionR1 = Math.max(1, Math.round((autoUpSociosR1 + autoCoSociosR1) * 0.25)) || 51;

  // Report 2 Communities
  const autoUpSociosR2 = viewData.sociosSummary.upconnect.count || 118;
  const autoCoSociosR2 = viewData.sociosSummary.connectors.count || 129;
  const autoFormacionR2 = Math.max(1, Math.round((autoUpSociosR2 + autoCoSociosR2) * 0.25)) || 51;

  const [customCommunities, setCustomCommunities] = useState<{
    upcontaSocios?: number;
    franquiciaVIP?: number;
    formacionComercial?: number;
  }>(() => {
    const ov = dataset.customOverrides || {};
    return {
      upcontaSocios: ov['upcontaSocios'] !== undefined ? Number(ov['upcontaSocios']) : undefined,
      franquiciaVIP: ov['franquiciaVIP'] !== undefined ? Number(ov['franquiciaVIP']) : undefined,
      formacionComercial: ov['formacionComercial'] !== undefined ? Number(ov['formacionComercial']) : undefined,
    };
  });

  const [isEditingCommunitiesModal, setIsEditingCommunitiesModal] = useState<boolean>(false);
  const [editingTargetPart, setEditingTargetPart] = useState<1 | 2>(1);
  const [tempUpconta, setTempUpconta] = useState<number>(0);
  const [tempFranquicia, setTempFranquicia] = useState<number>(0);
  const [tempFormacion, setTempFormacion] = useState<number>(0);

  // Active Member values reflecting overrides & custom state
  const upcontaMembersR1 = editValues['r1_upcontaSocios'] !== undefined 
    ? Number(editValues['r1_upcontaSocios']) 
    : (customCommunities.upcontaSocios ?? autoUpSociosR1);
  const franquiciaMembersR1 = editValues['r1_franquiciaVIP'] !== undefined 
    ? Number(editValues['r1_franquiciaVIP']) 
    : (customCommunities.franquiciaVIP ?? autoCoSociosR1);
  const formacionMembersR1 = editValues['r1_formacionComercial'] !== undefined 
    ? Number(editValues['r1_formacionComercial']) 
    : (customCommunities.formacionComercial ?? autoFormacionR1);

  const upcontaMembersR2 = editValues['r2_upcontaSocios'] !== undefined 
    ? Number(editValues['r2_upcontaSocios']) 
    : (customCommunities.upcontaSocios ?? autoUpSociosR2);
  const franquiciaMembersR2 = editValues['r2_franquiciaVIP'] !== undefined 
    ? Number(editValues['r2_franquiciaVIP']) 
    : (customCommunities.franquiciaVIP ?? autoCoSociosR2);
  const formacionMembersR2 = editValues['r2_formacionComercial'] !== undefined 
    ? Number(editValues['r2_formacionComercial']) 
    : (customCommunities.formacionComercial ?? autoFormacionR2);

  const handleSaveCommunities = (part: 1 | 2, up: number, vip: number, form: number) => {
    const prefix = part === 1 ? 'r1_' : 'r2_';
    const updated = {
      [`${prefix}upcontaSocios`]: up,
      [`${prefix}franquiciaVIP`]: vip,
      [`${prefix}formacionComercial`]: form,
      upcontaSocios: up,
      franquiciaVIP: vip,
      formacionComercial: form,
    };
    setCustomCommunities((prev) => ({ ...prev, upcontaSocios: up, franquiciaVIP: vip, formacionComercial: form }));
    setEditValues((prev) => ({ ...prev, ...updated }));
    onUpdateDataset({
      customOverrides: {
        ...(dataset.customOverrides || {}),
        ...updated,
      },
    });
    setIsEditingCommunitiesModal(false);
  };

  const handleResetCommunities = (part: 1 | 2) => {
    const prefix = part === 1 ? 'r1_' : 'r2_';
    const newOverrides = { ...(dataset.customOverrides || {}) };
    delete newOverrides[`${prefix}upcontaSocios`];
    delete newOverrides[`${prefix}franquiciaVIP`];
    delete newOverrides[`${prefix}formacionComercial`];
    delete newOverrides.upcontaSocios;
    delete newOverrides.franquiciaVIP;
    delete newOverrides.formacionComercial;
    setCustomCommunities({});
    setEditValues(newOverrides);
    onUpdateDataset({ customOverrides: newOverrides });
  };

  // Channel Table Summaries bound directly to members
  const communitiesTableSummaryR1 = useMemo(() => {
    const upCount = upcontaMembersR1;
    const coCount = franquiciaMembersR1;
    const totalMembers = upCount + coCount;
    const upPct = totalMembers > 0 ? parseFloat(((upCount / totalMembers) * 100).toFixed(1)) : 0;
    const coPct = totalMembers > 0 ? parseFloat(((coCount / totalMembers) * 100).toFixed(1)) : 0;

    const baseSummary = viewData.firmaSociosSummary;
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
        socios: viewData.firmaSocios.filter((s) => s.channel === 'UpConnect'),
      },
      connectors: {
        channel: 'Connectors' as const,
        channelLabel: 'Distribuidor Connect',
        count: coCount,
        percentage: coPct,
        operationsCount: coOps,
        totalSales: coNetSales,
        averageTicket: coAvgTicket,
        socios: viewData.firmaSocios.filter((s) => s.channel === 'Connectors'),
      },
      totalSociosCount: totalMembers,
      totalOperationsCount: totalOps,
      totalSales: allSales,
      averageTicket: avgTicket,
    };
  }, [upcontaMembersR1, franquiciaMembersR1, viewData.firmaSociosSummary, viewData.firmaSocios]);

  const communitiesTableSummaryR2 = useMemo(() => {
    const upCount = upcontaMembersR2;
    const coCount = franquiciaMembersR2;
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
        socios: viewData.socios.filter((s) => s.channel === 'UpConnect'),
      },
      connectors: {
        channel: 'Connectors' as const,
        channelLabel: 'Distribuidor Connect',
        count: coCount,
        percentage: coPct,
        operationsCount: coOps,
        totalSales: coNetSales,
        averageTicket: coAvgTicket,
        socios: viewData.socios.filter((s) => s.channel === 'Connectors'),
      },
      totalSociosCount: totalMembers,
      totalOperationsCount: totalOps,
      totalSales: allSales,
      averageTicket: avgTicket,
    };
  }, [upcontaMembersR2, franquiciaMembersR2, viewData.sociosSummary, viewData.socios]);

  // Handle Export Combined PPTX
  const handleExportCombinedPPTX = async () => {
    try {
      setIsExportingPPTX(true);
      await exportCombinedPresentationsToPPTX(dataset, customCommunities);
    } catch (e) {
      console.error('Error generating combined PPTX', e);
    } finally {
      setIsExportingPPTX(false);
    }
  };

  // Helper: Top 6 Products for Report 1
  const computeTop6Products = (channel: 'UpConnect' | 'Connectors') => {
    const map = new Map<string, { name: string; count: number; sales: number }>();
    viewData.transactions
      .filter((t) => t.channel === channel)
      .forEach((t) => {
        let dur = (t.duration || '1 año').trim();
        if (!dur.toLowerCase().includes('año') && !dur.toLowerCase().includes('día')) {
          dur = `${dur} año${dur === '1' ? '' : 's'}`;
        }
        let prodName = `Firma de ${dur}`;
        if (t.solutionCategory && !t.solutionCategory.toLowerCase().includes('firma') && t.solutionCategory !== 'Firma Electrónica') {
          prodName = `${t.solutionCategory} (Firma de ${dur})`;
        }
        const key = prodName;
        if (!map.has(key)) map.set(key, { name: key, count: 0, sales: 0 });
        const item = map.get(key)!;
        item.count += 1;
        item.sales += t.value || 0;
      });

    if (map.size === 0 && dataset.portfolioDurations && dataset.portfolioDurations.length > 0) {
      dataset.portfolioDurations
        .filter((p) => p.channel === channel)
        .forEach((p) => {
          map.set(p.duration, {
            name: `Firma de ${p.duration}`,
            count: p.count,
            sales: p.count * 15,
          });
        });
    }

    const list = Array.from(map.values()).sort((a, b) => b.count - a.count || b.sales - a.sales);
    const channelTotalCount = list.reduce((a, b) => a + b.count, 0) || 1;

    return list.slice(0, 6).map((item, idx) => ({
      rank: idx + 1,
      name: item.name,
      count: item.count,
      sales: parseFloat(item.sales.toFixed(2)),
      percentage: parseFloat(((item.count / channelTotalCount) * 100).toFixed(1)),
      label: idx === 0 ? 'Líder' : undefined,
    }));
  };

  const topUpProducts = computeTop6Products('UpConnect');
  const topCoProducts = computeTop6Products('Connectors');

  // Report 2 Monthly Metrics (Isolated to selected month if month filter active)
  const activeMonthlyMetricsR2 = useMemo(() => {
    const raw = (viewData.sistemasMonthlyMetrics && viewData.sistemasMonthlyMetrics.length > 0)
      ? viewData.sistemasMonthlyMetrics
      : (dataset.sistemasMonthlyMetrics || []);
    if (dataset.reportFilter?.type === 'month' && dataset.reportFilter.month) {
      const targetMonth = dataset.reportFilter.month.toUpperCase();
      const filtered = raw.filter((m) => m.month.toUpperCase() === targetMonth);
      if (filtered.length > 0) return filtered;
      return [{
        month: targetMonth,
        monthIndex: 9,
        upconnectSales: viewData.sociosSummary.upconnect.totalSales,
        connectorsSales: viewData.sociosSummary.connectors.totalSales,
        upconnectCount: viewData.sociosSummary.upconnect.operationsCount,
        connectorsCount: viewData.sociosSummary.connectors.operationsCount,
      }];
    }
    return raw;
  }, [viewData.sistemasMonthlyMetrics, dataset.sistemasMonthlyMetrics, dataset.reportFilter, viewData.sociosSummary]);

  const auditMonthlyR2 = activeMonthlyMetricsR2.length > 0
    ? activeMonthlyMetricsR2.map((m, _, arr) => {
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
  const maxAuditValR2 = auditMonthlyR2.length > 0 ? Math.max(...auditMonthlyR2.map((m) => m.total)) : 1;

  // Categories for Report 2
  const activeCategoriesR2: SolutionCategory[] = (viewData.solutionCategoriesReport2 && viewData.solutionCategoriesReport2.length > 0)
    ? viewData.solutionCategoriesReport2
    : (dataset.solutionCategories || []);

  return (
    <div className={`flex flex-col ${isFullscreen ? 'fixed inset-0 z-50 bg-[#070b14]' : 'w-full min-h-[calc(100vh-60px)] bg-[#070b14]'}`}>
      {/* Top Deck Action Bar */}
      <div className="no-print flex flex-wrap items-center justify-between px-4 sm:px-6 py-2.5 bg-[#0d1527] border-b border-slate-800 text-slate-300 gap-3">
        {/* Left Side: Back & Deck Badge */}
        <div className="flex items-center gap-2 sm:gap-3">
          <button
            onClick={onBackToDashboard}
            className="text-xs px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white transition-colors flex items-center gap-1.5 font-medium"
          >
            ← Dashboard
          </button>

          <div className="flex items-center gap-2">
            <span className="text-xs uppercase tracking-wider font-extrabold text-sky-400 flex items-center gap-1.5">
              <MonitorPlay className="w-4 h-4 text-sky-400" />
              <span>Reporte en Línea</span>
            </span>

            <span className="text-slate-600 hidden sm:inline">|</span>

            {/* Part indicator */}
            {currentSlideDef.reportPart === 1 ? (
              <span className="text-[11px] px-2.5 py-0.5 rounded-full font-bold bg-blue-500/20 text-blue-300 border border-blue-500/30">
                Parte I: Reporte 1 · Firmas (Pág. {currentSlideDef.slideNumberInPart}/7)
              </span>
            ) : (
              <span className="text-[11px] px-2.5 py-0.5 rounded-full font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                Parte II: Reporte 2 · Socios (Pág. {currentSlideDef.slideNumberInPart}/8)
              </span>
            )}
          </div>
        </div>

        {/* Center: Quick Part Switchers */}
        <div className="hidden lg:flex items-center gap-1 bg-slate-900/90 p-1 rounded-xl border border-slate-800 text-xs">
          <button
            onClick={() => setCurrentSlide(1)}
            className={`px-3 py-1 rounded-lg font-semibold transition-all ${
              currentSlideDef.reportPart === 1
                ? 'bg-blue-600 text-white shadow-sm'
                : 'text-slate-400 hover:text-white'
            }`}
          >
            Parte I: Reporte 1 (1–7)
          </button>
          <button
            onClick={() => setCurrentSlide(8)}
            className={`px-3 py-1 rounded-lg font-semibold transition-all ${
              currentSlideDef.reportPart === 2
                ? 'bg-emerald-600 text-white shadow-sm'
                : 'text-slate-400 hover:text-white'
            }`}
          >
            Parte II: Reporte 2 (8–15)
          </button>
        </div>

        {/* Right Side: Tools & Navigation */}
        <div className="flex items-center gap-2 text-xs">
          {/* Slide Navigator Dropdown */}
          <button
            onClick={() => setShowSlideDrawer(!showSlideDrawer)}
            className="px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 transition-colors flex items-center gap-1.5 font-medium"
            title="Ver índice de diapositivas unificadas"
          >
            <LayoutGrid className="w-3.5 h-3.5 text-sky-400" />
            <span className="hidden sm:inline">Índice ({currentSlide}/15)</span>
            <span className="sm:hidden">{currentSlide}/15</span>
          </button>

          {/* Edit Slide Button */}
          <button
            onClick={() => setIsEditorModalOpen(true)}
            className="px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-emerald-400 transition-colors flex items-center gap-1.5 font-semibold"
            title="Editar valores de esta diapositiva y guardarlos"
          >
            <Edit3 className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Editar</span>
          </button>

          {/* Download PPTX Button */}
          <button
            onClick={handleExportCombinedPPTX}
            disabled={isExportingPPTX}
            className="px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white font-bold transition-all shadow-md shadow-emerald-950/40 flex items-center gap-1.5 disabled:opacity-50"
            title="Descargar presentación unificada completa (.pptx)"
          >
            {isExportingPPTX ? (
              <Loader2 className="w-3.5 h-3.5 animate-spin" />
            ) : (
              <FileDown className="w-3.5 h-3.5" />
            )}
            <span className="hidden sm:inline">Descargar PPTX</span>
          </button>

          {/* Fullscreen Toggle */}
          <button
            onClick={toggleFullscreen}
            className="p-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white transition-colors"
            title={isFullscreen ? 'Salir de pantalla completa' : 'Pantalla completa'}
          >
            {isFullscreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
          </button>

          {/* Slide Arrow Navigation */}
          <div className="flex items-center gap-1 ml-1 bg-slate-900 border border-slate-800 rounded-lg p-0.5">
            <button
              onClick={() => setCurrentSlide((prev) => Math.max(prev - 1, 1))}
              disabled={currentSlide === 1}
              className="p-1 rounded text-slate-400 hover:text-white disabled:opacity-30 disabled:hover:text-slate-400 transition-colors"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
            <span className="px-1 text-[11px] font-bold text-slate-200 tabular-nums">
              {currentSlide}/{totalSlides}
            </span>
            <button
              onClick={() => setCurrentSlide((prev) => Math.min(prev + 1, totalSlides))}
              disabled={currentSlide === totalSlides}
              className="p-1 rounded text-slate-400 hover:text-white disabled:opacity-30 disabled:hover:text-slate-400 transition-colors"
            >
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>
        </div>
      </div>

      {/* Slide Drawer (Quick Index) */}
      {showSlideDrawer && (
        <div className="fixed inset-0 z-50 flex justify-end bg-black/60 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="w-full max-w-sm bg-slate-950 border-l border-slate-800 h-full flex flex-col p-5 overflow-y-auto">
            <div className="flex items-center justify-between pb-3 border-b border-slate-800 mb-4">
              <div className="flex items-center gap-2">
                <LayoutGrid className="w-4 h-4 text-sky-400" />
                <h3 className="text-sm font-bold text-white">Índice del Reporte en Línea</h3>
              </div>
              <button
                onClick={() => setShowSlideDrawer(false)}
                className="p-1 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-4">
              {/* Part 1 Header */}
              <div>
                <span className="text-[10px] font-bold uppercase tracking-wider text-sky-400 px-2 py-0.5 rounded bg-sky-500/10">
                  Parte I: Reporte 1 · Firmas y Emisiones
                </span>
                <div className="space-y-1 mt-2">
                  {unifiedSlides.slice(0, 7).map((s) => (
                    <button
                      key={s.id}
                      onClick={() => {
                        setCurrentSlide(s.globalIndex);
                        setShowSlideDrawer(false);
                      }}
                      className={`w-full text-left p-2 rounded-lg text-xs transition-colors flex items-center justify-between ${
                        currentSlide === s.globalIndex
                          ? 'bg-blue-600/30 text-white border border-blue-500/50 font-bold'
                          : 'text-slate-300 hover:bg-slate-900 hover:text-white'
                      }`}
                    >
                      <div className="flex items-center gap-2 truncate">
                        <span className="w-5 text-slate-500 font-mono text-[11px]">{s.globalIndex}.</span>
                        <span className="truncate">{s.title}</span>
                      </div>
                      <span className="text-[10px] text-slate-500 shrink-0">Pág. {s.slideNumberInPart}</span>
                    </button>
                  ))}
                </div>
              </div>

              {/* Part 2 Header */}
              <div>
                <span className="text-[10px] font-bold uppercase tracking-wider text-emerald-400 px-2 py-0.5 rounded bg-emerald-500/10">
                  Parte II: Reporte 2 · Socios y Sistemas UpConta
                </span>
                <div className="space-y-1 mt-2">
                  {unifiedSlides.slice(7).map((s) => (
                    <button
                      key={s.id}
                      onClick={() => {
                        setCurrentSlide(s.globalIndex);
                        setShowSlideDrawer(false);
                      }}
                      className={`w-full text-left p-2 rounded-lg text-xs transition-colors flex items-center justify-between ${
                        currentSlide === s.globalIndex
                          ? 'bg-emerald-600/30 text-white border border-emerald-500/50 font-bold'
                          : 'text-slate-300 hover:bg-slate-900 hover:text-white'
                      }`}
                    >
                      <div className="flex items-center gap-2 truncate">
                        <span className="w-5 text-slate-500 font-mono text-[11px]">{s.globalIndex}.</span>
                        <span className="truncate">{s.title}</span>
                      </div>
                      <span className="text-[10px] text-slate-500 shrink-0">Pág. {s.slideNumberInPart}</span>
                    </button>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Main Slide Stage (16:9 Presentation Frame) */}
      <div className="flex-1 flex items-center justify-center p-2 sm:p-4 md:p-6 overflow-hidden">
        <div className="w-full max-w-[1340px] aspect-[16/9] bg-[#0c1322] border border-slate-800 rounded-2xl shadow-2xl overflow-hidden relative flex flex-col justify-between">
          
          {/* ========================================================================= */}
          {/* SLIDE 1: Portada Oficial Reporte 1 */}
          {/* ========================================================================= */}
          {currentSlide === 1 && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-14 relative bg-gradient-to-br from-[#0c1322] via-[#09101d] to-[#040812]">
              <div className="absolute w-[450px] h-[450px] rounded-full bg-blue-600/10 blur-[120px] pointer-events-none -top-16 -left-16" />
              <div className="absolute w-[400px] h-[400px] rounded-full bg-amber-500/10 blur-[100px] pointer-events-none -bottom-16 -right-16" />

              <div className="z-10 max-w-4xl space-y-5 my-auto">
                <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-sky-500/10 border border-sky-500/20 text-sky-400 text-xs font-semibold tracking-wide">
                  <span>DOCUMENTO OFICIAL GERENCIAL · PARTE I</span>
                  <span>·</span>
                  <span>{viewData.periodLabel.toUpperCase()}</span>
                </div>

                <h1 className="text-3xl sm:text-5xl font-extrabold tracking-tight text-white leading-tight">
                  {editValues['titleReport1'] || editValues['r1_cover_title'] || viewData.titleReport1}
                </h1>

                <p className="text-slate-400 text-sm sm:text-base max-w-2xl pt-1 leading-relaxed">
                  Auditoría comparativa de emisión de certificados, rendimiento de canales directos y red externa de distribuidores con corte al {viewData.cutoffDateReport1}.
                </p>

                <div className="pt-6 flex flex-wrap items-center gap-6 text-xs text-slate-400">
                  <div className="flex items-center gap-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-sky-400" />
                    <span>Canal Directo: <strong>Distribuidor Upconnect</strong></span>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-amber-400" />
                    <span>Red Externa: <strong>Distribuidor Connect</strong></span>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-emerald-400" />
                    <span>Corte: <strong>{viewData.cutoffDateReport1}</strong></span>
                  </div>
                </div>
              </div>

              <div className="z-10 flex justify-between items-center text-xs text-slate-500 border-t border-slate-800/80 pt-4">
                <span>Upconnect / Connect Ecosystem · Reporte 1</span>
                <span>Diapositiva 01 de 15</span>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* SLIDE 2: Balance General de Facturación 2026 (Reporte 1) */}
          {/* ========================================================================= */}
          {currentSlide === 2 && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  BALANCE GENERAL DE EMISIONES · REPORTE 1
                </span>
                <h2 className="text-2xl sm:text-3xl font-black text-white tracking-tight mt-0.5">
                  Consolidado Total de Facturación de Firmas
                </h2>
                <p className="text-xs text-slate-400 mt-1">
                  Métricas oficiales conciliadas de ingresos por emisión directa y red distribuidores
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 my-auto">
                <div className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl relative overflow-hidden">
                  <div className="text-xs font-semibold uppercase tracking-wider text-slate-400">TOTAL FACTURADO (FIRMAS)</div>
                  <div className="text-3xl sm:text-4xl font-black text-white tabular-nums mt-1">
                    ${viewData.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                  <div className="text-xs text-emerald-400 mt-2 font-semibold">
                    {viewData.totalCount} emisiones conciliadas
                  </div>
                </div>

                <div className="p-6 rounded-2xl bg-slate-900/90 border-t-2 border-sky-500 border-x border-b border-slate-800 shadow-xl relative overflow-hidden">
                  <div className="text-xs font-semibold uppercase tracking-wider text-sky-400">DISTRIBUIDOR UPCONNECT</div>
                  <div className="text-3xl sm:text-4xl font-black text-sky-400 tabular-nums mt-1">
                    ${viewData.totalUpSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                  <div className="text-xs text-slate-400 mt-2 font-semibold">
                    {viewData.totalUpCount} firmas · {viewData.upSalesPct}% de la facturación
                  </div>
                </div>

                <div className="p-6 rounded-2xl bg-slate-900/90 border-t-2 border-amber-500 border-x border-b border-slate-800 shadow-xl relative overflow-hidden">
                  <div className="text-xs font-semibold uppercase tracking-wider text-amber-400">DISTRIBUIDOR CONNECT</div>
                  <div className="text-3xl sm:text-4xl font-black text-amber-400 tabular-nums mt-1">
                    ${viewData.totalCoSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                  <div className="text-xs text-slate-400 mt-2 font-semibold">
                    {viewData.totalCoCount} firmas · {viewData.coSalesPct}% de la facturación
                  </div>
                </div>
              </div>

              <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-sky-400">Estatus Operativo:</strong> Facturación promedio por firma de{' '}
                  <strong className="text-white">
                    ${(viewData.totalSales / (viewData.totalCount || 1)).toFixed(2)} USD
                  </strong>{' '}
                  en el período activo ({viewData.periodLabel}).
                </div>
                <span className="text-slate-500">Pág. 02/15</span>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* SLIDE 3: Desglose por Canal & Top 6 Productos */}
          {/* ========================================================================= */}
          {currentSlide === 3 && (
            <div className="h-full flex flex-col justify-between p-6 sm:p-8 relative bg-[#0c1322]">
              <div className="flex items-center justify-between border-b border-slate-800/80 pb-2.5">
                <div>
                  <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                    CANALES DE DISTRIBUCIÓN · REPORTE 1
                  </span>
                  <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                    Desglose de Ventas por Canal y Top Productos
                  </h2>
                </div>
                <div className="text-xs text-slate-400 font-semibold bg-slate-900 px-3 py-1 rounded-lg border border-slate-800">
                  {viewData.periodLabel}
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-5 my-auto">
                {/* UpConnect */}
                <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between">
                  <div className="flex items-center justify-between pb-2 border-b border-slate-800 mb-2">
                    <span className="text-xs font-bold text-sky-400">DISTRIBUIDOR UPCONNECT (DIRECTO)</span>
                    <span className="text-sm font-black text-white">${viewData.totalUpSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>
                  </div>
                  <div className="space-y-1.5">
                    {topUpProducts.map((p) => (
                      <div key={p.rank} className="flex items-center justify-between text-xs py-0.5 border-b border-slate-800/40">
                        <span className="text-slate-300 truncate max-w-[200px]">{p.rank}. {p.name}</span>
                        <div className="flex items-center gap-3">
                          <span className="text-slate-500 text-[11px]">{p.count} emi</span>
                          <span className="font-bold text-white font-mono">${p.sales.toFixed(2)}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Connectors */}
                <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between">
                  <div className="flex items-center justify-between pb-2 border-b border-slate-800 mb-2">
                    <span className="text-xs font-bold text-amber-400">DISTRIBUIDOR CONNECT (ALIADOS)</span>
                    <span className="text-sm font-black text-white">${viewData.totalCoSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>
                  </div>
                  <div className="space-y-1.5">
                    {topCoProducts.map((p) => (
                      <div key={p.rank} className="flex items-center justify-between text-xs py-0.5 border-b border-slate-800/40">
                        <span className="text-slate-300 truncate max-w-[200px]">{p.rank}. {p.name}</span>
                        <div className="flex items-center gap-3">
                          <span className="text-slate-500 text-[11px]">{p.count} emi</span>
                          <span className="font-bold text-white font-mono">${p.sales.toFixed(2)}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              <div className="p-3 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-400 flex justify-between items-center">
                <span>Distribución por participación: UpConnect {viewData.upSalesPct}% vs Connect {viewData.coSalesPct}%</span>
                <span>Pág. 03/15</span>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* SLIDE 4: Foco Semanal Reporte 1 */}
          {/* ========================================================================= */}
          {currentSlide === 4 && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  FOCO SEMANAL DE EMISIONES · REPORTE 1
                </span>
                <h2 className="text-2xl sm:text-3xl font-black text-white tracking-tight mt-0.5">
                  Foco de Rendimiento: Desglose Semana a Semana
                </h2>
                <p className="text-xs text-slate-400 mt-1">
                  Emisiones por semana del período activo ({viewData.periodLabel})
                </p>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4 my-auto">
                {viewData.weeklyBreakdownType1.map((w, idx) => {
                  const totalAmt = w.totalAmount;
                  const pct = viewData.totalSales > 0 ? ((totalAmt / viewData.totalSales) * 100).toFixed(1) : '25.0';
                  return (
                    <div key={idx} className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between">
                      <div>
                        <div className="flex items-center justify-between text-xs mb-1">
                          <span className="font-bold text-sky-400">{w.weekName}</span>
                          <span className="text-[11px] text-slate-400">{w.dateRange}</span>
                        </div>
                        <div className="text-2xl font-black text-white tabular-nums my-1">
                          ${totalAmt.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </div>
                        <div className="text-xs font-semibold text-emerald-400">{pct}% del total</div>
                      </div>
                      <div className="pt-3 border-t border-slate-800 text-[11px] text-slate-400">
                        {w.highlightText || `${w.totalCount} firmas (${w.upconnectCount} UP / ${w.connectorsCount} CO)`}
                      </div>
                    </div>
                  );
                })}
              </div>

              <div className="p-3.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-sky-400">Total Semanal Auditado:</strong>{' '}
                  <strong className="text-white">
                    ${viewData.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD
                  </strong>
                </div>
                <span>Pág. 04/15</span>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* SLIDE 5: Duración de la Cartera (Portafolio de Vigencias) */}
          {/* ========================================================================= */}
          {currentSlide === 5 && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  DURACIÓN DE CARTERA · REPORTE 1
                </span>
                <h2 className="text-2xl sm:text-3xl font-black text-white tracking-tight mt-0.5">
                  Portafolio de Duraciones de Firma
                </h2>
                <p className="text-xs text-slate-400 mt-1">
                  Distribución de certificados emitidos de 1 año a 5 años por canal de venta
                </p>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-3 my-auto">
                {['1 año', '2 años', '3 años', '4 años', '5 años', '15 días'].map((dur, idx) => {
                  const port = dataset.portfolioDurations || [];
                  const upCnt = port.find((p) => p.channel === 'UpConnect' && p.duration.includes(dur.slice(0, 2)))?.count || 0;
                  const coCnt = port.find((p) => p.channel === 'Connectors' && p.duration.includes(dur.slice(0, 2)))?.count || 0;
                  const tot = upCnt + coCnt;
                  return (
                    <div key={idx} className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between text-center">
                      <div className="text-xs font-bold text-slate-300">{dur}</div>
                      <div className="text-3xl font-black text-white tabular-nums my-2">{tot}</div>
                      <div className="text-[10px] text-slate-500">
                        UP: {upCnt} · CO: {coCnt}
                      </div>
                    </div>
                  );
                })}
              </div>

              <div className="p-3.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-sky-400">Vigencia Predominante:</strong> Mayor volumen en certificados de 1 año y 2 años.
                </div>
                <span>Pág. 05/15</span>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* SLIDE 6: Comunidades y Franquiciados (Reporte 1) */}
          {/* ========================================================================= */}
          {currentSlide === 6 && (
            <div className="h-full flex flex-col justify-between p-6 sm:p-8 relative bg-[#0c1322]">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                  <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                    RED HUMANA Y DISTRIBUCIÓN · REPORTE 1
                  </span>
                  <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                    Comunidades y Red de Franquiciados
                  </h2>
                  <p className="text-xs text-slate-400 mt-0.5">
                    Monitoreo activo de grupos de gestión, formación comercial y franquiciados
                  </p>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={() => handleResetCommunities(1)}
                    className="text-xs px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-amber-400 border border-amber-500/30 flex items-center gap-1.5 transition-colors font-medium"
                    title="Restablecer a valores automáticos del Excel"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    <span>Restablecer</span>
                  </button>
                  <button
                    onClick={() => {
                      setEditingTargetPart(1);
                      setTempUpconta(upcontaMembersR1);
                      setTempFranquicia(franquiciaMembersR1);
                      setTempFormacion(formacionMembersR1);
                      setIsEditingCommunitiesModal(true);
                    }}
                    className="text-xs px-3 py-1.5 rounded-lg bg-sky-600 hover:bg-sky-500 text-white font-semibold flex items-center gap-1.5 shadow-md shadow-sky-600/30 transition-colors"
                  >
                    <Edit3 className="w-3.5 h-3.5" />
                    <span>Modificar Miembros</span>
                  </button>
                </div>
              </div>

              {/* 3 Funnel Cards */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 my-2">
                <div className="p-4 rounded-xl bg-gradient-to-b from-[#0f1d38] to-[#0a1224] border border-sky-500/40 shadow-xl text-center">
                  <h3 className="text-xs font-bold text-sky-300">UpConta Socios (Distribuidor Upconnect)</h3>
                  <div className="text-3xl font-black text-sky-400 tabular-nums my-1">{upcontaMembersR1}</div>
                  <div className="text-[11px] text-slate-400">Socios Registrados</div>
                </div>

                <div className="p-4 rounded-xl bg-gradient-to-b from-[#241338] to-[#120a1f] border border-purple-500/40 shadow-xl text-center">
                  <h3 className="text-xs font-bold text-purple-300">Franquicia Contadores VIP (Distribuidor Connect)</h3>
                  <div className="text-3xl font-black text-purple-400 tabular-nums my-1">{franquiciaMembersR1}</div>
                  <div className="text-[11px] text-slate-400">Franquiciados Registrados</div>
                </div>

                <div className="p-4 rounded-xl bg-gradient-to-b from-[#0d2e26] to-[#071713] border border-emerald-500/40 shadow-xl text-center">
                  <h3 className="text-xs font-bold text-emerald-300">Formación Comercial Franquiciados</h3>
                  <div className="text-3xl font-black text-emerald-400 tabular-nums my-1">{formacionMembersR1}</div>
                  <div className="text-[11px] text-slate-400">Miembros en Formación</div>
                </div>
              </div>

              <div className="my-1">
                <SociosChannelTable
                  summary={communitiesTableSummaryR1}
                  variant="slide"
                  title="Distribución y Cantidad de Socios por Canal: Distribuidor Upconnect vs Distribuidor Connect"
                />
              </div>

              <div className="p-2.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex justify-between items-center">
                <span>Ecosistema: {communitiesTableSummaryR1.totalSociosCount} operadores ({communitiesTableSummaryR1.upconnect.count} UP / {communitiesTableSummaryR1.connectors.count} CO)</span>
                <span>Pág. 06/15</span>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* SLIDE 7: Siguiente Fase: Cruce Comercial (Reporte 1) */}
          {/* ========================================================================= */}
          {currentSlide === 7 && (
            <div className="h-full flex flex-col justify-center items-center text-center p-8 sm:p-14 relative bg-gradient-to-br from-[#0c1322] via-[#080d19] to-[#040812]">
              <div className="z-10 max-w-2xl space-y-5">
                <div className="w-12 h-12 rounded-2xl bg-blue-600 text-white flex items-center justify-center mx-auto shadow-lg shadow-blue-500/30">
                  <TrendingUp className="w-6 h-6" />
                </div>

                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  ESTRATEGIA INTEGRAL · FIN DE PARTE I
                </span>

                <h2 className="text-3xl sm:text-4xl font-black text-white tracking-tight">
                  CRUCE COMERCIAL: FUERZA DIRECTA Y RED CONNECT
                </h2>

                <div className="grid grid-cols-2 gap-6 pt-4 max-w-md mx-auto text-left">
                  <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800">
                    <span className="text-xs text-sky-400 font-bold block mb-1">CANAL 1 DIRECTO</span>
                    <span className="text-xl font-black text-white">${viewData.totalUpSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>
                    <span className="text-[11px] text-slate-400 block mt-1">{viewData.totalUpCount} transacciones directas</span>
                  </div>
                  <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800">
                    <span className="text-xs text-amber-400 font-bold block mb-1">CANAL 2 DISTRIBUIDOR</span>
                    <span className="text-xl font-black text-white">${viewData.totalCoSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>
                    <span className="text-[11px] text-slate-400 block mt-1">{viewData.totalCoCount} transacciones red</span>
                  </div>
                </div>

                <button
                  onClick={() => setCurrentSlide(8)}
                  className="mt-6 px-5 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs inline-flex items-center gap-2 transition-all shadow-lg shadow-emerald-950/40"
                >
                  <span>Continuar a Parte II: Reporte 2 (Sistemas & Socios)</span>
                  <ArrowRight className="w-4 h-4" />
                </button>
              </div>

              <div className="absolute bottom-4 left-8 right-8 flex justify-between items-center text-[11px] text-slate-500 border-t border-slate-800/60 pt-3">
                <span>Transición oficial de reporte</span>
                <span>Pág. 07/15</span>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* SLIDE 8: Portada Auditoría UpConta (Reporte 2) */}
          {/* ========================================================================= */}
          {currentSlide === 8 && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-14 relative bg-gradient-to-br from-[#0c1322] via-[#071318] to-[#040812]">
              <div className="absolute w-[450px] h-[450px] rounded-full bg-emerald-600/10 blur-[120px] pointer-events-none -top-16 -left-16" />
              <div className="absolute w-[400px] h-[400px] rounded-full bg-sky-500/10 blur-[100px] pointer-events-none -bottom-16 -right-16" />

              <div className="z-10 max-w-4xl space-y-5 my-auto">
                <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 text-xs font-semibold tracking-wide">
                  <span>AUDITORÍA UPCONTA OFICIAL · PARTE II</span>
                  <span>·</span>
                  <span>{viewData.periodLabel.toUpperCase()}</span>
                </div>

                <h1 className="text-3xl sm:text-5xl font-extrabold tracking-tight text-white leading-tight">
                  {editValues['r2_cover_title'] || editValues['titleReport2'] || viewData.titleReport2}
                </h1>

                <p className="text-slate-400 text-sm sm:text-base max-w-2xl pt-1 leading-relaxed">
                  Auditoría completa de recaudación, liquidación de ventas de sistemas por socio, análisis de Pareto y cruce comercial con corte al {viewData.cutoffDateReport2}.
                </p>

                <div className="pt-6 flex flex-wrap items-center gap-6 text-xs text-slate-400">
                  <div className="flex items-center gap-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-emerald-400" />
                    <span>Plataforma: <strong>UpConta Sistemas ERP</strong></span>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-sky-400" />
                    <span>Cartera Auditada: <strong>{viewData.socios.length} Socios Comerciales</strong></span>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="w-2.5 h-2.5 rounded-full bg-amber-400" />
                    <span>Corte: <strong>{viewData.cutoffDateReport2}</strong></span>
                  </div>
                </div>
              </div>

              <div className="z-10 flex justify-between items-center text-xs text-slate-500 border-t border-slate-800/80 pt-4">
                <span>UpConta Audit Deck · Reporte 2</span>
                <span>Diapositiva 08 de 15</span>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* SLIDE 9: Balance General de Facturación 2026 (Reporte 2: Neto, IVA, Bruto) */}
          {/* ========================================================================= */}
          {currentSlide === 9 && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                  BALANCE GENERAL UPCONTA · REPORTE 2
                </span>
                <h2 className="text-2xl sm:text-3xl font-black text-white tracking-tight mt-0.5">
                  Balance General de Facturación 2026
                </h2>
                <p className="text-xs text-slate-400 mt-1">
                  Desglose consolidado de ingresos netos, desglose impositivo IVA (15%) e ingreso bruto registrado ({viewData.periodLabel})
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 my-auto">
                <div className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl relative overflow-hidden">
                  <div className="text-xs font-semibold uppercase tracking-wider text-slate-400">VENTAS NETAS (SIN IVA)</div>
                  <div className="text-3xl sm:text-4xl font-black text-white tabular-nums mt-1">
                    ${viewData.netSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                  <p className="text-xs text-slate-400 mt-3 pt-3 border-t border-slate-800">
                    Base imponible real correspondiente a los servicios y sistemas emitidos en el período.
                  </p>
                </div>

                <div className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl relative overflow-hidden">
                  <div className="text-xs font-semibold uppercase tracking-wider text-slate-400">IVA ESTIMADO (15%)</div>
                  <div className="text-3xl sm:text-4xl font-black text-amber-400 tabular-nums mt-1">
                    ${viewData.ivaAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                  <p className="text-xs text-slate-400 mt-3 pt-3 border-t border-slate-800">
                    Retención impositiva fiscal proyectada para la conciliación tributaria oficial.
                  </p>
                </div>

                <div className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl relative overflow-hidden">
                  <div className="text-xs font-semibold uppercase tracking-wider text-slate-400">TOTAL FACTURADO CON IVA</div>
                  <div className="text-3xl sm:text-4xl font-black text-sky-400 tabular-nums mt-1">
                    ${viewData.grossSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                  <p className="text-xs text-slate-400 mt-3 pt-3 border-t border-slate-800">
                    Recaudación bruta total esperada en cartera y cuentas por cobrar.
                  </p>
                </div>
              </div>

              <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-emerald-400">Conclusión de Auditoría:</strong> El período{' '}
                  <strong className="text-white">{viewData.periodLabel}</strong> registra una base imponible de{' '}
                  <strong className="text-emerald-400">${viewData.netSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD</strong>.
                </div>
                <span>Pág. 09/15</span>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* SLIDE 10: Evolución Mes a Mes (Reporte 2) */}
          {/* ========================================================================= */}
          {currentSlide === 10 && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                  TENDENCIA TEMPORAL · REPORTE 2
                </span>
                <h2 className="text-2xl sm:text-3xl font-black text-white tracking-tight mt-0.5">
                  Evolución de Ventas Mes a Mes (Sin IVA)
                </h2>
                <p className="text-xs text-slate-400 mt-1">
                  Facturación neta y peso relativo por período ({viewData.periodLabel})
                </p>
              </div>

              <div className="space-y-4 my-auto max-w-4xl w-full">
                {auditMonthlyR2.length === 0 ? (
                  <div className="py-12 text-center text-slate-500 text-sm">
                    Sin datos mensuales registrados para el período seleccionado.
                  </div>
                ) : (
                  auditMonthlyR2.map((item, idx) => {
                    const isLeader = item.total === maxAuditValR2;
                    return (
                      <div key={idx} className="space-y-1">
                        <div className="flex justify-between items-center text-xs">
                          <div className="flex items-center gap-2">
                            <span className="font-bold text-white uppercase w-28 tracking-wide">
                              {item.month}
                            </span>
                            {isLeader && (
                              <span className="text-[10px] px-2 py-0.2 rounded-full bg-emerald-500/20 text-emerald-300 font-semibold border border-emerald-500/40">
                                Período Activo
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

                        <div className="h-5 bg-slate-900 rounded-lg overflow-hidden border border-slate-800 p-0.5">
                          <div
                            style={{ width: `${Math.max(4, (item.total / maxAuditValR2) * 100)}%` }}
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

              <div className="p-3.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-emerald-400">Total {viewData.isFiltered ? `del Período (${viewData.periodLabel})` : 'Acumulado Registrado'}:</strong>{' '}
                  <strong className="text-white">
                    ${auditMonthlyR2.reduce((a, b) => a + b.total, 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD
                  </strong>
                </div>
                <span>Pág. 10/15</span>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* SLIDE 11: Ranking Pareto & Canales en una Hoja (Reporte 2) */}
          {/* ========================================================================= */}
          {currentSlide === 11 && (
            <div className="h-full flex flex-col justify-between p-6 sm:p-8 relative bg-[#0c1322]">
              <div className="flex items-center justify-between border-b border-slate-800/80 pb-2.5">
                <div>
                  <span className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                    AUDITORÍA & RENDIMIENTO DE SOCIOS · {viewData.periodLabel.toUpperCase()}
                  </span>
                  <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                    Ranking de Socios: Gráfico Pareto & Distribución por Canales
                  </h2>
                </div>
                <div className="text-xs bg-slate-900/90 px-3 py-1.5 rounded-lg border border-slate-800 text-slate-300 font-semibold">
                  Cartera Auditada: <strong className="text-emerald-400">${viewData.socios.reduce((a, s) => a + s.totalSales, 0).toLocaleString('en-US', { minimumFractionDigits: 2 })} USD</strong>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-5 my-auto">
                {/* Pareto Column */}
                <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between">
                  <div className="flex items-center justify-between pb-2 border-b border-slate-800 mb-2">
                    <span className="text-xs font-bold uppercase tracking-wider text-emerald-300">
                      Regla de Oro Pareto: Top 1–8 Socios
                    </span>
                    <span className="text-[11px] text-slate-400">Mayor Volumen</span>
                  </div>

                  <div className="space-y-1.5">
                    {viewData.socios.slice(0, 8).map((s) => {
                      const maxVal = viewData.socios[0]?.totalSales || 1;
                      const barWidth = Math.max(5, (s.totalSales / maxVal) * 100);
                      const isTop1 = s.rank === 1;
                      return (
                        <div key={s.rank} className="flex items-center gap-2 text-xs">
                          <span className={`w-5 font-bold tabular-nums text-center ${isTop1 ? 'text-amber-400 font-black' : s.rank <= 3 ? 'text-emerald-400' : 'text-slate-500'}`}>
                            {s.rank}.
                          </span>
                          <span className="w-32 truncate font-semibold text-slate-200" title={s.name}>
                            {s.name}
                          </span>
                          <span className={`text-[10px] px-1.5 py-0.2 rounded font-medium shrink-0 ${s.channel === 'UpConnect' ? 'bg-sky-500/10 text-sky-300' : 'bg-amber-500/10 text-amber-300'}`}>
                            {s.channel}
                          </span>
                          <div className="flex-1 bg-slate-800/60 rounded h-3.5 overflow-hidden">
                            <div style={{ width: `${barWidth}%` }} className={`h-full ${isTop1 ? 'bg-amber-400' : s.rank <= 3 ? 'bg-emerald-500' : 'bg-blue-600'} rounded`} />
                          </div>
                          <span className="w-20 text-right font-bold text-white tabular-nums">
                            ${s.totalSales.toFixed(2)}
                          </span>
                        </div>
                      );
                    })}
                  </div>
                </div>

                {/* Channels Summary Column */}
                <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between">
                  <div className="flex items-center justify-between pb-2 border-b border-slate-800 mb-2">
                    <span className="text-xs font-bold uppercase tracking-wider text-sky-300">
                      Distribución por Canales de Socio
                    </span>
                    <span className="text-[11px] text-slate-400">UpConnect vs Connect</span>
                  </div>

                  <div className="grid grid-cols-2 gap-3 mb-3">
                    <div className="p-3 rounded-xl bg-slate-950/80 border-t-2 border-sky-500 border-x border-b border-slate-800">
                      <div className="text-xs font-bold text-sky-400 mb-1">Distribuidor Upconnect</div>
                      <div className="text-xl font-black text-white tabular-nums">${viewData.sociosSummary.upconnect.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}</div>
                      <div className="text-[11px] text-slate-400 mt-1 flex justify-between">
                        <span>{viewData.sociosSummary.upconnect.count} Socios</span>
                        <span className="font-bold text-sky-300">{viewData.sociosSummary.upconnect.percentage}%</span>
                      </div>
                    </div>

                    <div className="p-3 rounded-xl bg-slate-950/80 border-t-2 border-amber-500 border-x border-b border-slate-800">
                      <div className="text-xs font-bold text-amber-400 mb-1">Distribuidor Connect</div>
                      <div className="text-xl font-black text-white tabular-nums">${viewData.sociosSummary.connectors.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}</div>
                      <div className="text-[11px] text-slate-400 mt-1 flex justify-between">
                        <span>{viewData.sociosSummary.connectors.count} Socios</span>
                        <span className="font-bold text-amber-300">{viewData.sociosSummary.connectors.percentage}%</span>
                      </div>
                    </div>
                  </div>

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

              <div className="p-3 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex justify-between items-center">
                <span>Concentración Top 5: {viewData.socios.length > 0 ? ((viewData.socios.slice(0, 5).reduce((a, b) => a + b.totalSales, 0) / (viewData.socios.reduce((a, b) => a + b.totalSales, 0) || 1)) * 100).toFixed(1) : '0.0'}% de la recaudación</span>
                <span>Pág. 11/15</span>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* SLIDE 12: Desglose y Gráficos Lineales Semana a Semana (Reporte 2) */}
          {/* ========================================================================= */}
          {currentSlide === 12 && (() => {
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
            }));
            const erpPoints = diffs.map((d, i) => ({
              x: xPoints[i],
              y: getY(weeklyMetric === 'sales' ? d.erpSales : d.erpCount),
            }));
            const contPoints = diffs.map((d, i) => ({
              x: xPoints[i],
              y: getY(weeklyMetric === 'sales' ? d.contadoresSales : d.contadoresCount),
            }));

            const makePolyline = (pts: { x: number; y: number }[]) => pts.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
            const makeArea = (pts: { x: number; y: number }[]) => {
              const base = `${pts[0].x.toFixed(1)},${(padT + plotH).toFixed(1)}`;
              const line = pts.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
              const end = `${pts[pts.length - 1].x.toFixed(1)},${(padT + plotH).toFixed(1)}`;
              return `${base} ${line} ${end}`;
            };

            return (
              <div className="h-full flex flex-col justify-between p-6 sm:p-8 relative bg-[#0c1322]">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-800/80 pb-2.5">
                  <div>
                    <span className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                      COMPARATIVA SEMANA A SEMANA · REPORTE 2
                    </span>
                    <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                      Foco de Rendimiento: Desglose y Gráficos Lineales
                    </h2>
                    <p className="text-xs text-slate-400 mt-0.5">
                      Diferenciadores: Planes de Facturación, ERP y Contadores ({viewData.periodLabel})
                    </p>
                  </div>

                  <div className="flex items-center gap-2 bg-slate-900/90 p-1 rounded-xl border border-slate-800 text-xs">
                    <button
                      onClick={() => setWeeklyMetric('sales')}
                      className={`px-3 py-1 rounded-lg font-semibold flex items-center gap-1.5 transition-all ${weeklyMetric === 'sales' ? 'bg-emerald-500 text-white shadow-md' : 'text-slate-400 hover:text-white'}`}
                    >
                      <DollarSign className="w-3.5 h-3.5" />
                      <span>Facturación ($ USD)</span>
                    </button>
                    <button
                      onClick={() => setWeeklyMetric('count')}
                      className={`px-3 py-1 rounded-lg font-semibold flex items-center gap-1.5 transition-all ${weeklyMetric === 'count' ? 'bg-sky-500 text-white shadow-md' : 'text-slate-400 hover:text-white'}`}
                    >
                      <Activity className="w-3.5 h-3.5" />
                      <span>Volumen (Operaciones)</span>
                    </button>
                  </div>
                </div>

                {/* SVG Chart */}
                <div className="bg-slate-900/90 rounded-2xl border border-slate-800 p-4 shadow-xl my-2 flex flex-col justify-between">
                  <div className="flex flex-wrap items-center justify-between gap-3 pb-2 border-b border-slate-800/60 text-xs">
                    <div className="flex items-center gap-4">
                      <div className="flex items-center gap-1.5">
                        <span className="w-3 h-1 rounded bg-sky-400" />
                        <span className="text-slate-300">Facturación:</span>
                        <strong className="text-sky-300 tabular-nums">{weeklyMetric === 'sales' ? `$${totFact.toFixed(2)}` : `${totFact} ops`}</strong>
                        <span className="text-[10px] text-slate-500">({((totFact / totCombined) * 100).toFixed(1)}%)</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <span className="w-3 h-1 rounded bg-emerald-400" />
                        <span className="text-slate-300">Sistemas ERP:</span>
                        <strong className="text-emerald-300 tabular-nums">{weeklyMetric === 'sales' ? `$${totErp.toFixed(2)}` : `${totErp} ops`}</strong>
                        <span className="text-[10px] text-slate-500">({((totErp / totCombined) * 100).toFixed(1)}%)</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <span className="w-3 h-1 rounded bg-purple-400" />
                        <span className="text-slate-300">Planes Contadores:</span>
                        <strong className="text-purple-300 tabular-nums">{weeklyMetric === 'sales' ? `$${totCont.toFixed(2)}` : `${totCont} ops`}</strong>
                        <span className="text-[10px] text-slate-500">({((totCont / totCombined) * 100).toFixed(1)}%)</span>
                      </div>
                    </div>
                  </div>

                  <div className="relative w-full h-[175px] mt-2">
                    <svg viewBox={`0 0 ${svgW} ${svgH}`} className="w-full h-full overflow-visible">
                      <defs>
                        <linearGradient id="u_gradFact" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="#38bdf8" stopOpacity="0.25" />
                          <stop offset="100%" stopColor="#38bdf8" stopOpacity="0.0" />
                        </linearGradient>
                        <linearGradient id="u_gradErp" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="#10b981" stopOpacity="0.25" />
                          <stop offset="100%" stopColor="#10b981" stopOpacity="0.0" />
                        </linearGradient>
                        <linearGradient id="u_gradCont" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="#a855f7" stopOpacity="0.25" />
                          <stop offset="100%" stopColor="#a855f7" stopOpacity="0.0" />
                        </linearGradient>
                      </defs>

                      {[0, 0.33, 0.66, 1].map((step, sIdx) => {
                        const val = scaleMax * step;
                        const y = padT + plotH - step * plotH;
                        return (
                          <g key={sIdx}>
                            <line x1={padL} x2={svgW - padR} y1={y} y2={y} stroke="#1e293b" strokeDasharray="3 3" strokeWidth="1" />
                            <text x={padL - 10} y={y + 3.5} textAnchor="end" fill="#64748b" fontSize="9.5" fontFamily="monospace">
                              {weeklyMetric === 'sales' ? `$${Math.round(val)}` : `${Math.round(val)}`}
                            </text>
                          </g>
                        );
                      })}

                      <polygon points={makeArea(factPoints)} fill="url(#u_gradFact)" />
                      <polygon points={makeArea(erpPoints)} fill="url(#u_gradErp)" />
                      <polygon points={makeArea(contPoints)} fill="url(#u_gradCont)" />

                      <polyline points={makePolyline(factPoints)} fill="none" stroke="#38bdf8" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" />
                      <polyline points={makePolyline(erpPoints)} fill="none" stroke="#10b981" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" />
                      <polyline points={makePolyline(contPoints)} fill="none" stroke="#a855f7" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" />

                      {hoveredWeekIndex !== null && (
                        <line x1={xPoints[hoveredWeekIndex]} x2={xPoints[hoveredWeekIndex]} y1={padT} y2={padT + plotH} stroke="#94a3b8" strokeDasharray="2 2" strokeWidth="1.5" opacity="0.8" />
                      )}

                      {diffs.map((d, idx) => (
                        <g key={idx}>
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
                          <circle cx={factPoints[idx].x} cy={factPoints[idx].y} r={hoveredWeekIndex === idx ? 6 : 4} fill="#0f172a" stroke="#38bdf8" strokeWidth={hoveredWeekIndex === idx ? 3 : 2} />
                          <circle cx={erpPoints[idx].x} cy={erpPoints[idx].y} r={hoveredWeekIndex === idx ? 6 : 4} fill="#0f172a" stroke="#10b981" strokeWidth={hoveredWeekIndex === idx ? 3 : 2} />
                          <circle cx={contPoints[idx].x} cy={contPoints[idx].y} r={hoveredWeekIndex === idx ? 6 : 4} fill="#0f172a" stroke="#a855f7" strokeWidth={hoveredWeekIndex === idx ? 3 : 2} />

                          <text x={xPoints[idx]} y={padT + plotH + 15} textAnchor="middle" fill={hoveredWeekIndex === idx ? '#ffffff' : '#cbd5e1'} fontSize="11" fontWeight={hoveredWeekIndex === idx ? 'bold' : '600'}>
                            {d.weekName}
                          </text>
                          <text x={xPoints[idx]} y={padT + plotH + 28} textAnchor="middle" fill="#64748b" fontSize="9.5">
                            {d.dateRange}
                          </text>
                        </g>
                      ))}
                    </svg>

                    {hoveredWeekIndex !== null && diffs[hoveredWeekIndex] && (() => {
                      const hw = diffs[hoveredWeekIndex];
                      const leftPos = Math.min(Math.max(xPoints[hoveredWeekIndex] - 80, 20), svgW - 200);
                      return (
                        <div
                          style={{ left: `${(leftPos / svgW) * 100}%`, top: '-8px' }}
                          className="absolute z-20 pointer-events-none bg-slate-950/95 border border-slate-700 rounded-xl p-2.5 shadow-2xl backdrop-blur-md text-[11px] space-y-1 w-48"
                        >
                          <div className="font-bold text-white border-b border-slate-800 pb-1 flex justify-between">
                            <span>{hw.weekName}</span>
                            <span className="text-slate-400 font-normal">{hw.dateRange}</span>
                          </div>
                          <div className="flex justify-between text-sky-400">
                            <span>Facturación:</span>
                            <strong className="font-mono">{weeklyMetric === 'sales' ? `$${hw.facturacionSales.toFixed(2)}` : `${hw.facturacionCount} ops`}</strong>
                          </div>
                          <div className="flex justify-between text-emerald-400">
                            <span>ERP:</span>
                            <strong className="font-mono">{weeklyMetric === 'sales' ? `$${hw.erpSales.toFixed(2)}` : `${hw.erpCount} ops`}</strong>
                          </div>
                          <div className="flex justify-between text-purple-400">
                            <span>Contadores:</span>
                            <strong className="font-mono">{weeklyMetric === 'sales' ? `$${hw.contadoresSales.toFixed(2)}` : `${hw.contadoresCount} ops`}</strong>
                          </div>
                          <div className="border-t border-slate-800 pt-1 flex justify-between font-bold text-white">
                            <span>Total Semana:</span>
                            <span className="text-emerald-400 font-mono">{weeklyMetric === 'sales' ? `$${hw.totalSales.toFixed(2)} USD` : `${hw.totalCount} ops`}</span>
                          </div>
                        </div>
                      );
                    })()}
                  </div>
                </div>

                {/* 4 Cards */}
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 my-1">
                  {diffs.map((w, idx) => (
                    <div
                      key={idx}
                      className={`p-3 rounded-xl border transition-all ${
                        hoveredWeekIndex === idx ? 'bg-slate-900 border-sky-500/60 shadow-lg' : 'bg-slate-900/80 border-slate-800'
                      }`}
                    >
                      <div className="flex items-center justify-between text-xs mb-1">
                        <span className="font-bold text-white">{w.weekName}</span>
                        <span className="text-[10px] text-slate-400">{w.dateRange}</span>
                      </div>
                      <div className="text-xl font-black text-white tabular-nums my-1">
                        ${w.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                      </div>
                      <div className="space-y-0.5 text-[10px] border-t border-slate-800 pt-1 text-slate-300">
                        <div className="flex justify-between"><span className="text-sky-400">Fact:</span><span>${w.facturacionSales.toFixed(2)}</span></div>
                        <div className="flex justify-between"><span className="text-emerald-400">ERP:</span><span>${w.erpSales.toFixed(2)}</span></div>
                        <div className="flex justify-between"><span className="text-purple-400">Cont:</span><span>${w.contadoresSales.toFixed(2)}</span></div>
                      </div>
                    </div>
                  ))}
                </div>

                <div className="p-3 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                  <div>
                    <strong className="text-emerald-400">Total Auditado:</strong> ${diffs.reduce((a, b) => a + b.totalSales, 0).toLocaleString('en-US', { minimumFractionDigits: 2 })} USD
                  </div>
                  <span>Pág. 12/15</span>
                </div>
              </div>
            );
          })()}

          {/* ========================================================================= */}
          {/* SLIDE 13: Ventas por Categoría de Solución (Reporte 2) */}
          {/* ========================================================================= */}
          {currentSlide === 13 && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                  CATÁLOGO COMERCIAL · REPORTE 2
                </span>
                <h2 className="text-2xl sm:text-3xl font-black text-white tracking-tight mt-0.5">
                  Ventas por Categoría de Solución UpConta
                </h2>
                <p className="text-xs text-slate-400 mt-1">
                  Distribución de ingresos según planes de suscripción ERP y productos contratados ({viewData.periodLabel})
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 my-auto">
                {activeCategoriesR2.map((cat, idx) => (
                  <div key={idx} className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between">
                    <div>
                      <div className="flex items-center justify-between text-xs mb-2">
                        <span className="font-bold text-sky-400 uppercase tracking-wider text-[11px]">{cat.badge || 'Categoría'}</span>
                        <span className="font-bold text-white bg-slate-800 px-2 py-0.5 rounded text-[11px]">{cat.percentage}%</span>
                      </div>
                      <h3 className="text-base font-bold text-white">{cat.title}</h3>
                      <div className="text-2xl font-black text-emerald-400 tabular-nums my-2">
                        ${cat.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                      </div>
                      <p className="text-xs text-slate-400">{cat.plansDescription}</p>
                    </div>
                    <div className="pt-3 border-t border-slate-800 text-[11px] text-slate-400 flex items-center justify-between">
                      <span>Participación:</span>
                      <strong className="text-slate-200">{cat.percentage}% de ventas</strong>
                    </div>
                  </div>
                ))}
              </div>

              <div className="p-3.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-emerald-400">Estrategia de Socios:</strong> Foco en tickets medianos y altos (ERP + Planes Contador).
                </div>
                <span>Pág. 13/15</span>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* SLIDE 14: Canal de Distribución - Embudo de Socios (Reporte 2) */}
          {/* ========================================================================= */}
          {currentSlide === 14 && (
            <div className="h-full flex flex-col justify-between p-6 sm:p-8 relative bg-[#0c1322]">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                  <span className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                    CANAL DE DISTRIBUCIÓN · REPORTE 2
                  </span>
                  <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                    Embudo de Socios y Comunidades Activas
                  </h2>
                  <p className="text-xs text-slate-400 mt-0.5">
                    Conteo cuantitativo de socios, franquiciados UpConta y red de distribuidores Connect
                  </p>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={() => handleResetCommunities(2)}
                    className="text-xs px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-amber-400 border border-amber-500/30 flex items-center gap-1.5 transition-colors font-medium"
                    title="Restablecer a valores automáticos del Excel"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    <span>Restablecer</span>
                  </button>
                  <button
                    onClick={() => {
                      setEditingTargetPart(2);
                      setTempUpconta(upcontaMembersR2);
                      setTempFranquicia(franquiciaMembersR2);
                      setTempFormacion(formacionMembersR2);
                      setIsEditingCommunitiesModal(true);
                    }}
                    className="text-xs px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white font-semibold flex items-center gap-1.5 shadow-md shadow-emerald-600/30 transition-colors"
                  >
                    <Edit3 className="w-3.5 h-3.5" />
                    <span>Modificar Miembros</span>
                  </button>
                </div>
              </div>

              {/* 3 Funnel Cards */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 my-2">
                <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl text-center">
                  <h3 className="text-sm font-bold text-white">UpConta Socios</h3>
                  <div className="text-3xl font-black text-sky-400 tabular-nums my-1">{upcontaMembersR2}</div>
                  <div className="text-[11px] font-semibold text-slate-400">Socios Distribuidor Upconnect</div>
                </div>

                <div className="p-4 rounded-xl bg-slate-900/95 border-2 border-emerald-500 shadow-xl text-center">
                  <h3 className="text-sm font-bold text-white">Franquicia Contadores VIP</h3>
                  <div className="text-3xl font-black text-white tabular-nums my-1">{franquiciaMembersR2}</div>
                  <div className="text-[11px] font-semibold text-purple-400">Socios Distribuidor Connect</div>
                </div>

                <div className="p-4 rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl text-center">
                  <h3 className="text-sm font-bold text-white">Capacitación Franquicia</h3>
                  <div className="text-3xl font-black text-emerald-400 tabular-nums my-1">{formacionMembersR2}</div>
                  <div className="text-[11px] font-semibold text-slate-400">Miembros en Inducción</div>
                </div>
              </div>

              <div className="my-1">
                <SociosChannelTable
                  summary={communitiesTableSummaryR2}
                  variant="slide"
                  title="Distribución y Cantidad de Socios por Canal: Distribuidor Upconnect vs Distribuidor Connect"
                />
              </div>

              <div className="p-2.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-emerald-400">Total Cartera Auditada:</strong> {communitiesTableSummaryR2.totalSociosCount} socios registrados ({communitiesTableSummaryR2.upconnect.count} UP / {communitiesTableSummaryR2.connectors.count} CO) con {communitiesTableSummaryR2.totalOperationsCount} ventas por ${communitiesTableSummaryR2.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD.
                </div>
                <span>Pág. 14/15</span>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* SLIDE 15: Siguiente Fase: Cruce Comercial UpConta (Reporte 2) */}
          {/* ========================================================================= */}
          {currentSlide === 15 && (() => {
            const activeCross = viewData.commercialCrossReport2 || dataset.commercialCross;
            return (
              <div className="h-full flex flex-col justify-center items-center text-center p-8 sm:p-14 relative bg-gradient-to-br from-[#0c1322] via-[#080d19] to-[#040812]">
                <div className="z-10 max-w-2xl space-y-6">
                  <div className="w-12 h-12 rounded-2xl bg-emerald-500 text-white flex items-center justify-center mx-auto shadow-lg shadow-emerald-500/30">
                    <Activity className="w-6 h-6" />
                  </div>

                  <div className="text-xs font-bold uppercase tracking-wider text-emerald-400">
                    SIGUIENTE FASE: CRUCE COMERCIAL UPCONTA · FIN DE LA PRESENTACIÓN
                  </div>

                  <h2 className="text-3xl sm:text-4xl font-black text-white tracking-tight">
                    RESUMEN DE VENTAS UPCONTA Y SOCIOS
                  </h2>

                  <div className="space-y-2 py-4 text-base font-semibold text-slate-300">
                    <div>VENTAS EQUIPO COMERCIAL: <strong className="text-white">${activeCross?.commercialTeamSales ? activeCross.commercialTeamSales.toLocaleString('en-US', { minimumFractionDigits: 2 }) : '0.00'}</strong></div>
                    <div>VENTAS ORGÁNICAS: <strong className="text-white">${activeCross?.organicSales ? activeCross.organicSales.toLocaleString('en-US', { minimumFractionDigits: 2 }) : '0.00'}</strong></div>
                    <div>VENTAS SOCIOS: <strong className="text-emerald-400">${(activeCross?.partnersSales || viewData.netSales).toLocaleString('en-US', { minimumFractionDigits: 2 })}</strong></div>
                  </div>

                  <div className="grid grid-cols-2 gap-8 pt-6 border-t border-slate-800/80 max-w-lg mx-auto">
                    <div className="text-left">
                      <div className="text-[11px] font-bold text-slate-500 uppercase">{activeCross?.channel1Name || 'CANAL 1 AUDITADO'}</div>
                      <div className="text-sm sm:text-base font-bold text-sky-400 mt-1">
                        Socios UpConnect (${(activeCross?.channel1Amount || viewData.netSales).toLocaleString('en-US', { minimumFractionDigits: 2 })})
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
                  <span>Presentación Unificada Finalizada</span>
                  <span>Pág. 15 de 15</span>
                </div>
              </div>
            );
          })()}

        </div>
      </div>

      {/* Bottom Timeline Indicator */}
      <div className="no-print px-6 py-2.5 bg-[#09101d] border-t border-slate-800 flex items-center justify-between text-xs text-slate-400">
        <div className="flex items-center gap-2">
          <span className="font-semibold text-slate-300 truncate max-w-[280px] sm:max-w-md">
            {currentSlideDef.globalIndex}. {currentSlideDef.title}
          </span>
        </div>

        {/* 15 Progress Dots */}
        <div className="flex items-center gap-1.5">
          {unifiedSlides.map((s) => (
            <button
              key={s.id}
              onClick={() => setCurrentSlide(s.globalIndex)}
              className={`h-2 rounded-full transition-all ${
                currentSlide === s.globalIndex
                  ? s.reportPart === 1 ? 'w-6 bg-blue-500' : 'w-6 bg-emerald-500'
                  : s.reportPart === 1 ? 'w-2 bg-blue-900/60 hover:bg-blue-600' : 'w-2 bg-emerald-900/60 hover:bg-emerald-600'
              }`}
              title={`${s.globalIndex}. ${s.title}`}
            />
          ))}
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={() => setCurrentSlide((prev) => Math.max(prev - 1, 1))}
            disabled={currentSlide === 1}
            className="p-1 rounded text-slate-400 hover:text-white disabled:opacity-30"
          >
            <ChevronLeft className="w-4 h-4" />
          </button>
          <span className="font-bold text-white tabular-nums">{currentSlide} / 15</span>
          <button
            onClick={() => setCurrentSlide((prev) => Math.min(prev + 1, totalSlides))}
            disabled={currentSlide === totalSlides}
            className="p-1 rounded text-slate-400 hover:text-white disabled:opacity-30"
          >
            <ChevronRight className="w-4 h-4" />
          </button>
        </div>
      </div>

      {/* MODAL: Modificar Miembros de Comunidades / Embudo */}
      {isEditingCommunitiesModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="w-full max-w-md bg-slate-950 border border-slate-800 rounded-2xl p-6 shadow-2xl space-y-5">
            <div className="flex items-center justify-between pb-3 border-b border-slate-800">
              <div className="flex items-center gap-2">
                <Users className="w-5 h-5 text-emerald-400" />
                <h3 className="text-base font-bold text-white">
                  Modificar Miembros ({editingTargetPart === 1 ? 'Reporte 1: Firmas' : 'Reporte 2: Socios'})
                </h3>
              </div>
              <button
                onClick={() => setIsEditingCommunitiesModal(false)}
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
                onClick={() => handleResetCommunities(editingTargetPart)}
                className="text-xs px-3 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors flex items-center gap-1.5 font-medium"
              >
                <RotateCcw className="w-3.5 h-3.5" />
                <span>Restablecer</span>
              </button>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => setIsEditingCommunitiesModal(false)}
                  className="text-xs px-3 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white transition-colors"
                >
                  Cancelar
                </button>
                <button
                  onClick={() => handleSaveCommunities(editingTargetPart, tempUpconta, tempFranquicia, tempFormacion)}
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

      {/* Slide Editor Modal */}
      <SlideEditorModal
        isOpen={isEditorModalOpen}
        onClose={() => setIsEditorModalOpen(false)}
        reportType={currentSlideDef.reportPart}
        slideId={currentSlideDef.internalId}
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
