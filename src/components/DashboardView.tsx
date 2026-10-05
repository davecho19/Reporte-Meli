import React, { useState, useMemo } from 'react';
import { GlobalDataset, TransactionRecord, ReportTimeFilter } from '../types';
import { 
  TrendingUp, 
  DollarSign, 
  FileCheck2, 
  Users2, 
  Calendar, 
  Search, 
  Filter, 
  Download, 
  ArrowUpDown, 
  Layers, 
  Zap, 
  Share2,
  CheckCircle,
  PlusCircle,
  FileSignature,
  Clock,
  Sparkles,
  ChevronRight,
  RotateCcw,
  BarChart3,
  Trash2,
  AlertTriangle,
  Table,
  FileDown,
  Loader2
} from 'lucide-react';
import { 
  exportDatasetToExcel, 
  downloadTemplateFirmas, 
  downloadTemplateSocios 
} from '../services/excelService';
import { exportCombinedPresentationsToPPTX } from '../services/pptxExportService';
import { computeReportView, MONTH_NAMES_ES, extractTransactionMonth } from '../utils/reportFilters';
import { SociosChannelTable } from './SociosChannelTable';

interface Props {
  dataset: GlobalDataset;
  onOpenUploadFirma: () => void;
  onOpenUploadSocios: () => void;
  onClearData: () => void;
  onUpdateFilter: (filter: ReportTimeFilter) => void;
  onAddTransaction: (trx: TransactionRecord) => void;
  onToggleSocioChannel?: (socioName: string) => void;
}

export const DashboardView: React.FC<Props> = ({
  dataset,
  onOpenUploadFirma,
  onOpenUploadSocios,
  onClearData,
  onUpdateFilter,
  onAddTransaction,
  onToggleSocioChannel,
}) => {
  const currentFilter = dataset.reportFilter || { type: 'all' };
  const viewData = computeReportView(dataset);

  const [selectedChannel, setSelectedChannel] = useState<'ALL' | 'UpConnect' | 'Connectors'>('ALL');
  const [searchTerm, setSearchTerm] = useState('');
  const [sortField, setSortField] = useState<'date' | 'value' | 'clientName' | 'duration'>('date');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('desc');
  const [page, setPage] = useState(1);
  const pageSize = 10;
  const [showClearConfirm, setShowClearConfirm] = useState(false);
  const [isExportingPPTX, setIsExportingPPTX] = useState(false);

  const handleDownloadPPTX = async () => {
    setIsExportingPPTX(true);
    try {
      await exportCombinedPresentationsToPPTX(dataset);
    } catch (e) {
      console.error('Error generating unified PPTX from dashboard', e);
    } finally {
      setIsExportingPPTX(false);
    }
  };

  // Table tab: Socios vs Transactions
  const [activeTableTab, setActiveTableTab] = useState<'socios' | 'transactions'>('socios');
  const [socioSearchTerm, setSocioSearchTerm] = useState('');
  const [socioChannelFilter, setSocioChannelFilter] = useState<'ALL' | 'UpConnect' | 'Connectors'>('ALL');
  const [socioPage, setSocioPage] = useState(1);
  const socioPageSize = 10;

  // Local state for the filter panel
  const [filterMode, setFilterMode] = useState<'all' | 'month' | 'week' | 'range'>(currentFilter.type || 'all');
  const [selectedMonthName, setSelectedMonthName] = useState<string>(currentFilter.month || 'SEPTIEMBRE');
  const [selectedWeekId, setSelectedWeekId] = useState<string>(currentFilter.weekId || 'w2');
  const [startDate, setStartDate] = useState<string>(currentFilter.startDate || '2026-09-01');
  const [endDate, setEndDate] = useState<string>(currentFilter.endDate || '2026-09-20');

  const availableMonths = useMemo(() => {
    const monthsSet = new Set<string>();
    dataset.transactions.forEach((t) => {
      const { monthName } = extractTransactionMonth(t);
      if (monthName) monthsSet.add(monthName);
    });
    dataset.monthlyMetrics.forEach((m) => {
      if (m.month) monthsSet.add(m.month.toUpperCase());
    });
    const months = MONTH_NAMES_ES.filter((m) => monthsSet.has(m));
    return months.length > 0 ? months : ['AGOSTO', 'SEPTIEMBRE', 'OCTUBRE'];
  }, [dataset.transactions, dataset.monthlyMetrics]);

  const applyTimeFilter = (mode: 'all' | 'month' | 'week' | 'range') => {
    setFilterMode(mode);
    setPage(1);

    if (mode === 'all') {
      onUpdateFilter({ type: 'all' });
    } else if (mode === 'month') {
      onUpdateFilter({ type: 'month', month: selectedMonthName });
    } else if (mode === 'week') {
      onUpdateFilter({ type: 'week', weekId: selectedWeekId });
    } else if (mode === 'range') {
      onUpdateFilter({ type: 'range', startDate, endDate });
    }
  };

  // Filtered transactions (combining time viewData.transactions with search/channel filters)
  const filteredTransactions = useMemo(() => {
    return viewData.transactions.filter((trx) => {
      // Channel match
      if (selectedChannel !== 'ALL' && trx.channel !== selectedChannel) {
        return false;
      }
      // Search term
      if (searchTerm.trim()) {
        const term = searchTerm.toLowerCase();
        const fullName = `${trx.clientName} ${trx.clientLastName}`.toLowerCase();
        const matchId = trx.uniqueId.toLowerCase().includes(term);
        const matchSocio = (trx.socio || '').toLowerCase().includes(term);
        const matchRole = trx.role.toLowerCase().includes(term);
        const matchCategory = (trx.solutionCategory || '').toLowerCase().includes(term);
        return fullName.includes(term) || matchId || matchSocio || matchRole || matchCategory;
      }
      return true;
    }).sort((a, b) => {
      if (sortField === 'value') {
        return sortOrder === 'asc' ? a.value - b.value : b.value - a.value;
      }
      if (sortField === 'date') {
        return sortOrder === 'asc' ? a.date.localeCompare(b.date) : b.date.localeCompare(a.date);
      }
      if (sortField === 'clientName') {
        return sortOrder === 'asc' ? a.clientName.localeCompare(b.clientName) : b.clientName.localeCompare(a.clientName);
      }
      return 0;
    });
  }, [viewData.transactions, selectedChannel, searchTerm, sortField, sortOrder]);

  const totalPages = Math.ceil(filteredTransactions.length / pageSize) || 1;
  const paginatedTransactions = filteredTransactions.slice((page - 1) * pageSize, page * pageSize);

  // Filtered & paginated socios for the detailed socios table
  const filteredSociosList = useMemo(() => {
    return viewData.socios.filter((s) => {
      if (socioChannelFilter !== 'ALL' && s.channel !== socioChannelFilter) {
        return false;
      }
      if (!socioSearchTerm) return true;
      const term = socioSearchTerm.toLowerCase();
      return (
        s.name.toLowerCase().includes(term) ||
        (s.role && s.role.toLowerCase().includes(term)) ||
        (s.topPlan && s.topPlan.toLowerCase().includes(term))
      );
    });
  }, [viewData.socios, socioSearchTerm, socioChannelFilter]);

  const totalSocioPages = Math.ceil(filteredSociosList.length / socioPageSize) || 1;
  const paginatedSocios = filteredSociosList.slice((socioPage - 1) * socioPageSize, socioPage * socioPageSize);

  const avgTicket = viewData.totalCount > 0 ? (viewData.totalSales / viewData.totalCount) : 0;

  const handleToggleSort = (field: 'date' | 'value' | 'clientName' | 'duration') => {
    if (sortField === field) {
      setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortOrder('desc');
    }
  };

  return (
    <div className="w-full max-w-[1440px] mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8">
      {/* Top Welcome & Actions Header */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 pb-6 border-b border-slate-800">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-xs uppercase font-bold tracking-wider text-sky-400">
              PANEL GERENCIAL EJECUTIVO
            </span>
            <span className="text-slate-600">·</span>
            <span className="text-xs text-emerald-400 font-semibold">{viewData.periodLabel}</span>
          </div>
          <h1 className="text-2xl sm:text-3xl font-extrabold text-white tracking-tight mt-1">
            Dashboard Integral de Ventas & Emisiones
          </h1>
          <p className="text-sm text-slate-400 mt-1 max-w-2xl">
            Carga continua de archivos de firma y socios con filtrado dinámico por semana, mes y rangos de fecha que actualiza los títulos y métricas de las presentaciones en tiempo real.
          </p>
        </div>

        {/* Action Buttons: Carga Archivo Firma, Carga Archivo Socios & Limpiar Data */}
        <div className="flex flex-wrap items-center gap-2.5">
          <button
            onClick={onOpenUploadFirma}
            className="flex items-center gap-2 px-4 py-2.5 text-xs font-bold rounded-xl bg-blue-600 hover:bg-blue-500 text-white shadow-lg shadow-blue-600/25 transition-all"
            title="Cargar archivo Excel con reporte de firmas y emisiones"
          >
            <FileSignature className="w-4 h-4" />
            <span>Carga Archivo Firma</span>
          </button>

          <button
            onClick={onOpenUploadSocios}
            className="flex items-center gap-2 px-4 py-2.5 text-xs font-bold rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white shadow-lg shadow-emerald-600/25 transition-all"
            title="Cargar archivo Excel con ranking y cartera de socios contables"
          >
            <Users2 className="w-4 h-4" />
            <span>Carga Archivo Socios</span>
          </button>

          <button
            onClick={() => setShowClearConfirm(true)}
            className="flex items-center gap-2 px-4 py-2.5 text-xs font-bold rounded-xl bg-rose-500/10 hover:bg-rose-500/20 text-rose-400 hover:text-rose-300 border border-rose-500/30 transition-all shadow-sm"
            title="Borrar toda la información para volver a cargar datos desde cero"
          >
            <Trash2 className="w-4 h-4" />
            <span>Limpiar Data</span>
          </button>

          <button
            onClick={() => exportDatasetToExcel(dataset)}
            className="flex items-center gap-2 px-3.5 py-2.5 text-xs font-semibold rounded-xl bg-slate-800/90 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-colors"
          >
            <Download className="w-4 h-4" />
            <span>Exportar (.xlsx)</span>
          </button>

          <button
            onClick={handleDownloadPPTX}
            disabled={isExportingPPTX}
            className="flex items-center gap-2 px-3.5 py-2.5 text-xs font-bold rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white transition-all shadow-md shadow-emerald-950/40 disabled:opacity-50"
            title="Descargar presentación unificada con los reportes 1 y 2 consolidados (.pptx)"
          >
            {isExportingPPTX ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <FileDown className="w-4 h-4" />
            )}
            <span>Descargar PPTX Unificado</span>
          </button>
        </div>
      </div>

      {/* DYNAMIC TIME FILTER CONTROLLER BAR */}
      <div className="p-5 rounded-2xl bg-gradient-to-r from-[#0d1627] via-[#0b1220] to-[#0d1627] border border-sky-500/30 shadow-xl space-y-4">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-sky-500/20 text-sky-400 flex items-center justify-center">
              <Calendar className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-xs font-extrabold uppercase tracking-wider text-sky-400">
                Control de Periodo del Reporte
              </h3>
              <p className="text-xs text-slate-300">
                Selecciona la escala de tiempo para sincronizar automáticamente el <strong>Título del Reporte</strong> y las métricas
              </p>
            </div>
          </div>

          {/* Time mode selector pills */}
          <div className="flex items-center p-1 bg-slate-900 border border-slate-700 rounded-xl">
            <button
              onClick={() => applyTimeFilter('all')}
              className={`px-3 py-1.5 text-xs font-bold rounded-lg transition-colors ${
                filterMode === 'all' ? 'bg-sky-500 text-white shadow' : 'text-slate-400 hover:text-white'
              }`}
            >
              Consolidado 2026
            </button>
            <button
              onClick={() => applyTimeFilter('month')}
              className={`px-3 py-1.5 text-xs font-bold rounded-lg transition-colors ${
                filterMode === 'month' ? 'bg-sky-500 text-white shadow' : 'text-slate-400 hover:text-white'
              }`}
            >
              Por Mes
            </button>
            <button
              onClick={() => applyTimeFilter('week')}
              className={`px-3 py-1.5 text-xs font-bold rounded-lg transition-colors ${
                filterMode === 'week' ? 'bg-sky-500 text-white shadow' : 'text-slate-400 hover:text-white'
              }`}
            >
              Por Semana
            </button>
            <button
              onClick={() => applyTimeFilter('range')}
              className={`px-3 py-1.5 text-xs font-bold rounded-lg transition-colors ${
                filterMode === 'range' ? 'bg-sky-500 text-white shadow' : 'text-slate-400 hover:text-white'
              }`}
            >
              Por Rango
            </button>
          </div>
        </div>

        {/* Parameter pickers based on selected mode */}
        <div className="flex flex-wrap items-center gap-4 pt-1 border-t border-slate-800/80 text-xs">
          {filterMode === 'month' && (
            <div className="flex items-center gap-2">
              <span className="text-slate-400 font-semibold">Seleccionar Mes:</span>
              <select
                value={selectedMonthName}
                onChange={(e) => {
                  setSelectedMonthName(e.target.value);
                  onUpdateFilter({ type: 'month', month: e.target.value });
                }}
                className="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-1.5 text-xs focus:outline-none focus:border-sky-500 font-semibold"
              >
                {['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'].map((m) => (
                  <option key={m} value={m}>{m}</option>
                ))}
              </select>
            </div>
          )}

          {filterMode === 'week' && (
            <div className="flex items-center gap-2">
              <span className="text-slate-400 font-semibold">Seleccionar Semana:</span>
              <select
                value={selectedWeekId}
                onChange={(e) => {
                  setSelectedWeekId(e.target.value);
                  onUpdateFilter({ type: 'week', weekId: e.target.value });
                }}
                className="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-1.5 text-xs focus:outline-none focus:border-sky-500 font-semibold"
              >
                <option value="w1">Semana 1 (del 01 al 06 de septiembre)</option>
                <option value="w2">Semana 2 (del 07 al 13 de septiembre)</option>
                <option value="w3">Semana 3 (del 14 al 20 de septiembre)</option>
                <option value="w4">Semana 4 (del 21 al 30 de septiembre)</option>
              </select>
            </div>
          )}

          {filterMode === 'range' && (
            <div className="flex flex-wrap items-center gap-3">
              <div className="flex items-center gap-2">
                <span className="text-slate-400 font-semibold">Desde:</span>
                <input
                  type="date"
                  value={startDate}
                  onChange={(e) => {
                    setStartDate(e.target.value);
                    onUpdateFilter({ type: 'range', startDate: e.target.value, endDate });
                  }}
                  className="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-1 text-xs focus:outline-none focus:border-sky-500"
                />
              </div>

              <div className="flex items-center gap-2">
                <span className="text-slate-400 font-semibold">Hasta:</span>
                <input
                  type="date"
                  value={endDate}
                  onChange={(e) => {
                    setEndDate(e.target.value);
                    onUpdateFilter({ type: 'range', startDate, endDate: e.target.value });
                  }}
                  className="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-1 text-xs focus:outline-none focus:border-sky-500"
                />
              </div>
            </div>
          )}

          {/* Dynamic Generated Title Preview Pill */}
          <div className="ml-auto flex items-center gap-2 px-3 py-1 rounded-lg bg-sky-500/10 border border-sky-500/30 text-sky-300">
            <Sparkles className="w-3.5 h-3.5 text-sky-400" />
            <span className="text-[11px]">
              Título activo del reporte: <strong className="text-white">{viewData.titleReport1}</strong>
            </span>
          </div>
        </div>
      </div>

      {/* Top 4 KPI Metrics Grid (Dynamically updated) */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
        {/* KPI 1 */}
        <div className="p-5 rounded-2xl bg-[#0e1626] border border-slate-800/90 shadow-xl relative overflow-hidden group hover:border-sky-500/40 transition-colors">
          <div className="flex items-center justify-between text-xs text-slate-400 font-semibold mb-2">
            <span>FACTURACIÓN TOTAL ({viewData.periodLabel.toUpperCase()})</span>
            <div className="w-8 h-8 rounded-lg bg-blue-500/10 text-sky-400 flex items-center justify-center">
              <DollarSign className="w-4 h-4" />
            </div>
          </div>
          <div className="text-3xl font-black text-white tabular-nums tracking-tight">
            ${viewData.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
          </div>
          <div className="mt-2 flex items-center gap-2 text-xs">
            <span className="text-sky-400 font-bold tabular-nums">
              UP: ${viewData.totalUpSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}
            </span>
            <span className="text-slate-600">|</span>
            <span className="text-amber-400 font-bold tabular-nums">
              CO: ${viewData.totalCoSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}
            </span>
          </div>
          <div className="absolute bottom-0 left-0 right-0 h-1 bg-gradient-to-r from-blue-500 to-sky-400" />
        </div>

        {/* KPI 2 */}
        <div className="p-5 rounded-2xl bg-[#0e1626] border border-slate-800/90 shadow-xl relative overflow-hidden group hover:border-emerald-500/40 transition-colors">
          <div className="flex items-center justify-between text-xs text-slate-400 font-semibold mb-2">
            <span>TOTAL FIRMAS EMITIDAS</span>
            <div className="w-8 h-8 rounded-lg bg-emerald-500/10 text-emerald-400 flex items-center justify-center">
              <FileCheck2 className="w-4 h-4" />
            </div>
          </div>
          <div className="text-3xl font-black text-white tabular-nums tracking-tight">
            {viewData.totalCount}
          </div>
          <div className="mt-2 flex items-center gap-2 text-xs">
            <span className="text-emerald-400 font-bold tabular-nums">Upconnect: {viewData.totalUpCount}</span>
            <span className="text-slate-600">·</span>
            <span className="text-amber-400 font-bold tabular-nums">Connect: {viewData.totalCoCount}</span>
          </div>
          <div className="absolute bottom-0 left-0 right-0 h-1 bg-emerald-500" />
        </div>

        {/* KPI 3 */}
        <div className="p-5 rounded-2xl bg-[#0e1626] border border-slate-800/90 shadow-xl relative overflow-hidden group hover:border-purple-500/40 transition-colors">
          <div className="flex items-center justify-between text-xs text-slate-400 font-semibold mb-2">
            <span>SOCIOS EN CARTERA</span>
            <div className="w-8 h-8 rounded-lg bg-purple-500/10 text-purple-400 flex items-center justify-center">
              <Users2 className="w-4 h-4" />
            </div>
          </div>
          <div className="text-3xl font-black text-white tabular-nums tracking-tight">
            {viewData.sociosSummary.totalSociosCount}
          </div>
          <div className="mt-2 flex items-center gap-2 text-xs">
            <span className="text-sky-400 font-bold tabular-nums">Upconnect: {viewData.sociosSummary.upconnect.count}</span>
            <span className="text-slate-600">·</span>
            <span className="text-amber-400 font-bold tabular-nums">Connect: {viewData.sociosSummary.connectors.count}</span>
          </div>
          <div className="mt-1 text-[11px] text-slate-400">
            Ticket promedio: <strong className="text-white">${avgTicket.toFixed(2)}</strong> / firma
          </div>
          <div className="absolute bottom-0 left-0 right-0 h-1 bg-purple-500" />
        </div>

        {/* KPI 4 */}
        <div className="p-5 rounded-2xl bg-[#0e1626] border border-slate-800/90 shadow-xl relative overflow-hidden group hover:border-amber-500/40 transition-colors">
          <div className="flex items-center justify-between text-xs text-slate-400 font-semibold mb-2">
            <span>TOTAL FACTURADO CON IVA</span>
            <div className="w-8 h-8 rounded-lg bg-amber-500/10 text-amber-400 flex items-center justify-center">
              <TrendingUp className="w-4 h-4" />
            </div>
          </div>
          <div className="text-3xl font-black text-amber-400 tabular-nums tracking-tight">
            ${viewData.grossSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
          </div>
          <div className="mt-2 text-xs text-slate-400">
            IVA 15%: ${viewData.ivaAmount.toFixed(2)} USD
          </div>
          <div className="absolute bottom-0 left-0 right-0 h-1 bg-amber-500" />
        </div>
      </div>

      {/* Main Charts & Analytics Section */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left 2 Cols: Monthly Historical Comparison */}
        <div className="lg:col-span-2 p-6 rounded-2xl bg-[#0e1626] border border-slate-800/90 shadow-xl flex flex-col justify-between">
          <div className="flex items-center justify-between mb-6">
            <div>
              <span className="text-xs font-bold text-sky-400 uppercase tracking-wider">
                HISTÓRICO 2026
              </span>
              <h2 className="text-lg font-bold text-white mt-0.5">
                Evolución Comparativa Mensual de Facturación ($)
              </h2>
            </div>

            <div className="flex items-center gap-4 text-xs font-semibold">
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-3 rounded bg-blue-600" />
                <span className="text-slate-300">Upconnect</span>
              </div>
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-3 rounded bg-amber-500" />
                <span className="text-slate-300">Connect</span>
              </div>
            </div>
          </div>

          {/* Bar Chart Area */}
          {(() => {
            const activeMonthlyMetrics = (viewData.monthlyMetrics && viewData.monthlyMetrics.length > 0)
              ? viewData.monthlyMetrics
              : dataset.monthlyMetrics;

            if (activeMonthlyMetrics.length === 0) {
              return (
                <div className="h-64 flex flex-col items-center justify-center border border-dashed border-slate-800 rounded-xl p-6 text-center text-slate-400">
                  <BarChart3 className="w-10 h-10 text-slate-600 mb-2" />
                  <p className="font-semibold text-slate-300 text-sm">Sin datos mensuales para graficar</p>
                  <p className="text-xs text-slate-500 mt-1 max-w-sm">
                    Cargue el archivo de firmas o socios usando los botones de arriba para visualizar la evolución histórica mensual.
                  </p>
                </div>
              );
            }

            const maxChannelSale = Math.max(
              1,
              ...activeMonthlyMetrics.map((m) => Math.max(m.upconnectSales, m.connectorsSales, 1))
            );

            return (
              <div className="h-64 flex items-end justify-around gap-2 sm:gap-4 border-b border-slate-800/80 px-2 pb-2 pt-6">
                {activeMonthlyMetrics.map((item) => {
                  const upPct = Math.min(100, Math.max(item.upconnectSales > 0 ? 8 : 0, (item.upconnectSales / maxChannelSale) * 100));
                  const coPct = Math.min(100, Math.max(item.connectorsSales > 0 ? 8 : 0, (item.connectorsSales / maxChannelSale) * 100));

                  return (
                    <div key={item.month} className="flex-1 flex flex-col items-center justify-end h-full max-w-[85px] group">
                      <div className="w-full flex items-end justify-center gap-1.5 h-44 relative">
                        {/* UP Bar */}
                        <div className="h-full flex flex-col items-center justify-end flex-1 max-w-[36px]">
                          {item.upconnectSales > 0 && (
                            <span className="text-[10px] font-bold text-sky-400 tabular-nums mb-1 leading-none text-center">
                              ${Math.round(item.upconnectSales).toLocaleString()}
                            </span>
                          )}
                          <div
                            style={{ height: `${upPct}%` }}
                            className="w-full min-h-[4px] bg-gradient-to-t from-blue-700 to-blue-500 hover:from-blue-600 hover:to-blue-400 rounded-t-md transition-all shadow-md shadow-blue-900/40"
                            title={`UpConnect ${item.month}: $${item.upconnectSales.toFixed(2)} USD (${item.upconnectCount} firmas)`}
                          />
                        </div>

                        {/* CO Bar */}
                        <div className="h-full flex flex-col items-center justify-end flex-1 max-w-[36px]">
                          {item.connectorsSales > 0 && (
                            <span className="text-[10px] font-bold text-amber-400 tabular-nums mb-1 leading-none text-center">
                              ${Math.round(item.connectorsSales).toLocaleString()}
                            </span>
                          )}
                          <div
                            style={{ height: `${coPct}%` }}
                            className="w-full min-h-[4px] bg-gradient-to-t from-amber-600 to-amber-400 hover:from-amber-500 hover:to-amber-300 rounded-t-md transition-all shadow-md shadow-amber-900/40"
                            title={`Connectors ${item.month}: $${item.connectorsSales.toFixed(2)} USD (${item.connectorsCount} firmas)`}
                          />
                        </div>
                      </div>

                      <div className="mt-2 text-center">
                        <span className="text-xs font-semibold text-slate-300 capitalize">
                          {item.month.substring(0, 3).toLowerCase()}
                        </span>
                        <div className="text-[9px] text-slate-500 tabular-nums">
                          {item.upconnectCount + item.connectorsCount} f.
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            );
          })()}

          {/* Table under chart for exact figures */}
          {(() => {
            const activeMonthlyMetrics = (viewData.monthlyMetrics && viewData.monthlyMetrics.length > 0)
              ? viewData.monthlyMetrics
              : dataset.monthlyMetrics;

            if (activeMonthlyMetrics.length === 0) return null;

            return (
              <div className="mt-4 pt-3 border-t border-slate-800/60 overflow-x-auto">
                <table className="w-full text-xs text-left">
                  <thead>
                    <tr className="text-slate-400 border-b border-slate-800">
                      <th className="pb-2 font-semibold">Mes</th>
                      <th className="pb-2 font-semibold text-right">Upconnect ($)</th>
                      <th className="pb-2 font-semibold text-right">Distribuidor Connect ($)</th>
                      <th className="pb-2 font-semibold text-right">Total ($)</th>
                      <th className="pb-2 font-semibold text-right">Firmas</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/40 text-slate-300 tabular-nums">
                    {activeMonthlyMetrics.map((m) => (
                      <tr key={m.month} className="hover:bg-slate-800/30">
                        <td className="py-1.5 font-medium">{m.month}</td>
                        <td className="py-1.5 text-right text-sky-400">${m.upconnectSales.toFixed(2)}</td>
                        <td className="py-1.5 text-right text-amber-400">${m.connectorsSales.toFixed(2)}</td>
                        <td className="py-1.5 text-right font-bold text-white">
                          ${(m.upconnectSales + m.connectorsSales).toFixed(2)}
                        </td>
                        <td className="py-1.5 text-right font-semibold text-slate-400">
                          {m.upconnectCount + m.connectorsCount}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            );
          })()}
        </div>

        {/* Right 1 Col: Top Socios & Pareto Ranking (Fed by Carga Socios or Calculated) */}
        <div className="p-6 rounded-2xl bg-[#0e1626] border border-slate-800/90 shadow-xl flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between mb-4">
              <div>
                <span className="text-xs font-bold text-emerald-400 uppercase tracking-wider">
                  PARETO 80/20
                </span>
                <h3 className="text-lg font-bold text-white mt-0.5">Top Socios Productores</h3>
              </div>
              <span className="text-xs text-slate-400 font-semibold">{viewData.socios.length} Socios</span>
            </div>

            <div className="space-y-2.5 max-h-[360px] overflow-y-auto pr-1">
              {viewData.socios.length === 0 ? (
                <div className="py-8 text-center text-slate-500 text-xs border border-dashed border-slate-800 rounded-xl px-4">
                  <p className="font-semibold text-slate-400">Sin socios registrados</p>
                  <p className="text-[11px] text-slate-500 mt-1">
                    Cargue el archivo de socios (7 columnas) para visualizar el ranking y sumarlo al Reporte 2.
                  </p>
                </div>
              ) : (
                viewData.socios.slice(0, 8).map((s) => {
                  const maxVal = viewData.socios[0]?.totalSales || 1;
                  const pct = Math.max(8, (s.totalSales / maxVal) * 100);
                  return (
                    <div key={s.rank} className="p-2.5 rounded-xl bg-slate-900/70 border border-slate-800/80">
                      <div className="flex items-center justify-between text-xs mb-1.5">
                        <div className="flex items-center gap-2">
                          <span className="w-5 font-bold text-slate-500 tabular-nums">{s.rank}.</span>
                          <span className="font-semibold text-white">{s.name}</span>
                          {s.note && (
                            <span className="px-1.5 py-0.2 rounded text-[10px] bg-sky-500/20 text-sky-300 font-semibold">
                              {s.note}
                            </span>
                          )}
                        </div>
                        <span className="font-bold text-emerald-400 tabular-nums">
                          ${s.totalSales.toFixed(2)}
                        </span>
                      </div>

                      <div className="w-full bg-slate-800 rounded-full h-1.5 overflow-hidden">
                        <div
                          style={{ width: `${pct}%` }}
                          className={`h-full ${s.rank === 1 ? 'bg-emerald-400' : 'bg-blue-500'} rounded-full`}
                        />
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          </div>

          <div className="mt-4 pt-3 border-t border-slate-800/80 flex items-center justify-between">
            <span className="text-xs text-slate-400">
              Total Cartera: <strong className="text-white">${viewData.socios.reduce((a, s) => a + s.totalSales, 0).toLocaleString('en-US', { minimumFractionDigits: 2 })}</strong>
            </span>
            <button
              onClick={onOpenUploadSocios}
              className="text-xs text-emerald-400 hover:underline font-semibold"
            >
              + Actualizar Socios
            </button>
          </div>
        </div>
      </div>

      {/* Solutions Portfolio Grid */}
      {dataset.solutionCategories && dataset.solutionCategories.length > 0 && (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          {dataset.solutionCategories.map((cat) => (
            <div
              key={cat.id}
              className="p-6 rounded-2xl bg-[#0e1626] border border-slate-800/90 shadow-xl flex flex-col justify-between"
            >
              <div>
                <span className="text-[11px] font-bold text-sky-400 uppercase tracking-wider">
                  {cat.badge}
                </span>
                <h4 className="text-lg font-bold text-white mt-1 mb-2">{cat.title}</h4>
                <div className="text-2xl font-black text-white tabular-nums">
                  ${cat.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                </div>
                <div className="text-xs font-semibold text-emerald-400 mt-1">
                  {cat.percentage}% de participación
                </div>
              </div>

              <p className="text-xs text-slate-400 border-t border-slate-800/80 pt-3 mt-4">
                {cat.plansDescription}
              </p>
            </div>
          ))}
        </div>
      )}

      {/* TABLITA ADICIONAL: Cantidad de Socios UpConnect vs Connectors */}
      <SociosChannelTable
        summary={viewData.sociosSummary}
        monthlyMetrics={viewData.monthlyMetrics}
        transactions={viewData.transactions.length > 0 ? viewData.transactions : dataset.transactions}
        variant="full"
        title="Distribución y Cantidad de Socios por Canal"
        subtitle={`Resumen ejecutivo de cartera en ${viewData.periodLabel}: Conteo exacto de socios, operaciones y facturación`}
        onFilterChannel={(ch) => {
          setActiveTableTab('socios');
          setSocioChannelFilter(ch);
          setSocioPage(1);
        }}
        activeChannelFilter={socioChannelFilter}
      />

      {/* DUAL-TABBED TABLE SECTION: Tabla Detallada de Socios & Registro de Firmas */}
      <div className="p-6 rounded-2xl bg-[#0e1626] border border-slate-800/90 shadow-xl space-y-5">
        {/* Tab switcher header */}
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 border-b border-slate-800 pb-4">
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => setActiveTableTab('socios')}
              className={`flex items-center gap-2 px-4 py-2 text-xs font-bold rounded-xl transition-all ${
                activeTableTab === 'socios'
                  ? 'bg-emerald-600 text-white shadow-lg shadow-emerald-600/25 scale-[1.02]'
                  : 'bg-slate-900 text-slate-400 hover:text-white hover:bg-slate-800 border border-slate-800'
              }`}
            >
              <Users2 className="w-4 h-4" />
              <span>Tabla Detallada de Socios</span>
              <span className={`px-1.5 py-0.2 rounded-full text-[10px] font-bold ${
                activeTableTab === 'socios' ? 'bg-white/20 text-white' : 'bg-slate-800 text-slate-400'
              }`}>
                {viewData.socios.length}
              </span>
            </button>

            <button
              onClick={() => setActiveTableTab('transactions')}
              className={`flex items-center gap-2 px-4 py-2 text-xs font-bold rounded-xl transition-all ${
                activeTableTab === 'transactions'
                  ? 'bg-blue-600 text-white shadow-lg shadow-blue-600/25 scale-[1.02]'
                  : 'bg-slate-900 text-slate-400 hover:text-white hover:bg-slate-800 border border-slate-800'
              }`}
            >
              <FileSignature className="w-4 h-4" />
              <span>Registro de Firmas & Transacciones</span>
              <span className={`px-1.5 py-0.2 rounded-full text-[10px] font-bold ${
                activeTableTab === 'transactions' ? 'bg-white/20 text-white' : 'bg-slate-800 text-slate-400'
              }`}>
                {viewData.transactions.length}
              </span>
            </button>
          </div>

          {/* Search and action controls based on active tab */}
          {activeTableTab === 'socios' ? (
            <div className="flex flex-wrap items-center gap-3">
              <div className="relative min-w-[220px]">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  placeholder="Buscar socio, rol o plan..."
                  value={socioSearchTerm}
                  onChange={(e) => {
                    setSocioSearchTerm(e.target.value);
                    setSocioPage(1);
                  }}
                  className="w-full pl-9 pr-3 py-1.5 text-xs bg-slate-900 border border-slate-700 rounded-lg text-white placeholder-slate-500 focus:outline-none focus:border-emerald-500"
                />
              </div>

              <button
                onClick={onOpenUploadSocios}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 border border-emerald-500/30 transition-colors"
                title="Cargar archivo Excel con formato de 7 columnas para sumar al Reporte 2"
              >
                <Users2 className="w-3.5 h-3.5" />
                <span>Cargar Socios (7 Col.)</span>
              </button>
            </div>
          ) : (
            <div className="flex flex-wrap items-center gap-3">
              <div className="relative min-w-[220px]">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  placeholder="Buscar cliente, ID, socio..."
                  value={searchTerm}
                  onChange={(e) => {
                    setSearchTerm(e.target.value);
                    setPage(1);
                  }}
                  className="w-full pl-9 pr-3 py-1.5 text-xs bg-slate-900 border border-slate-700 rounded-lg text-white placeholder-slate-500 focus:outline-none focus:border-sky-500"
                />
              </div>

              {/* Channel filter */}
              <div className="flex items-center p-1 bg-slate-900 border border-slate-700 rounded-lg">
                <button
                  onClick={() => {
                    setSelectedChannel('ALL');
                    setPage(1);
                  }}
                  className={`px-2.5 py-1 text-xs font-semibold rounded-md transition-colors ${
                    selectedChannel === 'ALL' ? 'bg-blue-600 text-white' : 'text-slate-400 hover:text-white'
                  }`}
                >
                  Todos
                </button>
                <button
                  onClick={() => {
                    setSelectedChannel('UpConnect');
                    setPage(1);
                  }}
                  className={`px-2.5 py-1 text-xs font-semibold rounded-md transition-colors ${
                    selectedChannel === 'UpConnect' ? 'bg-blue-600 text-white' : 'text-slate-400 hover:text-white'
                  }`}
                >
                  UpConnect
                </button>
                <button
                  onClick={() => {
                    setSelectedChannel('Connectors');
                    setPage(1);
                  }}
                  className={`px-2.5 py-1 text-xs font-semibold rounded-md transition-colors ${
                    selectedChannel === 'Connectors' ? 'bg-amber-600 text-white' : 'text-slate-400 hover:text-white'
                  }`}
                >
                  Connect
                </button>
              </div>
            </div>
          )}
        </div>

        {/* TAB 1 CONTENT: TABLA DETALLADA DE SOCIOS */}
        {activeTableTab === 'socios' && (
          <div className="space-y-4">
            {/* Period Indicator & Month Selector Bar for Socios */}
            <div className="flex flex-wrap items-center justify-between gap-3 p-3 bg-slate-900/90 rounded-xl border border-slate-800 text-xs">
              <div className="flex items-center gap-2">
                <span className="text-slate-400">Periodo actual:</span>
                <span className="px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 font-bold border border-emerald-500/30">
                  {viewData.periodLabel}
                </span>
                <span className="text-slate-500">|</span>
                <span className="text-slate-300">
                  Total facturado: <strong className="text-emerald-400 font-bold">${viewData.socios.reduce((a, s) => a + s.totalSales, 0).toLocaleString('en-US', { minimumFractionDigits: 2 })} USD</strong>
                </span>
                <span className="text-slate-500">·</span>
                <span className="text-slate-300">
                  <strong className="text-white font-bold">{viewData.socios.reduce((a, s) => a + (s.operationsCount || 0), 0)}</strong> ventas conciliadas
                </span>
              </div>

              {/* Quick month switch pills */}
              <div className="flex items-center gap-1">
                <span className="text-[11px] text-slate-400 mr-1">Filtrar Mes:</span>
                <button
                  onClick={() => applyTimeFilter('all')}
                  className={`px-2 py-1 text-[11px] font-semibold rounded-md transition-colors ${
                    filterMode === 'all' ? 'bg-sky-600 text-white' : 'bg-slate-800 text-slate-400 hover:text-white'
                  }`}
                >
                  Todos
                </button>
                {availableMonths.map((m) => (
                  <button
                    key={m}
                    onClick={() => {
                      setSelectedMonthName(m);
                      applyTimeFilter('month');
                      onUpdateFilter({ type: 'month', month: m });
                    }}
                    className={`px-2 py-1 text-[11px] font-semibold rounded-md transition-colors ${
                      filterMode === 'month' && selectedMonthName === m
                        ? 'bg-emerald-600 text-white'
                        : 'bg-slate-800 text-slate-400 hover:text-white'
                    }`}
                  >
                    {m.charAt(0) + m.slice(1).toLowerCase()}
                  </button>
                ))}
              </div>
            </div>

            {/* TABLITA ADICIONAL: Cantidad de Socios UpConnect vs Connectors */}
            <SociosChannelTable
              summary={viewData.sociosSummary}
              variant="full"
              onFilterChannel={(ch) => {
                setSocioChannelFilter(ch);
                setSocioPage(1);
              }}
              activeChannelFilter={socioChannelFilter}
            />

            <div className="overflow-x-auto border border-slate-800 rounded-xl">
              <table className="w-full text-xs text-left">
                <thead className="bg-slate-900/95 text-slate-300 uppercase font-bold border-b border-slate-800 text-[11px]">
                  <tr>
                    <th className="py-3 px-3 w-12 text-center">#</th>
                    <th className="py-3 px-3">Nombre del Socio</th>
                    <th className="py-3 px-3">Canal / Rol</th>
                    <th className="py-3 px-3 text-center">Cant. Ventas Realizadas</th>
                    <th className="py-3 px-3 text-right">Monto Total de Venta</th>
                    <th className="py-3 px-3 text-right">Ticket Promedio</th>
                    <th className="py-3 px-3">Plan Más Vendido</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60 text-slate-300">
                  {paginatedSocios.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="py-8 text-center text-slate-500">
                        {viewData.socios.length === 0 
                          ? 'No hay socios registrados para este periodo. Use el botón "Carga Archivo Socios" para importar el archivo de 7 columnas.'
                          : socioChannelFilter !== 'ALL'
                          ? `No se encontraron socios para el canal ${socioChannelFilter} con los filtros activos.`
                          : 'No se encontraron socios que coincidan con la búsqueda.'}
                      </td>
                    </tr>
                  ) : (
                    paginatedSocios.map((s) => {
                      const avg = s.averageTicket ?? (s.operationsCount ? (s.totalSales / s.operationsCount) : s.totalSales);
                      const isUp = s.channel === 'UpConnect' || (!s.channel && !(s.role || '').toLowerCase().includes('connect'));
                      return (
                        <tr key={s.rank} className="hover:bg-slate-850/40 transition-colors">
                          <td className="py-2.5 px-3 text-center">
                            <span className={`inline-flex items-center justify-center w-6 h-6 rounded-md font-bold text-xs ${
                              s.rank === 1
                                ? 'bg-amber-500/20 text-amber-300 border border-amber-500/40'
                                : s.rank === 2
                                ? 'bg-slate-300/20 text-slate-200 border border-slate-300/40'
                                : s.rank === 3
                                ? 'bg-amber-700/20 text-amber-400 border border-amber-700/40'
                                : 'text-slate-500 font-mono'
                            }`}>
                              {s.rank}
                            </span>
                          </td>
                          <td className="py-2.5 px-3">
                            <div className="font-bold text-white flex items-center gap-1.5">
                              <span>{s.name}</span>
                              {s.rank === 1 && (
                                <span className="text-[10px] px-1.5 py-0.2 rounded bg-amber-500/20 text-amber-300 border border-amber-500/30 font-semibold">
                                  Top 1
                                </span>
                              )}
                            </div>
                          </td>
                          <td className="py-2.5 px-3">
                            <div className="flex flex-col gap-1 items-start">
                              <span className={`px-2 py-0.5 rounded text-[10px] font-bold border flex items-center gap-1 ${
                                isUp
                                  ? 'bg-sky-500/10 text-sky-300 border-sky-500/30'
                                  : 'bg-amber-500/10 text-amber-300 border-amber-500/30'
                              }`}>
                                <span className={`w-1.5 h-1.5 rounded-full ${isUp ? 'bg-sky-400' : 'bg-amber-400'}`} />
                                <span>{isUp ? 'Distribuidor Upconnect' : 'Distribuidor Connect'}</span>
                                {onToggleSocioChannel && (
                                  <button
                                    onClick={() => onToggleSocioChannel(s.name)}
                                    className="ml-1 text-[9px] text-slate-400 hover:text-white underline cursor-pointer"
                                    title="Alternar canal (Distribuidor Upconnect / Distribuidor Connect)"
                                  >
                                    Cambiar
                                  </button>
                                )}
                              </span>
                              <span className="text-[10px] text-slate-400">
                                {s.role || (isUp ? 'UpConnect Directo' : 'Distribuidor Connect')}
                              </span>
                            </div>
                          </td>
                          <td className="py-2.5 px-3 text-center font-bold text-slate-200 tabular-nums">
                            {s.operationsCount || 1}
                          </td>
                          <td className="py-2.5 px-3 text-right font-black text-emerald-400 tabular-nums">
                            ${s.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                          </td>
                          <td className="py-2.5 px-3 text-right font-bold text-sky-400 tabular-nums">
                            ${avg.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                          </td>
                          <td className="py-2.5 px-3">
                            <span className="px-2 py-0.5 rounded text-[11px] bg-slate-800 text-slate-300 border border-slate-700 font-medium inline-block max-w-[220px] truncate" title={s.topPlan}>
                              {s.topPlan || 'Firma Electrónica (1 año)'}
                            </span>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
                {viewData.socios.length > 0 && (
                  <tfoot className="bg-[#0b1220] border-t-2 border-slate-700 text-slate-200 font-bold text-xs">
                    <tr>
                      <td colSpan={3} className="py-2.5 px-3 text-left">
                        TOTAL CONSOLIDADO ({viewData.socios.length} SOCIOS)
                      </td>
                      <td className="py-2.5 px-3 text-center text-white tabular-nums">
                        {viewData.socios.reduce((a, s) => a + (s.operationsCount || 0), 0)} ventas
                      </td>
                      <td className="py-2.5 px-3 text-right text-emerald-400 font-black tabular-nums">
                        ${viewData.socios.reduce((a, s) => a + s.totalSales, 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                      </td>
                      <td className="py-2.5 px-3 text-right text-sky-400 font-black tabular-nums">
                        ${(
                          viewData.socios.reduce((a, s) => a + s.totalSales, 0) / 
                          (viewData.socios.reduce((a, s) => a + (s.operationsCount || 0), 0) || 1)
                        ).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                      </td>
                      <td className="py-2.5 px-3 text-slate-400 font-normal text-[11px]">
                        Ticket Promedio General
                      </td>
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>

            {/* Pagination for socios */}
            <div className="flex items-center justify-between pt-1 text-xs text-slate-400">
              <span>
                Mostrando <strong>{filteredSociosList.length > 0 ? (socioPage - 1) * socioPageSize + 1 : 0}</strong> a{' '}
                <strong>{Math.min(socioPage * socioPageSize, filteredSociosList.length)}</strong> de{' '}
                <strong>{filteredSociosList.length}</strong> socios
              </span>

              <div className="flex items-center gap-1">
                <button
                  onClick={() => setSocioPage((p) => Math.max(p - 1, 1))}
                  disabled={socioPage === 1}
                  className="px-3 py-1 rounded bg-slate-800 text-slate-300 hover:bg-slate-700 disabled:opacity-30 disabled:hover:bg-slate-800"
                >
                  Anterior
                </button>
                <span className="px-2 font-semibold text-white">
                  Página {socioPage} de {totalSocioPages}
                </span>
                <button
                  onClick={() => setSocioPage((p) => Math.min(p + 1, totalSocioPages))}
                  disabled={socioPage === totalSocioPages}
                  className="px-3 py-1 rounded bg-slate-800 text-slate-300 hover:bg-slate-700 disabled:opacity-30 disabled:hover:bg-slate-800"
                >
                  Siguiente
                </button>
              </div>
            </div>
          </div>
        )}

        {/* TAB 2 CONTENT: REGISTRO DE FIRMAS & TRANSACCIONES */}
        {activeTableTab === 'transactions' && (
          <div className="space-y-4">
            <div className="overflow-x-auto border border-slate-800 rounded-xl">
              <table className="w-full text-xs text-left">
                <thead className="bg-slate-900/90 text-slate-400 uppercase font-semibold border-b border-slate-800">
                  <tr>
                    <th className="py-3 px-3 w-12 text-center">N°</th>
                    <th className="py-3 px-3">ID Único</th>
                    <th 
                      className="py-3 px-3 cursor-pointer hover:text-white"
                      onClick={() => handleToggleSort('date')}
                    >
                      <div className="flex items-center gap-1">
                        <span>Fecha Solic</span>
                        <ArrowUpDown className="w-3 h-3" />
                      </div>
                    </th>
                    <th 
                      className="py-3 px-3 cursor-pointer hover:text-white"
                      onClick={() => handleToggleSort('clientName')}
                    >
                      <div className="flex items-center gap-1">
                        <span>Nombres y Apellidos</span>
                        <ArrowUpDown className="w-3 h-3" />
                      </div>
                    </th>
                    <th className="py-3 px-3">Estado</th>
                    <th className="py-3 px-3">Duración</th>
                    <th 
                      className="py-3 px-3 text-right cursor-pointer hover:text-white"
                      onClick={() => handleToggleSort('value')}
                    >
                      <div className="flex items-center justify-end gap-1">
                        <span>Valor ($)</span>
                        <ArrowUpDown className="w-3 h-3" />
                      </div>
                    </th>
                    <th className="py-3 px-3">Rol / Canal</th>
                    <th className="py-3 px-3">Socio Asignado</th>
                    <th className="py-3 px-3">Solución</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60 text-slate-300">
                  {paginatedTransactions.length === 0 ? (
                    <tr>
                      <td colSpan={10} className="py-8 text-center text-slate-500">
                        No se encontraron transacciones para el filtro de tiempo y búsqueda actual.
                      </td>
                    </tr>
                  ) : (
                    paginatedTransactions.map((trx, index) => (
                      <tr key={trx.id} className="hover:bg-slate-850/40 transition-colors">
                        <td className="py-2.5 px-3 text-center text-slate-500 font-mono">
                          {(page - 1) * pageSize + index + 1}
                        </td>
                        <td className="py-2.5 px-3 font-mono text-sky-400 font-semibold">
                          {trx.uniqueId}
                        </td>
                        <td className="py-2.5 px-3 font-medium text-slate-300">
                          {trx.date}
                        </td>
                        <td className="py-2.5 px-3 font-semibold text-white">
                          {trx.clientName} {trx.clientLastName}
                        </td>
                        <td className="py-2.5 px-3">
                          <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                            {trx.certificateStatus}
                          </span>
                        </td>
                        <td className="py-2.5 px-3 text-slate-300 font-medium">
                          {trx.duration}
                        </td>
                        <td className="py-2.5 px-3 text-right font-black text-white tabular-nums">
                          ${trx.value.toFixed(2)}
                        </td>
                        <td className="py-2.5 px-3">
                          <span className={`px-2 py-0.5 rounded text-[10px] font-semibold ${
                            trx.channel === 'Connectors' 
                              ? 'bg-amber-500/10 text-amber-300 border border-amber-500/20'
                              : 'bg-blue-500/10 text-sky-300 border border-blue-500/20'
                          }`}>
                            {trx.channel}
                          </span>
                        </td>
                        <td className="py-2.5 px-3 font-medium text-slate-300">
                          {trx.socio || 'Directo'}
                        </td>
                        <td className="py-2.5 px-3 text-slate-400 text-[11px]">
                          {trx.solutionCategory || 'ERP'}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination controls for transactions */}
            <div className="flex items-center justify-between pt-1 text-xs text-slate-400">
              <span>
                Mostrando <strong>{filteredTransactions.length > 0 ? (page - 1) * pageSize + 1 : 0}</strong> a{' '}
                <strong>{Math.min(page * pageSize, filteredTransactions.length)}</strong> de{' '}
                <strong>{filteredTransactions.length}</strong> registros
              </span>

              <div className="flex items-center gap-1">
                <button
                  onClick={() => setPage((p) => Math.max(p - 1, 1))}
                  disabled={page === 1}
                  className="px-3 py-1 rounded bg-slate-800 text-slate-300 hover:bg-slate-700 disabled:opacity-30 disabled:hover:bg-slate-800"
                >
                  Anterior
                </button>
                <span className="px-2 font-semibold text-white">
                  Página {page} de {totalPages}
                </span>
                <button
                  onClick={() => setPage((p) => Math.min(p + 1, totalPages))}
                  disabled={page === totalPages}
                  className="px-3 py-1 rounded bg-slate-800 text-slate-300 hover:bg-slate-700 disabled:opacity-30 disabled:hover:bg-slate-800"
                >
                  Siguiente
                </button>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Confirmation Modal to Limpiar Data */}
      {showClearConfirm && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="w-full max-w-md bg-[#0e1626] border border-rose-500/40 rounded-2xl shadow-2xl p-6 space-y-4">
            <div className="flex items-center gap-3">
              <div className="w-12 h-12 rounded-xl bg-rose-500/20 text-rose-400 flex items-center justify-center shrink-0">
                <Trash2 className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-base font-bold text-white">¿Limpiar toda la información?</h3>
                <p className="text-xs text-slate-400">Esta acción vaciará las firmas, transacciones y socios cargados.</p>
              </div>
            </div>

            <p className="text-xs text-slate-300 bg-slate-900/80 p-3.5 rounded-xl border border-slate-800 leading-relaxed">
              Se restablecerá la base de datos a su estado en blanco para que puedas <strong>volver a cargar todo desde cero</strong> con tus archivos de Excel actualizados.
            </p>

            <div className="flex items-center justify-end gap-3 pt-2">
              <button
                type="button"
                onClick={() => setShowClearConfirm(false)}
                className="px-4 py-2 text-xs font-semibold rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors"
              >
                Cancelar
              </button>
              <button
                type="button"
                onClick={() => {
                  setShowClearConfirm(false);
                  onClearData();
                }}
                className="flex items-center gap-1.5 px-4 py-2 text-xs font-bold rounded-lg bg-rose-600 hover:bg-rose-500 text-white shadow-lg shadow-rose-600/30 transition-all hover:scale-[1.02]"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Sí, Limpiar Data</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
