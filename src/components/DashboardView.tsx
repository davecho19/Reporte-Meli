import React, { useState, useMemo, useEffect } from 'react';
import { GlobalDataset, TransactionRecord, ReportTimeFilter } from '../types';
import { 
  TrendingUp, 
  DollarSign, 
  FileCheck2, 
  Users2, 
  Calendar, 
  FileSignature,
  Sparkles,
  BarChart3,
  Trash2,
  MonitorPlay,
  UserPlus
} from 'lucide-react';
import { computeReportView, MONTH_NAMES_ES, extractTransactionMonth } from '../utils/reportFilters';
import { SociosChannelTable } from './SociosChannelTable';

interface Props {
  dataset: GlobalDataset;
  onOpenUploadFirma: () => void;
  onOpenUploadSocios: () => void;
  onOpenUploadVendedores?: () => void;
  onClearData: () => void;
  onUpdateFilter: (filter: ReportTimeFilter) => void;
  onAddTransaction?: (trx: TransactionRecord) => void;
  onToggleSocioChannel?: (socioName: string) => void;
  onOpenOnlineReport?: () => void;
}

export const DashboardView: React.FC<Props> = ({
  dataset,
  onOpenUploadFirma,
  onOpenUploadSocios,
  onOpenUploadVendedores,
  onClearData,
  onUpdateFilter,
  onOpenOnlineReport,
}) => {
  const now = useMemo(() => new Date(), []);
  const currentActualMonthName = useMemo(() => {
    return MONTH_NAMES_ES[now.getMonth()] || 'OCTUBRE';
  }, [now]);

  // Si no hay filtro o está en 'all', usar el mes vigente como filtro predeterminado
  const currentFilter = dataset.reportFilter && dataset.reportFilter.month
    ? dataset.reportFilter
    : { type: 'month' as const, month: currentActualMonthName };

  const viewData = computeReportView(dataset);

  const [showClearConfirm, setShowClearConfirm] = useState(false);

  // Local state for the filter panel - siempre inicializado en el mes vigente
  const [filterMode, setFilterMode] = useState<'all' | 'month' | 'week' | 'range'>(
    (dataset.reportFilter && dataset.reportFilter.type !== 'all') ? dataset.reportFilter.type : 'month'
  );
  const [selectedMonthName, setSelectedMonthName] = useState<string>(
    dataset.reportFilter?.month || currentActualMonthName
  );
  const [selectedWeekId, setSelectedWeekId] = useState<string>(dataset.reportFilter?.weekId || 'w1');
  const [startDate, setStartDate] = useState<string>(
    currentFilter.startDate || `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-01`
  );
  const [endDate, setEndDate] = useState<string>(
    currentFilter.endDate || `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-20`
  );

  // Asegurar que al inicio el periodo esté activado en el mes actual si no había uno fijado
  useEffect(() => {
    if (!dataset.reportFilter || dataset.reportFilter.type === 'all') {
      onUpdateFilter({ type: 'month', month: currentActualMonthName });
    }
  }, [currentActualMonthName]);

  // Semanas calculadas dinámicamente según el mes vigente / seleccionado
  const currentMonthWeeks = useMemo(() => {
    const targetMonth = selectedMonthName || currentActualMonthName;
    const targetMonthIdx = MONTH_NAMES_ES.indexOf(targetMonth) + 1 || (now.getMonth() + 1);
    const targetMonthLower = targetMonth.toLowerCase();
    const targetYear = now.getFullYear();
    const daysInMonth = new Date(targetYear, targetMonthIdx, 0).getDate();

    return [
      { id: 'w1', label: `Semana 1 (del 01 al 06 de ${targetMonthLower})` },
      { id: 'w2', label: `Semana 2 (del 07 al 13 de ${targetMonthLower})` },
      { id: 'w3', label: `Semana 3 (del 14 al 20 de ${targetMonthLower})` },
      { id: 'w4', label: `Semana 4 (del 21 al ${daysInMonth} de ${targetMonthLower})` },
    ];
  }, [selectedMonthName, currentActualMonthName, now]);

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

    if (mode === 'all') {
      onUpdateFilter({ type: 'all' });
    } else if (mode === 'month') {
      onUpdateFilter({ type: 'month', month: selectedMonthName || currentActualMonthName });
    } else if (mode === 'week') {
      onUpdateFilter({ type: 'week', weekId: selectedWeekId, month: selectedMonthName || currentActualMonthName });
    } else if (mode === 'range') {
      onUpdateFilter({ type: 'range', startDate, endDate });
    }
  };

  const avgTicket = viewData.totalCount > 0 ? (viewData.totalSales / viewData.totalCount) : 0;

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

          {onOpenUploadVendedores && (
            <button
              onClick={onOpenUploadVendedores}
              className="flex items-center gap-2 px-4 py-2.5 text-xs font-bold rounded-xl bg-purple-600 hover:bg-purple-500 text-white shadow-lg shadow-purple-600/25 transition-all"
              title="Cargar archivo Excel de vendedores (Número y Fecha de Creación)"
            >
              <UserPlus className="w-4 h-4" />
              <span>Cargar Vendedores</span>
            </button>
          )}

          <button
            onClick={() => setShowClearConfirm(true)}
            className="flex items-center gap-2 px-4 py-2.5 text-xs font-bold rounded-xl bg-rose-500/10 hover:bg-rose-500/20 text-rose-400 hover:text-rose-300 border border-rose-500/30 transition-all shadow-sm"
            title="Borrar toda la información para volver a cargar datos desde cero"
          >
            <Trash2 className="w-4 h-4" />
            <span>Limpiar Data</span>
          </button>

          {onOpenOnlineReport && (
            <button
              onClick={onOpenOnlineReport}
              className="flex items-center gap-2 px-3.5 py-2.5 text-xs font-bold rounded-xl bg-gradient-to-r from-blue-600 via-indigo-600 to-sky-600 hover:from-blue-500 hover:to-sky-500 text-white transition-all shadow-md shadow-indigo-950/40"
              title="Ver presentación interactiva en línea con ambos reportes unificados (15 diapositivas)"
            >
              <MonitorPlay className="w-4 h-4 text-sky-200" />
              <span>Reporte en Línea</span>
            </button>
          )}
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
                {MONTH_NAMES_ES.map((m) => (
                  <option key={m} value={m}>
                    {m} {m === currentActualMonthName ? '(Mes Actual)' : ''}
                  </option>
                ))}
              </select>
            </div>
          )}

          {filterMode === 'week' && (
            <div className="flex flex-wrap items-center gap-4">
              <div className="flex items-center gap-2">
                <span className="text-slate-400 font-semibold">Mes de la Semana:</span>
                <select
                  value={selectedMonthName || currentActualMonthName}
                  onChange={(e) => {
                    const newMonth = e.target.value;
                    setSelectedMonthName(newMonth);
                    onUpdateFilter({ type: 'week', weekId: selectedWeekId, month: newMonth });
                  }}
                  className="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-1.5 text-xs focus:outline-none focus:border-sky-500 font-semibold"
                >
                  {MONTH_NAMES_ES.map((m) => (
                    <option key={m} value={m}>
                      {m} {m === currentActualMonthName ? '(Mes Actual)' : ''}
                    </option>
                  ))}
                </select>
              </div>

              <div className="flex items-center gap-2">
                <span className="text-slate-400 font-semibold">Seleccionar Semana:</span>
                <select
                  value={selectedWeekId}
                  onChange={(e) => {
                    setSelectedWeekId(e.target.value);
                    onUpdateFilter({ type: 'week', weekId: e.target.value, month: selectedMonthName || currentActualMonthName });
                  }}
                  className="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-1.5 text-xs focus:outline-none focus:border-sky-500 font-semibold"
                >
                  {currentMonthWeeks.map((w) => (
                    <option key={w.id} value={w.id}>
                      {w.label}
                    </option>
                  ))}
                </select>
              </div>
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
                            title={`Connect ${item.month}: $${item.connectorsSales.toFixed(2)} USD (${item.connectorsCount} firmas)`}
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

        {/* Right 1 Col: Top Socios & Pareto Ranking */}
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
                    Cargue el archivo de socios para visualizar el ranking y sumarlo al Reporte 2.
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

      {/* TABLA DE CANALES Y GRÁFICO LINEAL MES A MES */}
      <SociosChannelTable
        summary={viewData.sociosSummary}
        monthlyMetrics={viewData.monthlyMetrics}
        transactions={viewData.transactions.length > 0 ? viewData.transactions : dataset.transactions}
        variant="full"
        title="Distribución y Cantidad de Socios por Canal"
        subtitle={`Resumen ejecutivo de cartera en ${viewData.periodLabel}: Conteo exacto de socios, operaciones y facturación`}
      />

      {/* GRÁFICO ESTADÍSTICO: SOCIOS CREADOS MES A MES (ARCHIVO DE VENDEDORES) */}
      {(() => {
        const monthlyCreations = dataset.vendorMonthlyCreations || [];
        const hasData = monthlyCreations.length > 0;
        const totalCreaciones = monthlyCreations.reduce((acc, m) => acc + m.count, 0);
        const maxCount = hasData ? Math.max(...monthlyCreations.map((m) => m.count), 1) : 1;

        return (
          <div className="p-6 sm:p-7 rounded-2xl bg-[#0c1322] border border-slate-800 shadow-xl space-y-6">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-800/80">
              <div>
                <span className="text-xs uppercase font-bold tracking-wider text-purple-400">
                  ESTADÍSTICA DE VENDEDORES
                </span>
                <h3 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-1">
                  Socios Creados Mes a Mes
                </h3>
                <p className="text-xs text-slate-400 mt-1">
                  Cantidad de socios creados en cada mes según el archivo de vendedores.
                </p>
              </div>

              {hasData && (
                <div className="flex items-center gap-3">
                  <div className="px-4 py-2 rounded-xl bg-purple-950/40 border border-purple-500/30 text-right">
                    <div className="text-2xl font-black text-purple-400 tabular-nums">
                      {totalCreaciones}
                    </div>
                    <div className="text-[11px] text-slate-400 font-medium">
                      Total Creados
                    </div>
                  </div>
                  {onOpenUploadVendedores && (
                    <button
                      onClick={onOpenUploadVendedores}
                      className="px-3.5 py-2 text-xs font-semibold rounded-xl bg-slate-800 hover:bg-slate-700 text-purple-300 border border-purple-500/30 transition-all flex items-center gap-1.5"
                      title="Cargar o actualizar archivo de vendedores"
                    >
                      <UserPlus className="w-3.5 h-3.5" />
                      <span>Actualizar</span>
                    </button>
                  )}
                </div>
              )}
            </div>

            {hasData ? (
              <div className="space-y-4">
                {/* Gráfico de Barras Estadístico */}
                <div className="h-64 pt-8 pb-4 flex items-end gap-3 sm:gap-6 overflow-x-auto border-b border-slate-800/80 px-2">
                  {monthlyCreations.map((item, idx) => {
                    const heightPercent = Math.max(10, Math.round((item.count / maxCount) * 100));
                    return (
                      <div
                        key={idx}
                        className="flex-1 min-w-[56px] max-w-[90px] h-full flex flex-col items-center justify-end group"
                      >
                        {/* Cantidad exacta de socios creados */}
                        <span className="text-xs sm:text-sm font-black text-purple-300 tabular-nums mb-1.5 transition-transform group-hover:scale-110">
                          {item.count}
                        </span>

                        {/* Barra */}
                        <div className="w-full bg-slate-900/80 rounded-t-lg p-0.5 h-full flex items-end">
                          <div
                            style={{ height: `${heightPercent}%` }}
                            className="w-full rounded-t-md bg-gradient-to-t from-purple-700 via-purple-600 to-indigo-500 hover:to-purple-400 shadow-md shadow-purple-950/50 transition-all duration-300"
                            title={`${item.month}: ${item.count} socios creados`}
                          />
                        </div>

                        {/* Mes */}
                        <div className="mt-2 text-center">
                          <span className="text-[11px] font-bold text-slate-300 uppercase tracking-tight block">
                            {item.month.substring(0, 3)}
                          </span>
                          {item.year && (
                            <span className="text-[9px] text-slate-500 block">
                              {item.year}
                            </span>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>

                <div className="flex flex-col sm:flex-row sm:items-center justify-between text-xs text-slate-400 gap-2 pt-1">
                  <span>Indicador basado exclusivamente en el archivo de vendedores (Número y Fecha de Creación).</span>
                  <span className="font-medium text-slate-300">
                    {dataset.vendorCreations?.length || totalCreaciones} registros procesados
                  </span>
                </div>
              </div>
            ) : (
              <div className="py-12 px-4 rounded-xl bg-slate-900/50 border border-slate-800/80 text-center space-y-3">
                <div className="w-12 h-12 rounded-xl bg-purple-500/10 text-purple-400 flex items-center justify-center mx-auto">
                  <UserPlus className="w-6 h-6" />
                </div>
                <div>
                  <h4 className="text-sm font-bold text-white">Sin datos de vendedores cargados</h4>
                  <p className="text-xs text-slate-400 max-w-md mx-auto mt-1">
                    Cargue el archivo de Excel con 2 columnas (Número y Fecha de Creación) para visualizar la cantidad de socios creados mes a mes.
                  </p>
                </div>
                {onOpenUploadVendedores && (
                  <button
                    onClick={onOpenUploadVendedores}
                    className="inline-flex items-center gap-2 px-4 py-2 text-xs font-bold rounded-xl bg-purple-600 hover:bg-purple-500 text-white shadow-lg shadow-purple-600/25 transition-all"
                  >
                    <UserPlus className="w-4 h-4" />
                    <span>Cargar Vendedores</span>
                  </button>
                )}
              </div>
            )}
          </div>
        );
      })()}

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

