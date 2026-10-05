import React, { useState, useMemo } from 'react';
import { SociosChannelSummary, extractTransactionDayAndMonth } from '../utils/reportFilters';
import { MonthlyMetric, TransactionRecord } from '../types';
import { Zap, Share2, Users, DollarSign, Award, TrendingUp, BarChart2, Calendar } from 'lucide-react';

interface SociosChannelTableProps {
  summary: SociosChannelSummary;
  monthlyMetrics?: MonthlyMetric[];
  transactions?: TransactionRecord[];
  variant?: 'full' | 'slide' | 'compact';
  title?: string;
  subtitle?: string;
  onFilterChannel?: (channel: 'ALL' | 'UpConnect' | 'Connectors') => void;
  activeChannelFilter?: 'ALL' | 'UpConnect' | 'Connectors';
}

export const SociosChannelTable: React.FC<SociosChannelTableProps> = ({
  summary,
  monthlyMetrics,
  transactions,
  variant = 'full',
  title = 'Distribución y Cantidad de Socios por Canal',
  subtitle = 'Conteo exacto y volumen de ventas: Distribuidor Upconnect vs Distribuidor Connect',
  onFilterChannel,
  activeChannelFilter = 'ALL',
}) => {
  const { upconnect, connectors, totalSociosCount, totalOperationsCount, totalSales, averageTicket } = summary;

  const upSalesPct = totalSales > 0 ? ((upconnect.totalSales / totalSales) * 100).toFixed(1) : '0.0';
  const coSalesPct = totalSales > 0 ? ((connectors.totalSales / totalSales) * 100).toFixed(1) : '0.0';

  const [lineChartMetric, setLineChartMetric] = useState<'socios' | 'sales'>('socios');
  const [hoveredIndex, setHoveredIndex] = useState<number | null>(null);

  // Month-over-month comparison dataset
  const chartData = useMemo(() => {
    // 1. If monthlyMetrics provided and populated
    if (monthlyMetrics && monthlyMetrics.length > 0) {
      return monthlyMetrics.map((m, idx, arr) => {
        const prev = idx > 0 ? arr[idx - 1] : null;
        const totalSales = parseFloat((m.upconnectSales + m.connectorsSales).toFixed(2));
        const totalCount = m.upconnectCount + m.connectorsCount;

        const upGrowthSales = prev && prev.upconnectSales > 0 ? ((m.upconnectSales - prev.upconnectSales) / prev.upconnectSales) * 100 : 0;
        const coGrowthSales = prev && prev.connectorsSales > 0 ? ((m.connectorsSales - prev.connectorsSales) / prev.connectorsSales) * 100 : 0;
        const upGrowthCount = prev && prev.upconnectCount > 0 ? ((m.upconnectCount - prev.upconnectCount) / prev.upconnectCount) * 100 : 0;
        const coGrowthCount = prev && prev.connectorsCount > 0 ? ((m.connectorsCount - prev.connectorsCount) / prev.connectorsCount) * 100 : 0;

        return {
          month: m.month,
          shortMonth: m.month.slice(0, 3).toUpperCase(),
          monthIndex: m.monthIndex,
          upSales: m.upconnectSales,
          coSales: m.connectorsSales,
          totalSales,
          upCount: m.upconnectCount,
          coCount: m.connectorsCount,
          totalCount,
          upGrowthSales,
          coGrowthSales,
          upGrowthCount,
          coGrowthCount,
        };
      });
    }

    // 2. If transactions provided, aggregate by month
    if (transactions && transactions.length > 0) {
      const monthMap: Record<number, {
        month: string;
        monthIndex: number;
        upSales: number;
        coSales: number;
        upCount: number;
        coCount: number;
        upSocios: Set<string>;
        coSocios: Set<string>;
      }> = {};

      transactions.forEach((t) => {
        const { monthIndex, monthName } = extractTransactionDayAndMonth(t);
        if (!monthMap[monthIndex]) {
          monthMap[monthIndex] = {
            month: monthName,
            monthIndex,
            upSales: 0,
            coSales: 0,
            upCount: 0,
            coCount: 0,
            upSocios: new Set(),
            coSocios: new Set(),
          };
        }
        const val = t.value || 0;
        const sName = t.socio || t.clientName || 'Socio';
        if (t.channel === 'Connectors') {
          monthMap[monthIndex].coSales += val;
          monthMap[monthIndex].coCount += 1;
          monthMap[monthIndex].coSocios.add(sName);
        } else {
          monthMap[monthIndex].upSales += val;
          monthMap[monthIndex].upCount += 1;
          monthMap[monthIndex].upSocios.add(sName);
        }
      });

      const sortedIndices = Object.keys(monthMap).map(Number).sort((a, b) => a - b);
      return sortedIndices.map((mIdx, idx, arr) => {
        const item = monthMap[mIdx];
        const prevIdx = idx > 0 ? arr[idx - 1] : null;
        const prev = prevIdx ? monthMap[prevIdx] : null;

        const upCount = item.upSocios.size > 0 ? item.upSocios.size : item.upCount;
        const coCount = item.coSocios.size > 0 ? item.coSocios.size : item.coCount;
        const prevUpCount = prev ? (prev.upSocios.size > 0 ? prev.upSocios.size : prev.upCount) : 0;
        const prevCoCount = prev ? (prev.coSocios.size > 0 ? prev.coSocios.size : prev.coCount) : 0;

        return {
          month: item.month,
          shortMonth: item.month.slice(0, 3).toUpperCase(),
          monthIndex: mIdx,
          upSales: parseFloat(item.upSales.toFixed(2)),
          coSales: parseFloat(item.coSales.toFixed(2)),
          totalSales: parseFloat((item.upSales + item.coSales).toFixed(2)),
          upCount,
          coCount,
          totalCount: upCount + coCount,
          upGrowthSales: prev && prev.upSales > 0 ? ((item.upSales - prev.upSales) / prev.upSales) * 100 : 0,
          coGrowthSales: prev && prev.coSales > 0 ? ((item.coSales - prev.coSales) / prev.coSales) * 100 : 0,
          upGrowthCount: prevUpCount > 0 ? ((upCount - prevUpCount) / prevUpCount) * 100 : 0,
          coGrowthCount: prevCoCount > 0 ? ((coCount - prevCoCount) / prevCoCount) * 100 : 0,
        };
      });
    }

    return [];
  }, [monthlyMetrics, transactions]);

  // Compact Slide Variant (Designed for 16:9 Presentation Slides)
  if (variant === 'slide') {
    return (
      <div className="w-full rounded-xl bg-slate-900/90 border border-slate-800 shadow-xl overflow-hidden text-xs">
        {/* Header strip */}
        <div className="px-4 py-2.5 bg-[#091122] border-b border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Users className="w-4 h-4 text-emerald-400" />
            <span className="font-bold text-white text-xs uppercase tracking-wide">
              {title}
            </span>
          </div>
          <div className="text-[11px] font-semibold text-slate-300">
            Total Cartera:{' '}
            <strong className="text-emerald-400 font-extrabold">{totalSociosCount} Socios</strong>{' '}
            ({totalOperationsCount} ventas / ${totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD)
          </div>
        </div>

        {/* Table */}
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-[#0c1424] text-slate-400 uppercase text-[10px] tracking-wider border-b border-slate-800">
              <tr>
                <th className="py-2 px-3 font-bold">Canal de Distribución</th>
                <th className="py-2 px-3 font-bold text-center">Cantidad de Socios</th>
                <th className="py-2 px-3 font-bold text-center">% de Socios</th>
                <th className="py-2 px-3 font-bold text-center">Ventas / Firmas</th>
                <th className="py-2 px-3 font-bold text-right">Facturación Neta ($)</th>
                <th className="py-2 px-3 font-bold text-right">% Facturación</th>
                <th className="py-2 px-3 font-bold text-right">Ticket Promedio</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-800/80 text-slate-200 tabular-nums">
              {/* Row 1: UpConnect */}
              <tr className="hover:bg-blue-950/20 transition-colors">
                <td className="py-2 px-3 font-bold text-white flex items-center gap-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-sky-400 shrink-0 shadow-sm shadow-sky-400/50" />
                  <span className="text-sky-300 font-extrabold">Distribuidor Upconnect</span>
                  <span className="text-[10px] text-slate-400 font-normal hidden sm:inline">(Canal Directo)</span>
                </td>
                <td className="py-2 px-3 text-center">
                  <span className="inline-flex items-center justify-center px-2 py-0.5 rounded-md bg-sky-500/20 text-sky-300 font-black text-xs border border-sky-500/30">
                    {upconnect.count} Socios
                  </span>
                </td>
                <td className="py-2 px-3 text-center font-semibold text-slate-300">
                  {upconnect.percentage}%
                </td>
                <td className="py-2 px-3 text-center font-bold text-slate-200">
                  {upconnect.operationsCount} ventas
                </td>
                <td className="py-2 px-3 text-right font-black text-sky-400">
                  ${upconnect.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD
                </td>
                <td className="py-2 px-3 text-right font-semibold text-slate-300">
                  {upSalesPct}%
                </td>
                <td className="py-2 px-3 text-right font-medium text-slate-300">
                  ${upconnect.averageTicket.toFixed(2)} USD
                </td>
              </tr>

              {/* Row 2: Connectors */}
              <tr className="hover:bg-amber-950/20 transition-colors">
                <td className="py-2 px-3 font-bold text-white flex items-center gap-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-amber-400 shrink-0 shadow-sm shadow-amber-400/50" />
                  <span className="text-amber-300 font-extrabold">Distribuidor Connect</span>
                  <span className="text-[10px] text-slate-400 font-normal hidden sm:inline">(Red Distribuidores)</span>
                </td>
                <td className="py-2 px-3 text-center">
                  <span className="inline-flex items-center justify-center px-2 py-0.5 rounded-md bg-amber-500/20 text-amber-300 font-black text-xs border border-amber-500/30">
                    {connectors.count} Socios
                  </span>
                </td>
                <td className="py-2 px-3 text-center font-semibold text-slate-300">
                  {connectors.percentage}%
                </td>
                <td className="py-2 px-3 text-center font-bold text-slate-200">
                  {connectors.operationsCount} ventas
                </td>
                <td className="py-2 px-3 text-right font-black text-amber-400">
                  ${connectors.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD
                </td>
                <td className="py-2 px-3 text-right font-semibold text-slate-300">
                  {coSalesPct}%
                </td>
                <td className="py-2 px-3 text-right font-medium text-slate-300">
                  ${connectors.averageTicket.toFixed(2)} USD
                </td>
              </tr>
            </tbody>

            {/* Total Row */}
            <tfoot className="bg-[#070d18] border-t-2 border-slate-700 font-bold text-xs text-white">
              <tr>
                <td className="py-2.5 px-3 uppercase text-emerald-400 font-black flex items-center gap-1.5">
                  <Award className="w-3.5 h-3.5 text-emerald-400" />
                  <span>TOTAL CONSOLIDADO</span>
                </td>
                <td className="py-2.5 px-3 text-center">
                  <span className="inline-flex items-center justify-center px-2.5 py-0.5 rounded-md bg-emerald-500/20 text-emerald-300 font-black text-xs border border-emerald-500/30">
                    {totalSociosCount} Socios
                  </span>
                </td>
                <td className="py-2.5 px-3 text-center text-slate-300 font-bold">100%</td>
                <td className="py-2.5 px-3 text-center text-white font-bold">{totalOperationsCount} ventas</td>
                <td className="py-2.5 px-3 text-right text-emerald-400 font-black">
                  ${totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD
                </td>
                <td className="py-2.5 px-3 text-right text-slate-300 font-bold">100%</td>
                <td className="py-2.5 px-3 text-right text-sky-400 font-bold">
                  ${averageTicket.toFixed(2)} USD
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    );
  }

  // Full / Dashboard Variant
  return (
    <div className="w-full rounded-2xl bg-[#0e1626] border border-slate-800/90 shadow-xl overflow-hidden">
      {/* Header */}
      <div className="p-5 border-b border-slate-800 bg-[#0a1120] flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <div className="flex items-center gap-2">
            <span className="w-2.5 h-2.5 rounded-full bg-emerald-400 shadow-sm shadow-emerald-400/50" />
            <h3 className="text-sm font-bold text-white uppercase tracking-wider">{title}</h3>
          </div>
          <p className="text-xs text-slate-400 mt-0.5">{subtitle}</p>
        </div>

        {onFilterChannel && (
          <div className="flex items-center p-1 bg-slate-900 border border-slate-700 rounded-xl text-xs font-semibold">
            <button
              onClick={() => onFilterChannel('ALL')}
              className={`px-3 py-1 rounded-lg transition-colors ${
                activeChannelFilter === 'ALL'
                  ? 'bg-blue-600 text-white shadow-sm'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              Todos ({totalSociosCount})
            </button>
            <button
              onClick={() => onFilterChannel('UpConnect')}
              className={`px-3 py-1 rounded-lg transition-colors flex items-center gap-1.5 ${
                activeChannelFilter === 'UpConnect'
                  ? 'bg-sky-600 text-white shadow-sm'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <Zap className="w-3 h-3 text-sky-300" />
              <span>Distribuidor Upconnect ({upconnect.count})</span>
            </button>
            <button
              onClick={() => onFilterChannel('Connectors')}
              className={`px-3 py-1 rounded-lg transition-colors flex items-center gap-1.5 ${
                activeChannelFilter === 'Connectors'
                  ? 'bg-amber-600 text-white shadow-sm'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <Share2 className="w-3 h-3 text-amber-300" />
              <span>Distribuidor Connect ({connectors.count})</span>
            </button>
          </div>
        )}
      </div>

      {/* Top 3 Metric Cards for instant visual recognition */}
      <div className="grid grid-cols-1 sm:grid-cols-3 divide-y sm:divide-y-0 sm:divide-x divide-slate-800 bg-slate-950/60 border-b border-slate-800">
        {/* UpConnect Card */}
        <div className="p-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-blue-600/20 text-sky-400 border border-blue-500/30 flex items-center justify-center shrink-0">
              <Zap className="w-5 h-5" />
            </div>
            <div>
              <div className="text-[11px] font-bold text-sky-400 uppercase">Socios Distribuidor Upconnect</div>
              <div className="text-2xl font-black text-white tabular-nums">
                {upconnect.count} <span className="text-xs font-normal text-slate-400">socios ({upconnect.percentage}%)</span>
              </div>
            </div>
          </div>
          <div className="text-right">
            <div className="text-xs font-bold text-sky-300 tabular-nums">
              ${upconnect.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}
            </div>
            <div className="text-[10px] text-slate-500">{upconnect.operationsCount} ventas</div>
          </div>
        </div>

        {/* Connectors Card */}
        <div className="p-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-amber-600/20 text-amber-400 border border-amber-500/30 flex items-center justify-center shrink-0">
              <Share2 className="w-5 h-5" />
            </div>
            <div>
              <div className="text-[11px] font-bold text-amber-400 uppercase">Socios Distribuidor Connect</div>
              <div className="text-2xl font-black text-white tabular-nums">
                {connectors.count} <span className="text-xs font-normal text-slate-400">socios ({connectors.percentage}%)</span>
              </div>
            </div>
          </div>
          <div className="text-right">
            <div className="text-xs font-bold text-amber-300 tabular-nums">
              ${connectors.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}
            </div>
            <div className="text-[10px] text-slate-500">{connectors.operationsCount} ventas</div>
          </div>
        </div>

        {/* Total General Card */}
        <div className="p-4 flex items-center justify-between bg-emerald-950/20">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-emerald-600/20 text-emerald-400 border border-emerald-500/30 flex items-center justify-center shrink-0">
              <Users className="w-5 h-5" />
            </div>
            <div>
              <div className="text-[11px] font-bold text-emerald-400 uppercase">Total Cartera Activa</div>
              <div className="text-2xl font-black text-emerald-400 tabular-nums">
                {totalSociosCount} <span className="text-xs font-normal text-slate-400">socios registrados</span>
              </div>
            </div>
          </div>
          <div className="text-right">
            <div className="text-xs font-bold text-emerald-300 tabular-nums">
              ${totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}
            </div>
            <div className="text-[10px] text-slate-500">{totalOperationsCount} ventas totales</div>
          </div>
        </div>
      </div>

      {/* Main Table */}
      <div className="overflow-x-auto">
        <table className="w-full text-xs text-left">
          <thead className="bg-[#0b1220] text-slate-300 uppercase font-bold border-b border-slate-800 text-[11px]">
            <tr>
              <th className="py-3 px-4">Canal Oficial</th>
              <th className="py-3 px-3 text-center">Cantidad de Socios</th>
              <th className="py-3 px-3 text-center">% de Socios</th>
              <th className="py-3 px-3 text-center">Ventas Realizadas</th>
              <th className="py-3 px-4 text-right">Facturación Neta ($ USD)</th>
              <th className="py-3 px-4 text-right">% Facturación</th>
              <th className="py-3 px-4 text-right">Ticket Promedio</th>
              <th className="py-3 px-4">Participación en Cartera</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-800/70 text-slate-300 tabular-nums">
            {/* UpConnect */}
            <tr className="hover:bg-slate-800/40 transition-colors">
              <td className="py-3 px-4">
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 rounded-full bg-sky-400 shrink-0" />
                  <div>
                    <div className="font-bold text-white flex items-center gap-1.5">
                      <span>Distribuidor Upconnect</span>
                      <span className="text-[9px] px-1.5 py-0.2 rounded bg-sky-500/20 text-sky-300 border border-sky-500/30 font-semibold">
                        Canal 01
                      </span>
                    </div>
                    <div className="text-[10px] text-slate-400">Fuerza comercial interna y ejecutivos directos</div>
                  </div>
                </div>
              </td>
              <td className="py-3 px-3 text-center">
                <span className="inline-flex items-center justify-center px-3 py-1 rounded-lg bg-sky-500/20 text-sky-300 font-black text-xs border border-sky-500/30">
                  {upconnect.count} Socios
                </span>
              </td>
              <td className="py-3 px-3 text-center font-bold text-slate-200">
                {upconnect.percentage}%
              </td>
              <td className="py-3 px-3 text-center font-bold text-white">
                {upconnect.operationsCount} ventas
              </td>
              <td className="py-3 px-4 text-right font-black text-sky-400 text-sm">
                ${upconnect.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD
              </td>
              <td className="py-3 px-4 text-right font-bold text-slate-200">
                {upSalesPct}%
              </td>
              <td className="py-3 px-4 text-right font-semibold text-slate-300">
                ${upconnect.averageTicket.toFixed(2)} USD
              </td>
              <td className="py-3 px-4">
                <div className="w-full bg-slate-800 rounded-full h-2 overflow-hidden max-w-[120px]">
                  <div
                    style={{ width: `${upconnect.percentage}%` }}
                    className="h-full bg-sky-400 rounded-full transition-all"
                  />
                </div>
              </td>
            </tr>

            {/* Distribuidor Connect */}
            <tr className="hover:bg-slate-800/40 transition-colors">
              <td className="py-3 px-4">
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 rounded-full bg-amber-400 shrink-0" />
                  <div>
                    <div className="font-bold text-white flex items-center gap-1.5">
                      <span>Distribuidor Connect</span>
                      <span className="text-[9px] px-1.5 py-0.2 rounded bg-amber-500/20 text-amber-300 border border-amber-500/30 font-semibold">
                        Canal 02
                      </span>
                    </div>
                    <div className="text-[10px] text-slate-400">Contadores franquiciados, aliados y distribuidores</div>
                  </div>
                </div>
              </td>
              <td className="py-3 px-3 text-center">
                <span className="inline-flex items-center justify-center px-3 py-1 rounded-lg bg-amber-500/20 text-amber-300 font-black text-xs border border-amber-500/30">
                  {connectors.count} Socios
                </span>
              </td>
              <td className="py-3 px-3 text-center font-bold text-slate-200">
                {connectors.percentage}%
              </td>
              <td className="py-3 px-3 text-center font-bold text-white">
                {connectors.operationsCount} ventas
              </td>
              <td className="py-3 px-4 text-right font-black text-amber-400 text-sm">
                ${connectors.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD
              </td>
              <td className="py-3 px-4 text-right font-bold text-slate-200">
                {coSalesPct}%
              </td>
              <td className="py-3 px-4 text-right font-semibold text-slate-300">
                ${connectors.averageTicket.toFixed(2)} USD
              </td>
              <td className="py-3 px-4">
                <div className="w-full bg-slate-800 rounded-full h-2 overflow-hidden max-w-[120px]">
                  <div
                    style={{ width: `${connectors.percentage}%` }}
                    className="h-full bg-amber-400 rounded-full transition-all"
                  />
                </div>
              </td>
            </tr>
          </tbody>

          {/* Footer Total */}
          <tfoot className="bg-[#070d18] border-t-2 border-slate-700 text-white font-bold text-xs">
            <tr>
              <td className="py-3 px-4 uppercase text-emerald-400 font-black flex items-center gap-2">
                <Award className="w-4 h-4 text-emerald-400" />
                <span>TOTAL GENERAL ECOSISTEMA</span>
              </td>
              <td className="py-3 px-3 text-center">
                <span className="inline-flex items-center justify-center px-3 py-1 rounded-lg bg-emerald-500/20 text-emerald-300 font-black text-xs border border-emerald-500/30">
                  {totalSociosCount} Socios Totales
                </span>
              </td>
              <td className="py-3 px-3 text-center font-black text-white">100%</td>
              <td className="py-3 px-3 text-center font-black text-white">
                {totalOperationsCount} ventas
              </td>
              <td className="py-3 px-4 text-right font-black text-emerald-400 text-sm">
                ${totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD
              </td>
              <td className="py-3 px-4 text-right font-black text-white">100%</td>
              <td className="py-3 px-4 text-right font-black text-sky-400">
                ${averageTicket.toFixed(2)} USD
              </td>
              <td className="py-3 px-4 text-[11px] text-slate-400 font-normal">
                Base consolidada 100%
              </td>
            </tr>
          </tfoot>
        </table>
      </div>

      {/* GRÁFICO LINEAL: COMPARATIVO MES A MES POR CANAL (UpConnect vs Connectors) */}
      <div className="border-t border-slate-800 bg-[#0a1222] p-5">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 mb-4">
          <div>
            <div className="flex items-center gap-2">
              <TrendingUp className="w-4 h-4 text-sky-400" />
              <h4 className="text-xs font-bold text-white uppercase tracking-wider">
                Gráfico Lineal: Comparativo Mes a Mes por Canal
              </h4>
            </div>
            <p className="text-[11px] text-slate-400 mt-0.5">
              Evolución histórica y tendencia comparada entre el canal directo (Distribuidor Upconnect) y la red externa (Distribuidor Connect)
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {/* Metric Switcher */}
            <div className="flex items-center p-0.5 bg-slate-900 border border-slate-700/80 rounded-lg text-[11px] font-semibold">
              <button
                type="button"
                onClick={() => setLineChartMetric('socios')}
                className={`px-2.5 py-1 rounded-md transition-all ${
                  lineChartMetric === 'socios'
                    ? 'bg-blue-600 text-white shadow-sm'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                Cantidad de Socios / Firmas
              </button>
              <button
                type="button"
                onClick={() => setLineChartMetric('sales')}
                className={`px-2.5 py-1 rounded-md transition-all ${
                  lineChartMetric === 'sales'
                    ? 'bg-emerald-600 text-white shadow-sm'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                Facturación ($ USD)
              </button>
            </div>

            {/* Legend */}
            <div className="flex items-center gap-3 px-3 py-1 bg-slate-900/90 border border-slate-800 rounded-lg text-[11px]">
              <div className="flex items-center gap-1.5">
                <div className="w-2.5 h-2.5 rounded-full bg-sky-400" />
                <span className="font-bold text-sky-300">Distribuidor Upconnect</span>
              </div>
              <div className="flex items-center gap-1.5">
                <div className="w-2.5 h-2.5 rounded-full bg-amber-400" />
                <span className="font-bold text-amber-300">Distribuidor Connect</span>
              </div>
            </div>
          </div>
        </div>

        {chartData.length === 0 ? (
          <div className="py-10 text-center text-slate-500 border border-dashed border-slate-800 rounded-xl bg-slate-900/30">
            <Calendar className="w-8 h-8 text-slate-600 mx-auto mb-2" />
            <p className="text-xs font-semibold text-slate-400">Sin datos históricos mensuales</p>
            <p className="text-[11px] text-slate-500 mt-0.5">
              Cargue el archivo de firmas o socios para visualizar el comparativo mes a mes en este gráfico lineal.
            </p>
          </div>
        ) : (
          (() => {
            const isSocios = lineChartMetric === 'socios';
            const numPoints = chartData.length;
            const valuesUp = chartData.map((d) => (isSocios ? d.upCount : d.upSales));
            const valuesCo = chartData.map((d) => (isSocios ? d.coCount : d.coSales));

            const rawMax = Math.max(1, ...valuesUp, ...valuesCo);
            const maxVal = Math.ceil(rawMax * 1.15);

            const svgWidth = 800;
            const svgHeight = 220;
            const padLeft = 65;
            const padRight = 35;
            const padTop = 20;
            const padBottom = 40;
            const plotW = svgWidth - padLeft - padRight;
            const plotH = svgHeight - padTop - padBottom;

            const getX = (idx: number) => {
              if (numPoints <= 1) return padLeft + plotW / 2;
              return padLeft + (idx / (numPoints - 1)) * plotW;
            };

            const getY = (val: number) => {
              return padTop + plotH - (val / maxVal) * plotH;
            };

            const pointsUpStr = chartData
              .map((d, i) => `${getX(i)},${getY(isSocios ? d.upCount : d.upSales)}`)
              .join(' ');

            const pointsCoStr = chartData
              .map((d, i) => `${getX(i)},${getY(isSocios ? d.coCount : d.coSales)}`)
              .join(' ');

            // Area paths
            const areaUpStr = numPoints > 1
              ? `M ${getX(0)},${padTop + plotH} ` +
                chartData.map((d, i) => `L ${getX(i)},${getY(isSocios ? d.upCount : d.upSales)}`).join(' ') +
                ` L ${getX(numPoints - 1)},${padTop + plotH} Z`
              : '';

            const areaCoStr = numPoints > 1
              ? `M ${getX(0)},${padTop + plotH} ` +
                chartData.map((d, i) => `L ${getX(i)},${getY(isSocios ? d.coCount : d.coSales)}`).join(' ') +
                ` L ${getX(numPoints - 1)},${padTop + plotH} Z`
              : '';

            const gridLevels = [0, 0.33, 0.66, 1];

            const activeItem = hoveredIndex !== null ? chartData[hoveredIndex] : chartData[chartData.length - 1];

            return (
              <div className="space-y-3">
                <div className="relative rounded-xl bg-slate-950/70 border border-slate-800 p-2 overflow-hidden">
                  <svg
                    viewBox={`0 0 ${svgWidth} ${svgHeight}`}
                    className="w-full h-auto select-none overflow-visible"
                    style={{ minHeight: '180px', maxHeight: '250px' }}
                  >
                    <defs>
                      <linearGradient id="gradientUp" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#38bdf8" stopOpacity="0.28" />
                        <stop offset="100%" stopColor="#38bdf8" stopOpacity="0.0" />
                      </linearGradient>
                      <linearGradient id="gradientCo" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#fbbf24" stopOpacity="0.24" />
                        <stop offset="100%" stopColor="#fbbf24" stopOpacity="0.0" />
                      </linearGradient>
                    </defs>

                    {/* Horizontal Gridlines & Y-Axis Labels */}
                    {gridLevels.map((lvl, gIdx) => {
                      const yPos = padTop + plotH - lvl * plotH;
                      const labelVal = lvl * maxVal;
                      const formattedLabel = isSocios
                        ? Math.round(labelVal).toString()
                        : `$${Math.round(labelVal).toLocaleString('en-US')}`;

                      return (
                        <g key={gIdx}>
                          <line
                            x1={padLeft}
                            y1={yPos}
                            x2={svgWidth - padRight}
                            y2={yPos}
                            stroke="#334155"
                            strokeWidth="1"
                            strokeDasharray={lvl === 0 ? undefined : '3 3'}
                            opacity={lvl === 0 ? 0.8 : 0.4}
                          />
                          <text
                            x={padLeft - 8}
                            y={yPos + 3.5}
                            fill="#64748b"
                            fontSize="10"
                            textAnchor="end"
                            fontFamily="monospace"
                          >
                            {formattedLabel}
                          </text>
                        </g>
                      );
                    })}

                    {/* Filled Area UpConnect */}
                    {areaUpStr && (
                      <path d={areaUpStr} fill="url(#gradientUp)" />
                    )}

                    {/* Filled Area Connectors */}
                    {areaCoStr && (
                      <path d={areaCoStr} fill="url(#gradientCo)" />
                    )}

                    {/* Polyline UpConnect */}
                    <polyline
                      points={pointsUpStr}
                      fill="none"
                      stroke="#38bdf8"
                      strokeWidth="3"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    />

                    {/* Polyline Connectors */}
                    <polyline
                      points={pointsCoStr}
                      fill="none"
                      stroke="#fbbf24"
                      strokeWidth="3"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    />

                    {/* Hover vertical bar */}
                    {hoveredIndex !== null && (
                      <line
                        x1={getX(hoveredIndex)}
                        y1={padTop}
                        x2={getX(hoveredIndex)}
                        y2={padTop + plotH}
                        stroke="#94a3b8"
                        strokeWidth="1.5"
                        strokeDasharray="2 2"
                        opacity="0.6"
                      />
                    )}

                    {/* Data Points UpConnect */}
                    {chartData.map((d, i) => {
                      const cx = getX(i);
                      const cy = getY(isSocios ? d.upCount : d.upSales);
                      const isHovered = hoveredIndex === i;

                      return (
                        <g key={`up-${i}`} className="cursor-pointer">
                          <circle
                            cx={cx}
                            cy={cy}
                            r={isHovered ? 6.5 : 4.5}
                            fill="#0b1220"
                            stroke="#38bdf8"
                            strokeWidth={isHovered ? 3 : 2.5}
                            className="transition-all duration-150"
                          />
                          {/* Value label above dot */}
                          <text
                            x={cx}
                            y={cy - 8}
                            fill="#38bdf8"
                            fontSize="10"
                            fontWeight="bold"
                            textAnchor="middle"
                            fontFamily="monospace"
                          >
                            {isSocios ? d.upCount : `$${Math.round(d.upSales)}`}
                          </text>
                        </g>
                      );
                    })}

                    {/* Data Points Connectors */}
                    {chartData.map((d, i) => {
                      const cx = getX(i);
                      const cy = getY(isSocios ? d.coCount : d.coSales);
                      const isHovered = hoveredIndex === i;

                      return (
                        <g key={`co-${i}`} className="cursor-pointer">
                          <circle
                            cx={cx}
                            cy={cy}
                            r={isHovered ? 6.5 : 4.5}
                            fill="#0b1220"
                            stroke="#fbbf24"
                            strokeWidth={isHovered ? 3 : 2.5}
                            className="transition-all duration-150"
                          />
                          {/* Value label below dot */}
                          <text
                            x={cx}
                            y={cy + 16}
                            fill="#fbbf24"
                            fontSize="10"
                            fontWeight="bold"
                            textAnchor="middle"
                            fontFamily="monospace"
                          >
                            {isSocios ? d.coCount : `$${Math.round(d.coSales)}`}
                          </text>
                        </g>
                      );
                    })}

                    {/* X-Axis Month Labels & Interactive Invisible Hitboxes */}
                    {chartData.map((d, i) => {
                      const cx = getX(i);
                      const isHovered = hoveredIndex === i;

                      return (
                        <g
                          key={`axis-${i}`}
                          onMouseEnter={() => setHoveredIndex(i)}
                          onMouseLeave={() => setHoveredIndex(null)}
                          className="cursor-pointer"
                        >
                          {/* Invisible hit box for easier hover */}
                          <rect
                            x={cx - plotW / (numPoints * 2 || 1)}
                            y={padTop}
                            width={plotW / (numPoints || 1)}
                            height={plotH + padBottom}
                            fill="transparent"
                          />
                          <text
                            x={cx}
                            y={svgHeight - 12}
                            fill={isHovered ? '#ffffff' : '#94a3b8'}
                            fontSize="11"
                            fontWeight={isHovered ? 'bold' : '600'}
                            textAnchor="middle"
                          >
                            {d.shortMonth}
                          </text>
                        </g>
                      );
                    })}
                  </svg>
                </div>

                {/* Month-over-Month Comparative Pills below chart */}
                {activeItem && (
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                    {/* UpConnect MoM Card */}
                    <div className="p-3 rounded-xl bg-slate-900/80 border border-sky-500/30 flex items-center justify-between">
                      <div>
                        <div className="text-[10px] uppercase font-bold text-sky-400">
                          UpConnect · {activeItem.month}
                        </div>
                        <div className="text-base font-extrabold text-white tabular-nums mt-0.5">
                          {isSocios
                            ? `${activeItem.upCount} socios / firmas`
                            : `$${activeItem.upSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD`}
                        </div>
                      </div>
                      <div className="text-right">
                        <span className={`px-2 py-0.5 rounded text-[10px] font-bold inline-flex items-center ${
                          (isSocios ? activeItem.upGrowthCount : activeItem.upGrowthSales) >= 0
                            ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30'
                            : 'bg-rose-500/20 text-rose-300 border border-rose-500/30'
                        }`}>
                          {(isSocios ? activeItem.upGrowthCount : activeItem.upGrowthSales) >= 0 ? '+' : ''}
                          {(isSocios ? activeItem.upGrowthCount : activeItem.upGrowthSales).toFixed(1)}% MoM
                        </span>
                        <div className="text-[10px] text-slate-500 mt-0.5">vs mes anterior</div>
                      </div>
                    </div>

                    {/* Connectors MoM Card */}
                    <div className="p-3 rounded-xl bg-slate-900/80 border border-amber-500/30 flex items-center justify-between">
                      <div>
                        <div className="text-[10px] uppercase font-bold text-amber-400">
                          Distribuidor Connect · {activeItem.month}
                        </div>
                        <div className="text-base font-extrabold text-white tabular-nums mt-0.5">
                          {isSocios
                            ? `${activeItem.coCount} socios / firmas`
                            : `$${activeItem.coSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD`}
                        </div>
                      </div>
                      <div className="text-right">
                        <span className={`px-2 py-0.5 rounded text-[10px] font-bold inline-flex items-center ${
                          (isSocios ? activeItem.coGrowthCount : activeItem.coGrowthSales) >= 0
                            ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30'
                            : 'bg-rose-500/20 text-rose-300 border border-rose-500/30'
                        }`}>
                          {(isSocios ? activeItem.coGrowthCount : activeItem.coGrowthSales) >= 0 ? '+' : ''}
                          {(isSocios ? activeItem.coGrowthCount : activeItem.coGrowthSales).toFixed(1)}% MoM
                        </span>
                        <div className="text-[10px] text-slate-500 mt-0.5">vs mes anterior</div>
                      </div>
                    </div>

                    {/* Consolidated MoM Card */}
                    <div className="p-3 rounded-xl bg-slate-900/80 border border-emerald-500/30 flex items-center justify-between">
                      <div>
                        <div className="text-[10px] uppercase font-bold text-emerald-400">
                          Total General · {activeItem.month}
                        </div>
                        <div className="text-base font-extrabold text-emerald-400 tabular-nums mt-0.5">
                          {isSocios
                            ? `${activeItem.totalCount} operaciones`
                            : `$${activeItem.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD`}
                        </div>
                      </div>
                      <div className="text-right">
                        <div className="text-[11px] font-bold text-white tabular-nums">
                          {activeItem.totalSales > 0 ? ((activeItem.upSales / activeItem.totalSales) * 100).toFixed(0) : '0'}% Up / {activeItem.totalSales > 0 ? ((activeItem.coSales / activeItem.totalSales) * 100).toFixed(0) : '0'}% Co
                        </div>
                        <div className="text-[10px] text-slate-400 mt-0.5">Mix de canal</div>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            );
          })()
        )}
      </div>
    </div>
  );
};
