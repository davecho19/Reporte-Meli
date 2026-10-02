import React from 'react';
import { SociosChannelSummary } from '../utils/reportFilters';
import { Zap, Share2, Users, DollarSign, Award, TrendingUp, BarChart2 } from 'lucide-react';

interface SociosChannelTableProps {
  summary: SociosChannelSummary;
  variant?: 'full' | 'slide' | 'compact';
  title?: string;
  subtitle?: string;
  onFilterChannel?: (channel: 'ALL' | 'UpConnect' | 'Connectors') => void;
  activeChannelFilter?: 'ALL' | 'UpConnect' | 'Connectors';
}

export const SociosChannelTable: React.FC<SociosChannelTableProps> = ({
  summary,
  variant = 'full',
  title = 'Distribución y Cantidad de Socios por Canal',
  subtitle = 'Conteo exacto y volumen de ventas: UpConnect vs Connectors',
  onFilterChannel,
  activeChannelFilter = 'ALL',
}) => {
  const { upconnect, connectors, totalSociosCount, totalOperationsCount, totalSales, averageTicket } = summary;

  const upSalesPct = totalSales > 0 ? ((upconnect.totalSales / totalSales) * 100).toFixed(1) : '0.0';
  const coSalesPct = totalSales > 0 ? ((connectors.totalSales / totalSales) * 100).toFixed(1) : '0.0';

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
                  <span className="text-sky-300 font-extrabold">UpConnect</span>
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
                  <span className="text-amber-300 font-extrabold">Connectors</span>
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
              <span>UpConnect ({upconnect.count})</span>
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
              <span>Connectors ({connectors.count})</span>
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
              <div className="text-[11px] font-bold text-sky-400 uppercase">Socios UpConnect</div>
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
              <div className="text-[11px] font-bold text-amber-400 uppercase">Socios Connectors</div>
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
                      <span>UpConnect Directo</span>
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

            {/* Connectors */}
            <tr className="hover:bg-slate-800/40 transition-colors">
              <td className="py-3 px-4">
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 rounded-full bg-amber-400 shrink-0" />
                  <div>
                    <div className="font-bold text-white flex items-center gap-1.5">
                      <span>Connectors (Red Externa)</span>
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
    </div>
  );
};
