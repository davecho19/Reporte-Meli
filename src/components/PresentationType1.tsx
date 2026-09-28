import React, { useState, useEffect } from 'react';
import { GlobalDataset } from '../types';
import { computeReportView } from '../utils/reportFilters';
import { exportPresentationType1ToPPTX } from '../services/pptxExportService';
import { 
  ChevronLeft, 
  ChevronRight, 
  Maximize2, 
  Minimize2, 
  Printer, 
  TrendingUp, 
  Users, 
  DollarSign, 
  ShieldCheck, 
  Layers, 
  Zap, 
  BarChart3,
  Calendar,
  Share2,
  FileDown,
  Loader2,
  Table,
  Search,
  Award,
  FileSignature,
  Edit3,
  RotateCcw,
  Check,
  X,
  Package,
  Sparkles
} from 'lucide-react';

interface Props {
  dataset: GlobalDataset;
  onBackToDashboard?: () => void;
}

interface TopProductItem {
  rank: number;
  name: string;
  count: number;
  sales: number;
  percentage: number;
  label?: string;
}

export const PresentationType1: React.FC<Props> = ({ dataset, onBackToDashboard }) => {
  const [currentSlide, setCurrentSlide] = useState<number>(1);
  const [isFullscreen, setIsFullscreen] = useState<boolean>(false);
  const [sociosViewMode, setSociosViewMode] = useState<'table' | 'ranking'>('table');
  const [socioSearch, setSocioSearch] = useState<string>('');
  const [isExportingPPTX, setIsExportingPPTX] = useState<boolean>(false);

  const viewData = computeReportView(dataset);
  const activeSocios = viewData.socios;

  // Dynamic filter check: When filter is 'all' (Consolidado), hide weekly evolution slide!
  const filterType = dataset.reportFilter?.type || 'all';
  const showWeeklySlide = filterType !== 'all';

  interface SlideConfig {
    id: 'cover' | 'balance' | 'monthly' | 'channels' | 'weekly' | 'products' | 'communities' | 'socios';
    title: string;
  }

  const activeSlides: SlideConfig[] = [
    { id: 'cover', title: 'Portada' },
    { id: 'balance', title: 'Balance General' },
    { id: 'monthly', title: 'Evolución Comparativa' },
    { id: 'channels', title: 'Resumen por Canal' },
    ...(showWeeklySlide ? [{ id: 'weekly' as const, title: 'Evolución Semana a Semana' }] : []),
    { id: 'products', title: 'Top 5 Productos más Vendidos' },
    { id: 'communities', title: 'Comunidades y Franquiciados' },
    { id: 'socios', title: 'Rendimiento de Socios' },
  ];

  const totalSlides = activeSlides.length;

  // Clamp currentSlide if slide count changes
  useEffect(() => {
    if (currentSlide > totalSlides) {
      setCurrentSlide(totalSlides);
    }
  }, [totalSlides, currentSlide]);

  // Keyboard navigation
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

  // Communities: Auto calculation based on registered socios & manual customization
  const autoUpSocios = activeSocios.filter(
    (s) => !(s.role?.toLowerCase().includes('connect') || s.role?.toLowerCase().includes('distribuidor'))
  ).length;
  const autoCoSocios = activeSocios.filter(
    (s) => s.role?.toLowerCase().includes('connect') || s.role?.toLowerCase().includes('distribuidor')
  ).length;
  const autoFormacion = Math.max(1, Math.round((autoUpSocios + autoCoSocios) * 0.25)) || 51;

  const [customCommunities, setCustomCommunities] = useState<{
    upcontaSocios?: number;
    franquiciaVIP?: number;
    formacionComercial?: number;
  }>({});
  const [isEditingCommunities, setIsEditingCommunities] = useState<boolean>(false);
  const [tempUpconta, setTempUpconta] = useState<number>(0);
  const [tempFranquicia, setTempFranquicia] = useState<number>(0);
  const [tempFormacion, setTempFormacion] = useState<number>(0);

  const upcontaMembers = customCommunities.upcontaSocios ?? (autoUpSocios > 0 ? autoUpSocios : 118);
  const franquiciaMembers = customCommunities.franquiciaVIP ?? (autoCoSocios > 0 ? autoCoSocios : 129);
  const formacionMembers = customCommunities.formacionComercial ?? (autoCoSocios > 0 ? autoFormacion : 51);
  const isManuallyModified =
    customCommunities.upcontaSocios !== undefined ||
    customCommunities.franquiciaVIP !== undefined ||
    customCommunities.formacionComercial !== undefined;

  const handlePrint = () => {
    window.print();
  };

  const handleExportPPTX = async () => {
    try {
      setIsExportingPPTX(true);
      await exportPresentationType1ToPPTX(dataset, customCommunities);
    } catch (e) {
      console.error('Error generating PowerPoint file', e);
    } finally {
      setIsExportingPPTX(false);
    }
  };

  const totalUpSales = viewData.totalUpSales;
  const totalCoSales = viewData.totalCoSales;
  const totalSales = viewData.totalSales;
  const upSalesPct = viewData.upSalesPct;
  const coSalesPct = viewData.coSalesPct;

  const totalUpCount = viewData.totalUpCount;
  const totalCoCount = viewData.totalCoCount;
  const totalCount = viewData.totalCount;

  const activeMonthlyMetrics = (viewData.monthlyMetrics && viewData.monthlyMetrics.length > 0)
    ? viewData.monthlyMetrics
    : dataset.monthlyMetrics;

  // Helper: Compute Top 5 Products per channel
  const computeTop5Products = (channel: 'UpConnect' | 'Connectors'): TopProductItem[] => {
    const map = new Map<string, { name: string; count: number; sales: number }>();

    // 1. From transactions
    viewData.transactions
      .filter((t) => t.channel === channel)
      .forEach((t) => {
        let prodName = t.solutionCategory ? t.solutionCategory.trim() : 'Firma Electrónica';
        if (t.duration && !prodName.toLowerCase().includes(t.duration.toLowerCase())) {
          prodName = `${prodName} (${t.duration})`;
        }
        const key = prodName;
        if (!map.has(key)) {
          map.set(key, { name: key, count: 0, sales: 0 });
        }
        const item = map.get(key)!;
        item.count += 1;
        item.sales += t.value || 0;
      });

    // 2. Fallback to portfolioDurations if transactions are not loaded
    if (map.size === 0 && dataset.portfolioDurations && dataset.portfolioDurations.length > 0) {
      dataset.portfolioDurations
        .filter((p) => p.channel === channel)
        .forEach((p) => {
          map.set(p.duration, {
            name: `Firma Electrónica (${p.duration})`,
            count: p.count,
            sales: p.count * 15,
          });
        });
    }

    const list = Array.from(map.values()).sort((a, b) => b.count - a.count || b.sales - a.sales);
    const channelTotalCount = list.reduce((a, b) => a + b.count, 0) || 1;

    return list.slice(0, 5).map((item, idx) => ({
      rank: idx + 1,
      name: item.name,
      count: item.count,
      sales: parseFloat(item.sales.toFixed(2)),
      percentage: parseFloat(((item.count / channelTotalCount) * 100).toFixed(1)),
      label: idx === 0 ? 'Líder' : undefined,
    }));
  };

  const topUpProducts = computeTop5Products('UpConnect');
  const topCoProducts = computeTop5Products('Connectors');

  // Jump helper for socios slide
  const jumpToSociosSlide = () => {
    const idx = activeSlides.findIndex((s) => s.id === 'socios');
    if (idx !== -1) setCurrentSlide(idx + 1);
  };

  // Reusable Slide Content Renderer (used both for screen and print)
  const renderSlideContent = (slideId: string, pageNum: number, isPrint: boolean = false) => {
    switch (slideId) {
      case 'cover':
        return (
          <div className="h-full flex flex-col justify-between p-8 sm:p-14 relative bg-gradient-to-br from-[#0c1322] via-[#09101d] to-[#040812]">
            <div className="absolute w-[450px] h-[450px] rounded-full bg-blue-600/10 blur-[120px] pointer-events-none -top-16 -left-16" />
            <div className="absolute w-[400px] h-[400px] rounded-full bg-amber-500/10 blur-[100px] pointer-events-none -bottom-16 -right-16" />

            <div className="z-10 max-w-4xl space-y-5 my-auto">
              <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-sky-500/10 border border-sky-500/20 text-sky-400 text-xs font-semibold tracking-wide">
                <span>DOCUMENTO OFICIAL GERENCIAL</span>
                <span>·</span>
                <span>{viewData.periodLabel.toUpperCase()}</span>
              </div>

              <h1 className="text-3xl sm:text-5xl font-extrabold tracking-tight text-white leading-tight">
                {viewData.titleReport1}
              </h1>

              <p className="text-slate-400 text-sm sm:text-base max-w-2xl pt-1 leading-relaxed">
                Auditoría comparativa de emisión de certificados, rendimiento de canales directos y red externa de distribuidores con corte al {viewData.subtitleDate}.
              </p>

              <div className="pt-6 flex flex-wrap items-center gap-6 text-xs text-slate-400">
                <div className="flex items-center gap-2">
                  <span className="w-2.5 h-2.5 rounded-full bg-sky-400" />
                  <span>Canal Directo: <strong>UpConnect (UP)</strong></span>
                </div>
                <div className="flex items-center gap-2">
                  <span className="w-2.5 h-2.5 rounded-full bg-amber-400" />
                  <span>Red Externa: <strong>Connectors (CO)</strong></span>
                </div>
                <div className="flex items-center gap-2">
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-400" />
                  <span>Total Firmas: <strong>{totalCount}</strong></span>
                </div>
                <div className="flex items-center gap-2">
                  <span className="w-2.5 h-2.5 rounded-full bg-purple-400" />
                  <span>Total Facturado: <strong>${totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD</strong></span>
                </div>
              </div>
            </div>

            <div className="flex justify-between items-center text-[11px] text-slate-500 border-t border-slate-800/60 pt-3">
              <span>UpConnect · Dirección Comercial y Estrategia</span>
              <span>Pág. 0{pageNum}</span>
            </div>
          </div>
        );

      case 'balance':
        return (
          <div className="h-full flex flex-col justify-between p-6 sm:p-10 relative bg-[#0c1322]">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-800/80 pb-3">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  RESUMEN ACUMULADO 2026
                </span>
                <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                  Balance General de Emisiones y Rendimiento por Operador
                </h2>
              </div>

              <div className="flex items-center gap-2.5">
                <div className="px-3.5 py-1.5 rounded-xl bg-slate-900/90 border border-slate-700 text-right">
                  <div className="text-[9px] uppercase font-semibold text-slate-400">TOTAL FIRMAS</div>
                  <div className="text-sm sm:text-base font-bold text-white tabular-nums">{totalCount} Operaciones</div>
                </div>
                <div className="px-3.5 py-1.5 rounded-xl bg-slate-900/90 border border-slate-700 text-right">
                  <div className="text-[9px] uppercase font-semibold text-slate-400">FACTURACIÓN TOTAL</div>
                  <div className="text-sm sm:text-base font-bold text-emerald-400 tabular-nums">
                    ${totalSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                </div>
              </div>
            </div>

            {/* Two Big Channel Comparison Cards */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 my-auto">
              {/* UpConnect Card */}
              <div className="p-6 rounded-2xl bg-slate-900/80 border border-slate-800 shadow-xl relative overflow-hidden flex flex-col justify-between">
                <div className="flex items-center justify-between pb-3 border-b border-slate-800">
                  <div className="flex items-center gap-2.5">
                    <div className="w-9 h-9 rounded-xl bg-blue-500/20 text-sky-400 flex items-center justify-center font-bold">
                      <Zap className="w-4 h-4" />
                    </div>
                    <div>
                      <h3 className="text-base font-bold text-white">UpConnect (UP)</h3>
                      <p className="text-[11px] text-slate-400">Canal Propio / Directo</p>
                    </div>
                  </div>
                  <span className="px-2.5 py-0.5 rounded-md bg-blue-500/20 text-sky-300 text-xs font-semibold">
                    {upSalesPct}% de Ventas
                  </span>
                </div>

                <div className="grid grid-cols-3 gap-3 py-5 text-center">
                  <div>
                    <div className="text-xs text-slate-400">Ventas Totales</div>
                    <div className="text-lg sm:text-xl font-black text-sky-400 tabular-nums mt-1">
                      ${totalUpSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-400">Firmas Emitidas</div>
                    <div className="text-lg sm:text-xl font-black text-white tabular-nums mt-1">
                      {totalUpCount}
                    </div>
                    <div className="text-[10px] text-slate-500">({totalUpCount} en el periodo)</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-400">Capacidad</div>
                    <div className="text-lg sm:text-xl font-black text-slate-200 tabular-nums mt-1">
                      {totalUpCount > 0 ? (totalUpCount > 131 ? totalUpCount : 131) : 0}
                    </div>
                    <div className="text-[10px] text-slate-500">Operadores ({totalUpCount > 0 ? '58.0%' : '0%'})</div>
                  </div>
                </div>

                <div className="pt-3 border-t border-slate-800/80 flex items-center justify-between">
                  <span className="text-xs text-slate-400">Rendimiento por Operador:</span>
                  <span className="text-sm font-bold text-sky-400 tabular-nums">
                    ${totalUpCount > 0 ? (totalUpSales / (totalUpCount > 131 ? totalUpCount : 131)).toFixed(2) : '0.00'} / operador
                  </span>
                </div>
                <div className="absolute bottom-0 left-0 right-0 h-1 bg-blue-500" />
              </div>

              {/* Connectors Card */}
              <div className="p-6 rounded-2xl bg-slate-900/80 border border-slate-800 shadow-xl relative overflow-hidden flex flex-col justify-between">
                <div className="flex items-center justify-between pb-3 border-b border-slate-800">
                  <div className="flex items-center gap-2.5">
                    <div className="w-9 h-9 rounded-xl bg-amber-500/20 text-amber-400 flex items-center justify-center font-bold">
                      <Share2 className="w-4 h-4" />
                    </div>
                    <div>
                      <h3 className="text-base font-bold text-white">Connectors (CO)</h3>
                      <p className="text-[11px] text-slate-400">Red Externa / Aliados</p>
                    </div>
                  </div>
                  <span className="px-2.5 py-0.5 rounded-md bg-amber-500/20 text-amber-300 text-xs font-semibold">
                    {coSalesPct}% de Ventas
                  </span>
                </div>

                <div className="grid grid-cols-3 gap-3 py-5 text-center">
                  <div>
                    <div className="text-xs text-slate-400">Ventas Totales</div>
                    <div className="text-lg sm:text-xl font-black text-amber-400 tabular-nums mt-1">
                      ${totalCoSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-400">Firmas Emitidas</div>
                    <div className="text-lg sm:text-xl font-black text-white tabular-nums mt-1">
                      {totalCoCount}
                    </div>
                    <div className="text-[10px] text-slate-500">({totalCoCount} en el periodo)</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-400">Capacidad</div>
                    <div className="text-lg sm:text-xl font-black text-slate-200 tabular-nums mt-1">
                      {totalCoCount > 0 ? (totalCoCount > 95 ? totalCoCount : 95) : 0}
                    </div>
                    <div className="text-[10px] text-slate-500">Operadores ({totalCoCount > 0 ? '42.0%' : '0%'})</div>
                  </div>
                </div>

                <div className="pt-3 border-t border-slate-800/80 flex items-center justify-between">
                  <span className="text-xs text-slate-400">Rendimiento por Operador:</span>
                  <span className="text-sm font-bold text-amber-400 tabular-nums">
                    ${totalCoCount > 0 ? (totalCoSales / (totalCoCount > 95 ? totalCoCount : 95)).toFixed(2) : '0.00'} / operador
                  </span>
                </div>
                <div className="absolute bottom-0 left-0 right-0 h-1 bg-amber-500" />
              </div>
            </div>

            <div className="flex justify-between items-center text-[11px] text-slate-500 border-t border-slate-800/60 pt-3">
              <div className="flex items-center gap-3">
                <span>Rendimiento normalizado por canal directo vs distribuidores aliados</span>
                {!isPrint && (
                  <button
                    onClick={jumpToSociosSlide}
                    className="text-sky-400 hover:text-sky-300 hover:underline font-semibold flex items-center gap-1"
                  >
                    <span>Ver Tabla de Socios / Vendedores →</span>
                  </button>
                )}
              </div>
              <span>Pág. 0{pageNum}</span>
            </div>
          </div>
        );

      case 'monthly':
        return (
          <div className="h-full flex flex-col justify-between p-6 sm:p-10 relative bg-[#0c1322]">
            <div>
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  EVOLUCIÓN VISUAL COMPARATIVA
                </span>
                <div className="flex items-center gap-4 text-xs font-medium">
                  <div className="flex items-center gap-2">
                    <span className="w-3 h-3 rounded bg-blue-600" />
                    <span className="text-slate-300">UpConnect (UP)</span>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="w-3 h-3 rounded bg-amber-500" />
                    <span className="text-slate-300">Connectors (CO)</span>
                  </div>
                </div>
              </div>
              <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                Ventas Acumuladas año 2026: UpConnect vs Connectors
              </h2>
            </div>

            {/* Interactive Bar Chart matching Slide 3 */}
            <div className="my-auto pt-3 pb-2">
              {activeMonthlyMetrics.length === 0 ? (
                <div className="h-52 flex flex-col items-center justify-center border border-dashed border-slate-800 rounded-xl p-8 text-center text-slate-500">
                  <BarChart3 className="w-10 h-10 text-slate-700 mb-2" />
                  <p className="font-semibold text-slate-400 text-sm">Sin datos de facturación mensual</p>
                  <p className="text-xs text-slate-500 mt-1 max-w-md">
                    Cargue el archivo de firmas o socios en el Dashboard para renderizar el comparativo de ventas mensuales.
                  </p>
                </div>
              ) : (
                (() => {
                  const maxChannelSale = Math.max(
                    1,
                    ...activeMonthlyMetrics.map((m) => Math.max(m.upconnectSales, m.connectorsSales, 1))
                  );

                  return (
                    <div className="h-56 flex items-end justify-around gap-2 sm:gap-6 border-b border-slate-800 px-4 pb-2 pt-4">
                      {activeMonthlyMetrics.map((item) => {
                        const upPct = Math.min(100, Math.max(item.upconnectSales > 0 ? 8 : 0, (item.upconnectSales / maxChannelSale) * 100));
                        const coPct = Math.min(100, Math.max(item.connectorsSales > 0 ? 8 : 0, (item.connectorsSales / maxChannelSale) * 100));

                        return (
                          <div key={item.month} className="flex-1 flex flex-col items-center justify-end h-full max-w-[95px] group">
                            <div className="w-full flex items-end justify-center gap-1.5 sm:gap-2 h-40 relative">
                              {/* UP Bar */}
                              <div className="h-full flex flex-col items-center justify-end flex-1 max-w-[42px]">
                                {item.upconnectSales > 0 && (
                                  <span className="text-[10px] font-bold text-sky-400 tabular-nums mb-1 leading-none text-center">
                                    ${Math.round(item.upconnectSales).toLocaleString()}
                                  </span>
                                )}
                                <div 
                                  style={{ height: `${upPct}%` }} 
                                  className="w-full min-h-[4px] bg-gradient-to-t from-blue-700 to-blue-500 rounded-t-md transition-all duration-300 shadow-md shadow-blue-900/40"
                                  title={`UpConnect ${item.month}: $${item.upconnectSales.toFixed(2)} USD (${item.upconnectCount} firmas)`}
                                />
                              </div>

                              {/* CO Bar */}
                              <div className="h-full flex flex-col items-center justify-end flex-1 max-w-[42px]">
                                {item.connectorsSales > 0 && (
                                  <span className="text-[10px] font-bold text-amber-400 tabular-nums mb-1 leading-none text-center">
                                    ${Math.round(item.connectorsSales).toLocaleString()}
                                  </span>
                                )}
                                <div 
                                  style={{ height: `${coPct}%` }} 
                                  className="w-full min-h-[4px] bg-gradient-to-t from-amber-600 to-amber-400 rounded-t-md transition-all duration-300 shadow-md shadow-amber-900/40"
                                  title={`Connectors ${item.month}: $${item.connectorsSales.toFixed(2)} USD (${item.connectorsCount} firmas)`}
                                />
                              </div>
                            </div>

                            <div className="mt-2.5 text-center">
                              <div className="text-xs font-semibold text-slate-200 capitalize">
                                {item.month.toLowerCase()}
                              </div>
                              <div className="text-[10px] text-slate-400 tabular-nums mt-0.5">
                                {item.connectorsCount > 0 
                                  ? `UP: ${item.upconnectCount} | CO: ${item.connectorsCount}`
                                  : `${item.upconnectCount} firmas`}
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  );
                })()
              )}
            </div>

            {/* Callout Conclusion Box */}
            <div className="p-3 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
              <div>
                <strong className="text-sky-400">Conclusión del Periodo:</strong>{' '}
                {totalSales > 0 ? (
                  <>
                    Facturación consolidada de{' '}
                    <strong className="text-white">${totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD</strong>{' '}
                    con <strong className="text-emerald-400">{totalCount}</strong> certificados emitidos en total (UpConnect: ${totalUpSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD / Connectors: ${totalCoSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD).
                  </>
                ) : (
                  <span>Sin registros de facturación cargados. La plataforma y los reportes están vacíos.</span>
                )}
              </div>
              <span className="text-[10px] text-slate-500 hidden sm:inline">Pág. 0{pageNum}</span>
            </div>
          </div>
        );

      case 'channels':
        return (
          <div className="h-full flex flex-col justify-between p-6 sm:p-10 relative bg-[#0c1322]">
            <div>
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  {viewData.periodLabel.toUpperCase()} · CORTE AUDITADO
                </span>
                <span className="px-2.5 py-0.5 rounded text-[11px] bg-slate-800 text-slate-400">
                  {viewData.subtitleDate}
                </span>
              </div>
              <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                {viewData.isFiltered ? viewData.titleReport1 : 'Resumen de Ventas por Canal y Total General'}
              </h2>
            </div>

            {/* Two Horizontal Channel Panels */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-5 my-auto">
              {/* Canal 01 UP */}
              <div className="p-5 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-lg relative">
                <div className="flex items-center justify-between pb-2.5">
                  <span className="text-xs font-bold text-sky-400 uppercase tracking-wider">CANAL 01</span>
                  <Zap className="w-4 h-4 text-sky-400" />
                </div>
                <div className="flex items-baseline justify-between mb-3">
                  <h3 className="text-lg font-black text-white">UPCONNECT</h3>
                  <div className="text-right">
                    <span className="text-[11px] text-slate-400">FIRMAS ({viewData.periodLabel})</span>
                    <div className="text-base font-bold text-white tabular-nums">{totalUpCount} Firmas</div>
                  </div>
                </div>

                <div className="pt-3 border-t border-slate-800">
                  <div className="text-xs text-slate-400">RECAUDADO ACUM.</div>
                  <div className="text-2xl font-black text-sky-400 tabular-nums mt-0.5">${totalUpSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD</div>
                  <div className="text-xs font-semibold text-slate-300 mt-1">{upSalesPct}% de participación</div>
                </div>
              </div>

              {/* Canal 02 CONNECTORS */}
              <div className="p-5 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-lg relative">
                <div className="flex items-center justify-between pb-2.5">
                  <span className="text-xs font-bold text-amber-400 uppercase tracking-wider">CANAL 02</span>
                  <Share2 className="w-4 h-4 text-amber-400" />
                </div>
                <div className="flex items-baseline justify-between mb-3">
                  <h3 className="text-lg font-black text-white">CONNECTORS</h3>
                  <div className="text-right">
                    <span className="text-[11px] text-slate-400">FIRMAS ({viewData.periodLabel})</span>
                    <div className="text-base font-bold text-white tabular-nums">{totalCoCount} Firmas</div>
                  </div>
                </div>

                <div className="pt-3 border-t border-slate-800">
                  <div className="text-xs text-slate-400">RECAUDADO ACUM.</div>
                  <div className="text-2xl font-black text-amber-400 tabular-nums mt-0.5">${totalCoSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD</div>
                  <div className="text-xs font-semibold text-slate-300 mt-1">{coSalesPct}% de participación</div>
                </div>
              </div>
            </div>

            {/* General Consolidated Green Strip */}
            <div className="p-4 sm:p-5 rounded-2xl bg-gradient-to-r from-emerald-950/70 via-emerald-900/60 to-slate-900 border border-emerald-500/30 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div>
                <span className="text-xs text-emerald-400 font-semibold uppercase tracking-wider">
                  Consolidado General ({viewData.periodLabel})
                </span>
                <h4 className="text-base font-bold text-white">
                  SUMA TOTAL (UPCONNECT + CONNECTORS)
                </h4>
              </div>

              <div className="flex items-center gap-6">
                <div className="text-right">
                  <div className="text-xs text-slate-400">TOTAL FIRMAS</div>
                  <div className="text-xl font-black text-white tabular-nums">{totalCount} Firmas</div>
                </div>
                <div className="h-8 w-[1px] bg-slate-700 hidden sm:block" />
                <div className="text-right">
                  <div className="text-xs text-slate-400">RECAUDACIÓN TOTAL</div>
                  <div className="text-xl font-black text-emerald-400 tabular-nums">${totalSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD</div>
                </div>
              </div>
            </div>

            <div className="flex justify-between items-center text-[11px] text-slate-500 pt-2 border-t border-slate-800/60">
              <div className="flex items-center gap-3">
                <span>Los canales se presentan de manera 100% independiente garantizando trazabilidad contable.</span>
                {!isPrint && (
                  <button
                    onClick={jumpToSociosSlide}
                    className="text-sky-400 hover:text-sky-300 hover:underline font-semibold flex items-center gap-1"
                  >
                    <span>Ver Tabla de Socios / Vendedores →</span>
                  </button>
                )}
              </div>
              <span>Pág. 0{pageNum}</span>
            </div>
          </div>
        );

      case 'weekly':
        return (
          <div className="h-full flex flex-col justify-between p-6 sm:p-10 relative bg-[#0c1322]">
            <div>
              <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                FOCO DE RENDIMIENTO OPERATIVO
              </span>
              <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                Evolución Semana a Semana: {viewData.periodLabel}
              </h2>
            </div>

            {/* Weekly Top Cards or Empty state */}
            {dataset.weeklyBreakdownType1.length === 0 ? (
              <div className="my-auto py-12 text-center text-slate-500 border border-dashed border-slate-800 rounded-xl max-w-2xl mx-auto px-6">
                <p className="font-semibold text-slate-400">Sin desglose semanal registrado</p>
                <p className="text-xs text-slate-500 mt-1">
                  Cargue el archivo de firmas o socios para visualizar la evolución semana a semana.
                </p>
              </div>
            ) : (
              <>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 my-2">
                  {dataset.weeklyBreakdownType1.map((w) => (
                    <div key={w.id} className="p-3.5 rounded-xl bg-slate-900/80 border border-slate-800 text-center">
                      <span className="text-[11px] font-bold text-sky-400 uppercase">
                        {w.weekName} ({w.dateRange})
                      </span>
                      <div className="text-base font-black text-white tabular-nums mt-0.5">
                        {w.totalCount} Firmas | ${w.totalAmount.toFixed(2)} USD
                      </div>
                      <div className="text-[11px] text-slate-400 mt-0.5">
                        UP: {w.upconnectCount} (${w.upconnectAmount.toFixed(2)}) • CO: {w.connectorsCount} (${w.connectorsAmount.toFixed(2)})
                      </div>
                    </div>
                  ))}
                </div>

                {/* Dual Bar Chart for Weeks */}
                <div className="my-auto py-1">
                  <div className="flex justify-end gap-4 text-xs font-medium mb-2">
                    <div className="flex items-center gap-2">
                      <span className="w-2.5 h-2.5 rounded bg-blue-500" />
                      <span className="text-slate-300">Upconnect (UP)</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <span className="w-2.5 h-2.5 rounded bg-amber-500" />
                      <span className="text-slate-300">Connectors (CO)</span>
                    </div>
                  </div>

                  <div className="h-44 flex items-end justify-around border-b border-slate-800 px-6 pb-2">
                    {(() => {
                      const maxWeekCount = Math.max(
                        1,
                        ...dataset.weeklyBreakdownType1.map((w) => Math.max(w.upconnectCount, w.connectorsCount, 1))
                      );

                      return dataset.weeklyBreakdownType1.map((w) => {
                        const upPct = Math.min(100, Math.max(w.upconnectCount > 0 ? 8 : 0, (w.upconnectCount / maxWeekCount) * 100));
                        const coPct = Math.min(100, Math.max(w.connectorsCount > 0 ? 8 : 0, (w.connectorsCount / maxWeekCount) * 100));

                        return (
                          <div key={w.id} className="flex flex-col items-center w-48">
                            <div className="flex items-end justify-center gap-2 w-full h-32 relative">
                              {/* UP Bar */}
                              <div className="h-full flex flex-col items-center justify-end w-14">
                                {w.upconnectCount > 0 && (
                                  <span className="text-[11px] font-bold text-sky-400 tabular-nums mb-1 leading-none text-center">
                                    {w.upconnectCount}
                                  </span>
                                )}
                                <div 
                                  style={{ height: `${upPct}%` }}
                                  className="w-full min-h-[4px] bg-gradient-to-t from-blue-600 to-blue-400 rounded-t-md transition-all shadow-md shadow-blue-900/20"
                                  title={`UpConnect ${w.weekName}: ${w.upconnectCount} firmas ($${w.upconnectAmount.toFixed(2)} USD)`}
                                />
                              </div>
                              {/* CO Bar */}
                              <div className="h-full flex flex-col items-center justify-end w-14">
                                {w.connectorsCount > 0 && (
                                  <span className="text-[11px] font-bold text-amber-400 tabular-nums mb-1 leading-none text-center">
                                    {w.connectorsCount}
                                  </span>
                                )}
                                <div 
                                  style={{ height: `${coPct}%` }}
                                  className="w-full min-h-[4px] bg-gradient-to-t from-amber-600 to-amber-400 rounded-t-md transition-all shadow-md shadow-amber-900/20"
                                  title={`Connectors ${w.weekName}: ${w.connectorsCount} firmas ($${w.connectorsAmount.toFixed(2)} USD)`}
                                />
                              </div>
                            </div>
                            <span className="text-xs text-slate-300 font-medium mt-2.5 text-center">
                              {w.weekName} ({w.dateRange})
                            </span>
                          </div>
                        );
                      });
                    })()}
                  </div>
                </div>
              </>
            )}

            {/* Green highlight banner */}
            <div className="p-2.5 rounded-xl bg-emerald-500/20 border border-emerald-500/30 text-center text-xs text-emerald-300 font-bold">
              {totalCount > 0 
                ? `TOTAL CONSOLIDADO ${viewData.periodLabel.toUpperCase()}: ${totalCount} FIRMAS | $${totalSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD (Upconnect: ${totalUpCount} firmas | Connectors: ${totalCoCount} firmas)`
                : 'TOTAL CONSOLIDADO: 0 FIRMAS | $0.00 USD (Sin datos cargados)'
              }
            </div>
          </div>
        );

      case 'products':
        return (
          <div className="h-full flex flex-col justify-between p-6 sm:p-10 relative bg-[#0c1322]">
            <div>
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  ESTRUCTURA DE PORTAFOLIO Y CATÁLOGO
                </span>
                <span className="px-2.5 py-0.5 rounded text-[11px] bg-blue-500/10 text-sky-300 border border-blue-500/20 font-semibold">
                  TOP 5 POR CANAL
                </span>
              </div>
              <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                Top 5 de Producto Más Vendido: UpConnect vs Connectors
              </h2>
            </div>

            {/* Two Side-by-side Top 5 Product Columns */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 my-auto">
              {/* UpConnect Top 5 Column */}
              <div className="p-5 rounded-2xl bg-slate-900/90 border-t-4 border-blue-500 border-x border-b border-slate-800 shadow-xl flex flex-col justify-between">
                <div className="flex items-center justify-between pb-3 border-b border-slate-800">
                  <div className="flex items-center gap-2">
                    <Zap className="w-4 h-4 text-sky-400" />
                    <h3 className="text-base font-bold text-white">UpConnect (Canal Propio)</h3>
                  </div>
                  <span className="text-xs text-sky-400 font-semibold tabular-nums">
                    {totalUpCount} Firmas Totales
                  </span>
                </div>

                <div className="divide-y divide-slate-800/70 py-1 space-y-2 mt-2">
                  {topUpProducts.length === 0 ? (
                    <div className="py-8 text-center text-slate-500 text-xs">
                      Sin datos de productos para UpConnect
                    </div>
                  ) : (
                    topUpProducts.map((p) => (
                      <div key={p.rank} className="pt-2">
                        <div className="flex items-center justify-between text-xs mb-1">
                          <div className="flex items-center gap-2">
                            <span className={`w-5 h-5 rounded-full flex items-center justify-center font-bold text-[10px] ${
                              p.rank === 1 ? 'bg-blue-600 text-white shadow-sm' : 'bg-slate-800 text-slate-300'
                            }`}>
                              {p.rank}
                            </span>
                            <span className="font-semibold text-slate-200">
                              {p.name} {p.label && <span className="text-sky-400 font-bold ml-1">({p.label})</span>}
                            </span>
                          </div>
                          <div className="text-right">
                            <strong className="text-white tabular-nums">{p.count} firmas</strong>
                            <span className="text-slate-400 text-[11px] ml-2">(${p.sales.toFixed(2)})</span>
                          </div>
                        </div>

                        <div className="w-full bg-slate-800 rounded-full h-1.5 overflow-hidden">
                          <div
                            style={{ width: `${Math.max(5, p.percentage)}%` }}
                            className={`h-full ${p.rank === 1 ? 'bg-sky-400' : 'bg-blue-600'} rounded-full transition-all`}
                          />
                        </div>
                      </div>
                    ))
                  )}
                </div>

                <div className="pt-3 mt-2 border-t border-slate-800/80 flex items-center justify-between text-xs text-slate-400">
                  <span>Concentración Top 5:</span>
                  <strong className="text-sky-300 tabular-nums">
                    {topUpProducts.reduce((a, b) => a + b.count, 0)} firmas ({topUpProducts.reduce((a, b) => a + b.percentage, 0).toFixed(1)}%)
                  </strong>
                </div>
              </div>

              {/* Connectors Top 5 Column */}
              <div className="p-5 rounded-2xl bg-slate-900/90 border-t-4 border-amber-500 border-x border-b border-slate-800 shadow-xl flex flex-col justify-between">
                <div className="flex items-center justify-between pb-3 border-b border-slate-800">
                  <div className="flex items-center gap-2">
                    <Share2 className="w-4 h-4 text-amber-400" />
                    <h3 className="text-base font-bold text-white">Connectors (Red Externa)</h3>
                  </div>
                  <span className="text-xs text-amber-400 font-semibold tabular-nums">
                    {totalCoCount} Firmas Totales
                  </span>
                </div>

                <div className="divide-y divide-slate-800/70 py-1 space-y-2 mt-2">
                  {topCoProducts.length === 0 ? (
                    <div className="py-8 text-center text-slate-500 text-xs">
                      Sin datos de productos para Connectors
                    </div>
                  ) : (
                    topCoProducts.map((p) => (
                      <div key={p.rank} className="pt-2">
                        <div className="flex items-center justify-between text-xs mb-1">
                          <div className="flex items-center gap-2">
                            <span className={`w-5 h-5 rounded-full flex items-center justify-center font-bold text-[10px] ${
                              p.rank === 1 ? 'bg-amber-600 text-white shadow-sm' : 'bg-slate-800 text-slate-300'
                            }`}>
                              {p.rank}
                            </span>
                            <span className="font-semibold text-slate-200">
                              {p.name} {p.label && <span className="text-amber-400 font-bold ml-1">({p.label})</span>}
                            </span>
                          </div>
                          <div className="text-right">
                            <strong className="text-white tabular-nums">{p.count} firmas</strong>
                            <span className="text-slate-400 text-[11px] ml-2">(${p.sales.toFixed(2)})</span>
                          </div>
                        </div>

                        <div className="w-full bg-slate-800 rounded-full h-1.5 overflow-hidden">
                          <div
                            style={{ width: `${Math.max(5, p.percentage)}%` }}
                            className={`h-full ${p.rank === 1 ? 'bg-amber-400' : 'bg-amber-600'} rounded-full transition-all`}
                          />
                        </div>
                      </div>
                    ))
                  )}
                </div>

                <div className="pt-3 mt-2 border-t border-slate-800/80 flex items-center justify-between text-xs text-slate-400">
                  <span>Concentración Top 5:</span>
                  <strong className="text-amber-300 tabular-nums">
                    {topCoProducts.reduce((a, b) => a + b.count, 0)} firmas ({topCoProducts.reduce((a, b) => a + b.percentage, 0).toFixed(1)}%)
                  </strong>
                </div>
              </div>
            </div>

            <div className="flex justify-between items-center text-[11px] text-slate-500 pt-2 border-t border-slate-800/60">
              <span>El producto líder concentra la mayor preferencia del cliente final con alta recurrencia de renovación.</span>
              <span>Pág. 0{pageNum}</span>
            </div>
          </div>
        );

      case 'communities':
        return (
          <div className="h-full flex flex-col justify-between p-6 sm:p-10 relative bg-[#0c1322]">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  RED HUMANA Y DISTRIBUCIÓN
                </span>
                <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                  Comunidades y Red de Franquiciados
                </h2>
                <p className="text-xs text-slate-400 mt-0.5">
                  Monitoreo activo de grupos de gestión, formación comercial y franquiciados contables
                </p>
              </div>

              {!isPrint && (
                <div className="flex items-center gap-2">
                  {isManuallyModified && (
                    <button
                      onClick={() => setCustomCommunities({})}
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
                    className="text-xs px-3 py-1.5 rounded-lg bg-sky-600 hover:bg-sky-500 text-white font-semibold flex items-center gap-1.5 shadow-md shadow-sky-600/30 transition-colors"
                  >
                    <Edit3 className="w-3.5 h-3.5" />
                    <span>Modificar Miembros</span>
                  </button>
                </div>
              )}
            </div>

            {/* 3 Large Vibrant Community Cards */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-5 my-auto">
              {/* Community 1: UpConta Socios */}
              <div className="p-6 rounded-2xl bg-gradient-to-b from-[#0f1d38] to-[#0a1224] border border-sky-500/40 shadow-xl flex flex-col justify-between relative group">
                <div>
                  <div className="flex items-center justify-between">
                    <h3 className="text-base font-bold text-white">UpConta Socios</h3>
                    <span className="text-[10px] px-2 py-0.5 rounded bg-sky-500/20 text-sky-300 font-semibold">
                      UpConnect Directo
                    </span>
                  </div>
                  <div className="text-4xl font-black text-sky-400 tabular-nums my-4">
                    {upcontaMembers}
                  </div>
                  <div className="text-xs text-slate-400 mb-2 font-medium">Miembros Registrados</div>
                </div>
                <p className="text-xs text-slate-300 leading-relaxed border-t border-slate-800 pt-3">
                  Red principal de socios estratégicos y contadores vinculados a la plataforma UpConnect.
                </p>
              </div>

              {/* Community 2: Franquicia Contadores VIP */}
              <div className="p-6 rounded-2xl bg-gradient-to-b from-[#241338] to-[#120a1f] border border-purple-500/40 shadow-xl flex flex-col justify-between relative group">
                <div>
                  <div className="flex items-center justify-between">
                    <h3 className="text-base font-bold text-white">Franquicia Contadores VIP</h3>
                    <span className="text-[10px] px-2 py-0.5 rounded bg-purple-500/20 text-purple-300 font-semibold">
                      Connectors Aliados
                    </span>
                  </div>
                  <div className="text-4xl font-black text-purple-400 tabular-nums my-4">
                    {franquiciaMembers}
                  </div>
                  <div className="text-xs text-slate-400 mb-2 font-medium">Miembros Registrados</div>
                </div>
                <p className="text-xs text-slate-300 leading-relaxed border-t border-slate-800 pt-3">
                  Grupo élite de contadores franquiciados y red externa de distribución comercial Connectors.
                </p>
              </div>

              {/* Community 3: Formación Comercial */}
              <div className="p-6 rounded-2xl bg-gradient-to-b from-[#0d2e26] to-[#071713] border border-emerald-500/40 shadow-xl flex flex-col justify-between relative group">
                <div>
                  <div className="flex items-center justify-between">
                    <h3 className="text-base font-bold text-white">Formación Comercial</h3>
                    <span className="text-[10px] px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 font-semibold">
                      Capacitación
                    </span>
                  </div>
                  <div className="text-4xl font-black text-emerald-400 tabular-nums my-4">
                    {formacionMembers}
                  </div>
                  <div className="text-xs text-slate-400 mb-2 font-medium">Miembros en Formación</div>
                </div>
                <p className="text-xs text-slate-300 leading-relaxed border-t border-slate-800 pt-3">
                  Programa activo de capacitación comercial, habilitación técnica y nuevo grupo de formación.
                </p>
              </div>
            </div>

            {/* Bottom official ecosystem banner */}
            <div className="p-3 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 text-center">
              Ecosistema oficial: <strong className="text-white">{upcontaMembers + franquiciaMembers} operadores registrados</strong> ({upcontaMembers} Upconnect / {franquiciaMembers} Connectors) y{' '}
              <strong className="text-sky-400">{upcontaMembers + franquiciaMembers + formacionMembers} miembros en comunidades oficiales</strong>.
              {isManuallyModified && (
                <span className="ml-2 text-amber-400 font-medium text-[11px]">(Valores personalizados manualmente)</span>
              )}
            </div>
          </div>
        );

      case 'socios':
        return (
          <div className="h-full flex flex-col justify-between p-6 sm:p-10 relative bg-[#0c1322]">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 border-b border-slate-800/80 pb-3">
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                    REPORTE DE FIRMAS · AUDITORÍA DE VENDEDORES & SOCIOS
                  </span>
                  <span className="text-slate-600">·</span>
                  <span className="text-xs text-emerald-400 font-semibold">{viewData.periodLabel}</span>
                </div>
                <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                  ¿Quién Vende Más?: Rendimiento Comercial y Cartera de Vendedores
                </h2>
              </div>

              <div className="flex items-center gap-3">
                <div className="px-3.5 py-1.5 rounded-xl bg-slate-900/90 border border-slate-700 text-right">
                  <div className="text-[9px] uppercase font-semibold text-slate-400">CARTERA ACTIVA</div>
                  <div className="text-sm font-bold text-white tabular-nums">{activeSocios.length} Vendedores</div>
                </div>
                <div className="px-3.5 py-1.5 rounded-xl bg-slate-900/90 border border-slate-700 text-right">
                  <div className="text-[9px] uppercase font-semibold text-slate-400">FACTURACIÓN CARTERA</div>
                  <div className="text-sm font-bold text-emerald-400 tabular-nums">
                    ${activeSocios.reduce((a, s) => a + s.totalSales, 0).toLocaleString('en-US', { minimumFractionDigits: 2 })}
                  </div>
                </div>
              </div>
            </div>

            {/* Table or Ranking View */}
            <div className="my-auto flex-1 overflow-hidden flex flex-col justify-center py-2">
              <div className="overflow-x-auto rounded-xl border border-slate-800 bg-slate-900/70 max-h-[360px]">
                <table className="w-full text-xs text-left">
                  <thead className="sticky top-0 bg-[#0d1627] text-slate-400 border-b border-slate-800 z-10">
                    <tr>
                      <th className="py-2.5 px-3 font-semibold text-center w-12">#</th>
                      <th className="py-2.5 px-3 font-semibold">Nombre Socio / Vendedor</th>
                      <th className="py-2.5 px-3 font-semibold">Rol / Canal</th>
                      <th className="py-2.5 px-3 font-semibold text-center">Operaciones</th>
                      <th className="py-2.5 px-3 font-semibold text-right">Monto Total ($)</th>
                      <th className="py-2.5 px-3 font-semibold text-right">Ticket Prom. ($)</th>
                      <th className="py-2.5 px-3 font-semibold">Plan Más Vendido</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/50 text-slate-300 tabular-nums">
                    {activeSocios.length === 0 ? (
                      <tr>
                        <td colSpan={7} className="py-8 text-center text-slate-500">
                          Sin vendedores registrados. Cargue el archivo de firmas o socios en el Dashboard.
                        </td>
                      </tr>
                    ) : (
                      activeSocios.slice(0, 10).map((s) => {
                        const avg = s.averageTicket ?? (s.operationsCount ? s.totalSales / s.operationsCount : s.totalSales);
                        const isTop3 = s.rank <= 3;
                        return (
                          <tr key={s.rank} className="hover:bg-slate-800/40 transition-colors">
                            <td className="py-2 px-3 text-center">
                              <span className={`inline-flex items-center justify-center w-5 h-5 rounded-full text-[10px] font-bold ${
                                s.rank === 1 ? 'bg-amber-500 text-slate-950 font-black' : isTop3 ? 'bg-slate-700 text-amber-300' : 'text-slate-500'
                              }`}>
                                {s.rank}
                              </span>
                            </td>
                            <td className="py-2 px-3 font-medium text-white flex items-center gap-1.5">
                              <span>{s.name}</span>
                              {s.note && (
                                <span className="px-1.5 py-0.2 rounded text-[9px] bg-sky-500/20 text-sky-300 font-semibold">
                                  {s.note}
                                </span>
                              )}
                            </td>
                            <td className="py-2 px-3 text-slate-400">
                              <span className="px-2 py-0.5 rounded text-[10px] bg-slate-800 border border-slate-700 text-slate-300">
                                {s.role || 'Distribuidor Connect'}
                              </span>
                            </td>
                            <td className="py-2 px-3 text-center font-semibold text-slate-200">
                              {s.operationsCount || 1}
                            </td>
                            <td className="py-2 px-3 text-right font-bold text-emerald-400">
                              ${s.totalSales.toFixed(2)}
                            </td>
                            <td className="py-2 px-3 text-right font-medium text-sky-400">
                              ${avg.toFixed(2)}
                            </td>
                            <td className="py-2 px-3 text-slate-300 max-w-[180px] truncate">
                              {s.topPlan || 'Firma Electrónica (1 año)'}
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="flex justify-between items-center text-[11px] text-slate-500 pt-2 border-t border-slate-800/60">
              <span>Auditoría de rendimiento comercial individual por vendedor con cálculo de ticket promedio real por firma.</span>
              <span>Pág. 0{pageNum}</span>
            </div>
          </div>
        );

      default:
        return null;
    }
  };

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
          <span className="text-xs uppercase tracking-wider font-semibold text-sky-400">
            Presentación Tipo 1 · Reporte Upconnect & Connectors
          </span>
          <span className="text-slate-600">|</span>
          <span className="text-xs px-2 py-0.5 rounded-full bg-sky-950/80 text-sky-300 border border-sky-800/60 font-medium">
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
                    ? 'bg-sky-500 text-white font-bold'
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
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded bg-amber-600 hover:bg-amber-500 disabled:opacity-50 text-white shadow-sm transition-colors"
            title="Descargar presentación editable en formato PowerPoint (.pptx)"
          >
            {isExportingPPTX ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <FileDown className="w-4 h-4" />
            )}
            <span>Descargar PPTX</span>
          </button>

          <button
            onClick={handlePrint}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded bg-slate-800 hover:bg-slate-700 text-slate-200 transition-colors"
            title="Exportar a PDF o Imprimir diapositivas"
          >
            <Printer className="w-4 h-4" />
            <span className="hidden sm:inline">PDF / Imprimir</span>
          </button>
        </div>
      </div>

      {/* 1. SCREEN VIEW: Active slide in 16:9 container */}
      <div className="print:hidden flex-1 flex items-center justify-center p-2 sm:p-6 bg-[#070b14] overflow-auto">
        <div className="w-full max-w-[1240px] aspect-[16/9] min-h-[580px] bg-[#0c1322] border border-slate-800/80 rounded-2xl shadow-2xl relative overflow-hidden flex flex-col justify-between slide-container">
          {renderSlideContent(currentSlideDef.id, currentSlide, false)}
        </div>
      </div>

      {/* 2. PRINT-ONLY VIEW: Full multi-page landscape print deck */}
      <div className="hidden print:block print:w-full space-y-0">
        {activeSlides.map((slide, idx) => (
          <div key={slide.id} className="print-slide-page">
            {renderSlideContent(slide.id, idx + 1, true)}
          </div>
        ))}
      </div>

      {/* MODAL: Modificar Miembros de Comunidades Manualmente */}
      {isEditingCommunities && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm no-print">
          <div className="w-full max-w-md bg-[#0e1626] border border-sky-500/40 rounded-2xl p-6 shadow-2xl space-y-5">
            <div className="flex items-center justify-between pb-3 border-b border-slate-800">
              <div className="flex items-center gap-2">
                <Edit3 className="w-4 h-4 text-sky-400" />
                <h3 className="text-base font-bold text-white">Modificar Miembros de Comunidades</h3>
              </div>
              <button
                onClick={() => setIsEditingCommunities(false)}
                className="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <p className="text-xs text-slate-300">
              Ajuste manualmente el número de miembros para las comunidades. Los valores se sincronizarán en la diapositiva, en el resumen y en las exportaciones a PDF y PowerPoint.
            </p>

            <div className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-slate-300 mb-1">
                  UpConta Socios (UpConnect Directo)
                </label>
                <div className="text-[11px] text-slate-500 mb-1.5">
                  Socios calculados en Excel: <strong className="text-sky-400">{autoUpSocios}</strong>
                </div>
                <input
                  type="number"
                  min="0"
                  value={tempUpconta}
                  onChange={(e) => setTempUpconta(parseInt(e.target.value, 10) || 0)}
                  className="w-full px-3 py-2 text-sm bg-slate-900 border border-slate-700 rounded-xl text-white focus:outline-none focus:border-sky-500"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-300 mb-1">
                  Franquicia Contadores VIP (Connectors Aliados)
                </label>
                <div className="text-[11px] text-slate-500 mb-1.5">
                  Socios calculados en Excel: <strong className="text-amber-400">{autoCoSocios}</strong>
                </div>
                <input
                  type="number"
                  min="0"
                  value={tempFranquicia}
                  onChange={(e) => setTempFranquicia(parseInt(e.target.value, 10) || 0)}
                  className="w-full px-3 py-2 text-sm bg-slate-900 border border-slate-700 rounded-xl text-white focus:outline-none focus:border-sky-500"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-300 mb-1">
                  Formación Comercial Franquiciados
                </label>
                <div className="text-[11px] text-slate-500 mb-1.5">
                  Estimado automático: <strong className="text-emerald-400">{autoFormacion}</strong>
                </div>
                <input
                  type="number"
                  min="0"
                  value={tempFormacion}
                  onChange={(e) => setTempFormacion(parseInt(e.target.value, 10) || 0)}
                  className="w-full px-3 py-2 text-sm bg-slate-900 border border-slate-700 rounded-xl text-white focus:outline-none focus:border-sky-500"
                />
              </div>
            </div>

            <div className="flex items-center justify-between pt-3 border-t border-slate-800 gap-2">
              <button
                onClick={() => {
                  setCustomCommunities({});
                  setIsEditingCommunities(false);
                }}
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
                  onClick={() => {
                    setCustomCommunities({
                      upcontaSocios: tempUpconta,
                      franquiciaVIP: tempFranquicia,
                      formacionComercial: tempFormacion,
                    });
                    setIsEditingCommunities(false);
                  }}
                  className="text-xs px-4 py-2 rounded-xl bg-sky-600 hover:bg-sky-500 text-white font-bold transition-colors shadow-lg shadow-sky-600/30 flex items-center gap-1.5"
                >
                  <Check className="w-3.5 h-3.5" />
                  <span>Guardar Cambios</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
