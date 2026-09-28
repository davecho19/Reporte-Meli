import React, { useState, useEffect } from 'react';
import { GlobalDataset } from '../types';
import { computeReportView } from '../utils/reportFilters';
import { exportPresentationType2ToPPTX } from '../services/pptxExportService';
import { 
  ChevronLeft, 
  ChevronRight, 
  Maximize2, 
  Minimize2, 
  Printer, 
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
  Search
} from 'lucide-react';

interface Props {
  dataset: GlobalDataset;
  onBackToDashboard?: () => void;
}

export const PresentationType2: React.FC<Props> = ({ dataset, onBackToDashboard }) => {
  const [currentSlide, setCurrentSlide] = useState<number>(1);
  const [isFullscreen, setIsFullscreen] = useState<boolean>(false);
  const [sociosViewMode, setSociosViewMode] = useState<'table' | 'pareto'>('table');
  const [socioSearch, setSocioSearch] = useState<string>('');
  const totalSlides = 8;

  const viewData = computeReportView(dataset);

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
  }, [isFullscreen]);

  const toggleFullscreen = () => {
    setIsFullscreen(!isFullscreen);
  };

  const [isExportingPPTX, setIsExportingPPTX] = useState<boolean>(false);

  const handlePrint = () => {
    window.print();
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

  // Monthly values for audit (derived dynamically from dataset)
  const activeMonthlyMetrics = (viewData.monthlyMetrics && viewData.monthlyMetrics.length > 0)
    ? viewData.monthlyMetrics
    : dataset.monthlyMetrics;

  const auditMonthly = activeMonthlyMetrics.length > 0
    ? activeMonthlyMetrics.map((m, _, arr) => {
        const totalAmount = m.upconnectSales + m.connectorsSales;
        const maxVal = Math.max(...arr.map((x) => x.upconnectSales + x.connectorsSales));
        return {
          period: `${m.month.charAt(0) + m.month.slice(1).toLowerCase()} 2026`,
          amount: totalAmount,
          color: totalAmount === maxVal && totalAmount > 0 ? 'bg-emerald-500' : 'bg-blue-600',
          isRecord: totalAmount === maxVal && totalAmount > 0,
        };
      })
    : [];

  return (
    <div className={`flex flex-col ${isFullscreen ? 'fixed inset-0 z-50 bg-[#070b14]' : 'w-full'}`}>
      {/* Top Controls */}
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
            Presentación Tipo 2 · Auditoría Comercial & Gestión de Socios UpConta
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
          {/* Jump buttons */}
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
            title="Exportar a PDF o Imprimir"
          >
            <Printer className="w-4 h-4" />
            <span className="hidden sm:inline">PDF / Imprimir</span>
          </button>
        </div>
      </div>

      {/* Main Slide Stage (16:9) */}
      <div className="flex-1 flex items-center justify-center p-2 sm:p-6 bg-[#070b14] overflow-auto">
        <div className="w-full max-w-[1240px] aspect-[16/9] min-h-[580px] bg-[#0c1322] border border-slate-800/80 rounded-2xl shadow-2xl relative overflow-hidden flex flex-col justify-between slide-container">
          
          {/* SLIDE 1: Cover */}
          {currentSlide === 1 && (
            <div className="h-full flex flex-col justify-center items-center text-center p-8 sm:p-16 relative bg-gradient-to-br from-[#0a1120] via-[#09101d] to-[#040812]">
              <div className="absolute w-[500px] h-[500px] rounded-full bg-emerald-500/10 blur-[130px] pointer-events-none -top-24 -right-24" />
              <div className="absolute w-[450px] h-[450px] rounded-full bg-blue-500/10 blur-[110px] pointer-events-none -bottom-24 -left-24" />

              <div className="z-10 max-w-4xl space-y-6">
                <div className="flex items-center justify-center gap-3">
                  <span className="w-10 h-10 rounded-xl bg-blue-600 text-white font-black flex items-center justify-center text-base shadow-lg shadow-blue-500/30">
                    Up
                  </span>
                  <span className="text-xs uppercase tracking-widest font-bold text-slate-400">
                    UPCONNECT · UPCONTA ERP
                  </span>
                  <span className="text-slate-600">·</span>
                  <span className="text-xs font-semibold text-sky-400">
                    Cierre Oficial 2026 · Actualizado
                  </span>
                </div>

                <div className="text-xs font-bold uppercase tracking-wider text-sky-400 pt-2">
                  AUDITORÍA COMERCIAL Y GESTIÓN DE SOCIOS · {viewData.periodLabel.toUpperCase()}
                </div>

                <h1 className="text-3xl sm:text-5xl lg:text-6xl font-black tracking-tight text-white leading-tight">
                  <span className="text-transparent bg-clip-text bg-gradient-to-r from-sky-400 to-emerald-400">
                    {viewData.titleReport2}
                  </span>
                </h1>

                <div className="pt-8 grid grid-cols-1 sm:grid-cols-3 gap-6 max-w-2xl mx-auto">
                  <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800">
                    <div className="text-[11px] uppercase font-bold text-slate-400">FACTURACIÓN ACUMULADA</div>
                    <div className="text-2xl font-black text-white tabular-nums mt-1">${netSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD</div>
                  </div>
                  <div className="p-4 rounded-xl bg-slate-900/80 border border-emerald-500/30">
                    <div className="text-[11px] uppercase font-bold text-emerald-400">RÉCORD / CIERRE</div>
                    <div className="text-2xl font-black text-emerald-400 tabular-nums mt-1">${netSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD</div>
                  </div>
                  <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800">
                    <div className="text-[11px] uppercase font-bold text-slate-400">TRANSACCIONES AUDITADAS</div>
                    <div className="text-2xl font-black text-white tabular-nums mt-1">{viewData.totalCount || activeSocios.reduce((a, s) => a + (s.operationsCount || 0), 0) || 0} Operaciones</div>
                  </div>
                </div>
              </div>

              <div className="absolute bottom-4 left-8 right-8 flex justify-between items-center text-[11px] text-slate-500 border-t border-slate-800/60 pt-3">
                <span>Auditoría de Socios Franquiciados y Contabilidad ERP · {viewData.subtitleDate}</span>
                <span>Pág. 01</span>
              </div>
            </div>
          )}

          {/* SLIDE 2: Balance General de Facturación 2026 */}
          {currentSlide === 2 && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  RESUMEN GERENCIAL CONSOLIDADO
                </span>
                <h2 className="text-2xl font-bold text-white tracking-tight mt-1">
                  Balance General de Facturación 2026
                </h2>
              </div>

              {/* 4 Clean Metric Cards */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6 my-auto">
                {/* Net Sales */}
                <div className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between">
                  <div className="text-xs font-bold text-slate-400 uppercase">VENTAS NETAS (SIN IVA)</div>
                  <div className="my-3">
                    <div className="text-3xl font-black text-white tabular-nums">
                      ${netSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                    </div>
                    <div className="text-xs font-semibold text-emerald-400 mt-1">
                      +26.0% vs corte previo ($4,265)
                    </div>
                  </div>
                  <div className="text-[10px] text-slate-500">Liquidación auditada</div>
                </div>

                {/* Gross Sales */}
                <div className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between">
                  <div className="text-xs font-bold text-slate-400 uppercase">TOTAL FACTURADO (CON 15% IVA)</div>
                  <div className="my-3">
                    <div className="text-3xl font-black text-sky-400 tabular-nums">
                      ${grossSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                    </div>
                    <div className="text-xs font-medium text-slate-300 mt-1">
                      IVA recaudado: ${ivaAmount.toFixed(2)} USD
                    </div>
                  </div>
                  <div className="text-[10px] text-slate-500">Tributación 15% vigente</div>
                </div>

                {/* Operations */}
                <div className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between">
                  <div className="text-xs font-bold text-slate-400 uppercase">OPERACIONES TOTALES</div>
                  <div className="my-3">
                    <div className="text-3xl font-black text-white tabular-nums">
                      {activeSocios.reduce((a, s) => a + (s.operationsCount || 0), 0) || viewData.transactions.length || 0}
                    </div>
                    <div className="text-xs font-medium text-slate-300 mt-1">
                      {activeSocios.length} socios registrados en cartera
                    </div>
                  </div>
                  <div className="text-[10px] text-slate-500">Transacciones comerciales</div>
                </div>

                {/* Monthly Record */}
                <div className="p-6 rounded-2xl bg-slate-900/90 border border-emerald-500/40 shadow-xl flex flex-col justify-between">
                  <div className="text-xs font-bold text-emerald-400 uppercase">VOLUMEN TOTAL ({viewData.periodLabel.toUpperCase()})</div>
                  <div className="my-3">
                    <div className="text-3xl font-black text-emerald-400 tabular-nums">
                      ${netSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </div>
                    <div className="text-xs font-medium text-emerald-300 mt-1">
                      Cierre auditado de ventas
                    </div>
                  </div>
                  <div className="text-[10px] text-slate-500">Liquidación verificada</div>
                </div>
              </div>

              {/* Bottom banner */}
              <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 flex items-center justify-between text-xs text-slate-300">
                <div className="flex items-center gap-2">
                  <CheckCircle2 className="w-4 h-4 text-sky-400" />
                  <span>Base consolidada: Liquidación de comisiones de socios verificada e integrada al cierre de Septiembre.</span>
                </div>
                <span className="text-slate-500 font-semibold">Canal Socios UpConnect</span>
              </div>
            </div>
          )}

          {/* SLIDE 3: Evolución de Ventas Mes a Mes (Sin IVA) */}
          {currentSlide === 3 && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  HISTÓRICO AUDITADO
                </span>
                <h2 className="text-2xl font-bold text-white tracking-tight mt-1">
                  Evolución de Ventas Mes a Mes (Sin IVA)
                </h2>
              </div>

              {/* Horizontal Bar Chart matching Slide 3 of PDF 2 */}
              <div className="space-y-3.5 my-auto max-w-4xl mx-auto w-full">
                {auditMonthly.length === 0 ? (
                  <div className="py-12 text-center text-slate-500 border border-dashed border-slate-800 rounded-xl max-w-2xl mx-auto px-6">
                    <p className="font-semibold text-slate-400">Sin histórico de ventas mes a mes</p>
                    <p className="text-xs text-slate-500 mt-1">
                      Cargue el archivo de firmas o socios en el Dashboard para visualizar el histórico auditado.
                    </p>
                  </div>
                ) : (
                  auditMonthly.map((m) => {
                    const maxAudit = Math.max(100, ...auditMonthly.map((x) => x.amount));
                    const pct = Math.max(6, (m.amount / maxAudit) * 100);

                    return (
                      <div key={m.period} className="flex items-center gap-4">
                        <div className="w-36 sm:w-44 text-right text-xs font-medium text-slate-300 shrink-0">
                          {m.period}
                        </div>

                        <div className="flex-1 bg-slate-900/80 rounded-lg h-7 overflow-hidden relative flex items-center">
                          <div
                            style={{ width: `${pct}%` }}
                            className={`h-full ${m.isRecord ? 'bg-emerald-500' : 'bg-blue-600'} rounded-lg transition-all duration-500`}
                          />
                        </div>

                        <div className="w-28 text-left text-xs font-bold tabular-nums shrink-0">
                          <span className={m.isRecord ? 'text-emerald-400 font-black' : 'text-slate-200'}>
                            ${m.amount.toFixed(2)} USD
                          </span>
                        </div>
                      </div>
                    );
                  })
                )}
              </div>

              {/* Bottom footer text */}
              <div className="p-3.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-400 flex items-center justify-between">
                <div>
                  <strong className="text-sky-400">Tendencia:</strong> Crecimiento sostenido desde Abril (+753% de expansión en 5 meses).
                </div>
                <span>Abril = Base Inicial de Comparativa</span>
              </div>
            </div>
          )}

          {/* SLIDE 4: Ranking General de Socios & Nueva Tabla Detallada */}
          {currentSlide === 4 && (
            <div className="h-full flex flex-col justify-between p-6 sm:p-10 relative bg-[#0c1322]">
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 border-b border-slate-800/80 pb-3">
                <div>
                  <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                    DESEMPEÑO INDIVIDUAL DE SOCIOS · AUDITORÍA CONSOLIDADA
                  </span>
                  <h2 className="text-xl sm:text-2xl font-bold text-white tracking-tight mt-0.5">
                    ¿Quién Vende Más?: Cartera y Rendimiento de Socios
                  </h2>
                </div>

                <div className="flex flex-wrap items-center gap-3">
                  {/* View Mode Toggle: Tabla Detallada vs Gráfico Pareto */}
                  <div className="flex items-center p-1 bg-slate-900 border border-slate-700/80 rounded-xl">
                    <button
                      onClick={() => setSociosViewMode('table')}
                      className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-lg transition-all ${
                        sociosViewMode === 'table'
                          ? 'bg-emerald-600 text-white shadow-md'
                          : 'text-slate-400 hover:text-white'
                      }`}
                    >
                      <Table className="w-3.5 h-3.5" />
                      <span>Tabla Detallada</span>
                    </button>
                    <button
                      onClick={() => setSociosViewMode('pareto')}
                      className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-lg transition-all ${
                        sociosViewMode === 'pareto'
                          ? 'bg-blue-600 text-white shadow-md'
                          : 'text-slate-400 hover:text-white'
                      }`}
                    >
                      <BarChart3 className="w-3.5 h-3.5" />
                      <span>Gráfico Pareto</span>
                    </button>
                  </div>

                  {/* Search box for table */}
                  {sociosViewMode === 'table' && activeSocios.length > 5 && (
                    <div className="relative">
                      <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                      <input
                        type="text"
                        placeholder="Buscar socio, rol o plan..."
                        value={socioSearch}
                        onChange={(e) => setSocioSearch(e.target.value)}
                        className="pl-8 pr-2.5 py-1 text-xs bg-slate-900 border border-slate-700 rounded-lg text-white placeholder-slate-500 focus:outline-none focus:border-emerald-500 w-44"
                      />
                    </div>
                  )}

                  <div className="text-xs font-semibold text-slate-300 bg-slate-900/90 px-3 py-1.5 rounded-lg border border-slate-800">
                    Cartera Total: <strong className="text-emerald-400 font-bold">${activeSocios.reduce((a, s) => a + s.totalSales, 0).toLocaleString('en-US', { minimumFractionDigits: 2 })} USD</strong> ({activeSocios.length} Socios)
                  </div>
                </div>
              </div>

              {/* Empty state */}
              {activeSocios.length === 0 ? (
                <div className="my-auto py-12 text-center text-slate-500 border border-dashed border-slate-800 rounded-xl max-w-2xl mx-auto px-6">
                  <Users className="w-12 h-12 text-slate-700 mx-auto mb-2" />
                  <p className="font-semibold text-slate-400">Sin cartera de socios registrada</p>
                  <p className="text-xs text-slate-500 mt-1">
                    Cargue el archivo de socios en el Dashboard para visualizar automáticamente la tabla detallada (ventas, ticket promedio, planes más vendidos y roles).
                  </p>
                </div>
              ) : sociosViewMode === 'table' ? (
                /* TABLA DETALLADA DE SOCIOS */
                <div className="my-auto flex-1 flex flex-col justify-center py-2">
                  <div className="border border-slate-800 rounded-xl overflow-hidden bg-slate-900/70 shadow-2xl">
                    <div className="max-h-[380px] overflow-y-auto">
                      <table className="w-full text-xs text-left">
                        <thead className="bg-[#0b1220] text-slate-300 uppercase tracking-wider font-bold border-b border-slate-800 sticky top-0 z-10 text-[11px]">
                          <tr>
                            <th className="py-2.5 px-3 text-center w-12">#</th>
                            <th className="py-2.5 px-3">Nombre del Socio</th>
                            <th className="py-2.5 px-3">Rol</th>
                            <th className="py-2.5 px-3 text-center">Ventas Realizadas</th>
                            <th className="py-2.5 px-3 text-right">Monto Total de Venta</th>
                            <th className="py-2.5 px-3 text-right">Ticket Promedio</th>
                            <th className="py-2.5 px-3">Plan Más Vendido</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-800/70 text-slate-300">
                          {activeSocios
                            .filter((s) => 
                              !socioSearch || 
                              s.name.toLowerCase().includes(socioSearch.toLowerCase()) ||
                              (s.role && s.role.toLowerCase().includes(socioSearch.toLowerCase())) ||
                              (s.topPlan && s.topPlan.toLowerCase().includes(socioSearch.toLowerCase()))
                            )
                            .map((s) => {
                              const avgTicket = s.averageTicket ?? (s.operationsCount ? (s.totalSales / s.operationsCount) : s.totalSales);
                              const isTop3 = s.rank <= 3;
                              return (
                                <tr key={s.rank} className="hover:bg-slate-800/40 transition-colors">
                                  {/* Rank */}
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

                                  {/* Nombre del Socio */}
                                  <td className="py-2.5 px-3">
                                    <div className="font-bold text-white flex items-center gap-1.5">
                                      <span>{s.name}</span>
                                      {s.rank === 1 && (
                                        <span className="text-[10px] px-1.5 py-0.2 rounded bg-amber-500/20 text-amber-300 border border-amber-500/30 font-semibold">
                                          Líder
                                        </span>
                                      )}
                                    </div>
                                  </td>

                                  {/* Rol */}
                                  <td className="py-2.5 px-3">
                                    <span className={`px-2 py-0.5 rounded text-[10px] font-semibold border ${
                                      (s.role || '').toLowerCase().includes('connect')
                                        ? 'bg-blue-500/10 text-sky-300 border-blue-500/20'
                                        : (s.role || '').toLowerCase().includes('vip') || (s.role || '').toLowerCase().includes('franquicia')
                                        ? 'bg-purple-500/10 text-purple-300 border-purple-500/20'
                                        : 'bg-emerald-500/10 text-emerald-300 border-emerald-500/20'
                                    }`}>
                                      {s.role || 'Distribuidor Connect'}
                                    </span>
                                  </td>

                                  {/* Cantidades Ventas Realizadas */}
                                  <td className="py-2.5 px-3 text-center font-bold text-slate-200 tabular-nums">
                                    {s.operationsCount || 1}
                                  </td>

                                  {/* Monto Total de Venta */}
                                  <td className="py-2.5 px-3 text-right font-black text-emerald-400 tabular-nums">
                                    ${s.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD
                                  </td>

                                  {/* Ticket Promedio */}
                                  <td className="py-2.5 px-3 text-right font-bold text-sky-400 tabular-nums">
                                    ${avgTicket.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD
                                  </td>

                                  {/* Plan Más Vendido */}
                                  <td className="py-2.5 px-3">
                                    <span className="px-2 py-0.5 rounded text-[11px] bg-slate-800 text-slate-300 border border-slate-700 font-medium inline-block max-w-[220px] truncate" title={s.topPlan}>
                                      {s.topPlan || 'Firma Electrónica (1 año)'}
                                    </span>
                                  </td>
                                </tr>
                              );
                            })}
                        </tbody>
                        <tfoot className="bg-[#080d19] border-t-2 border-slate-700 text-slate-200 font-bold text-xs sticky bottom-0">
                          <tr>
                            <td colSpan={3} className="py-2.5 px-3 text-left">
                              TOTAL GENERAL ({activeSocios.length} SOCIOS)
                            </td>
                            <td className="py-2.5 px-3 text-center text-white tabular-nums">
                              {activeSocios.reduce((a, s) => a + (s.operationsCount || 0), 0)} ventas
                            </td>
                            <td className="py-2.5 px-3 text-right text-emerald-400 font-black tabular-nums">
                              ${activeSocios.reduce((a, s) => a + s.totalSales, 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD
                            </td>
                            <td className="py-2.5 px-3 text-right text-sky-400 font-black tabular-nums">
                              ${(
                                activeSocios.reduce((a, s) => a + s.totalSales, 0) / 
                                (activeSocios.reduce((a, s) => a + (s.operationsCount || 0), 0) || 1)
                              ).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD
                            </td>
                            <td className="py-2.5 px-3 text-slate-400 font-normal text-[11px]">
                              Ticket Promedio General
                            </td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                </div>
              ) : (
                /* GRÁFICO PARETO TOP 1-10 vs 11-20 */
                <div className="grid grid-cols-1 md:grid-cols-2 gap-8 my-auto">
                  {/* TOP 1 - 10 */}
                  <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800">
                    <div className="text-xs font-bold text-sky-400 uppercase tracking-wider mb-2">
                      TOP 1 – 10: MAYOR VOLUMEN COLOCADO
                    </div>
                    <div className="space-y-1.5">
                      {activeSocios.slice(0, 10).map((s) => {
                        const maxVal = activeSocios[0]?.totalSales || 1;
                        const barWidth = Math.max(5, (s.totalSales / maxVal) * 100);
                        return (
                          <div key={s.rank} className="flex items-center gap-2 text-xs">
                            <span className="w-5 text-slate-500 font-bold tabular-nums">{s.rank}.</span>
                            <span className="w-36 truncate font-medium text-slate-200" title={s.name}>
                              {s.name}
                            </span>
                            <div className="flex-1 bg-slate-800/60 rounded h-4 overflow-hidden">
                              <div
                                style={{ width: `${barWidth}%` }}
                                className={`h-full ${s.rank === 1 ? 'bg-sky-400' : 'bg-blue-600'} rounded`}
                              />
                            </div>
                            <span className="w-20 text-right font-bold text-white tabular-nums">
                              ${s.totalSales.toFixed(2)}
                            </span>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* TOP 11 - 20 */}
                  <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800">
                    <div className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">
                      TOP 11 – 20: PENETRACIÓN Y RED BASE
                    </div>
                    <div className="space-y-1.5">
                      {activeSocios.slice(10, 20).map((s) => {
                        const maxSub = activeSocios[10]?.totalSales || 1;
                        const barWidth = Math.max(5, (s.totalSales / maxSub) * 100);
                        return (
                          <div key={s.rank} className="flex items-center gap-2 text-xs">
                            <span className="w-5 text-slate-500 font-bold tabular-nums">{s.rank}.</span>
                            <span className="w-36 truncate font-medium text-slate-200" title={s.name}>
                              {s.name}
                            </span>
                            <div className="flex-1 bg-slate-800/60 rounded h-4 overflow-hidden">
                              <div
                                style={{ width: `${barWidth}%` }}
                                className="h-full bg-slate-500 rounded"
                              />
                            </div>
                            <span className="w-20 text-right font-bold text-slate-300 tabular-nums">
                              ${s.totalSales.toFixed(2)}
                            </span>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                </div>
              )}

              {/* Callout Footer */}
              <div className="p-3 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex flex-col sm:flex-row items-center justify-between gap-2">
                <div>
                  <strong className="text-sky-400">Concentración Pareto:</strong> El Top 5 representa el{' '}
                  <strong className="text-white">
                    {activeSocios.length > 0
                      ? ((activeSocios.slice(0, 5).reduce((a, b) => a + b.totalSales, 0) / (activeSocios.reduce((a, b) => a + b.totalSales, 0) || 1)) * 100).toFixed(1)
                      : '0.0'}%
                  </strong> del total facturado por socios.
                </div>
                {activeSocios.length > 0 && (
                  <div className="text-amber-400 font-bold">
                    Líder en Ventas: {activeSocios[0]?.name} (${activeSocios[0]?.totalSales.toFixed(2)} USD · {activeSocios[0]?.role || 'Distribuidor Connect'})
                  </div>
                )}
              </div>
            </div>
          )}

          {/* SLIDE 5: Foco de Rendimiento Mensual: Septiembre Desglose Semana a Semana */}
          {currentSlide === 5 && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  FOCO DE RENDIMIENTO MENSUAL
                </span>
                <h2 className="text-2xl font-bold text-white tracking-tight mt-1">
                  Septiembre 2026: Desglose Semana a Semana
                </h2>
              </div>

              {/* 4 Week Cards */}
              {dataset.weeklyBreakdownType2.length === 0 ? (
                <div className="my-auto py-12 text-center text-slate-500 border border-dashed border-slate-800 rounded-xl max-w-2xl mx-auto px-6">
                  <p className="font-semibold text-slate-400">Sin desglose semanal registrado</p>
                  <p className="text-xs text-slate-500 mt-1">
                    Cargue el archivo de firmas o socios en el Dashboard para visualizar el rendimiento semana a semana.
                  </p>
                </div>
              ) : (
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6 my-auto">
                  {dataset.weeklyBreakdownType2.map((wb, idx) => {
                    const isWeek4 = idx === 3;
                    return (
                      <div
                        key={wb.weekName}
                        className={`p-6 rounded-2xl bg-slate-900/90 border flex flex-col justify-between shadow-xl ${
                          isWeek4 ? 'border-emerald-500/60 shadow-emerald-950/40' : 'border-slate-800'
                        }`}
                      >
                        <div>
                          <div className="text-[11px] font-bold text-slate-400 uppercase">
                            {wb.weekName} ({wb.dateRange})
                          </div>
                          <div className={`text-3xl font-black tabular-nums my-3 ${isWeek4 ? 'text-emerald-400' : 'text-white'}`}>
                            ${wb.amount.toFixed(2)}
                          </div>
                          <div className="text-xs font-semibold text-slate-300 mb-3">
                            {wb.percentage}% del mes
                          </div>
                        </div>

                        <p className="text-xs text-slate-400 border-t border-slate-800/80 pt-3 leading-relaxed">
                          {wb.description}
                        </p>
                      </div>
                    );
                  })}
                </div>
              )}

              {/* Bottom black bar */}
              <div className="p-4 rounded-xl bg-slate-950 border border-slate-800 flex items-center justify-between text-xs">
                <span className="text-slate-400 font-semibold">
                  Total Auditado ({dataset.weeklyBreakdownType2.reduce((acc, w) => acc + (w.operationsCount || 0), 0) || viewData.totalCount} Operaciones Verificadas):
                </span>
                <span className="text-xl font-black text-sky-400 tabular-nums">
                  ${netSales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} USD
                </span>
              </div>
            </div>
          )}

          {/* SLIDE 6: Ventas por Categoría de Solución */}
          {currentSlide === 6 && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  ESTRUCTURA DE PORTAFOLIO
                </span>
                <h2 className="text-2xl font-bold text-white tracking-tight mt-1">
                  Ventas por Categoría de Solución
                </h2>
              </div>

              {/* 3 Large Solution Cards */}
              {dataset.solutionCategories.length === 0 ? (
                <div className="my-auto py-12 text-center text-slate-500 border border-dashed border-slate-800 rounded-xl max-w-2xl mx-auto px-6">
                  <p className="font-semibold text-slate-400">Sin categorías de solución registradas</p>
                  <p className="text-xs text-slate-500 mt-1">
                    Cargue el archivo de firmas o socios en el Dashboard para desglosar las ventas por solución o plan.
                  </p>
                </div>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 my-auto">
                  {dataset.solutionCategories.map((cat) => (
                    <div
                      key={cat.id}
                      className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between"
                    >
                      <div>
                        <span className="text-[11px] font-bold text-sky-400 uppercase tracking-wider">
                          {cat.badge}
                        </span>
                        <h3 className="text-lg font-bold text-white mt-1 mb-3">{cat.title}</h3>

                        <div className="text-3xl font-black text-sky-400 tabular-nums">
                          ${cat.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                        </div>
                        <div className="text-xs font-semibold text-slate-300 mt-1 mb-4">
                          {cat.percentage}% del total facturado
                        </div>
                      </div>

                      <p className="text-xs text-slate-400 border-t border-slate-800 pt-3 leading-relaxed">
                        {cat.plansDescription}
                      </p>
                    </div>
                  ))}
                </div>
              )}

              {/* Bottom callout */}
              <div className="p-3.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-sky-400">Estrategia de Socios:</strong> Foco en tickets medianos y altos (ERP + Planes Contador representan el 88.5% del ingreso).
                </div>
                <span className="text-emerald-400 font-bold">Alta Retención</span>
              </div>
            </div>
          )}

          {/* SLIDE 7: Canal de Distribución - Embudo de Socios */}
          {currentSlide === 7 && (
            <div className="h-full flex flex-col justify-between p-8 sm:p-12 relative bg-[#0c1322]">
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  CANAL DE DISTRIBUCIÓN
                </span>
                <h2 className="text-2xl font-bold text-white tracking-tight mt-1">
                  Embudo de Socios y Comunidades Activas
                </h2>
              </div>

              {/* 3 Clean Funnel Columns */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 my-auto">
                {/* Stage 1 */}
                <div className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between text-center">
                  <div>
                    <div className="w-12 h-12 rounded-xl bg-blue-600/20 text-sky-400 font-bold flex items-center justify-center mx-auto mb-3">
                      Up
                    </div>
                    <h3 className="text-lg font-bold text-white">UpConta Socios</h3>
                    <div className="text-4xl font-black text-sky-400 tabular-nums my-3">
                      {activeSocios.length > 0 ? (activeSocios.length > 118 ? activeSocios.length : 118) : 0}
                    </div>
                    <div className="text-xs font-semibold text-slate-400 mb-2">Miembros Registrados</div>
                  </div>
                  <p className="text-xs text-slate-400 border-t border-slate-800 pt-3">
                    Red general de difusión, novedades y leads de webinars.
                  </p>
                </div>

                {/* Stage 2 (Hero highlighted) */}
                <div className="p-6 rounded-2xl bg-slate-900/95 border-2 border-sky-500 shadow-2xl shadow-sky-950/50 flex flex-col justify-between text-center relative">
                  <div>
                    <div className="w-12 h-12 rounded-xl bg-purple-600/20 text-purple-400 font-bold flex items-center justify-center mx-auto mb-3">
                      👑
                    </div>
                    <h3 className="text-lg font-bold text-white">Franquicia Contadores VIP</h3>
                    <div className="text-4xl font-black text-white tabular-nums my-3">
                      {activeSocios.length > 0 ? (activeSocios.length > 121 ? activeSocios.length : 121) : 0}
                    </div>
                    <div className="text-xs font-semibold text-purple-400 mb-2">Miembros Estratégicos</div>
                  </div>
                  <p className="text-xs text-slate-300 border-t border-slate-800 pt-3">
                    Contadores activos con potencial de distribución masiva.
                  </p>
                </div>

                {/* Stage 3 */}
                <div className="p-6 rounded-2xl bg-slate-900/90 border border-slate-800 shadow-xl flex flex-col justify-between text-center">
                  <div>
                    <div className="w-12 h-12 rounded-xl bg-emerald-600/20 text-emerald-400 font-bold flex items-center justify-center mx-auto mb-3">
                      🎓
                    </div>
                    <h3 className="text-lg font-bold text-white">Capacitación Franquicia</h3>
                    <div className="text-4xl font-black text-emerald-400 tabular-nums my-3">
                      {activeSocios.length > 0 ? (activeSocios.length > 51 ? activeSocios.length : 51) : 0}
                    </div>
                    <div className="text-xs font-semibold text-slate-400 mb-2">Miembros en Inducción</div>
                  </div>
                  <p className="text-xs text-slate-400 border-t border-slate-800 pt-3">
                    Núcleo operativo en formación técnica y comercial directa.
                  </p>
                </div>
              </div>

              {/* Bottom callout */}
              <div className="p-3.5 rounded-xl bg-slate-900/90 border border-slate-800 text-xs text-slate-300 flex items-center justify-between">
                <div>
                  <strong className="text-sky-400">Ratio de Activación:</strong> {activeSocios.length} socios con ventas registradas sobre {activeSocios.length > 0 ? 127 : 0} miembros en comunidad ({activeSocios.length > 0 ? ((activeSocios.length / 127) * 100).toFixed(1) : '0.0'}% de conversión).
                </div>
                <span className="text-emerald-400 font-bold">Potencial Alto de Escalamiento</span>
              </div>
            </div>
          )}

          {/* SLIDE 8: Siguiente Fase: Cruce Comercial */}
          {currentSlide === 8 && (
            <div className="h-full flex flex-col justify-center items-center text-center p-8 sm:p-16 relative bg-gradient-to-br from-[#0c1322] via-[#080d19] to-[#040812]">
              <div className="z-10 max-w-2xl space-y-6">
                <div className="w-12 h-12 rounded-2xl bg-sky-500 text-white flex items-center justify-center mx-auto shadow-lg shadow-sky-500/30">
                  <Activity className="w-6 h-6" />
                </div>

                <div className="text-xs font-bold uppercase tracking-wider text-sky-400">
                  SIGUIENTE FASE: CRUCE COMERCIAL
                </div>

                <h2 className="text-3xl sm:text-4xl font-black text-white tracking-tight">
                  RESUMEN DE VENTAS UPCONTA Y SOCIOS
                </h2>

                <div className="space-y-2 py-4 text-base font-semibold text-slate-300">
                  <div>VENTAS EQUIPO COMERCIAL: <strong className="text-white">${dataset.commercialCross?.commercialTeamSales ? dataset.commercialCross.commercialTeamSales.toLocaleString('en-US', { minimumFractionDigits: 2 }) : '0.00'}</strong></div>
                  <div>VENTAS ORGÁNICAS: <strong className="text-white">${dataset.commercialCross?.organicSales ? dataset.commercialCross.organicSales.toLocaleString('en-US', { minimumFractionDigits: 2 }) : '0.00'}</strong></div>
                  <div>VENTAS SOCIOS: <strong className="text-sky-400">${(dataset.commercialCross?.partnersSales || netSales).toLocaleString('en-US', { minimumFractionDigits: 2 })}</strong></div>
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
    </div>
  );
};
