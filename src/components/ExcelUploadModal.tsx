import React, { useState, useRef, useEffect } from 'react';
import { 
  X, 
  Upload, 
  FileSpreadsheet, 
  CheckCircle2, 
  AlertCircle, 
  Download, 
  ArrowRight,
  FileSignature,
  Users2,
  Table,
  RefreshCw
} from 'lucide-react';
import { 
  parseExcelFile, 
  ParsedExcelResult, 
  downloadTemplateFirmas, 
  downloadTemplateSocios,
  downloadTemplateExcel 
} from '../services/excelService';
import { TransactionRecord, MonthlyMetric, SocioRecord } from '../types';

export type UploadTargetType = 'firma' | 'socios' | 'auto';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  targetType: UploadTargetType;
  onChangeTargetType?: (type: UploadTargetType) => void;
  onApplyData: (result: {
    type: 'transactions' | 'monthly_summary' | 'socios';
    transactions?: TransactionRecord[];
    monthlyMetrics?: MonthlyMetric[];
    socios?: SocioRecord[];
    detectedMonth?: string;
    detectedCutoffDate?: string;
    dateRangeStr?: string;
    mode: 'append' | 'replace';
  }) => void;
}

export const ExcelUploadModal: React.FC<Props> = ({
  isOpen,
  onClose,
  targetType: initialTargetType,
  onChangeTargetType,
  onApplyData,
}) => {
  const [activeType, setActiveType] = useState<UploadTargetType>(initialTargetType || 'firma');
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [parsedResult, setParsedResult] = useState<ParsedExcelResult | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string>('');
  const [importMode, setImportMode] = useState<'append' | 'replace'>('replace');
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const nextType = initialTargetType || 'firma';
    setActiveType(nextType);
    setParsedResult(null);
    setErrorMessage(null);
    // User requested: "para que lo tomes como referencia para que se sume al reporte 2"
    // Default to 'append' for socios so it automatically sums into Reporte 2
    if (nextType === 'socios') {
      setImportMode('append');
    }
  }, [initialTargetType, isOpen]);

  if (!isOpen) return null;

  const handleFileSelected = async (file: File) => {
    if (!file) return;
    setFileName(file.name);
    setIsProcessing(true);
    setErrorMessage(null);
    setParsedResult(null);

    try {
      const forced = activeType === 'auto' ? undefined : activeType;
      const result = await parseExcelFile(file, forced);
      setParsedResult(result);
    } catch (err: any) {
      setErrorMessage(err.message || 'Error al procesar el archivo Excel.');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFileSelected(e.dataTransfer.files[0]);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleConfirmImport = () => {
    if (!parsedResult) return;

    if (activeType === 'socios' || (activeType === 'auto' && parsedResult.recognizedType === 'socios' && parsedResult.socios)) {
      onApplyData({
        type: 'socios',
        socios: parsedResult.socios,
        transactions: parsedResult.transactions,
        detectedMonth: parsedResult.detectedMonth,
        detectedCutoffDate: parsedResult.detectedCutoffDate,
        dateRangeStr: parsedResult.dateRangeStr,
        mode: importMode,
      });
    } else if (parsedResult.recognizedType === 'monthly_summary' && parsedResult.monthlyMetrics) {
      onApplyData({
        type: 'monthly_summary',
        monthlyMetrics: parsedResult.monthlyMetrics,
        mode: importMode,
      });
    } else {
      // Por defecto o activeType === 'firma': Afecta exclusivamente al Reporte 1
      onApplyData({
        type: 'transactions',
        transactions: parsedResult.transactions,
        detectedMonth: parsedResult.detectedMonth,
        detectedCutoffDate: parsedResult.detectedCutoffDate,
        dateRangeStr: parsedResult.dateRangeStr,
        mode: importMode,
      });
    }
    onClose();
  };

  const switchTab = (type: UploadTargetType) => {
    setActiveType(type);
    if (onChangeTargetType) onChangeTargetType(type);
    setParsedResult(null);
    setErrorMessage(null);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm animate-in fade-in duration-200">
      <div className="w-full max-w-3xl bg-[#0e1626] border border-slate-800 rounded-2xl shadow-2xl flex flex-col max-h-[92vh] overflow-hidden">
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-800 bg-[#09101d]">
          <div className="flex items-center gap-3">
            <div className={`w-10 h-10 rounded-xl flex items-center justify-center font-bold ${
              activeType === 'firma'
                ? 'bg-blue-500/20 text-sky-400'
                : 'bg-emerald-500/20 text-emerald-400'
            }`}>
              {activeType === 'firma' ? <FileSignature className="w-5 h-5" /> : <Users2 className="w-5 h-5" />}
            </div>
            <div>
              <h2 className="text-base font-bold text-white">
                {activeType === 'firma' ? 'Carga Archivo Firma (Afecta únicamente al Reporte 1)' : 'Carga Archivo Socios (Afecta únicamente al Reporte 2)'}
              </h2>
              <p className="text-xs text-slate-400">
                {activeType === 'firma' 
                  ? 'Alimenta las emisiones y cálculos de canales del Reporte 1 (no afecta la cartera del Reporte 2)'
                  : 'Alimenta la nómina de socios, auditoría y Pareto del Reporte 2 (no altera las ventas del Reporte 1)'}
              </p>
            </div>
          </div>

          <button
            onClick={onClose}
            className="p-1 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* 2-Option Tabs: Carga Archivo Firma vs Carga Archivo Socios */}
        <div className="flex border-b border-slate-800 bg-slate-900/60 px-6 pt-3 gap-2">
          <button
            type="button"
            onClick={() => switchTab('firma')}
            className={`flex items-center gap-2 px-4 py-2.5 text-xs font-bold rounded-t-xl transition-all border-t border-x ${
              activeType === 'firma'
                ? 'bg-[#0e1626] text-sky-400 border-slate-700 border-b-transparent shadow-md'
                : 'bg-transparent text-slate-400 border-transparent hover:text-white'
            }`}
          >
            <FileSignature className="w-4 h-4" />
            <span>Carga Archivo Firma (Reporte 1)</span>
          </button>

          <button
            type="button"
            onClick={() => switchTab('socios')}
            className={`flex items-center gap-2 px-4 py-2.5 text-xs font-bold rounded-t-xl transition-all border-t border-x ${
              activeType === 'socios'
                ? 'bg-[#0e1626] text-emerald-400 border-slate-700 border-b-transparent shadow-md'
                : 'bg-transparent text-slate-400 border-transparent hover:text-white'
            }`}
          >
            <Users2 className="w-4 h-4" />
            <span>Carga Archivo Socios (Reporte 2)</span>
          </button>
        </div>

        {/* Content */}
        <div className="flex-1 overflow-y-auto p-6 space-y-6">
          {/* Dropzone */}
          {!parsedResult && (
            <div
              onDrop={handleDrop}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onClick={() => fileInputRef.current?.click()}
              className={`border-2 border-dashed rounded-2xl p-8 text-center cursor-pointer transition-all ${
                isDragging
                  ? 'border-sky-400 bg-sky-500/10'
                  : 'border-slate-700 hover:border-slate-600 bg-slate-900/60'
              }`}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx, .xls, .csv"
                className="hidden"
                onChange={(e) => {
                  if (e.target.files && e.target.files.length > 0) {
                    handleFileSelected(e.target.files[0]);
                  }
                }}
              />

              <div className={`w-14 h-14 rounded-2xl border flex items-center justify-center mx-auto mb-4 ${
                activeType === 'firma'
                  ? 'bg-blue-600/10 border-blue-500/20 text-sky-400'
                  : 'bg-emerald-600/10 border-emerald-500/20 text-emerald-400'
              }`}>
                <Upload className="w-7 h-7" />
              </div>

              <h3 className="text-sm font-bold text-white mb-1">
                {isProcessing
                  ? 'Procesando archivo...'
                  : activeType === 'firma'
                  ? 'Arrastra tu archivo Excel de Firmas / Emisiones aquí'
                  : 'Arrastra tu archivo Excel de Socios (7 Columnas) aquí'}
              </h3>
              <p className="text-xs text-slate-400 max-w-md mx-auto mb-4">
                {activeType === 'firma'
                  ? 'Procesa el reporte de emisiones extrayendo el valor de cada firma de la Columna BH (USD) y el canal (Distribuidor Connect / Upconnect) de la Columna CB.'
                  : 'Formato oficial de 7 columnas: SOCIO, FECHA, MES, ID CLIENTE, NOMBRE CLIENTE, TIPO DE PLAN, PRECIO. Los datos se sumarán directamente al Reporte 2.'}
              </p>

              <div className="flex flex-wrap items-center justify-center gap-2">
                {activeType === 'firma' ? (
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      downloadTemplateFirmas();
                    }}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-sky-300 border border-slate-700 shadow-sm"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Descargar Plantilla Archivo Firma (.xlsx)</span>
                  </button>
                ) : (
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      downloadTemplateSocios();
                    }}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-emerald-300 border border-slate-700 shadow-sm"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Descargar Plantilla Archivo Socios (7 Columnas .xlsx)</span>
                  </button>
                )}
              </div>

              {/* Decimal & format compatibility badge */}
              <div className="mt-3.5 py-1.5 px-3 rounded-lg bg-slate-800/70 border border-slate-700/60 text-[11px] text-slate-300 inline-flex items-center gap-1.5 max-w-lg mx-auto">
                {activeType === 'firma' ? (
                  <>
                    <span className="text-emerald-400 font-bold">✓ Columna BH y CB:</span>
                    <span>Montos extraídos de la <strong>Columna BH</strong> (USD) y clasificación automática de <strong>Distribuidor Connect vs Upconnect</strong> de la <strong>Columna CB</strong>.</span>
                  </>
                ) : (
                  <>
                    <span className="text-emerald-400 font-bold">✓ Integración Reporte 2:</span>
                    <span>Suma automáticamente las ventas, ticket promedio, planes más vendidos y evolución mensual de socios.</span>
                  </>
                )}
              </div>
            </div>
          )}

          {/* Error display */}
          {errorMessage && (
            <div className="p-4 rounded-xl bg-red-500/10 border border-red-500/30 flex items-start gap-3 text-xs text-red-300">
              <AlertCircle className="w-5 h-5 text-red-400 shrink-0 mt-0.5" />
              <div>
                <strong className="font-bold">Error de lectura:</strong> {errorMessage}
              </div>
            </div>
          )}

          {/* Parsed Preview */}
          {parsedResult && (
            <div className="space-y-5">
              <div className="p-4 rounded-xl bg-emerald-500/10 border border-emerald-500/30 flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <CheckCircle2 className="w-6 h-6 text-emerald-400 shrink-0" />
                  <div>
                    <h4 className="text-xs font-bold text-white">Archivo procesado con éxito: {fileName}</h4>
                    <p className="text-[11px] text-emerald-300">
                      Tipo: {
                        parsedResult.recognizedType === 'socios' 
                          ? 'Ranking y Ventas de Socios' 
                          : parsedResult.recognizedType === 'monthly_summary'
                          ? 'Resumen Mensual Consolidado'
                          : 'Detalle de Transacciones / Firmas'
                      } ({parsedResult.rawRowsCount} filas)
                    </p>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setParsedResult(null)}
                    className="px-2.5 py-1.5 text-xs text-slate-400 hover:text-white bg-slate-800 rounded-lg border border-slate-700 transition-colors"
                  >
                    Cambiar archivo
                  </button>
                  <button
                    onClick={handleConfirmImport}
                    className="flex items-center gap-1.5 px-4 py-1.5 text-xs font-bold rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white shadow-md transition-all hover:scale-[1.02]"
                    title="Cargar y actualizar información en el sistema"
                  >
                    <RefreshCw className="w-3.5 h-3.5" />
                    <span>Actualizar</span>
                  </button>
                </div>
              </div>

              {/* Month & Date Reconciliation Badge */}
              {(parsedResult.detectedMonth || parsedResult.detectedCutoffDate) && (
                <div className="p-3.5 rounded-xl bg-[#091122] border border-emerald-500/40 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 shadow-lg">
                  <div className="flex items-center gap-2.5">
                    <div className="w-3 h-3 rounded-full bg-emerald-400 shadow-sm shadow-emerald-400/50 shrink-0" />
                    <div>
                      <div className="text-xs font-bold text-white flex items-center gap-2">
                        <span>Mes Identificado:</span>
                        <span className="px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 font-extrabold uppercase">
                          {parsedResult.detectedMonth || 'SEPTIEMBRE'} {parsedResult.detectedYear || 2026}
                        </span>
                      </div>
                      {parsedResult.dateRangeStr && (
                        <div className="text-[11px] text-slate-400 mt-0.5">
                          Rango de emisiones: <strong className="text-slate-200">{parsedResult.dateRangeStr}</strong>
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="text-[11px] font-semibold text-sky-300 bg-sky-950/80 px-3 py-1.5 rounded-lg border border-sky-800/60 shrink-0">
                    Corte Oficial: {parsedResult.detectedCutoffDate || '20 de septiembre de 2026'}
                  </div>
                </div>
              )}

              {/* Action Banner to inform user to click Actualizar */}
              <div className="p-3.5 rounded-xl bg-gradient-to-r from-blue-950/60 via-slate-900/80 to-blue-950/60 border border-sky-500/40 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-lg">
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-lg bg-sky-500/20 text-sky-400 flex items-center justify-center shrink-0">
                    <RefreshCw className="w-4 h-4" />
                  </div>
                  <div>
                    <div className="text-xs font-bold text-white">
                      Archivo listo para cargar: <span className="text-sky-400 font-mono">{fileName}</span>
                    </div>
                    <div className="text-[11px] text-slate-300">
                      Presiona el botón <strong className="text-white">Actualizar</strong> para sincronizar e incorporar la información al sistema.
                    </div>
                  </div>
                </div>

                <button
                  onClick={handleConfirmImport}
                  className="w-full sm:w-auto flex items-center justify-center gap-2 px-5 py-2 text-xs font-extrabold rounded-xl bg-gradient-to-r from-blue-600 to-sky-500 hover:from-blue-500 hover:to-sky-400 text-white shadow-lg shadow-sky-600/30 transition-all hover:scale-[1.02] uppercase tracking-wider shrink-0"
                >
                  <RefreshCw className="w-4 h-4" />
                  <span>Actualizar</span>
                </button>
              </div>

              {/* Mapped columns */}
              <div>
                <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">
                  Columnas Mapeadas
                </span>
                <div className="flex flex-wrap gap-1.5 mt-2">
                  {Object.entries(parsedResult.recognizedColumns).map(([stdName, originalHeader]) => (
                    <span
                      key={stdName}
                      className="px-2 py-0.5 rounded text-[10px] font-medium bg-slate-800 text-sky-300 border border-slate-700"
                    >
                      {stdName}: <strong>{originalHeader}</strong>
                    </span>
                  ))}
                </div>
              </div>

              {/* Decimal & currency detection indicator */}
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1.5 px-3.5 py-2 rounded-xl bg-slate-900 border border-slate-800 text-xs">
                <div className="flex items-center gap-2 text-slate-300 text-[11px]">
                  <span className="text-emerald-400 font-bold">✓ Formato Numérico:</span>
                  <span>Interpretación flexible activa (soporta <code>29.99</code> con punto y <code>29,99</code> con coma, enteros y <code>$</code>).</span>
                </div>
                <span className="text-[10px] text-slate-400 font-mono">
                  {parsedResult.recognizedColumns['Valor / Monto'] || parsedResult.recognizedColumns['Valor'] 
                    ? `Columna Monto/Valor detectada: "${parsedResult.recognizedColumns['Valor / Monto'] || parsedResult.recognizedColumns['Valor']}"` 
                    : 'Columna Monto/Valor: auto-detectada'}
                </span>
              </div>

              {/* Socios Preview */}
              {parsedResult.socios && parsedResult.socios.length > 0 && (
                <div>
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-xs font-bold text-white">
                      Vista previa de Socios / Vendedores ({parsedResult.socios.length} registrados · Afecta Reporte 1 y 2)
                    </span>
                    <span className="text-xs text-emerald-400 font-bold tabular-nums">
                      Monto total vendedores: $
                      {parsedResult.socios.reduce((a, s) => a + s.totalSales, 0).toLocaleString('en-US', { minimumFractionDigits: 2 })} USD
                    </span>
                  </div>

                  <div className="border border-slate-800 rounded-xl overflow-hidden max-h-56 overflow-y-auto">
                    <table className="w-full text-xs text-left">
                      <thead className="bg-slate-900 text-slate-400 border-b border-slate-800">
                        <tr>
                          <th className="py-2 px-2.5 text-center">#</th>
                          <th className="py-2 px-2.5">Socio</th>
                          <th className="py-2 px-2.5">Rol</th>
                          <th className="py-2 px-2.5 text-center">Ventas</th>
                          <th className="py-2 px-2.5 text-right">Monto Total ($)</th>
                          <th className="py-2 px-2.5 text-right">Ticket Prom. ($)</th>
                          <th className="py-2 px-2.5">Plan Más Vendido</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-800 text-slate-300 tabular-nums">
                        {parsedResult.socios.slice(0, 10).map((s) => (
                          <tr key={s.rank} className="hover:bg-slate-850">
                            <td className="py-1.5 px-2.5 text-center font-bold text-slate-400">{s.rank}</td>
                            <td className="py-1.5 px-2.5 font-semibold text-white">{s.name}</td>
                            <td className="py-1.5 px-2.5 text-[11px] text-sky-300">{s.role || 'Distribuidor Connect'}</td>
                            <td className="py-1.5 px-2.5 text-center font-semibold text-slate-200">{s.operationsCount || 1}</td>
                            <td className="py-1.5 px-2.5 text-right font-black text-emerald-400">${s.totalSales.toFixed(2)}</td>
                            <td className="py-1.5 px-2.5 text-right font-bold text-sky-400">
                              ${(s.averageTicket ?? (s.totalSales / (s.operationsCount || 1))).toFixed(2)}
                            </td>
                            <td className="py-1.5 px-2.5 text-[11px] text-slate-300 truncate max-w-[150px]" title={s.topPlan}>
                              {s.topPlan || 'Firma Electrónica'}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* Transactions Preview */}
              {parsedResult.recognizedType === 'transactions' && (
                <div>
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-xs font-bold text-white">
                      Vista previa de Firmas ({parsedResult.transactions.length} registros identificados)
                    </span>
                    <span className="text-xs text-sky-400 font-bold tabular-nums">
                      Monto total: $
                      {parsedResult.transactions
                        .reduce((a, b) => a + b.value, 0)
                        .toFixed(2)} USD
                    </span>
                  </div>

                  <div className="border border-slate-800 rounded-xl overflow-hidden max-h-48 overflow-y-auto">
                    <table className="w-full text-[11px] text-left">
                      <thead className="bg-slate-900 text-slate-400 border-b border-slate-800">
                        <tr>
                          <th className="py-2 px-2.5">ID Único</th>
                          <th className="py-2 px-2.5">Fecha</th>
                          <th className="py-2 px-2.5">Cliente</th>
                          <th className="py-2 px-2.5">Duración</th>
                          <th className="py-2 px-2.5 text-right">Valor</th>
                          <th className="py-2 px-2.5">Canal</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-800/60 text-slate-300">
                        {parsedResult.transactions.slice(0, 5).map((trx, idx) => (
                          <tr key={idx} className="hover:bg-slate-850">
                            <td className="py-1.5 px-2.5 font-mono text-sky-400">{trx.uniqueId}</td>
                            <td className="py-1.5 px-2.5">{trx.date}</td>
                            <td className="py-1.5 px-2.5 font-medium text-white">{trx.clientName} {trx.clientLastName}</td>
                            <td className="py-1.5 px-2.5">{trx.duration}</td>
                            <td className="py-1.5 px-2.5 text-right font-bold text-emerald-400">${trx.value.toFixed(2)}</td>
                            <td className="py-1.5 px-2.5">{trx.channel}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* Append vs Replace Choice */}
              <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 space-y-2">
                <span className="text-xs font-bold text-slate-300">Modo de integración al sistema:</span>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
                  <label
                    className={`flex items-start gap-3 p-3 rounded-lg border cursor-pointer transition-colors ${
                      importMode === 'append'
                        ? 'border-emerald-500 bg-emerald-500/10'
                        : 'border-slate-800 hover:border-slate-700 bg-slate-950/40'
                    }`}
                  >
                    <input
                      type="radio"
                      name="importMode"
                      checked={importMode === 'append'}
                      onChange={() => setImportMode('append')}
                      className="mt-1"
                    />
                    <div>
                      <div className="text-xs font-bold text-white flex items-center gap-1.5">
                        <span>{activeType === 'socios' ? 'Sumar al Reporte 2 (Carga Progresiva)' : 'Carga Progresiva (Incremental)'}</span>
                        {activeType === 'socios' && (
                          <span className="text-[10px] px-1.5 py-0.2 rounded bg-emerald-500/20 text-emerald-300 font-semibold border border-emerald-500/30">
                            Recomendado
                          </span>
                        )}
                      </div>
                      <div className="text-[11px] text-slate-400">
                        {activeType === 'socios'
                          ? 'Suma y acumula las ventas de los socios a la base actual para consolidar el Reporte 2.'
                          : 'Suma y acumula a la base actual para alimentar continuamente el reporte.'}
                      </div>
                    </div>
                  </label>

                  <label
                    className={`flex items-start gap-3 p-3 rounded-lg border cursor-pointer transition-colors ${
                      importMode === 'replace'
                        ? 'border-sky-500 bg-sky-500/10'
                        : 'border-slate-800 hover:border-slate-700 bg-slate-950/40'
                    }`}
                  >
                    <input
                      type="radio"
                      name="importMode"
                      checked={importMode === 'replace'}
                      onChange={() => setImportMode('replace')}
                      className="mt-1"
                    />
                    <div>
                      <div className="text-xs font-bold text-white">Reemplazar Todo el Conjunto</div>
                      <div className="text-[11px] text-slate-400">
                        {activeType === 'socios'
                          ? 'Sobrescribe los socios y ventas del Reporte 2 únicamente con este archivo.'
                          : 'Sobrescribe y actualiza con los datos del archivo cargado.'}
                      </div>
                    </div>
                  </label>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="flex items-center justify-between px-6 py-4 border-t border-slate-800 bg-[#09101d]">
          <button
            onClick={onClose}
            className="px-4 py-2 text-xs font-semibold rounded-lg bg-slate-800 text-slate-300 hover:bg-slate-700 transition-colors"
          >
            Cancelar
          </button>

          <button
            onClick={handleConfirmImport}
            disabled={!parsedResult}
            className="flex items-center gap-2 px-6 py-2.5 text-xs font-black rounded-xl bg-blue-600 hover:bg-blue-500 disabled:opacity-40 disabled:hover:bg-blue-600 text-white shadow-lg shadow-blue-600/25 transition-all uppercase tracking-wider"
          >
            <RefreshCw className="w-4 h-4" />
            <span>Actualizar</span>
          </button>
        </div>
      </div>
    </div>
  );
};
