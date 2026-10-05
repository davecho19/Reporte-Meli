import React, { useState } from 'react';
import { X, Check, Edit3, DollarSign, Type, Hash, Sparkles, Building2, MessageSquare, Layers } from 'lucide-react';
import { GlobalDataset } from '../types';
import { FilteredReportView } from '../utils/reportFilters';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  reportType: 1 | 2;
  slideId: string;
  slideTitle: string;
  viewData: FilteredReportView;
  dataset: GlobalDataset;
  editValues: Record<string, string | number>;
  onChangeValue: (key: string, value: string | number) => void;
  onSave: () => void;
}

export const SlideEditorModal: React.FC<Props> = ({
  isOpen,
  onClose,
  reportType,
  slideId,
  slideTitle,
  viewData,
  dataset,
  editValues,
  onChangeValue,
  onSave,
}) => {
  if (!isOpen) return null;

  const [activeTab, setActiveTab] = useState<'slide' | 'general' | 'kpis' | 'notes'>('slide');

  const getValue = (key: string, fallback: string | number) => {
    return editValues[key] !== undefined ? editValues[key] : fallback;
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-5 bg-black/80 backdrop-blur-sm no-print animate-in fade-in duration-200">
      <div className="w-full max-w-2xl bg-[#0c1322] border border-emerald-500/40 rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[92vh]">
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-800 bg-[#09101d]">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-emerald-500/20 border border-emerald-500/30 flex items-center justify-center text-emerald-400">
              <Edit3 className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-base font-bold text-white tracking-tight">
                Editor: {slideTitle}
              </h3>
              <p className="text-xs text-slate-400">
                Reporte {reportType} · Edite títulos, valores, montos y textos de la diapositiva y el reporte
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="text-slate-400 hover:text-white p-1.5 rounded-lg hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Tab Selector */}
        <div className="flex items-center gap-2 px-6 pt-3 pb-2 border-b border-slate-800 bg-[#0a1120] text-xs">
          <button
            onClick={() => setActiveTab('slide')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg font-semibold transition-colors ${
              activeTab === 'slide'
                ? 'bg-emerald-600 text-white shadow-sm'
                : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
          >
            <Layers className="w-3.5 h-3.5" />
            <span>Datos de esta Diapositiva ({slideTitle})</span>
          </button>
          <button
            onClick={() => setActiveTab('general')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg font-semibold transition-colors ${
              activeTab === 'general'
                ? 'bg-emerald-600 text-white shadow-sm'
                : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
          >
            <Type className="w-3.5 h-3.5" />
            <span>Títulos y Fechas</span>
          </button>
          <button
            onClick={() => setActiveTab('kpis')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg font-semibold transition-colors ${
              activeTab === 'kpis'
                ? 'bg-emerald-600 text-white shadow-sm'
                : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
          >
            <DollarSign className="w-3.5 h-3.5" />
            <span>Métricas Globales</span>
          </button>
          <button
            onClick={() => setActiveTab('notes')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg font-semibold transition-colors ${
              activeTab === 'notes'
                ? 'bg-emerald-600 text-white shadow-sm'
                : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
          >
            <MessageSquare className="w-3.5 h-3.5" />
            <span>Conclusiones</span>
          </button>
        </div>

        {/* Form Body */}
        <div className="p-6 overflow-y-auto space-y-5 flex-1 text-xs">
          {/* TAB 1: SLIDE-SPECIFIC DATA */}
          {activeTab === 'slide' && (
            <div className="space-y-4">
              <div className="p-3 bg-emerald-950/40 border border-emerald-500/30 rounded-xl text-emerald-300 text-xs">
                Editando campos específicos para <strong>{slideTitle}</strong> (Diapositiva: {slideId}). Todos los cambios se reflejan inmediatamente en pantalla y en la exportación PowerPoint (.pptx).
              </div>

              {/* COVER SLIDE */}
              {slideId === 'cover' && (
                <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 space-y-3">
                  <div className="text-sky-400 font-bold uppercase tracking-wider text-[11px] flex items-center gap-1.5">
                    <Type className="w-3.5 h-3.5" />
                    <span>Textos de Portada</span>
                  </div>
                  <div>
                    <label className="block text-slate-300 font-semibold mb-1">Título Principal de la Portada</label>
                    <input
                      type="text"
                      value={String(getValue(reportType === 1 ? 'r1_cover_title' : 'r2_cover_title', reportType === 1 ? viewData.titleReport1 : viewData.titleReport2))}
                      onChange={(e) => {
                        const val = e.target.value;
                        if (reportType === 1) {
                          onChangeValue('r1_cover_title', val);
                          onChangeValue('titleReport1', val);
                        } else {
                          onChangeValue('r2_cover_title', val);
                          onChangeValue('titleReport2', val);
                        }
                      }}
                      className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-medium focus:border-emerald-500 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-slate-300 font-semibold mb-1">Subtítulo Descriptivo</label>
                    <input
                      type="text"
                      value={String(getValue(reportType === 1 ? 'r1_cover_subtitle' : 'r2_cover_subtitle', reportType === 1 ? 'Auditoría comparativa de emisión de firmas electrónicas' : 'Análisis detallado de ventas por socio, ticket promedio y conciliación'))}
                      onChange={(e) => {
                        const val = e.target.value;
                        if (reportType === 1) onChangeValue('r1_cover_subtitle', val);
                        else onChangeValue('r2_cover_subtitle', val);
                      }}
                      className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-medium focus:border-emerald-500 focus:outline-none"
                    />
                  </div>
                </div>
              )}

              {/* BALANCE GENERAL SLIDE */}
              {slideId === 'balance' && (
                <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 space-y-3">
                  <div className="text-emerald-400 font-bold uppercase tracking-wider text-[11px] flex items-center gap-1.5">
                    <DollarSign className="w-3.5 h-3.5" />
                    <span>Montos Balance General</span>
                  </div>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-slate-300 font-semibold mb-1">Ventas Upconnect ($ USD)</label>
                      <input
                        type="number"
                        step="any"
                        value={Number(getValue('totalUpSales', viewData.totalUpSales))}
                        onChange={(e) => {
                          const n = parseFloat(e.target.value) || 0;
                          onChangeValue('totalUpSales', n);
                          onChangeValue('r1_kpi_up_sales', n);
                        }}
                        className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-sky-400 font-bold focus:border-emerald-500 focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-300 font-semibold mb-1">Ventas Distribuidor Connect ($ USD)</label>
                      <input
                        type="number"
                        step="any"
                        value={Number(getValue('totalCoSales', viewData.totalCoSales))}
                        onChange={(e) => {
                          const n = parseFloat(e.target.value) || 0;
                          onChangeValue('totalCoSales', n);
                          onChangeValue('r1_kpi_co_sales', n);
                        }}
                        className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-amber-400 font-bold focus:border-emerald-500 focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-300 font-semibold mb-1">Firmas Upconnect (Cant.)</label>
                      <input
                        type="number"
                        value={Number(getValue('totalUpCount', viewData.totalUpCount))}
                        onChange={(e) => {
                          const n = parseInt(e.target.value, 10) || 0;
                          onChangeValue('totalUpCount', n);
                          onChangeValue('r1_kpi_up_count', n);
                        }}
                        className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-bold focus:border-emerald-500 focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-300 font-semibold mb-1">Firmas Distribuidor Connect (Cant.)</label>
                      <input
                        type="number"
                        value={Number(getValue('totalCoCount', viewData.totalCoCount))}
                        onChange={(e) => {
                          const n = parseInt(e.target.value, 10) || 0;
                          onChangeValue('totalCoCount', n);
                          onChangeValue('r1_kpi_co_count', n);
                        }}
                        className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-bold focus:border-emerald-500 focus:outline-none"
                      />
                    </div>
                  </div>
                </div>
              )}

              {/* WEEKLY BREAKDOWN SLIDE */}
              {slideId === 'weekly' && (
                <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 space-y-3">
                  <div className="flex items-center gap-2 text-amber-400 font-bold uppercase tracking-wider text-[11px]">
                    <Hash className="w-3.5 h-3.5" />
                    <span>Desglose por Semana (Valores Exactos)</span>
                  </div>
                  <div className="space-y-2">
                    {(viewData.weeklyBreakdownType1 || []).map((w) => {
                      const wKey = `week_${w.id}_amount`;
                      const wCountKey = `week_${w.id}_count`;
                      const curVal = getValue(wKey, w.totalAmount);
                      const curCount = getValue(wCountKey, w.totalCount);
                      return (
                        <div key={w.id} className="flex items-center justify-between gap-3 p-2 rounded-lg bg-slate-950/60 border border-slate-800">
                          <div>
                            <span className="font-bold text-white">{w.weekName}</span>
                            <span className="text-slate-400 ml-2">({w.dateRange})</span>
                          </div>
                          <div className="flex items-center gap-3">
                            <div className="flex items-center gap-1">
                              <span className="text-slate-400 text-[10px]">Firmas:</span>
                              <input
                                type="number"
                                value={Number(curCount)}
                                onChange={(e) => onChangeValue(wCountKey, parseInt(e.target.value, 10) || 0)}
                                className="w-16 bg-slate-900 border border-slate-700 rounded px-1.5 py-1 text-white text-center text-xs"
                              />
                            </div>
                            <div className="flex items-center gap-1">
                              <span className="text-slate-400">$</span>
                              <input
                                type="number"
                                step="any"
                                value={Number(curVal)}
                                onChange={(e) => onChangeValue(wKey, parseFloat(e.target.value) || 0)}
                                className="w-24 bg-slate-900 border border-emerald-500 rounded px-2 py-1 text-emerald-400 font-bold text-right text-xs"
                              />
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* PRODUCTS SLIDE (TOP 6) */}
              {slideId === 'products' && (
                <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 space-y-3">
                  <div className="flex items-center gap-2 text-sky-400 font-bold uppercase tracking-wider text-[11px]">
                    <Sparkles className="w-3.5 h-3.5" />
                    <span>Top 6 Productos (Vigencia 1 año, 2 años, etc.)</span>
                  </div>
                  <p className="text-slate-400 text-[11px]">
                    Modifique los nombres y montos para el canal Distribuidor Upconnect y Distribuidor Connect.
                  </p>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div className="space-y-2">
                      <span className="font-bold text-sky-400 block border-b border-sky-900/40 pb-1">Distribuidor Upconnect</span>
                      {[1, 2, 3, 4, 5, 6].map((num) => (
                        <div key={num} className="flex items-center gap-2">
                          <span className="w-4 text-slate-500 font-bold">{num}.</span>
                          <input
                            type="text"
                            placeholder={`Firma de ${num} año(s)`}
                            value={String(getValue(`top_up_${num}_name`, `Firma de ${num} año${num === 1 ? '' : 's'}`))}
                            onChange={(e) => onChangeValue(`top_up_${num}_name`, e.target.value)}
                            className="flex-1 bg-slate-950 border border-slate-700 rounded px-2 py-1 text-white text-xs"
                          />
                          <input
                            type="number"
                            step="any"
                            placeholder="$ USD"
                            value={Number(getValue(`top_up_${num}_amount`, num === 1 ? 1650 : num === 2 ? 840 : 420))}
                            onChange={(e) => onChangeValue(`top_up_${num}_amount`, parseFloat(e.target.value) || 0)}
                            className="w-20 bg-slate-950 border border-sky-600/50 rounded px-1.5 py-1 text-sky-400 font-bold text-right text-xs"
                          />
                        </div>
                      ))}
                    </div>

                    <div className="space-y-2">
                      <span className="font-bold text-amber-400 block border-b border-amber-900/40 pb-1">Distribuidor Connect</span>
                      {[1, 2, 3, 4, 5, 6].map((num) => (
                        <div key={num} className="flex items-center gap-2">
                          <span className="w-4 text-slate-500 font-bold">{num}.</span>
                          <input
                            type="text"
                            placeholder={`Firma de ${num} año(s)`}
                            value={String(getValue(`top_co_${num}_name`, `Firma de ${num} año${num === 1 ? '' : 's'}`))}
                            onChange={(e) => onChangeValue(`top_co_${num}_name`, e.target.value)}
                            className="flex-1 bg-slate-950 border border-slate-700 rounded px-2 py-1 text-white text-xs"
                          />
                          <input
                            type="number"
                            step="any"
                            placeholder="$ USD"
                            value={Number(getValue(`top_co_${num}_amount`, num === 1 ? 2100 : num === 2 ? 1150 : 530))}
                            onChange={(e) => onChangeValue(`top_co_${num}_amount`, parseFloat(e.target.value) || 0)}
                            className="w-20 bg-slate-950 border border-amber-600/50 rounded px-1.5 py-1 text-amber-400 font-bold text-right text-xs"
                          />
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {/* COMMUNITIES SLIDE */}
              {slideId === 'communities' && (
                <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 space-y-3">
                  <div className="flex items-center gap-2 text-indigo-400 font-bold uppercase tracking-wider text-[11px]">
                    <Building2 className="w-3.5 h-3.5" />
                    <span>Comunidades y Franquiciados</span>
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div>
                      <label className="block text-slate-300 font-semibold mb-1">UpConta Socios (%)</label>
                      <input
                        type="number"
                        step="any"
                        value={Number(getValue('upcontaSocios', 52))}
                        onChange={(e) => onChangeValue('upcontaSocios', parseFloat(e.target.value) || 0)}
                        className="w-full bg-slate-950 border border-slate-700 rounded px-2.5 py-1.5 text-sky-400 font-bold text-xs"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-300 font-semibold mb-1">Franquicia VIP (%)</label>
                      <input
                        type="number"
                        step="any"
                        value={Number(getValue('franquiciaVIP', 31))}
                        onChange={(e) => onChangeValue('franquiciaVIP', parseFloat(e.target.value) || 0)}
                        className="w-full bg-slate-950 border border-slate-700 rounded px-2.5 py-1.5 text-amber-400 font-bold text-xs"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-300 font-semibold mb-1">Formación Comercial (%)</label>
                      <input
                        type="number"
                        step="any"
                        value={Number(getValue('formacionComercial', 17))}
                        onChange={(e) => onChangeValue('formacionComercial', parseFloat(e.target.value) || 0)}
                        className="w-full bg-slate-950 border border-slate-700 rounded px-2.5 py-1.5 text-emerald-400 font-bold text-xs"
                      />
                    </div>
                  </div>
                </div>
              )}

              {/* SOCIOS / PARETO SLIDE */}
              {slideId === 'socios' && (
                <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 space-y-3">
                  <div className="flex items-center gap-2 text-emerald-400 font-bold uppercase tracking-wider text-[11px]">
                    <Layers className="w-3.5 h-3.5" />
                    <span>Auditoría de Socios & Pareto</span>
                  </div>
                  <p className="text-slate-400 text-[11px]">
                    Modifique las ventas o nombres de los principales socios comerciales.
                  </p>
                  <div className="space-y-2 max-h-60 overflow-y-auto pr-1">
                    {(viewData.socios || []).slice(0, 8).map((s, idx) => {
                      const socioKey = s.rank || (idx + 1);
                      return (
                        <div key={socioKey} className="flex items-center gap-2 p-2 rounded bg-slate-950 border border-slate-800">
                          <span className="w-6 font-bold text-emerald-400">#{s.rank}</span>
                          <input
                            type="text"
                            value={String(getValue(`socio_${socioKey}_name`, s.name))}
                            onChange={(e) => onChangeValue(`socio_${socioKey}_name`, e.target.value)}
                            className="flex-1 bg-transparent border-0 text-white font-medium text-xs focus:ring-1 focus:ring-emerald-500 rounded px-1"
                          />
                          <div className="flex items-center gap-1">
                            <span className="text-slate-400">$</span>
                            <input
                              type="number"
                              step="any"
                              value={Number(getValue(`socio_${socioKey}_sales`, s.totalSales))}
                              onChange={(e) => onChangeValue(`socio_${socioKey}_sales`, parseFloat(e.target.value) || 0)}
                              className="w-24 bg-slate-900 border border-emerald-600 rounded px-1.5 py-1 text-emerald-400 font-bold text-right text-xs"
                            />
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          )}

          {/* TAB 2: GENERAL TITLES & DATES */}
          {activeTab === 'general' && (
            <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 space-y-3">
              <div className="flex items-center gap-2 text-sky-400 font-bold uppercase tracking-wider text-[11px]">
                <Type className="w-3.5 h-3.5" />
                <span>Títulos y Encabezados Globales</span>
              </div>

              <div>
                <label className="block text-slate-300 font-semibold mb-1">
                  Título de Portada Reporte 1 (Distribuidor Upconnect / Distribuidor Connect)
                </label>
                <input
                  type="text"
                  value={String(getValue('titleReport1', viewData.titleReport1))}
                  onChange={(e) => {
                    onChangeValue('titleReport1', e.target.value);
                    onChangeValue('r1_cover_title', e.target.value);
                  }}
                  className="w-full bg-slate-950 border border-slate-700 focus:border-emerald-500 rounded-lg px-3 py-2 text-white font-medium text-xs focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-slate-300 font-semibold mb-1">
                  Título de Portada Reporte 2 (Auditoría & Socios)
                </label>
                <input
                  type="text"
                  value={String(getValue('titleReport2', viewData.titleReport2))}
                  onChange={(e) => {
                    onChangeValue('titleReport2', e.target.value);
                    onChangeValue('r2_cover_title', e.target.value);
                  }}
                  className="w-full bg-slate-950 border border-slate-700 focus:border-emerald-500 rounded-lg px-3 py-2 text-white font-medium text-xs focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-slate-300 font-semibold mb-1">
                  Subtítulo / Fecha de Corte
                </label>
                <input
                  type="text"
                  value={String(getValue('subtitleDate', viewData.subtitleDate))}
                  onChange={(e) => {
                    onChangeValue('subtitleDate', e.target.value);
                    onChangeValue('r1_cover_subtitle', e.target.value);
                  }}
                  className="w-full bg-slate-950 border border-slate-700 focus:border-emerald-500 rounded-lg px-3 py-2 text-white font-medium text-xs focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-slate-300 font-semibold mb-1">
                  Etiqueta de Período Activo (ej. SEPTIEMBRE 2026, CORTE W2)
                </label>
                <input
                  type="text"
                  value={String(getValue('periodLabel', viewData.periodLabel))}
                  onChange={(e) => onChangeValue('periodLabel', e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 focus:border-emerald-500 rounded-lg px-3 py-2 text-white font-medium text-xs focus:outline-none"
                />
              </div>
            </div>
          )}

          {/* TAB 3: KEY KPIS */}
          {activeTab === 'kpis' && (
            <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 space-y-3">
              <div className="flex items-center gap-2 text-emerald-400 font-bold uppercase tracking-wider text-[11px]">
                <DollarSign className="w-3.5 h-3.5" />
                <span>Valores y Facturación ($ USD)</span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-slate-300 font-semibold mb-1">
                    Ventas Upconnect ($ USD)
                  </label>
                  <input
                    type="number"
                    step="any"
                    value={Number(getValue('totalUpSales', viewData.totalUpSales))}
                    onChange={(e) => {
                      const num = parseFloat(e.target.value) || 0;
                      onChangeValue('totalUpSales', num);
                      onChangeValue('r1_kpi_up_sales', num);
                    }}
                    className="w-full bg-slate-950 border border-slate-700 focus:border-emerald-500 rounded-lg px-3 py-2 text-emerald-400 font-bold text-xs focus:outline-none"
                  />
                </div>

                <div>
                  <label className="block text-slate-300 font-semibold mb-1">
                    Ventas Distribuidor Connect ($ USD)
                  </label>
                  <input
                    type="number"
                    step="any"
                    value={Number(getValue('totalCoSales', viewData.totalCoSales))}
                    onChange={(e) => {
                      const num = parseFloat(e.target.value) || 0;
                      onChangeValue('totalCoSales', num);
                      onChangeValue('r1_kpi_co_sales', num);
                    }}
                    className="w-full bg-slate-950 border border-slate-700 focus:border-emerald-500 rounded-lg px-3 py-2 text-amber-400 font-bold text-xs focus:outline-none"
                  />
                </div>

                <div>
                  <label className="block text-slate-300 font-semibold mb-1">
                    Total Facturado Consolidado ($ USD)
                  </label>
                  <input
                    type="number"
                    step="any"
                    value={Number(getValue('totalSales', viewData.totalSales))}
                    onChange={(e) => {
                      const num = parseFloat(e.target.value) || 0;
                      onChangeValue('totalSales', num);
                      onChangeValue('r1_kpi_total_sales', num);
                    }}
                    className="w-full bg-slate-950 border border-slate-700 focus:border-emerald-500 rounded-lg px-3 py-2 text-white font-black text-xs focus:outline-none"
                  />
                </div>

                <div>
                  <label className="block text-slate-300 font-semibold mb-1">
                    Total Operaciones / Firmas
                  </label>
                  <input
                    type="number"
                    value={Number(getValue('totalCount', viewData.totalCount))}
                    onChange={(e) => {
                      const num = parseInt(e.target.value, 10) || 0;
                      onChangeValue('totalCount', num);
                      onChangeValue('r1_kpi_total_count', num);
                    }}
                    className="w-full bg-slate-950 border border-slate-700 focus:border-emerald-500 rounded-lg px-3 py-2 text-white font-bold text-xs focus:outline-none"
                  />
                </div>
              </div>
            </div>
          )}

          {/* TAB 4: CONCLUSIONS */}
          {activeTab === 'notes' && (
            <div className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 space-y-3">
              <div className="flex items-center gap-2 text-sky-400 font-bold uppercase tracking-wider text-[11px]">
                <MessageSquare className="w-3.5 h-3.5" />
                <span>Conclusiones y Notas del Informe</span>
              </div>

              <div>
                <label className="block text-slate-300 font-semibold mb-1">
                  Conclusión Reporte 1 (Distribuidor Upconnect vs Distribuidor Connect)
                </label>
                <textarea
                  rows={3}
                  value={String(getValue('r1_conclusion', `Facturación consolidada de $${viewData.totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD con ${viewData.totalCount} certificados emitidos en total.`))}
                  onChange={(e) => onChangeValue('r1_conclusion', e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 focus:border-emerald-500 rounded-lg px-3 py-2 text-white font-medium text-xs focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-slate-300 font-semibold mb-1">
                  Conclusión Reporte 2 (Auditoría de Socios & Pareto)
                </label>
                <textarea
                  rows={3}
                  value={String(getValue('r2_conclusion', 'Regla de Oro Pareto: El Top 10 concentra la gran mayoría de la recaudación comercial de la plataforma UpConta.'))}
                  onChange={(e) => onChangeValue('r2_conclusion', e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 focus:border-emerald-500 rounded-lg px-3 py-2 text-white font-medium text-xs focus:outline-none"
                />
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="flex items-center justify-between px-6 py-4 border-t border-slate-800 bg-[#09101d]">
          <span className="text-[11px] text-slate-400">
            Los cambios se guardan y se sincronizan en pantalla y en la exportación PowerPoint.
          </span>
          <div className="flex items-center gap-2">
            <button
              onClick={onClose}
              className="px-4 py-2 text-xs font-medium rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors"
            >
              Cancelar
            </button>
            <button
              onClick={() => {
                onSave();
                onClose();
              }}
              className="flex items-center gap-1.5 px-4 py-2 text-xs font-bold rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white transition-colors shadow-lg shadow-emerald-950/40"
            >
              <Check className="w-4 h-4" />
              <span>Guardar y Aplicar</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
