/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { GlobalDataset, TransactionRecord, MonthlyMetric, SocioRecord, ReportTimeFilter } from './types';
import { initialDataset } from './data/initialData';
import { Navigation, NavView } from './components/Navigation';
import { DashboardView } from './components/DashboardView';
import { PresentationType1 } from './components/PresentationType1';
import { PresentationType2 } from './components/PresentationType2';
import { ExcelUploadModal, UploadTargetType } from './components/ExcelUploadModal';
import { 
  deriveMonthlyMetricsFromTransactions, 
  derivePortfolioDurationsFromTransactions,
  deriveWeeklyBreakdownFromTransactions,
  deriveSolutionCategoriesFromTransactions,
  deriveCommercialCrossFromTransactions,
  formatSpanishDateCutoff
} from './utils/reportFilters';
import { CheckCircle2, AlertCircle, Info, X } from 'lucide-react';
import { exportCombinedPresentationsToPPTX } from './services/pptxExportService';

const DATASET_STORAGE_KEY = 'upconnect_global_dataset_v2';

export default function App() {
  const [currentView, setCurrentView] = useState<NavView>('dashboard');
  const [isUploadOpen, setIsUploadOpen] = useState(false);
  const [uploadTargetType, setUploadTargetType] = useState<UploadTargetType>('firma');
  const [toast, setToast] = useState<{ 
    type: 'success' | 'error' | 'info'; 
    message: string;
    actionLabel?: string;
    onAction?: () => void;
  } | null>(null);

  // Initialize dataset from localStorage or fallback to audited initialDataset
  const [dataset, setDataset] = useState<GlobalDataset>(() => {
    try {
      const saved = localStorage.getItem(DATASET_STORAGE_KEY);
      if (saved) {
        return JSON.parse(saved);
      }
    } catch (e) {
      console.error('Error loading dataset from local storage', e);
    }
    return initialDataset;
  });

  // Save dataset to localStorage whenever it changes
  useEffect(() => {
    try {
      localStorage.setItem(DATASET_STORAGE_KEY, JSON.stringify(dataset));
    } catch (e) {
      console.error('Failed to save dataset in local storage', e);
    }
  }, [dataset]);

  // Toast auto dismiss
  useEffect(() => {
    if (!toast) return;
    const timer = setTimeout(() => {
      setToast(null);
    }, 6000);
    return () => clearTimeout(timer);
  }, [toast]);

  const showToast = (type: 'success' | 'error' | 'info', message: string, actionLabel?: string, onAction?: () => void) => {
    setToast({ type, message, actionLabel, onAction });
  };

  const handleOpenUploadFirma = () => {
    setUploadTargetType('firma');
    setIsUploadOpen(true);
  };

  const handleOpenUploadSocios = () => {
    setUploadTargetType('socios');
    setIsUploadOpen(true);
  };

  const handleUpdateFilter = (filter: ReportTimeFilter) => {
    const updated = {
      ...dataset,
      reportFilter: filter,
      updatedAt: new Date().toISOString(),
    };
    setDataset(updated);

    let filterMsg = 'Histórico Completo';
    if (filter.type === 'month') {
      filterMsg = `Mes: ${filter.month || 'SEPTIEMBRE'}`;
    } else if (filter.type === 'week') {
      filterMsg = `Semana: ${filter.weekId?.toUpperCase() || 'W2'} (hasta el 20 de septiembre)`;
    } else if (filter.type === 'range') {
      filterMsg = `Rango: ${filter.startDate || 'Inicio'} al ${filter.endDate || 'Fin'}`;
    }
    showToast('info', `Filtro temporal activo: ${filterMsg}`);
  };

  // Handle incoming data from Excel import
  const handleApplyData = async (result: {
    type: 'transactions' | 'monthly_summary' | 'socios';
    transactions?: TransactionRecord[];
    monthlyMetrics?: MonthlyMetric[];
    socios?: SocioRecord[];
    detectedMonth?: string;
    detectedCutoffDate?: string;
    dateRangeStr?: string;
    mode: 'append' | 'replace';
  }) => {
    let updatedDataset: GlobalDataset = { ...dataset, updatedAt: new Date().toISOString() };

    if (result.type === 'transactions' && result.transactions) {
      // REPORTE 1: Se actualiza EXCLUSIVAMENTE con el archivo de firmas
      const newTrxList = result.transactions;
      let finalTrx: TransactionRecord[] = [];

      if (result.mode === 'append') {
        const existingIds = new Set(dataset.transactions.map((t) => t.uniqueId));
        const filteredNew = newTrxList.filter((t) => !existingIds.has(t.uniqueId));
        finalTrx = [...filteredNew, ...dataset.transactions];
        showToast('success', `Reporte 1: Se añadieron ${filteredNew.length} nuevas emisiones de firmas.`);
      } else {
        finalTrx = newTrxList;
        showToast('success', `Reporte 1: Se cargaron ${newTrxList.length} emisiones de firmas.`);
      }

      updatedDataset.transactions = finalTrx;
      updatedDataset.monthlyMetrics = deriveMonthlyMetricsFromTransactions(finalTrx);
      updatedDataset.portfolioDurations = derivePortfolioDurationsFromTransactions(finalTrx);
      const weekly = deriveWeeklyBreakdownFromTransactions(finalTrx);
      updatedDataset.weeklyBreakdownType1 = weekly.type1;
      updatedDataset.solutionCategories = deriveSolutionCategoriesFromTransactions(finalTrx);
      updatedDataset.commercialCross = deriveCommercialCrossFromTransactions(finalTrx);

      // NO se modifica dataset.socios: el Reporte 2 permanece intacto con su archivo de socios

      if (result.detectedCutoffDate) {
        updatedDataset.cutoffDate = result.detectedCutoffDate;
      } else {
        const validDates = finalTrx
          .map((t) => new Date(t.date))
          .filter((d) => !isNaN(d.getTime()))
          .sort((a, b) => b.getTime() - a.getTime());
        if (validDates.length > 0) {
          updatedDataset.cutoffDate = formatSpanishDateCutoff(validDates[0]);
        }
      }

      if (result.detectedMonth) {
        updatedDataset.reportFilter = {
          type: 'month',
          month: result.detectedMonth,
        };
      }
    } else if (result.type === 'socios') {
      // REPORTE 2: Se actualiza EXCLUSIVAMENTE con el archivo de socios
      let incomingSocios = result.socios || [];

      // Si el archivo de socios trajo registros de socios directamente
      if (incomingSocios.length > 0) {
        if (result.mode === 'append') {
          const existingMap = new Map(dataset.socios.map((s) => [s.name.toLowerCase().trim(), { ...s }]));
          incomingSocios.forEach((newS) => {
            const key = newS.name.toLowerCase().trim();
            if (existingMap.has(key)) {
              const existing = existingMap.get(key)!;
              const combinedOps = (existing.operationsCount || 0) + (newS.operationsCount || 0);
              const combinedSales = parseFloat(((existing.totalSales || 0) + (newS.totalSales || 0)).toFixed(2));
              existingMap.set(key, {
                ...existing,
                totalSales: combinedSales,
                operationsCount: combinedOps,
                averageTicket: combinedOps > 0 
                  ? parseFloat((combinedSales / combinedOps).toFixed(2)) 
                  : combinedSales,
              });
            } else {
              existingMap.set(key, { ...newS });
            }
          });
          const merged = Array.from(existingMap.values()).sort((a, b) => b.totalSales - a.totalSales);
          merged.forEach((s, idx) => {
            s.rank = idx + 1;
            s.group = idx < 10 ? 'TOP 1-10' : 'TOP 11-20';
            if (idx === 0) s.note = 'Líder en Ventas';
          });
          updatedDataset.socios = merged;
          showToast('success', `Reporte 2: Se anexaron ${incomingSocios.length} registros a la cartera de socios.`);
        } else {
          updatedDataset.socios = incomingSocios;
          showToast('success', `Reporte 2: Se cargaron ${incomingSocios.length} socios en la cartera.`);
        }
      }

      // Si el archivo trajo liquidaciones semanales para el Reporte 2
      if (result.transactions && result.transactions.length > 0) {
        const weekly = deriveWeeklyBreakdownFromTransactions(result.transactions, result.detectedMonth);
        if (weekly.type2 && weekly.type2.length > 0) {
          updatedDataset.weeklyBreakdownType2 = weekly.type2;
        }
      }

      // NO se modifican transactions ni monthlyMetrics de Reporte 1: Reporte 1 permanece intacto

      if (result.detectedCutoffDate) {
        updatedDataset.cutoffDate = result.detectedCutoffDate;
      }

      // Automatically focus on the identified month!
      if (result.detectedMonth) {
        updatedDataset.reportFilter = {
          type: 'month',
          month: result.detectedMonth,
        };
      }

      showToast(
        'success',
        `Se actualizó el Reporte 2 con las ventas de socios (${incomingSocios.length} socios registrados).`,
        'Ver Reporte 2',
        () => setCurrentView('presentation2')
      );
    } else if (result.type === 'monthly_summary' && result.monthlyMetrics) {
      updatedDataset.monthlyMetrics = result.monthlyMetrics;
      showToast('success', 'Matriz de facturación mensual actualizada exitosamente.');
    }

    setDataset(updatedDataset);
  };

  const handleClearData = () => {
    const clearedDataset: GlobalDataset = {
      ...initialDataset,
      updatedAt: new Date().toISOString(),
      cutoffDate: 'Sin corte definido',
      reportFilter: { type: 'all' },
      monthlyMetrics: [],
      socios: [],
      transactions: [],
      weeklyBreakdownType1: [],
      weeklyBreakdownType2: [],
      portfolioDurations: [],
      solutionCategories: [],
      communities: [],
      commercialCross: {
        commercialTeamSales: 0,
        organicSales: 0,
        partnersSales: 0,
        channel1Name: 'Canal Directo',
        channel1Amount: 0,
        channel2Name: 'Red Distribuidores',
        channel2Status: 'Sin datos',
      },
    };
    setDataset(clearedDataset);
    try {
      localStorage.removeItem(DATASET_STORAGE_KEY);
      localStorage.setItem(DATASET_STORAGE_KEY, JSON.stringify(clearedDataset));
    } catch (e) {
      console.error('Error clearing dataset from localStorage', e);
    }
    showToast('info', 'Toda la información ha sido eliminada. El dashboard y los reportes están completamente vacíos.');
  };

  const handleAddManualTransaction = (trx: TransactionRecord) => {
    const updated = {
      ...dataset,
      transactions: [trx, ...dataset.transactions],
      updatedAt: new Date().toISOString(),
    };
    setDataset(updated);
    showToast('success', `Transacción ${trx.uniqueId} agregada.`);
  };

  const handleToggleSocioChannel = (socioName: string) => {
    const updatedSocios = dataset.socios.map((s) => {
      if (s.name.toLowerCase().trim() === socioName.toLowerCase().trim()) {
        const currentCh = s.channel || (s.role?.toLowerCase().includes('connect') ? 'Connectors' : 'UpConnect');
        const newChannel: 'UpConnect' | 'Connectors' = currentCh === 'UpConnect' ? 'Connectors' : 'UpConnect';
        const newRole = newChannel === 'Connectors' ? 'Distribuidor Connect' : 'UpConnect Directo';
        return { ...s, channel: newChannel, role: newRole };
      }
      return s;
    });

    const updatedTrx = dataset.transactions.map((t) => {
      if ((t.socio || '').toLowerCase().trim() === socioName.toLowerCase().trim()) {
        const currentCh = t.channel || 'Connectors';
        const newChannel: 'UpConnect' | 'Connectors' = currentCh === 'UpConnect' ? 'Connectors' : 'UpConnect';
        return { ...t, channel: newChannel };
      }
      return t;
    });

    const updated = {
      ...dataset,
      socios: updatedSocios,
      transactions: updatedTrx,
      updatedAt: new Date().toISOString(),
    };
    setDataset(updated);
    showToast('info', `Canal del socio "${socioName}" alternado.`);
  };

  const [isExportingCombinedPPTX, setIsExportingCombinedPPTX] = useState<boolean>(false);

  const handleDownloadCombinedPPTX = async () => {
    setIsExportingCombinedPPTX(true);
    try {
      await exportCombinedPresentationsToPPTX(dataset);
      showToast('success', 'Presentación consolidada de los 2 reportes descargada con éxito (.pptx)');
    } catch (e: any) {
      console.error('Error generating combined PPTX', e);
      showToast('error', 'Error al generar la presentación combinada PowerPoint.');
    } finally {
      setIsExportingCombinedPPTX(false);
    }
  };

  const handleUpdateDataset = (updates: Partial<GlobalDataset>) => {
    const updated: GlobalDataset = {
      ...dataset,
      ...updates,
      updatedAt: new Date().toISOString(),
    };
    setDataset(updated);
    try {
      localStorage.setItem(DATASET_STORAGE_KEY, JSON.stringify(updated));
    } catch (e) {
      console.error('Error saving updated dataset to localStorage', e);
    }
    showToast('success', 'Cambios guardados exitosamente.');
  };

  return (
    <div className="min-h-screen bg-[#070b14] text-slate-100 flex flex-col font-sans selection:bg-sky-500 selection:text-white">
      {/* Top Bar Navigation */}
      <Navigation
        currentView={currentView}
        onSelectView={setCurrentView}
        onDownloadCombinedPPTX={handleDownloadCombinedPPTX}
        isExportingPPTX={isExportingCombinedPPTX}
      />

      {/* Main View Port */}
      <main className="flex-1 flex flex-col">
        {currentView === 'dashboard' && (
          <DashboardView
            dataset={dataset}
            onOpenUploadFirma={handleOpenUploadFirma}
            onOpenUploadSocios={handleOpenUploadSocios}
            onClearData={handleClearData}
            onUpdateFilter={handleUpdateFilter}
            onAddTransaction={handleAddManualTransaction}
            onToggleSocioChannel={handleToggleSocioChannel}
          />
        )}

        {currentView === 'presentation1' && (
          <PresentationType1 
            dataset={dataset}
            onBackToDashboard={() => setCurrentView('dashboard')}
            onUpdateDataset={handleUpdateDataset}
          />
        )}

        {currentView === 'presentation2' && (
          <PresentationType2 
            dataset={dataset}
            onBackToDashboard={() => setCurrentView('dashboard')}
            onUpdateDataset={handleUpdateDataset}
          />
        )}
      </main>

      {/* Excel Upload Modal */}
      <ExcelUploadModal
        isOpen={isUploadOpen}
        onClose={() => setIsUploadOpen(false)}
        targetType={uploadTargetType}
        onChangeTargetType={setUploadTargetType}
        onApplyData={handleApplyData}
      />

      {/* Floating Notification Toast */}
      {toast && (
        <div className="no-print fixed bottom-5 right-5 z-50 flex items-center gap-3 px-4 py-3 rounded-xl bg-slate-900 border border-slate-700 shadow-2xl text-xs max-w-md animate-in slide-in-from-bottom-2 duration-200">
          {toast.type === 'success' && <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0" />}
          {toast.type === 'error' && <AlertCircle className="w-5 h-5 text-red-400 shrink-0" />}
          {toast.type === 'info' && <Info className="w-5 h-5 text-sky-400 shrink-0" />}
          <span className="flex-1 text-slate-200 font-medium">{toast.message}</span>
          {toast.actionLabel && toast.onAction && (
            <button
              onClick={() => {
                toast.onAction?.();
                setToast(null);
              }}
              className="px-2.5 py-1 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-[11px] whitespace-nowrap transition-colors shadow-sm"
            >
              {toast.actionLabel}
            </button>
          )}
          <button
            onClick={() => setToast(null)}
            className="text-slate-500 hover:text-white p-0.5"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      )}
    </div>
  );
}
