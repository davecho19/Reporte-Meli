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
      const newTrxList = result.transactions;
      let finalTrx: TransactionRecord[] = [];

      if (result.mode === 'append') {
        const existingIds = new Set(dataset.transactions.map((t) => t.uniqueId));
        const filteredNew = newTrxList.filter((t) => !existingIds.has(t.uniqueId));
        finalTrx = [...filteredNew, ...dataset.transactions];
        showToast('success', `Se añadieron ${filteredNew.length} nuevas emisiones de firmas al reporte.`);
      } else {
        finalTrx = newTrxList;
        showToast('success', `Se reemplazó el conjunto con ${newTrxList.length} emisiones de firmas.`);
      }

      updatedDataset.transactions = finalTrx;
      updatedDataset.monthlyMetrics = deriveMonthlyMetricsFromTransactions(finalTrx);
      updatedDataset.portfolioDurations = derivePortfolioDurationsFromTransactions(finalTrx);
      const weekly = deriveWeeklyBreakdownFromTransactions(finalTrx);
      updatedDataset.weeklyBreakdownType1 = weekly.type1;
      updatedDataset.weeklyBreakdownType2 = weekly.type2;
      updatedDataset.solutionCategories = deriveSolutionCategoriesFromTransactions(finalTrx);
      updatedDataset.commercialCross = deriveCommercialCrossFromTransactions(finalTrx);

      // Reconcile socios directly from final transactions to prevent double-counting
      const socioMap: Record<string, { total: number; count: number; role: string; plans: Record<string, number> }> = {};
      finalTrx.forEach((t) => {
        const sName = (t.socio || (t.channel === 'Connectors' ? 'Distribuidor Connect' : 'UpConnect Directo')).trim();
        if (!socioMap[sName]) {
          socioMap[sName] = { total: 0, count: 0, role: t.role || 'Distribuidor Connect', plans: {} };
        }
        socioMap[sName].total = parseFloat((socioMap[sName].total + t.value).toFixed(2));
        socioMap[sName].count += 1;
        if (t.role && t.role !== 'Distribuidor Connect') socioMap[sName].role = t.role;
        const planName = t.solutionCategory ? `${t.solutionCategory} (${t.duration})` : (t.duration || 'Plan Estándar');
        socioMap[sName].plans[planName] = (socioMap[sName].plans[planName] || 0) + 1;
      });

      const ranked: SocioRecord[] = Object.entries(socioMap)
        .map(([name, data]) => {
          let topPlan = 'Firma Electrónica (1 año)';
          let maxCount = -1;
          for (const [pName, count] of Object.entries(data.plans)) {
            if (count > maxCount) {
              maxCount = count;
              topPlan = pName;
            }
          }
          const averageTicket = data.count > 0 ? (data.total / data.count) : 0;
          return {
            rank: 0,
            name,
            totalSales: parseFloat(data.total.toFixed(2)),
            group: 'TOP 1-10' as const,
            operationsCount: data.count,
            averageTicket: parseFloat(averageTicket.toFixed(2)),
            topPlan,
            role: data.role || 'Distribuidor Connect',
            note: undefined,
          };
        })
        .sort((a, b) => b.totalSales - a.totalSales);

      ranked.forEach((s, idx) => {
        s.rank = idx + 1;
        s.group = idx < 10 ? 'TOP 1-10' : 'TOP 11-20';
        if (idx === 0) s.note = 'Líder en Ventas';
      });

      if (ranked.length > 0) {
        updatedDataset.socios = ranked;
      }

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
      let finalTrx: TransactionRecord[] = [];

      if (result.transactions && result.transactions.length > 0) {
        if (result.mode === 'append') {
          const existingIds = new Set(dataset.transactions.map((t) => t.uniqueId));
          const filteredNew = result.transactions.filter((t) => !existingIds.has(t.uniqueId));
          finalTrx = [...filteredNew, ...dataset.transactions];
        } else {
          finalTrx = result.transactions;
        }
        updatedDataset.transactions = finalTrx;
        updatedDataset.monthlyMetrics = deriveMonthlyMetricsFromTransactions(finalTrx);
        updatedDataset.portfolioDurations = derivePortfolioDurationsFromTransactions(finalTrx);
        const weekly = deriveWeeklyBreakdownFromTransactions(finalTrx);
        updatedDataset.weeklyBreakdownType1 = weekly.type1;
        updatedDataset.weeklyBreakdownType2 = weekly.type2;
        updatedDataset.solutionCategories = deriveSolutionCategoriesFromTransactions(finalTrx);
        updatedDataset.commercialCross = deriveCommercialCrossFromTransactions(finalTrx);

        // Derive socios strictly from final transactions to guarantee 100% exact math per month
        const socioMap: Record<string, {
          total: number;
          count: number;
          role: string;
          channel?: 'UpConnect' | 'Connectors';
          upOps: number;
          coOps: number;
          plans: Record<string, number>;
        }> = {};

        finalTrx.forEach((t) => {
          const sName = (t.socio || (t.channel === 'Connectors' ? 'Distribuidor Connect' : 'UpConnect Directo')).trim();
          if (!socioMap[sName]) {
            socioMap[sName] = { 
              total: 0, 
              count: 0, 
              role: t.role || (t.channel === 'Connectors' ? 'Distribuidor Connect' : 'UpConnect Directo'), 
              channel: t.channel,
              upOps: 0,
              coOps: 0,
              plans: {} 
            };
          }
          socioMap[sName].total = parseFloat((socioMap[sName].total + t.value).toFixed(2));
          socioMap[sName].count += 1;
          if (t.channel === 'Connectors') socioMap[sName].coOps += 1;
          else socioMap[sName].upOps += 1;

          if (t.role && t.role !== 'Distribuidor Connect' && t.role !== 'UpConnect Directo') socioMap[sName].role = t.role;
          const planName = t.solutionCategory
            ? (t.solutionCategory.includes('(') ? t.solutionCategory : (t.duration && t.duration !== 'Un año' ? `${t.solutionCategory} (${t.duration})` : t.solutionCategory))
            : (t.duration || 'UP INTERMEDIO ($17.25)');
          socioMap[sName].plans[planName] = (socioMap[sName].plans[planName] || 0) + 1;
        });

        const ranked: SocioRecord[] = Object.entries(socioMap)
          .map(([name, data]) => {
            let topPlan = 'UP INTERMEDIO ($17.25)';
            let maxCount = -1;
            for (const [pName, count] of Object.entries(data.plans)) {
              if (count > maxCount) {
                maxCount = count;
                topPlan = pName;
              }
            }
            const averageTicket = data.count > 0 ? (data.total / data.count) : 0;
            const channel: 'UpConnect' | 'Connectors' = data.coOps > data.upOps 
              ? 'Connectors' 
              : data.upOps > data.coOps 
              ? 'UpConnect' 
              : (data.channel || 'Connectors');
            const role = data.role || (channel === 'Connectors' ? 'Distribuidor Connect' : 'UpConnect Directo');

            return {
              rank: 0,
              name,
              totalSales: parseFloat(data.total.toFixed(2)),
              group: 'TOP 1-10' as const,
              operationsCount: data.count,
              averageTicket: parseFloat(averageTicket.toFixed(2)),
              topPlan,
              role,
              channel,
              note: undefined,
            };
          })
          .sort((a, b) => b.totalSales - a.totalSales);

        ranked.forEach((s, idx) => {
          s.rank = idx + 1;
          s.group = idx < 10 ? 'TOP 1-10' : 'TOP 11-20';
          if (idx === 0) s.note = 'Líder en Ventas';
        });

        updatedDataset.socios = ranked;
      } else if (result.socios && result.socios.length > 0) {
        updatedDataset.socios = result.socios;
      }

      if (result.detectedCutoffDate) {
        updatedDataset.cutoffDate = result.detectedCutoffDate;
      } else if (finalTrx.length > 0) {
        const validDates = finalTrx
          .map((t) => new Date(t.date))
          .filter((d) => !isNaN(d.getTime()))
          .sort((a, b) => b.getTime() - a.getTime());
        if (validDates.length > 0) {
          updatedDataset.cutoffDate = formatSpanishDateCutoff(validDates[0]);
        }
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
        `Se sincronizó el Reporte 2 con las ventas de socios del mes de ${result.detectedMonth || 'SEPTIEMBRE'} (${result.transactions?.length || 0} emisiones conciliadas).`,
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

  return (
    <div className="min-h-screen bg-[#070b14] text-slate-100 flex flex-col font-sans selection:bg-sky-500 selection:text-white">
      {/* Top Bar Navigation */}
      <Navigation
        currentView={currentView}
        onSelectView={setCurrentView}
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
          />
        )}

        {currentView === 'presentation2' && (
          <PresentationType2 
            dataset={dataset}
            onBackToDashboard={() => setCurrentView('dashboard')}
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
