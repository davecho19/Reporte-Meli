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
  const [toast, setToast] = useState<{ type: 'success' | 'error' | 'info'; message: string } | null>(null);

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
    }, 4500);
    return () => clearTimeout(timer);
  }, [toast]);

  const showToast = (type: 'success' | 'error' | 'info', message: string) => {
    setToast({ type, message });
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
      // Derive monthly metrics and all presentation models from transactions so monthly charts, tables and reports are always synchronized
      updatedDataset.monthlyMetrics = deriveMonthlyMetricsFromTransactions(finalTrx);
      updatedDataset.portfolioDurations = derivePortfolioDurationsFromTransactions(finalTrx);
      const weekly = deriveWeeklyBreakdownFromTransactions(finalTrx);
      updatedDataset.weeklyBreakdownType1 = weekly.type1;
      updatedDataset.weeklyBreakdownType2 = weekly.type2;
      updatedDataset.solutionCategories = deriveSolutionCategoriesFromTransactions(finalTrx);
      updatedDataset.commercialCross = deriveCommercialCrossFromTransactions(finalTrx);

      // Compute cutoff date from the latest transaction date
      const validDates = finalTrx
        .map((t) => new Date(t.date))
        .filter((d) => !isNaN(d.getTime()))
        .sort((a, b) => b.getTime() - a.getTime());
      if (validDates.length > 0) {
        updatedDataset.cutoffDate = formatSpanishDateCutoff(validDates[0]);
      }

      // Always update socios with full detailed metrics (name, role, operations, totalSales, averageTicket, topPlan)
      if (result.socios && result.socios.length > 0) {
        if (result.mode === 'replace') {
          updatedDataset.socios = result.socios;
        } else {
          // Merge socios in append mode
          const socioMap = new Map<string, SocioRecord>();
          dataset.socios.forEach((s) => socioMap.set(s.name.toLowerCase().trim(), { ...s }));
          result.socios.forEach((ns) => {
            const key = ns.name.toLowerCase().trim();
            if (socioMap.has(key)) {
              const ex = socioMap.get(key)!;
              ex.totalSales = parseFloat((ex.totalSales + ns.totalSales).toFixed(2));
              ex.operationsCount = (ex.operationsCount || 0) + (ns.operationsCount || 1);
              ex.averageTicket = ex.operationsCount > 0 ? parseFloat((ex.totalSales / ex.operationsCount).toFixed(2)) : ex.totalSales;
              if (ns.topPlan) ex.topPlan = ns.topPlan;
              if (ns.role) ex.role = ns.role;
            } else {
              socioMap.set(key, ns);
            }
          });
          const reRanked = Array.from(socioMap.values()).sort((a, b) => b.totalSales - a.totalSales);
          reRanked.forEach((s, idx) => {
            s.rank = idx + 1;
            s.group = idx < 10 ? 'TOP 1-10' : 'TOP 11-20';
            if (idx === 0) s.note = 'Líder en Ventas';
          });
          updatedDataset.socios = reRanked;
        }
      } else {
        const socioMap: Record<string, { total: number; count: number; role: string; plans: Record<string, number> }> = {};
        finalTrx.forEach((t) => {
          const sName = t.socio || (t.channel === 'Connectors' ? 'Distribuidor Connect' : 'UpConnect Directo');
          if (!socioMap[sName]) {
            socioMap[sName] = { total: 0, count: 0, role: t.role || 'Distribuidor Connect', plans: {} };
          }
          socioMap[sName].total += t.value;
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
      }
    } else if (result.type === 'socios' && result.socios) {
      if (result.mode === 'append') {
        const socioMap = new Map<string, SocioRecord>();
        dataset.socios.forEach((s) => socioMap.set(s.name.toLowerCase().trim(), { ...s }));
        result.socios.forEach((ns) => {
          const key = ns.name.toLowerCase().trim();
          if (socioMap.has(key)) {
            const ex = socioMap.get(key)!;
            ex.totalSales = parseFloat((ex.totalSales + ns.totalSales).toFixed(2));
            ex.operationsCount = (ex.operationsCount || 0) + (ns.operationsCount || 1);
            ex.averageTicket = ex.operationsCount > 0 ? parseFloat((ex.totalSales / ex.operationsCount).toFixed(2)) : ex.totalSales;
            if (ns.topPlan) ex.topPlan = ns.topPlan;
            if (ns.role) ex.role = ns.role;
          } else {
            socioMap.set(key, ns);
          }
        });
        const reRanked = Array.from(socioMap.values()).sort((a, b) => b.totalSales - a.totalSales);
        reRanked.forEach((s, idx) => {
          s.rank = idx + 1;
          s.group = idx < 10 ? 'TOP 1-10' : 'TOP 11-20';
          if (idx === 0) s.note = 'Líder en Ventas';
        });
        updatedDataset.socios = reRanked;
        showToast('success', `Se sincronizaron ${result.socios.length} socios en la cartera.`);
      } else {
        updatedDataset.socios = result.socios;
        showToast('success', `Se actualizó la cartera con ${result.socios.length} socios.`);
      }

      // If transactions were extracted from the 9 columns of the socios file, incorporate them
      if (result.transactions && result.transactions.length > 0) {
        if (result.mode === 'append') {
          updatedDataset.transactions = [...result.transactions, ...dataset.transactions];
        } else {
          updatedDataset.transactions = result.transactions;
        }

        updatedDataset.monthlyMetrics = deriveMonthlyMetricsFromTransactions(updatedDataset.transactions);
        updatedDataset.portfolioDurations = derivePortfolioDurationsFromTransactions(updatedDataset.transactions);
        const weekly = deriveWeeklyBreakdownFromTransactions(updatedDataset.transactions);
        updatedDataset.weeklyBreakdownType1 = weekly.type1;
        updatedDataset.weeklyBreakdownType2 = weekly.type2;
        updatedDataset.solutionCategories = deriveSolutionCategoriesFromTransactions(updatedDataset.transactions);
        updatedDataset.commercialCross = deriveCommercialCrossFromTransactions(updatedDataset.transactions);

        const validDates = updatedDataset.transactions
          .map((t) => new Date(t.date))
          .filter((d) => !isNaN(d.getTime()))
          .sort((a, b) => b.getTime() - a.getTime());
        if (validDates.length > 0) {
          updatedDataset.cutoffDate = formatSpanishDateCutoff(validDates[0]);
        }
      }
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
