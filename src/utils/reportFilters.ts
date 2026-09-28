import { 
  GlobalDataset, 
  ReportTimeFilter, 
  TransactionRecord, 
  MonthlyMetric, 
  SocioRecord, 
  PortfolioDuration, 
  WeeklySalesBreakdown, 
  SolutionCategory, 
  CommercialCrossPhase 
} from '../types';

export interface FilteredReportView {
  titleReport1: string;
  titleReport2: string;
  subtitleDate: string;
  periodLabel: string;
  transactions: TransactionRecord[];
  socios: SocioRecord[];
  totalUpSales: number;
  totalCoSales: number;
  totalSales: number;
  totalUpCount: number;
  totalCoCount: number;
  totalCount: number;
  upSalesPct: string;
  coSalesPct: string;
  netSales: number;
  ivaAmount: number;
  grossSales: number;
  monthlyMetrics: MonthlyMetric[];
  isFiltered: boolean;
}

const MONTH_NAMES_ES = [
  'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
  'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'
];

export const formatSpanishDate = (dateStr: string): string => {
  if (!dateStr) return '';
  const parts = dateStr.split(/[-/]/);
  if (parts.length === 3) {
    const y = parts[0].length === 4 ? parseInt(parts[0], 10) : parseInt(parts[2], 10);
    const m = parts[0].length === 4 ? parseInt(parts[1], 10) : parseInt(parts[0], 10);
    const d = parts[0].length === 4 ? parseInt(parts[2], 10) : parseInt(parts[1], 10);
    const monthName = MONTH_NAMES_ES[m - 1] || 'Septiembre';
    return `${d} de ${monthName.toLowerCase()} de ${y}`;
  }
  return dateStr;
};

export const deriveMonthlyMetricsFromTransactions = (transactions: TransactionRecord[]): MonthlyMetric[] => {
  if (transactions.length === 0) return [];
  const map: Record<number, { upSales: number; coSales: number; upCount: number; coCount: number }> = {};
  
  transactions.forEach((t) => {
    let m = 9;
    if (t.date) {
      const dateStr = String(t.date).trim();
      const parts = dateStr.split(/[-/]/);
      if (parts.length >= 2) {
        if (parts[0].length === 4) {
          // Format YYYY-MM-DD
          m = parseInt(parts[1], 10);
        } else if (parts.length >= 3 && parts[2].length === 4) {
          // Format DD/MM/YYYY or DD-MM-YYYY -> parts[1] is the month!
          m = parseInt(parts[1], 10);
        } else {
          // Fallback to parts[1] or parts[0]
          const cand = parseInt(parts[1], 10);
          m = !isNaN(cand) && cand >= 1 && cand <= 12 ? cand : parseInt(parts[0], 10);
        }
      }
    }
    if (isNaN(m) || m < 1 || m > 12) m = 9;

    if (!map[m]) {
      map[m] = { upSales: 0, coSales: 0, upCount: 0, coCount: 0 };
    }
    if (t.channel === 'Connectors') {
      map[m].coSales += t.value;
      map[m].coCount += 1;
    } else {
      map[m].upSales += t.value;
      map[m].upCount += 1;
    }
  });

  return Object.keys(map)
    .map(Number)
    .sort((a, b) => a - b)
    .map((mIndex) => ({
      month: MONTH_NAMES_ES[mIndex - 1] || 'MES',
      monthIndex: mIndex,
      upconnectSales: parseFloat(map[mIndex].upSales.toFixed(2)),
      connectorsSales: parseFloat(map[mIndex].coSales.toFixed(2)),
      upconnectCount: map[mIndex].upCount,
      connectorsCount: map[mIndex].coCount,
    }));
};

export const derivePortfolioDurationsFromTransactions = (transactions: TransactionRecord[]): PortfolioDuration[] => {
  if (transactions.length === 0) return [];

  const counts: Record<string, { UpConnect: number; Connectors: number }> = {
    '15 días': { UpConnect: 0, Connectors: 0 },
    'Un año': { UpConnect: 0, Connectors: 0 },
    'Dos años': { UpConnect: 0, Connectors: 0 },
    'Tres años': { UpConnect: 0, Connectors: 0 },
    'Cuatro años': { UpConnect: 0, Connectors: 0 },
    'Cinco años': { UpConnect: 0, Connectors: 0 },
  };

  let upTotal = 0;
  let coTotal = 0;

  transactions.forEach((t) => {
    const durLower = (t.duration || '').toLowerCase();
    let norm = 'Un año';
    if (durLower.includes('15') || durLower.includes('quince')) norm = '15 días';
    else if (durLower.includes('5') || durLower.includes('cinco')) norm = 'Cinco años';
    else if (durLower.includes('4') || durLower.includes('cuatro')) norm = 'Cuatro años';
    else if (durLower.includes('3') || durLower.includes('tres')) norm = 'Tres años';
    else if (durLower.includes('2') || durLower.includes('dos')) norm = 'Dos años';
    else norm = 'Un año';

    if (!counts[norm]) {
      counts[norm] = { UpConnect: 0, Connectors: 0 };
    }

    if (t.channel === 'Connectors') {
      counts[norm].Connectors += 1;
      coTotal += 1;
    } else {
      counts[norm].UpConnect += 1;
      upTotal += 1;
    }
  });

  const result: PortfolioDuration[] = [];

  // UpConnect durations
  const upDurations = Object.entries(counts)
    .filter(([_, data]) => data.UpConnect > 0)
    .sort((a, b) => b[1].UpConnect - a[1].UpConnect);

  upDurations.forEach(([duration, data], idx) => {
    result.push({
      duration,
      channel: 'UpConnect',
      count: data.UpConnect,
      percentage: upTotal > 0 ? parseFloat(((data.UpConnect / upTotal) * 100).toFixed(1)) : 0,
      label: idx === 0 ? 'Líder' : undefined,
    });
  });

  // Connectors durations
  const coDurations = Object.entries(counts)
    .filter(([_, data]) => data.Connectors > 0)
    .sort((a, b) => b[1].Connectors - a[1].Connectors);

  coDurations.forEach(([duration, data], idx) => {
    result.push({
      duration,
      channel: 'Connectors',
      count: data.Connectors,
      percentage: coTotal > 0 ? parseFloat(((data.Connectors / coTotal) * 100).toFixed(1)) : 0,
      label: idx === 0 ? 'Líder' : undefined,
    });
  });

  return result;
};

export const deriveWeeklyBreakdownFromTransactions = (
  transactions: TransactionRecord[]
): {
  type1: WeeklySalesBreakdown[];
  type2: {
    weekName: string;
    dateRange: string;
    amount: number;
    percentage: number;
    description: string;
    operationsCount?: number;
  }[];
} => {
  if (transactions.length === 0) return { type1: [], type2: [] };

  const weekDefs = [
    { id: 'w1', weekName: 'Semana 1', dateRange: '01 al 06 de septiembre', startDay: 1, endDay: 6 },
    { id: 'w2', weekName: 'Semana 2', dateRange: '07 al 13 de septiembre', startDay: 7, endDay: 13 },
    { id: 'w3', weekName: 'Semana 3', dateRange: '14 al 20 de septiembre', startDay: 14, endDay: 20 },
    { id: 'w4', weekName: 'Semana 4', dateRange: '21 al 30 de septiembre', startDay: 21, endDay: 31 },
  ];

  const totalMonthSales = transactions.reduce((acc, t) => acc + (t.value || 0), 0) || 1;

  const type1: WeeklySalesBreakdown[] = [];
  const type2: {
    weekName: string;
    dateRange: string;
    amount: number;
    percentage: number;
    description: string;
    operationsCount?: number;
  }[] = [];

  weekDefs.forEach((w) => {
    let upAmount = 0;
    let upCount = 0;
    let coAmount = 0;
    let coCount = 0;

    transactions.forEach((t) => {
      const parts = t.date.split(/[-/]/);
      let day = 1;
      if (parts.length === 3) {
        day = parts[0].length === 4 ? parseInt(parts[2], 10) : parseInt(parts[0], 10);
      }
      if (isNaN(day)) day = 1;

      if (day >= w.startDay && day <= w.endDay) {
        if (t.channel === 'Connectors') {
          coAmount += t.value;
          coCount += 1;
        } else {
          upAmount += t.value;
          upCount += 1;
        }
      }
    });

    const totalAmount = parseFloat((upAmount + coAmount).toFixed(2));
    const totalCount = upCount + coCount;
    const pct = parseFloat(((totalAmount / totalMonthSales) * 100).toFixed(1));

    type1.push({
      id: w.id,
      weekName: w.weekName,
      dateRange: w.dateRange,
      totalAmount,
      totalCount,
      upconnectAmount: parseFloat(upAmount.toFixed(2)),
      upconnectCount: upCount,
      connectorsAmount: parseFloat(coAmount.toFixed(2)),
      connectorsCount: coCount,
    });

    type2.push({
      weekName: w.weekName,
      dateRange: w.dateRange,
      amount: totalAmount,
      percentage: pct,
      description: `${totalCount} emisiones registradas en este corte operativo (${totalAmount.toFixed(2)} USD).`,
      operationsCount: totalCount,
    });
  });

  return { type1, type2 };
};

export const formatSpanishDateCutoff = (d: Date): string => {
  const months = [
    'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
    'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
  ];
  return `${d.getDate()} de ${months[d.getMonth()]} de ${d.getFullYear()}`;
};

export const deriveSolutionCategoriesFromTransactions = (transactions: TransactionRecord[]): SolutionCategory[] => {
  if (transactions.length === 0) return [];

  const catMap: Record<string, number> = {};
  let totalSales = 0;

  transactions.forEach((t) => {
    const cat = t.solutionCategory || 'ERP Contables y Módulos';
    catMap[cat] = (catMap[cat] || 0) + (t.value || 0);
    totalSales += t.value || 0;
  });

  if (totalSales === 0) totalSales = 1;

  const badges = ['Alta Demanda', 'Mayor Margen', 'Ticket Accesible'];
  return Object.entries(catMap)
    .sort((a, b) => b[1] - a[1])
    .map(([title, amount], idx) => ({
      id: `cat-${idx + 1}`,
      title,
      badge: badges[idx % badges.length],
      amount: parseFloat(amount.toFixed(2)),
      percentage: parseFloat(((amount / totalSales) * 100).toFixed(1)),
      plansDescription: `Distribución de planes y emisiones de ${title}.`,
    }));
};

export const deriveCommercialCrossFromTransactions = (transactions: TransactionRecord[]): CommercialCrossPhase => {
  let partnersSales = 0;
  let commercialTeamSales = 0;
  let organicSales = 0;

  transactions.forEach((t) => {
    if (t.channel === 'Connectors') {
      partnersSales += t.value;
    } else if (t.role && t.role.toLowerCase().includes('direct')) {
      commercialTeamSales += t.value;
    } else {
      organicSales += t.value;
    }
  });

  return {
    commercialTeamSales: parseFloat(commercialTeamSales.toFixed(2)),
    organicSales: parseFloat(organicSales.toFixed(2)),
    partnersSales: parseFloat(partnersSales.toFixed(2)),
    channel1Name: 'Socios UpConnect',
    channel1Amount: parseFloat(partnersSales.toFixed(2)),
    channel2Name: 'Canal Directo UpConnect',
    channel2Status: 'Verificado',
  };
};

/**
 * Computes filtered stats and dynamic executive titles based on filter selection:
 * - month: "Reporte del mes de Septiembre 2026"
 * - week: "hasta el 20 de septiembre semana 2 (del 07 al 13 de septiembre)"
 * - range: "Reporte del 1 al 20 de septiembre de 2026"
 */
export const computeReportView = (dataset: GlobalDataset): FilteredReportView => {
  const filter: ReportTimeFilter = dataset.reportFilter || { type: 'all' };

  let titleReport1 = 'Reporte Gerencial de Ventas: Upconnect / Connectors';
  let titleReport2 = 'REPORTE GERENCIAL VENTAS HASTA EL 20 DE SEPTIEMBRE DE 2026';
  let subtitleDate = dataset.cutoffDate || '20 de Septiembre de 2026';
  let periodLabel = 'Consolidado General 2026';
  let isFiltered = false;

  let filteredTrx = [...dataset.transactions];
  let filteredSocios = [...dataset.socios];

  if (filter.type === 'month' && filter.month) {
    isFiltered = true;
    const monthName = filter.month.toUpperCase();
    titleReport1 = `Reporte del mes de ${monthName.charAt(0) + monthName.slice(1).toLowerCase()} 2026: Upconnect / Connectors`;
    titleReport2 = `REPORTE GERENCIAL VENTAS DEL MES DE ${monthName} 2026`;
    subtitleDate = `Mes completo de ${monthName.charAt(0) + monthName.slice(1).toLowerCase()} 2026`;
    periodLabel = `Mes de ${monthName.charAt(0) + monthName.slice(1).toLowerCase()}`;

    // Filter transactions by month index
    const mIndex = MONTH_NAMES_ES.indexOf(monthName) + 1;
    if (mIndex > 0) {
      filteredTrx = dataset.transactions.filter((t) => {
        const parts = t.date.split(/[-/]/);
        const m = parts[0].length === 4 ? parseInt(parts[1], 10) : parseInt(parts[0], 10);
        return m === mIndex;
      });
    }
  } else if (filter.type === 'week' && filter.weekId) {
    isFiltered = true;
    const weekMap: Record<string, { name: string; range: string; start: string; end: string }> = {
      w1: { name: 'Semana 1', range: '01 al 06 de septiembre', start: '2026-09-01', end: '2026-09-06' },
      w2: { name: 'Semana 2', range: '07 al 13 de septiembre', start: '2026-09-07', end: '2026-09-13' },
      w3: { name: 'Semana 3', range: '14 al 20 de septiembre', start: '2026-09-14', end: '2026-09-20' },
      w4: { name: 'Semana 4', range: '21 al 30 de septiembre', start: '2026-09-21', end: '2026-09-30' },
    };
    const targetWeek = weekMap[filter.weekId] || weekMap.w2;

    titleReport1 = `Reporte hasta el 20 de septiembre ${targetWeek.name} (del ${targetWeek.range}): Upconnect / Connectors`;
    titleReport2 = `REPORTE GERENCIAL VENTAS HASTA EL 20 DE SEPTIEMBRE ${targetWeek.name.toUpperCase()} (DEL ${targetWeek.range.toUpperCase()})`;
    subtitleDate = `${targetWeek.name} (${targetWeek.range} de 2026)`;
    periodLabel = `${targetWeek.name} (${targetWeek.range})`;

    filteredTrx = dataset.transactions.filter((t) => {
      return t.date >= targetWeek.start && t.date <= targetWeek.end;
    });
  } else if (filter.type === 'range' && filter.startDate && filter.endDate) {
    isFiltered = true;
    const sParts = filter.startDate.split('-');
    const eParts = filter.endDate.split('-');
    const startDay = parseInt(sParts[2] || '1', 10);
    const endDay = parseInt(eParts[2] || '20', 10);
    const endMonthIdx = parseInt(eParts[1] || '9', 10);
    const endMonthName = MONTH_NAMES_ES[endMonthIdx - 1]?.toLowerCase() || 'septiembre';

    titleReport1 = `Reporte del ${startDay} al ${endDay} de ${endMonthName}: Upconnect / Connectors`;
    titleReport2 = `REPORTE GERENCIAL VENTAS DEL ${startDay} AL ${endDay} DE ${endMonthName.toUpperCase()} DE 2026`;
    subtitleDate = `Del ${startDay} al ${endDay} de ${endMonthName} de 2026`;
    periodLabel = `Rango del ${startDay} al ${endDay} de ${endMonthName}`;

    filteredTrx = dataset.transactions.filter((t) => {
      return t.date >= filter.startDate! && t.date <= filter.endDate!;
    });
  }

  // If filtered transactions exist, re-tally stats; otherwise fallback to month metrics
  let totalUpSales = 0;
  let totalCoSales = 0;
  let totalUpCount = 0;
  let totalCoCount = 0;

  if (filteredTrx.length > 0) {
    filteredTrx.forEach((t) => {
      if (t.channel === 'Connectors') {
        totalCoSales += t.value;
        totalCoCount += 1;
      } else {
        totalUpSales += t.value;
        totalUpCount += 1;
      }
    });

    // Recompute socios ranking based on filtered transactions if available
    const socioMap: Record<string, { total: number; count: number; role?: string; plans: Record<string, number> }> = {};
    filteredTrx.forEach((t) => {
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
      if (idx === 0) s.note = 'Líder en Periodo';
    });

    if (ranked.length > 0 && (filteredSocios.length === 0 || isFiltered)) {
      filteredSocios = ranked;
    }
  } else {
    // Standard totals from monthly metrics
    totalUpSales = dataset.monthlyMetrics.reduce((a, m) => a + m.upconnectSales, 0);
    totalCoSales = dataset.monthlyMetrics.reduce((a, m) => a + m.connectorsSales, 0);
    totalUpCount = dataset.monthlyMetrics.reduce((a, m) => a + m.upconnectCount, 0);
    totalCoCount = dataset.monthlyMetrics.reduce((a, m) => a + m.connectorsCount, 0);
  }

  let totalSales = totalUpSales + totalCoSales;
  let totalCount = totalUpCount + totalCoCount;

  // If transactions didn't provide sales (e.g. socios summary upload), synchronize from socios!
  const sociosTotalSales = filteredSocios.reduce((acc, s) => acc + (s.totalSales || 0), 0);
  const sociosTotalCount = filteredSocios.reduce((acc, s) => acc + (s.operationsCount || 0), 0);

  if (totalSales === 0 && sociosTotalSales > 0) {
    totalSales = sociosTotalSales;
    totalCount = sociosTotalCount > 0 ? sociosTotalCount : filteredSocios.length;

    let upS = 0;
    let coS = 0;
    let upC = 0;
    let coC = 0;

    filteredSocios.forEach((s) => {
      const isCo = (s.role || '').toLowerCase().includes('connect') || (s.role || '').toLowerCase().includes('distribuidor');
      if (isCo) {
        coS += s.totalSales || 0;
        coC += s.operationsCount || 1;
      } else {
        upS += s.totalSales || 0;
        upC += s.operationsCount || 1;
      }
    });

    totalUpSales = upS;
    totalCoSales = coS;
    totalUpCount = upC;
    totalCoCount = coC;
  }

  const upSalesPct = totalSales > 0 ? ((totalUpSales / totalSales) * 100).toFixed(1) : '0.0';
  const coSalesPct = totalSales > 0 ? ((totalCoSales / totalSales) * 100).toFixed(1) : '0.0';

  // Calculate actual netSales from filtered/aggregated transactions or socios
  const netSales = totalSales;
  const ivaAmount = netSales * 0.15;
  const grossSales = netSales + ivaAmount;

  const effectiveMonthlyMetrics = dataset.monthlyMetrics.length > 0
    ? dataset.monthlyMetrics
    : deriveMonthlyMetricsFromTransactions(dataset.transactions);

  return {
    titleReport1,
    titleReport2,
    subtitleDate,
    periodLabel,
    transactions: filteredTrx,
    socios: filteredSocios,
    totalUpSales,
    totalCoSales,
    totalSales,
    totalUpCount,
    totalCoCount,
    totalCount,
    upSalesPct,
    coSalesPct,
    netSales,
    ivaAmount,
    grossSales,
    monthlyMetrics: effectiveMonthlyMetrics,
    isFiltered,
  };
};
