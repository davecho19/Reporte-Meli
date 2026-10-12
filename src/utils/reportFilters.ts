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

export interface SociosChannelBreakdown {
  channel: 'UpConnect' | 'Connectors';
  channelLabel: string;
  count: number;             // Cantidad de socios (conteo exacto)
  percentage: number;        // % de socios sobre el total
  operationsCount: number;   // Cantidad de ventas/firmas realizadas
  totalSales: number;        // Monto total facturado ($ USD)
  averageTicket: number;     // Ticket promedio ($ USD)
  topPlan?: string;
  socios: SocioRecord[];
}

export interface SociosChannelSummary {
  upconnect: SociosChannelBreakdown;
  connectors: SociosChannelBreakdown;
  totalSociosCount: number;
  totalOperationsCount: number;
  totalSales: number;
  averageTicket: number;
}

export interface WeeklyDifferentiatorItem {
  weekId: string;
  weekName: string;
  dateRange: string;
  totalSales: number;
  totalCount: number;
  facturacionSales: number;
  facturacionCount: number;
  erpSales: number;
  erpCount: number;
  contadoresSales: number;
  contadoresCount: number;
}

export interface FilteredReportView {
  titleReport1: string;
  titleReport2: string;
  subtitleDate: string;
  periodLabel: string;
  cutoffDateReport1: string;
  cutoffDateReport2: string;
  activeMonthName?: string;
  activeMonthIndex?: number;
  transactions: TransactionRecord[];
  socios: SocioRecord[];
  sociosSummary: SociosChannelSummary;
  firmaSocios: SocioRecord[];
  firmaSociosSummary: SociosChannelSummary;
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
  sistemasMonthlyMetrics: MonthlyMetric[];
  isFiltered: boolean;
  weeklyBreakdownType1: WeeklySalesBreakdown[];
  weeklyBreakdownType2: {
    weekName: string;
    dateRange: string;
    amount: number;
    percentage: number;
    description: string;
    operationsCount?: number;
  }[];
  weeklyDifferentiators: WeeklyDifferentiatorItem[];
  solutionCategoriesReport2?: SolutionCategory[];
  commercialCrossReport2?: CommercialCrossPhase;
}

export const MONTH_NAMES_ES = [
  'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
  'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'
];

/**
 * Robustly extracts the month index (1-12) and standardized month name ('SEPTIEMBRE')
 * from any TransactionRecord using its explicit month fields, notes tag, or date string.
 */
export const extractTransactionMonth = (t: TransactionRecord): { monthIndex: number; monthName: string } => {
  if (t.monthIndex && t.monthIndex >= 1 && t.monthIndex <= 12) {
    return {
      monthIndex: t.monthIndex,
      monthName: t.month || MONTH_NAMES_ES[t.monthIndex - 1],
    };
  }

  const monthLookup: Record<string, number> = {
    enero: 1, january: 1, ene: 1, jan: 1,
    febrero: 2, february: 2, feb: 2,
    marzo: 3, march: 3, mar: 3,
    abril: 4, april: 4, abr: 4, apr: 4,
    mayo: 5, may: 5,
    junio: 6, june: 6, jun: 6,
    julio: 7, july: 7, jul: 7,
    agosto: 8, august: 8, ago: 8, aug: 8,
    septiembre: 9, setiembre: 9, september: 9, sep: 9, sept: 9, set: 9,
    octubre: 10, october: 10, oct: 10,
    noviembre: 11, november: 11, nov: 11,
    diciembre: 12, december: 12, dic: 12, dec: 12,
  };

  if (t.month) {
    const cleanMonth = t.month.toLowerCase().trim();
    if (monthLookup[cleanMonth]) {
      const idx = monthLookup[cleanMonth];
      return { monthIndex: idx, monthName: MONTH_NAMES_ES[idx - 1] };
    }
  }

  if (t.notes) {
    const lowerNotes = t.notes.toLowerCase();
    for (const [name, idx] of Object.entries(monthLookup)) {
      if (lowerNotes.includes(name)) {
        return { monthIndex: idx, monthName: MONTH_NAMES_ES[idx - 1] };
      }
    }
  }

  if (t.date) {
    const dateStr = String(t.date).trim();
    const parts = dateStr.split(/[-/]/);
    if (parts.length >= 2) {
      let m = -1;
      if (parts[0].length === 4) {
        // Format YYYY-MM-DD
        m = parseInt(parts[1], 10);
      } else if (parts.length >= 3 && parts[2].length === 4) {
        // Format DD/MM/YYYY or MM/DD/YYYY
        const p1 = parseInt(parts[0], 10);
        const p2 = parseInt(parts[1], 10);
        if (p1 > 12 && p2 <= 12) m = p2;
        else if (p2 > 12 && p1 <= 12) m = p1;
        else m = p2; // Default Latin DD/MM/YYYY
      } else {
        const cand = parseInt(parts[1], 10);
        m = !isNaN(cand) && cand >= 1 && cand <= 12 ? cand : parseInt(parts[0], 10);
      }
      if (!isNaN(m) && m >= 1 && m <= 12) {
        return { monthIndex: m, monthName: MONTH_NAMES_ES[m - 1] };
      }
    }
  }

  return { monthIndex: 9, monthName: 'SEPTIEMBRE' };
};

export const formatSpanishDate = (dateStr: string): string => {
  if (!dateStr) return '';
  const parts = dateStr.split(/[-/]/);
  if (parts.length === 3) {
    let y = 2026, m = 9, d = 1;
    if (parts[0].length === 4) {
      // YYYY-MM-DD
      y = parseInt(parts[0], 10);
      m = parseInt(parts[1], 10);
      d = parseInt(parts[2], 10);
    } else {
      // DD/MM/YYYY or MM/DD/YYYY
      y = parseInt(parts[2], 10);
      const p0 = parseInt(parts[0], 10);
      const p1 = parseInt(parts[1], 10);
      if (p0 > 12 && p1 <= 12) {
        // Definitely DD/MM/YYYY
        d = p0;
        m = p1;
      } else if (p1 > 12 && p0 <= 12) {
        // MM/DD/YYYY
        m = p0;
        d = p1;
      } else {
        // Standard Latin DD/MM/YYYY
        d = p0;
        m = p1;
      }
    }
    const monthName = MONTH_NAMES_ES[m - 1] || 'Septiembre';
    return `${d} de ${monthName.toLowerCase()} de ${y}`;
  }
  return dateStr;
};

export const deriveMonthlyMetricsFromTransactions = (transactions: TransactionRecord[]): MonthlyMetric[] => {
  if (transactions.length === 0) return [];
  const map: Record<number, { upSales: number; coSales: number; upCount: number; coCount: number }> = {};

  transactions.forEach((t) => {
    const { monthIndex: m } = extractTransactionMonth(t);

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

export const extractTransactionDayAndMonth = (
  t: TransactionRecord
): { day: number; monthIndex: number; monthName: string; year: number } => {
  let day = 1;
  let mIndex = 9;
  let year = 2026;

  if (t.monthIndex && t.monthIndex >= 1 && t.monthIndex <= 12) {
    mIndex = t.monthIndex;
  } else if (t.month || t.notes) {
    const extracted = extractTransactionMonth(t);
    if (extracted.monthIndex >= 1 && extracted.monthIndex <= 12) {
      mIndex = extracted.monthIndex;
    }
  }

  if (t.date) {
    const raw = String(t.date).trim().split(/\s+/)[0];
    const parts = raw.split(/[-/]/);
    if (parts.length >= 3) {
      if (parts[0].length === 4) {
        // Format YYYY-MM-DD
        year = parseInt(parts[0], 10) || 2026;
        mIndex = parseInt(parts[1], 10) || mIndex;
        day = parseInt(parts[2], 10) || 1;
      } else {
        // Format DD/MM/YYYY or MM/DD/YYYY
        year = parseInt(parts[2], 10) || 2026;
        if (year < 100) year += 2000;
        const p0 = parseInt(parts[0], 10);
        const p1 = parseInt(parts[1], 10);
        if (p0 > 12 && p1 <= 12) {
          day = p0;
          mIndex = p1;
        } else if (p1 > 12 && p0 <= 12) {
          day = p1;
          mIndex = p0;
        } else {
          // Standard Latin DD/MM/YYYY
          day = p0;
          mIndex = p1;
        }
      }
    } else if (parts.length === 1 && /^\d{1,2}$/.test(parts[0])) {
      day = parseInt(parts[0], 10) || 1;
    }
  }

  if (day < 1 || isNaN(day)) day = 1;
  if (day > 31) day = 31;
  if (mIndex < 1 || mIndex > 12 || isNaN(mIndex)) mIndex = 9;

  return { day, monthIndex: mIndex, monthName: MONTH_NAMES_ES[mIndex - 1], year };
};

export const deriveWeeklyBreakdownFromTransactions = (
  transactions: TransactionRecord[],
  targetMonthIndexOrName?: number | string
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
  const now = new Date();
  const currentActualMonthIdx = now.getMonth() + 1;
  const currentMonthFallback = MONTH_NAMES_ES[now.getMonth()] || 'OCTUBRE';

  // 1. Identify target month (defaults to current actual month)
  let targetMonthIndex = currentActualMonthIdx;
  if (typeof targetMonthIndexOrName === 'number' && targetMonthIndexOrName >= 1 && targetMonthIndexOrName <= 12) {
    targetMonthIndex = targetMonthIndexOrName;
  } else if (typeof targetMonthIndexOrName === 'string' && targetMonthIndexOrName.trim()) {
    const cleanStr = targetMonthIndexOrName.toUpperCase().trim();
    const idx = MONTH_NAMES_ES.indexOf(cleanStr) + 1;
    if (idx > 0) {
      targetMonthIndex = idx;
    } else {
      const found = MONTH_NAMES_ES.findIndex((m) => m.includes(cleanStr) || cleanStr.includes(m));
      if (found !== -1) targetMonthIndex = found + 1;
    }
  }

  const targetMonthNameLower = (MONTH_NAMES_ES[targetMonthIndex - 1] || currentMonthFallback).toLowerCase();
  const daysInMonthCalc = new Date(now.getFullYear(), targetMonthIndex, 0).getDate();

  const weekDefs = [
    { id: 'w1', weekName: 'Semana 1', dateRange: `01 al 06 de ${targetMonthNameLower}`, startDay: 1, endDay: 6 },
    { id: 'w2', weekName: 'Semana 2', dateRange: `07 al 13 de ${targetMonthNameLower}`, startDay: 7, endDay: 13 },
    { id: 'w3', weekName: 'Semana 3', dateRange: `14 al 20 de ${targetMonthNameLower}`, startDay: 14, endDay: 20 },
    { id: 'w4', weekName: 'Semana 4', dateRange: `21 al ${daysInMonthCalc} de ${targetMonthNameLower}`, startDay: 21, endDay: daysInMonthCalc },
  ];

  if (!transactions || transactions.length === 0) {
    const emptyType1: WeeklySalesBreakdown[] = weekDefs.map((w) => ({
      id: w.id,
      weekName: w.weekName,
      dateRange: w.dateRange,
      totalAmount: 0,
      totalCount: 0,
      upconnectAmount: 0,
      upconnectCount: 0,
      connectorsAmount: 0,
      connectorsCount: 0,
    }));
    const emptyType2 = weekDefs.map((w) => ({
      weekName: w.weekName,
      dateRange: w.dateRange,
      amount: 0,
      percentage: 0,
      description: `0 emisiones registradas en este corte operativo (0.00 USD).`,
      operationsCount: 0,
    }));
    return { type1: emptyType1, type2: emptyType2 };
  }

  // 2. Filter transactions strictly to this target month!
  const monthTransactions = transactions.filter((t) => {
    const { monthIndex } = extractTransactionDayAndMonth(t);
    return monthIndex === targetMonthIndex;
  });

  // CRITICAL: ONLY use transactions belonging strictly to the target month!
  // If a week in this month hasn't occurred yet, it MUST remain at 0.
  const activeTrx = monthTransactions;

  // Exact total month sales from transactions of this month
  const totalMonthSales = parseFloat(activeTrx.reduce((acc, t) => acc + (t.value || 0), 0).toFixed(2));

  const weekBuckets = [
    { upAmount: 0, upCount: 0, coAmount: 0, coCount: 0 },
    { upAmount: 0, upCount: 0, coAmount: 0, coCount: 0 },
    { upAmount: 0, upCount: 0, coAmount: 0, coCount: 0 },
    { upAmount: 0, upCount: 0, coAmount: 0, coCount: 0 },
  ];

  activeTrx.forEach((t) => {
    const { day } = extractTransactionDayAndMonth(t);
    let weekIndex = 0;
    if (day >= 1 && day <= 6) weekIndex = 0;
    else if (day >= 7 && day <= 13) weekIndex = 1;
    else if (day >= 14 && day <= 20) weekIndex = 2;
    else weekIndex = 3; // 21 to end of month

    const val = t.value || 0;
    if (t.channel === 'Connectors') {
      weekBuckets[weekIndex].coAmount += val;
      weekBuckets[weekIndex].coCount += 1;
    } else {
      weekBuckets[weekIndex].upAmount += val;
      weekBuckets[weekIndex].upCount += 1;
    }
  });

  const type1: WeeklySalesBreakdown[] = [];
  const type2: {
    weekName: string;
    dateRange: string;
    amount: number;
    percentage: number;
    description: string;
    operationsCount?: number;
  }[] = [];

  let accumulatedAmount = 0;

  weekDefs.forEach((w, idx) => {
    const bucket = weekBuckets[idx];
    const upAmount = parseFloat(bucket.upAmount.toFixed(2));
    const coAmount = parseFloat(bucket.coAmount.toFixed(2));
    let weekTotal = parseFloat((upAmount + coAmount).toFixed(2));
    const weekCount = bucket.upCount + bucket.coCount;

    // Reconciliation on last week only if week 4 actually had sales recorded
    if (idx === 3 && totalMonthSales > 0 && weekTotal > 0) {
      const remainingExpected = parseFloat((totalMonthSales - accumulatedAmount).toFixed(2));
      if (Math.abs(remainingExpected - weekTotal) <= 0.05 && Math.abs(remainingExpected - weekTotal) > 0) {
        weekTotal = remainingExpected;
      }
    } else {
      accumulatedAmount = parseFloat((accumulatedAmount + weekTotal).toFixed(2));
    }

    const pct = totalMonthSales > 0
      ? parseFloat(((weekTotal / totalMonthSales) * 100).toFixed(1))
      : 0;

    type1.push({
      id: w.id,
      weekName: w.weekName,
      dateRange: w.dateRange,
      totalAmount: weekTotal,
      totalCount: weekCount,
      upconnectAmount: upAmount,
      upconnectCount: bucket.upCount,
      connectorsAmount: coAmount,
      connectorsCount: bucket.coCount,
    });

    type2.push({
      weekName: w.weekName,
      dateRange: w.dateRange,
      amount: weekTotal,
      percentage: pct,
      description: `${weekCount} emisiones registradas en este corte operativo (${weekTotal.toFixed(2)} USD).`,
      operationsCount: weekCount,
    });
  });

  return { type1, type2 };
};

export const formatSpanishDateCutoff = (d: Date | string): string => {
  if (!d) return '20 de septiembre de 2026';
  if (typeof d === 'string') {
    return formatSpanishDate(d);
  }
  const months = [
    'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
    'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
  ];
  // Use UTC to prevent subtracting 1 day in western hemisphere timezones like Ecuador (UTC-5)
  const day = d.getUTCDate();
  const month = months[d.getUTCMonth()] || 'septiembre';
  const year = d.getUTCFullYear() || 2026;
  return `${day} de ${month} de ${year}`;
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

export const deriveSellersFromFirmaTransactions = (transactions: TransactionRecord[]): SocioRecord[] => {
  if (!transactions || transactions.length === 0) return [];
  const sellerMap: Record<string, {
    name: string;
    totalSales: number;
    count: number;
    channel: 'UpConnect' | 'Connectors';
    role: string;
    plans: Record<string, number>;
  }> = {};

  transactions.forEach((t) => {
    const rawName = (t.socio || (t.clientLastName && t.clientLastName !== '(Socio)' ? `${t.clientName} ${t.clientLastName}` : t.clientName) || 'Distribuidor').trim();
    if (!rawName || rawName.toLowerCase() === 'total' || rawName.toLowerCase().includes('total general')) return;

    const ch: 'UpConnect' | 'Connectors' = t.channel === 'Connectors' ? 'Connectors' : 'UpConnect';
    const role = t.role || (ch === 'Connectors' ? 'Distribuidor Connect' : 'Distribuidor Upconnect');
    const plan = t.duration ? `Firma ${t.duration}` : 'Firma Electrónica';

    if (!sellerMap[rawName]) {
      sellerMap[rawName] = {
        name: rawName,
        totalSales: 0,
        count: 0,
        channel: ch,
        role,
        plans: {},
      };
    }

    sellerMap[rawName].totalSales += t.value || 0;
    sellerMap[rawName].count += 1;
    sellerMap[rawName].plans[plan] = (sellerMap[rawName].plans[plan] || 0) + 1;
  });

  const sorted = Object.values(sellerMap).sort((a, b) => b.totalSales - a.totalSales);
  return sorted.map((s, idx) => {
    let topPlan = 'Firma Electrónica (1 año)';
    let maxPCount = -1;
    for (const [pName, cnt] of Object.entries(s.plans)) {
      if (cnt > maxPCount) {
        maxPCount = cnt;
        topPlan = pName;
      }
    }
    const avgTicket = s.count > 0 ? s.totalSales / s.count : s.totalSales;
    return {
      rank: idx + 1,
      name: s.name,
      totalSales: parseFloat(s.totalSales.toFixed(2)),
      group: idx < 10 ? 'TOP 1-10' : 'TOP 11-20',
      operationsCount: s.count,
      averageTicket: parseFloat(avgTicket.toFixed(2)),
      topPlan,
      role: s.role,
      channel: s.channel,
      note: idx === 0 ? 'Líder en Emisiones' : undefined,
    };
  });
};

export const deriveSociosFromSistemasTransactions = (transactions: TransactionRecord[]): SocioRecord[] => {
  if (!transactions || transactions.length === 0) return [];
  const socioMap: Record<string, {
    name: string;
    totalSales: number;
    count: number;
    channel: 'UpConnect' | 'Connectors';
    role: string;
    plans: Record<string, number>;
  }> = {};

  transactions.forEach((t) => {
    const rawName = (t.socio || `${t.clientName} ${t.clientLastName}` || 'Socio').trim();
    if (!rawName || rawName.toLowerCase() === 'total' || rawName.toLowerCase().includes('total general')) return;

    const ch: 'UpConnect' | 'Connectors' = t.channel === 'Connectors' ? 'Connectors' : 'UpConnect';
    const role = t.role || (ch === 'Connectors' ? 'Distribuidor Connect' : 'Distribuidor Upconnect');
    const plan = t.solutionCategory || 'UP INTERMEDIO ($17.25)';

    if (!socioMap[rawName]) {
      socioMap[rawName] = {
        name: rawName,
        totalSales: 0,
        count: 0,
        channel: ch,
        role,
        plans: {},
      };
    }
    socioMap[rawName].totalSales += t.value || 0;
    socioMap[rawName].count += 1;
    socioMap[rawName].plans[plan] = (socioMap[rawName].plans[plan] || 0) + 1;
  });

  const sorted = Object.values(socioMap).sort((a, b) => b.totalSales - a.totalSales);
  return sorted.map((s, idx) => {
    let topPlan = 'UP INTERMEDIO ($17.25)';
    let maxCnt = -1;
    for (const [pName, cnt] of Object.entries(s.plans)) {
      if (cnt > maxCnt) {
        maxCnt = cnt;
        topPlan = pName;
      }
    }
    const avgTicket = s.count > 0 ? s.totalSales / s.count : s.totalSales;
    return {
      rank: idx + 1,
      name: s.name,
      totalSales: parseFloat(s.totalSales.toFixed(2)),
      group: idx < 10 ? 'TOP 1-10' : 'TOP 11-20',
      operationsCount: s.count,
      averageTicket: parseFloat(avgTicket.toFixed(2)),
      topPlan,
      role: s.role,
      channel: s.channel,
      note: idx === 0 ? 'Líder en Ventas' : undefined,
    };
  });
};

export const classifyTransactionDifferentiator = (
  t: TransactionRecord
): 'facturacion' | 'erp' | 'contadores' => {
  const cat = (t.solutionCategory || '').toLowerCase();
  const plan = (t.notes || '').toLowerCase();
  const role = (t.role || '').toLowerCase();
  const dur = (t.duration || '').toLowerCase();

  // 1. Contadores (Planes Contador, Franquicias, VIP, Aliados)
  if (
    cat.includes('contador') ||
    cat.includes('franquicia') ||
    cat.includes('estudio') ||
    role.includes('contador') ||
    role.includes('franquicia') ||
    plan.includes('contador')
  ) {
    return 'contadores';
  }

  // 2. ERP Contables y Módulos
  if (
    cat.includes('erp') ||
    cat.includes('módulo') ||
    cat.includes('modulo') ||
    cat.includes('avanzado') ||
    cat.includes('corporativo') ||
    cat.includes('sistema') ||
    cat.includes('inventario') ||
    plan.includes('erp')
  ) {
    return 'erp';
  }

  // 3. Planes de Facturación (Facturación electrónica, emisión de firmas, planes básicos)
  if (
    cat.includes('factur') ||
    cat.includes('firma') ||
    cat.includes('emisi') ||
    cat.includes('básico') ||
    cat.includes('basico') ||
    cat.includes('intermedio') ||
    dur.includes('año') ||
    dur.includes('firma')
  ) {
    return 'facturacion';
  }

  // Default fallback if not matched
  if (t.channel === 'Connectors') return 'contadores';
  return 'facturacion';
};

export const deriveWeeklyDifferentiators = (
  transactions: TransactionRecord[],
  targetMonthIndexOrName?: number | string,
  baseWeekly?: { weekName: string; dateRange: string; amount: number; operationsCount?: number }[]
): WeeklyDifferentiatorItem[] => {
  const now = new Date();
  let mName = MONTH_NAMES_ES[now.getMonth()] || 'OCTUBRE';
  if (typeof targetMonthIndexOrName === 'number' && targetMonthIndexOrName >= 1 && targetMonthIndexOrName <= 12) {
    mName = MONTH_NAMES_ES[targetMonthIndexOrName - 1];
  } else if (typeof targetMonthIndexOrName === 'string' && targetMonthIndexOrName.trim()) {
    mName = targetMonthIndexOrName.toUpperCase().trim();
  }
  const mLower = mName.toLowerCase();
  const mIdx = MONTH_NAMES_ES.indexOf(mName) + 1 || (now.getMonth() + 1);
  const daysInDiffMonth = new Date(now.getFullYear(), mIdx, 0).getDate();

  const weekDefs = [
    { id: 'w1', weekName: 'Semana 1', startDay: 1, endDay: 6, defaultRange: `01 al 06 de ${mLower}` },
    { id: 'w2', weekName: 'Semana 2', startDay: 7, endDay: 13, defaultRange: `07 al 13 de ${mLower}` },
    { id: 'w3', weekName: 'Semana 3', startDay: 14, endDay: 20, defaultRange: `14 al 20 de ${mLower}` },
    { id: 'w4', weekName: 'Semana 4', startDay: 21, endDay: daysInDiffMonth, defaultRange: `21 al ${daysInDiffMonth} de ${mLower}` },
  ];

  const results: WeeklyDifferentiatorItem[] = weekDefs.map((w, idx) => {
    const matchedBase = baseWeekly && baseWeekly[idx];
    return {
      weekId: w.id,
      weekName: w.weekName,
      dateRange: matchedBase?.dateRange || w.defaultRange,
      totalSales: 0,
      totalCount: 0,
      facturacionSales: 0,
      facturacionCount: 0,
      erpSales: 0,
      erpCount: 0,
      contadoresSales: 0,
      contadoresCount: 0,
    };
  });

  // Strictly filter transactions to this target month!
  const monthTransactions = (transactions || []).filter((t) => {
    const { monthIndex } = extractTransactionDayAndMonth(t);
    return monthIndex === mIdx;
  });

  if (monthTransactions.length > 0) {
    monthTransactions.forEach((t) => {
      const { day } = extractTransactionDayAndMonth(t);
      let wIdx = 0;
      if (day >= 1 && day <= 6) wIdx = 0;
      else if (day >= 7 && day <= 13) wIdx = 1;
      else if (day >= 14 && day <= 20) wIdx = 2;
      else wIdx = 3;

      const diff = classifyTransactionDifferentiator(t);
      const val = t.value || 0;

      results[wIdx].totalSales += val;
      results[wIdx].totalCount += 1;

      if (diff === 'facturacion') {
        results[wIdx].facturacionSales += val;
        results[wIdx].facturacionCount += 1;
      } else if (diff === 'erp') {
        results[wIdx].erpSales += val;
        results[wIdx].erpCount += 1;
      } else {
        results[wIdx].contadoresSales += val;
        results[wIdx].contadoresCount += 1;
      }
    });

    results.forEach((r) => {
      r.totalSales = parseFloat(r.totalSales.toFixed(2));
      r.facturacionSales = parseFloat(r.facturacionSales.toFixed(2));
      r.erpSales = parseFloat(r.erpSales.toFixed(2));
      r.contadoresSales = parseFloat(r.contadoresSales.toFixed(2));
    });
  } else if (baseWeekly && baseWeekly.length > 0) {
    baseWeekly.forEach((w, idx) => {
      if (idx < 4) {
        const tot = w.amount || 0;
        const cnt = w.operationsCount || 0;
        const erpVal = tot > 0 ? parseFloat((tot * 0.44).toFixed(2)) : 0;
        const contVal = tot > 0 ? parseFloat((tot * 0.36).toFixed(2)) : 0;
        const factVal = tot > 0 ? parseFloat((tot - erpVal - contVal).toFixed(2)) : 0;

        results[idx].totalSales = tot;
        results[idx].totalCount = cnt;
        results[idx].erpSales = erpVal;
        results[idx].erpCount = cnt > 0 ? Math.round(cnt * 0.42) : 0;
        results[idx].contadoresSales = contVal;
        results[idx].contadoresCount = cnt > 0 ? Math.round(cnt * 0.36) : 0;
        results[idx].facturacionSales = factVal;
        results[idx].facturacionCount = cnt > 0 ? Math.max(0, cnt - results[idx].erpCount - results[idx].contadoresCount) : 0;
        results[idx].dateRange = w.dateRange;
      }
    });
  }

  return results;
};

/**
 * Computes exact channel breakdown and partner counts (Distribuidor Upconnect vs Distribuidor Connect):
 * - count of partners (cantidad de socios)
 * - operations/sales count
 * - total revenue in USD
 * - average ticket
 */
export const computeSociosChannelSummary = (
  socios: SocioRecord[],
  transactions: TransactionRecord[] = []
): SociosChannelSummary => {
  const upSocios: SocioRecord[] = [];
  const coSocios: SocioRecord[] = [];

  let upSales = 0;
  let coSales = 0;
  let upOps = 0;
  let coOps = 0;

  socios.forEach((s) => {
    let ch = s.channel;
    if (!ch) {
      const sName = (s.name || '').toLowerCase().trim();
      const sTrx = transactions.filter((t) => (t.socio || '').toLowerCase().trim() === sName);
      if (sTrx.length > 0) {
        let uCnt = 0;
        let cCnt = 0;
        sTrx.forEach((t) => {
          if (t.channel === 'Connectors') cCnt++;
          else uCnt++;
        });
        ch = cCnt > uCnt ? 'Connectors' : 'UpConnect';
      } else {
        const role = (s.role || '').toLowerCase();
        if (role.includes('up') || role.includes('direct') || sName.includes('upconnect')) {
          ch = 'UpConnect';
        } else {
          ch = 'Connectors';
        }
      }
    }

    // Mutate/ensure channel is present on record
    s.channel = ch;

    if (ch === 'UpConnect') {
      upSocios.push(s);
      upSales += s.totalSales || 0;
      upOps += s.operationsCount || 1;
    } else {
      coSocios.push(s);
      coSales += s.totalSales || 0;
      coOps += s.operationsCount || 1;
    }
  });

  const totalSocios = upSocios.length + coSocios.length;
  const totalSales = parseFloat((upSales + coSales).toFixed(2));
  const totalOps = upOps + coOps;

  const upPct = totalSocios > 0 ? parseFloat(((upSocios.length / totalSocios) * 100).toFixed(1)) : 0;
  const coPct = totalSocios > 0 ? parseFloat(((coSocios.length / totalSocios) * 100).toFixed(1)) : 0;

  return {
    upconnect: {
      channel: 'UpConnect',
      channelLabel: 'Distribuidor Upconnect',
      count: upSocios.length,
      percentage: upPct,
      operationsCount: upOps,
      totalSales: parseFloat(upSales.toFixed(2)),
      averageTicket: upOps > 0 ? parseFloat((upSales / upOps).toFixed(2)) : 0,
      socios: upSocios,
    },
    connectors: {
      channel: 'Connectors',
      channelLabel: 'Distribuidor Connect',
      count: coSocios.length,
      percentage: coPct,
      operationsCount: coOps,
      totalSales: parseFloat(coSales.toFixed(2)),
      averageTicket: coOps > 0 ? parseFloat((coSales / coOps).toFixed(2)) : 0,
      socios: coSocios,
    },
    totalSociosCount: totalSocios,
    totalOperationsCount: totalOps,
    totalSales,
    averageTicket: totalOps > 0 ? parseFloat((totalSales / totalOps).toFixed(2)) : 0,
  };
};

/**
 * Computes filtered stats and dynamic executive titles based on filter selection:
 * - month: "Reporte del mes de Septiembre 2026"
 * - week: "hasta el 20 de septiembre semana 2 (del 07 al 13 de septiembre)"
 * - range: "Reporte del 1 al 20 de septiembre de 2026"
 */
export const computeReportView = (dataset: GlobalDataset): FilteredReportView => {
  const now = new Date();
  const currentActualMonthName = MONTH_NAMES_ES[now.getMonth()] || 'OCTUBRE';
  const filter: ReportTimeFilter = dataset.reportFilter || { type: 'month', month: currentActualMonthName };

  let titleReport1 = 'Reporte Gerencial de Ventas: Distribuidor Upconnect / Distribuidor Connect';
  let titleReport2 = 'REPORTE GERENCIAL VENTAS HASTA EL 20 DE SEPTIEMBRE DE 2026';
  let subtitleDate = dataset.cutoffDate || '20 de Septiembre de 2026';
  let periodLabel = 'Consolidado General 2026';
  let isFiltered = false;

  const targetWeekMonthName = filter.month ? filter.month.toUpperCase() : currentActualMonthName;
  const targetWeekMonthIdx = MONTH_NAMES_ES.indexOf(targetWeekMonthName) + 1 || (now.getMonth() + 1);
  const targetWeekMonthLower = targetWeekMonthName.toLowerCase();
  const targetWeekYear = now.getFullYear();
  const daysInTargetMonth = new Date(targetWeekYear, targetWeekMonthIdx, 0).getDate();
  const padMonth = String(targetWeekMonthIdx).padStart(2, '0');

  const weekMap: Record<string, { name: string; range: string; start: string; end: string; startDay: number; endDay: number }> = {
    w1: {
      name: 'Semana 1',
      range: `01 al 06 de ${targetWeekMonthLower}`,
      start: `${targetWeekYear}-${padMonth}-01`,
      end: `${targetWeekYear}-${padMonth}-06`,
      startDay: 1,
      endDay: 6,
    },
    w2: {
      name: 'Semana 2',
      range: `07 al 13 de ${targetWeekMonthLower}`,
      start: `${targetWeekYear}-${padMonth}-07`,
      end: `${targetWeekYear}-${padMonth}-13`,
      startDay: 7,
      endDay: 13,
    },
    w3: {
      name: 'Semana 3',
      range: `14 al 20 de ${targetWeekMonthLower}`,
      start: `${targetWeekYear}-${padMonth}-14`,
      end: `${targetWeekYear}-${padMonth}-20`,
      startDay: 14,
      endDay: 20,
    },
    w4: {
      name: 'Semana 4',
      range: `21 al ${daysInTargetMonth} de ${targetWeekMonthLower}`,
      start: `${targetWeekYear}-${padMonth}-21`,
      end: `${targetWeekYear}-${padMonth}-${daysInTargetMonth}`,
      startDay: 21,
      endDay: daysInTargetMonth,
    },
  };

  let filteredTrx = [...dataset.transactions];
  let filteredSocios = [...dataset.socios];

  let activeMonthName: string | undefined = undefined;
  let activeMonthIndex: number | undefined = undefined;

  if (filter.type === 'month' && filter.month) {
    isFiltered = true;
    const monthName = filter.month.toUpperCase();
    activeMonthName = monthName;
    activeMonthIndex = MONTH_NAMES_ES.indexOf(monthName) + 1;

    titleReport1 = `Reporte del mes de ${monthName.charAt(0) + monthName.slice(1).toLowerCase()} 2026: Distribuidor Upconnect / Distribuidor Connect`;
    titleReport2 = `REPORTE GERENCIAL VENTAS DEL MES DE ${monthName} 2026`;
    subtitleDate = `Mes completo de ${monthName.charAt(0) + monthName.slice(1).toLowerCase()} 2026`;
    periodLabel = `Mes de ${monthName.charAt(0) + monthName.slice(1).toLowerCase()}`;

    // Filter transactions by month index using exact extractTransactionMonth
    const mIndex = MONTH_NAMES_ES.indexOf(monthName) + 1;
    if (mIndex > 0) {
      filteredTrx = dataset.transactions.filter((t) => {
        const { monthIndex } = extractTransactionMonth(t);
        return monthIndex === mIndex;
      });
    }
  } else if (filter.type === 'week' && filter.weekId) {
    isFiltered = true;
    activeMonthName = targetWeekMonthName;
    activeMonthIndex = targetWeekMonthIdx;
    const targetWeek = weekMap[filter.weekId] || weekMap.w1;

    titleReport1 = `Reporte ${targetWeek.name} (del ${targetWeek.range}): Distribuidor Upconnect / Distribuidor Connect`;
    titleReport2 = `REPORTE GERENCIAL VENTAS ${targetWeek.name.toUpperCase()} (DEL ${targetWeek.range.toUpperCase()})`;
    subtitleDate = `${targetWeek.name} (${targetWeek.range} de ${targetWeekYear})`;
    periodLabel = `${targetWeek.name} (${targetWeek.range})`;

    filteredTrx = dataset.transactions.filter((t) => {
      const { day, monthIndex } = extractTransactionDayAndMonth(t);
      if (monthIndex === targetWeekMonthIdx) {
        return day >= targetWeek.startDay && day <= targetWeek.endDay;
      }
      return false;
    });
  } else if (filter.type === 'range' && filter.startDate && filter.endDate) {
    isFiltered = true;
    const sParts = filter.startDate.split('-');
    const eParts = filter.endDate.split('-');
    const startDay = parseInt(sParts[2] || '1', 10);
    const endDay = parseInt(eParts[2] || '20', 10);
    const endMonthIdx = parseInt(eParts[1] || '9', 10);
    const endMonthName = MONTH_NAMES_ES[endMonthIdx - 1]?.toLowerCase() || 'septiembre';

    titleReport1 = `Reporte del ${startDay} al ${endDay} de ${endMonthName}: Distribuidor Upconnect / Distribuidor Connect`;
    titleReport2 = `REPORTE GERENCIAL VENTAS DEL ${startDay} AL ${endDay} DE ${endMonthName.toUpperCase()} DE 2026`;
    subtitleDate = `Del ${startDay} al ${endDay} de ${endMonthName} de 2026`;
    periodLabel = `Rango del ${startDay} al ${endDay} de ${endMonthName}`;

    filteredTrx = dataset.transactions.filter((t) => {
      return t.date >= filter.startDate! && t.date <= filter.endDate!;
    });
  }

  // =========================================================================
  // REPORTE 1: Calculado EXCLUSIVAMENTE con el archivo de firmas (transactions / monthlyMetrics)
  // El archivo de socios NO debe afectar los totales ni métricas del Reporte 1.
  // =========================================================================
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
  } else {
    // When filtered by month but no individual transaction rows match, check target monthly metric
    if (filter.type === 'month' && filter.month) {
      const targetMetric = dataset.monthlyMetrics.find(
        (m) => m.month.toUpperCase() === filter.month!.toUpperCase()
      );
      if (targetMetric) {
        totalUpSales = targetMetric.upconnectSales;
        totalCoSales = targetMetric.connectorsSales;
        totalUpCount = targetMetric.upconnectCount;
        totalCoCount = targetMetric.connectorsCount;
      }
    } else if (filter.type === 'week' || filter.type === 'range') {
      // In weeks or ranges that have no transactions yet, totals MUST remain at 0!
      totalUpSales = 0;
      totalCoSales = 0;
      totalUpCount = 0;
      totalCoCount = 0;
    } else {
      // Standard totals from monthly metrics for 'all' (Consolidado)
      totalUpSales = dataset.monthlyMetrics.reduce((a, m) => a + m.upconnectSales, 0);
      totalCoSales = dataset.monthlyMetrics.reduce((a, m) => a + m.connectorsSales, 0);
      totalUpCount = dataset.monthlyMetrics.reduce((a, m) => a + m.upconnectCount, 0);
      totalCoCount = dataset.monthlyMetrics.reduce((a, m) => a + m.connectorsCount, 0);
    }
  }

  let totalSales = parseFloat((totalUpSales + totalCoSales).toFixed(2));
  let totalCount = totalUpCount + totalCoCount;
  const upSalesPct = totalSales > 0 ? ((totalUpSales / totalSales) * 100).toFixed(1) : '0.0';
  const coSalesPct = totalSales > 0 ? ((totalCoSales / totalSales) * 100).toFixed(1) : '0.0';

  // =========================================================================
  // REPORTE 2: Calculado EXCLUSIVAMENTE con el archivo de socios (dataset.socios / sistemasTransactions)
  // El archivo de firmas NO debe sobreescribir ni modificar la cartera de socios.
  // Cuando se selecciona filtro de mes/semana/rango, SOLO se toman los valores del período elegido (NO acumulados).
  // =========================================================================
  let filteredSistemasTrx: TransactionRecord[] = [];
  if (dataset.sistemasTransactions && dataset.sistemasTransactions.length > 0) {
    if (filter.type === 'month' && filter.month) {
      const mIndex = MONTH_NAMES_ES.indexOf(filter.month.toUpperCase()) + 1;
      filteredSistemasTrx = dataset.sistemasTransactions.filter((t) => {
        const { monthIndex, monthName } = extractTransactionMonth(t);
        return (mIndex > 0 && monthIndex === mIndex) || (monthName && monthName.toUpperCase() === filter.month!.toUpperCase());
      });
    } else if (filter.type === 'week' && filter.weekId) {
      const targetWeek = weekMap[filter.weekId] || weekMap.w1;
      filteredSistemasTrx = (dataset.sistemasTransactions || []).filter((t) => {
        const { day, monthIndex } = extractTransactionDayAndMonth(t);
        if (monthIndex === targetWeekMonthIdx) {
          return day >= targetWeek.startDay && day <= targetWeek.endDay;
        }
        return false;
      });
    } else if (filter.type === 'range' && filter.startDate && filter.endDate) {
      filteredSistemasTrx = dataset.sistemasTransactions.filter((t) => {
        return t.date >= filter.startDate! && t.date <= filter.endDate!;
      });
    } else {
      filteredSistemasTrx = [...dataset.sistemasTransactions];
    }
  }

  let r2Socios: SocioRecord[] = [];
  let netSales = 0;

  if (isFiltered) {
    if (filteredSistemasTrx.length > 0) {
      r2Socios = deriveSociosFromSistemasTransactions(filteredSistemasTrx);
      const sociosTotalSales = r2Socios.reduce((acc, s) => acc + (s.totalSales || 0), 0);
      netSales = parseFloat(sociosTotalSales.toFixed(2));
    } else if (filter.type === 'month' && filter.month) {
      // Check target monthly metric if individual transactions are not present
      const targetSistemasMetric = (dataset.sistemasMonthlyMetrics || []).find(
        (m) => m.month.toUpperCase() === filter.month!.toUpperCase()
      );
      if (targetSistemasMetric) {
        netSales = parseFloat((targetSistemasMetric.upconnectSales + targetSistemasMetric.connectorsSales).toFixed(2));
        r2Socios = dataset.socios;
      } else {
        r2Socios = [];
        netSales = 0;
      }
    } else {
      r2Socios = [];
      netSales = 0;
    }
  } else {
    // Consolidado 2026: todos los socios acumulados
    r2Socios = dataset.socios && dataset.socios.length > 0 
      ? [...dataset.socios] 
      : deriveSociosFromSistemasTransactions(dataset.sistemasTransactions || []);
    const sociosTotalSales = r2Socios.reduce((acc, s) => acc + (s.totalSales || 0), 0);
    netSales = parseFloat(sociosTotalSales.toFixed(2));
  }

  let ivaAmount = parseFloat((netSales * 0.15).toFixed(2));
  let grossSales = parseFloat((netSales + ivaAmount).toFixed(2));

  // Vendedores y resumen para Reporte 1 (obtenidos EXCLUSIVAMENTE del archivo de firmas)
  const rawFirmaSocios = (dataset.firmaSocios && dataset.firmaSocios.length > 0)
    ? dataset.firmaSocios
    : deriveSellersFromFirmaTransactions(dataset.transactions);
  const firmaSociosSummary = computeSociosChannelSummary(rawFirmaSocios, dataset.transactions);

  // Fechas de corte independientes para evitar mezclar información
  const cutoffDateReport1 = dataset.cutoffDateFirma || dataset.cutoffDate || '20 de Septiembre de 2026';
  const cutoffDateReport2 = dataset.cutoffDateSocios || dataset.cutoffDate || '20 de Septiembre de 2026';

  // Métricas mensuales de firmas para Reporte 1
  const effectiveMonthlyMetrics = dataset.monthlyMetrics.length > 0
    ? dataset.monthlyMetrics
    : deriveMonthlyMetricsFromTransactions(dataset.transactions);

  // Métricas mensuales de sistemas / socios para Reporte 2
  const rawSistemasMonthly = (dataset.sistemasMonthlyMetrics && dataset.sistemasMonthlyMetrics.length > 0)
    ? dataset.sistemasMonthlyMetrics
    : (dataset.sistemasTransactions && dataset.sistemasTransactions.length > 0
        ? deriveMonthlyMetricsFromTransactions(dataset.sistemasTransactions)
        : []);

  // Derive exact partners count breakdown (Distribuidor Upconnect vs Distribuidor Connect) para Reporte 2
  const sociosSummary = computeSociosChannelSummary(r2Socios, filteredSistemasTrx);

  // When filtered by month, only return that chosen month's metric so Slide 3 does NOT show accumulated other months
  let sistemasMonthlyMetrics: MonthlyMetric[] = [];
  if (isFiltered && filter.type === 'month' && filter.month) {
    const matched = rawSistemasMonthly.filter((m) => m.month.toUpperCase() === filter.month!.toUpperCase());
    if (matched.length > 0) {
      sistemasMonthlyMetrics = matched;
    } else {
      sistemasMonthlyMetrics = [{
        month: filter.month.toUpperCase(),
        monthIndex: MONTH_NAMES_ES.indexOf(filter.month.toUpperCase()) + 1 || 9,
        upconnectSales: sociosSummary.upconnect.totalSales,
        connectorsSales: sociosSummary.connectors.totalSales,
        upconnectCount: sociosSummary.upconnect.operationsCount,
        connectorsCount: sociosSummary.connectors.operationsCount,
      }];
    }
  } else {
    sistemasMonthlyMetrics = rawSistemasMonthly;
  }

  // Derive exact weekly breakdown matching filtered transactions and active period/month for Report 2
  const effectiveMonthIdx = activeMonthIndex || targetWeekMonthIdx;
  const effectiveMonthName = activeMonthName || targetWeekMonthName;

  const monthSistemasTrx = (dataset.sistemasTransactions || []).filter((t) => {
    const { monthIndex } = extractTransactionDayAndMonth(t);
    return monthIndex === effectiveMonthIdx;
  });

  const weeklyR2 = deriveWeeklyBreakdownFromTransactions(
    monthSistemasTrx,
    effectiveMonthIdx || effectiveMonthName
  );

  const weeklyDifferentiators = deriveWeeklyDifferentiators(
    monthSistemasTrx,
    effectiveMonthIdx || effectiveMonthName,
    weeklyR2.type2
  );

  // Categorías de solución y cruce comercial para Reporte 2 ajustados al período
  const targetTrxForR2Categories = filteredSistemasTrx.length > 0
    ? filteredSistemasTrx
    : (isFiltered ? [] : (dataset.sistemasTransactions || []));

  const solutionCategoriesReport2 = targetTrxForR2Categories.length > 0
    ? deriveSolutionCategoriesFromTransactions(targetTrxForR2Categories)
    : [];

  const commercialCrossReport2 = targetTrxForR2Categories.length > 0
    ? deriveCommercialCrossFromTransactions(targetTrxForR2Categories)
    : (isFiltered ? undefined : dataset.commercialCross);

  // Derive exact weekly breakdown matching the selected month for Report 1
  const monthTrxForWeekly = dataset.transactions.filter((t) => {
    const { monthIndex } = extractTransactionDayAndMonth(t);
    return monthIndex === effectiveMonthIdx;
  });

  const weekly = deriveWeeklyBreakdownFromTransactions(
    monthTrxForWeekly,
    effectiveMonthIdx || effectiveMonthName
  );

  // Apply manual customizations / edits that affect reports
  const overrides = dataset.customOverrides || {};
  if (overrides.titleReport1) titleReport1 = String(overrides.titleReport1);
  else if (overrides.r1_cover_title) titleReport1 = String(overrides.r1_cover_title);

  if (overrides.titleReport2) titleReport2 = String(overrides.titleReport2);
  else if (overrides.r2_cover_title) titleReport2 = String(overrides.r2_cover_title);

  if (overrides.subtitleDate) subtitleDate = String(overrides.subtitleDate);
  else if (overrides.r1_cover_subtitle) subtitleDate = String(overrides.r1_cover_subtitle);

  if (overrides.periodLabel) periodLabel = String(overrides.periodLabel);
  else if (overrides.r1_cover_period) periodLabel = String(overrides.r1_cover_period);

  // Overrides de Reporte 1 (NO afectan a Reporte 2)
  const rawTotalSales = overrides.totalSales !== undefined ? overrides.totalSales : overrides.r1_kpi_total_sales;
  if (rawTotalSales !== undefined) {
    const val = Number(rawTotalSales);
    if (!isNaN(val)) {
      totalSales = val;
    }
  }

  // Overrides específicos de Reporte 2
  const rawR2NetSales = overrides.r2_net_sales !== undefined 
    ? overrides.r2_net_sales 
    : (overrides.netSales !== undefined ? overrides.netSales : undefined);
  if (rawR2NetSales !== undefined) {
    const valR2 = Number(rawR2NetSales);
    if (!isNaN(valR2)) {
      netSales = valR2;
      ivaAmount = parseFloat((valR2 * 0.15).toFixed(2));
      grossSales = parseFloat((valR2 + ivaAmount).toFixed(2));
    }
  }

  const rawUpSales = overrides.totalUpSales !== undefined ? overrides.totalUpSales : overrides.r1_kpi_up_sales;
  if (rawUpSales !== undefined) {
    const val = Number(rawUpSales);
    if (!isNaN(val)) totalUpSales = val;
  }

  const rawCoSales = overrides.totalCoSales !== undefined ? overrides.totalCoSales : overrides.r1_kpi_co_sales;
  if (rawCoSales !== undefined) {
    const val = Number(rawCoSales);
    if (!isNaN(val)) totalCoSales = val;
  }

  const rawTotalCount = overrides.totalCount !== undefined ? overrides.totalCount : overrides.r1_kpi_total_count;
  if (rawTotalCount !== undefined) {
    const val = Number(rawTotalCount);
    if (!isNaN(val)) totalCount = val;
  }

  const rawUpCount = overrides.totalUpCount !== undefined ? overrides.totalUpCount : overrides.r1_kpi_up_count;
  if (rawUpCount !== undefined) {
    const val = Number(rawUpCount);
    if (!isNaN(val)) totalUpCount = val;
  }

  const rawCoCount = overrides.totalCoCount !== undefined ? overrides.totalCoCount : overrides.r1_kpi_co_count;
  if (rawCoCount !== undefined) {
    const val = Number(rawCoCount);
    if (!isNaN(val)) totalCoCount = val;
  }

  const finalUpSalesPct = totalSales > 0 ? ((totalUpSales / totalSales) * 100).toFixed(1) : upSalesPct;
  const finalCoSalesPct = totalSales > 0 ? ((totalCoSales / totalSales) * 100).toFixed(1) : coSalesPct;

  // Apply weekly breakdown overrides so landing, reports and PPTX reflect exact edits
  let finalWeeklyType1 = weekly.type1.length > 0 ? weekly.type1 : dataset.weeklyBreakdownType1;
  if (finalWeeklyType1 && finalWeeklyType1.length > 0) {
    finalWeeklyType1 = finalWeeklyType1.map((w) => {
      const wKey = `week_${w.id}_amount`;
      const wCountKey = `week_${w.id}_count`;
      const newAmt = overrides[wKey] !== undefined ? Number(overrides[wKey]) : w.totalAmount;
      const newCount = overrides[wCountKey] !== undefined ? Number(overrides[wCountKey]) : w.totalCount;
      return {
        ...w,
        totalAmount: isNaN(newAmt) ? w.totalAmount : newAmt,
        totalCount: isNaN(newCount) ? w.totalCount : newCount,
      };
    });
  }

  let finalWeeklyType2 = weekly.type2.length > 0 ? weekly.type2 : dataset.weeklyBreakdownType2;
  if (finalWeeklyType2 && finalWeeklyType2.length > 0) {
    finalWeeklyType2 = finalWeeklyType2.map((w, idx) => {
      const wKey = `r2_week_${idx}_amount`;
      const newAmt = overrides[wKey] !== undefined ? Number(overrides[wKey]) : w.amount;
      return {
        ...w,
        amount: isNaN(newAmt) ? w.amount : newAmt,
      };
    });
  }

  // Apply socios overrides
  let finalSociosList = r2Socios.map((s) => {
    const nameKey = `socio_${s.rank}_name`;
    const salesKey = `socio_${s.rank}_sales`;
    const newName = overrides[nameKey] !== undefined ? String(overrides[nameKey]) : s.name;
    const newSales = overrides[salesKey] !== undefined ? Number(overrides[salesKey]) : s.totalSales;
    return {
      ...s,
      name: newName,
      totalSales: isNaN(newSales) ? s.totalSales : newSales,
    };
  });

  return {
    titleReport1,
    titleReport2,
    subtitleDate,
    periodLabel,
    cutoffDateReport1,
    cutoffDateReport2,
    activeMonthName,
    activeMonthIndex,
    transactions: filteredTrx,
    socios: finalSociosList,
    sociosSummary,
    firmaSocios: rawFirmaSocios,
    firmaSociosSummary,
    totalUpSales,
    totalCoSales,
    totalSales,
    totalUpCount,
    totalCoCount,
    totalCount,
    upSalesPct: finalUpSalesPct,
    coSalesPct: finalCoSalesPct,
    netSales,
    ivaAmount,
    grossSales,
    monthlyMetrics: effectiveMonthlyMetrics,
    sistemasMonthlyMetrics,
    isFiltered,
    weeklyBreakdownType1: finalWeeklyType1,
    weeklyBreakdownType2: finalWeeklyType2,
    weeklyDifferentiators,
    solutionCategoriesReport2,
    commercialCrossReport2,
  };
};
