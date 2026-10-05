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
  if (!transactions || transactions.length === 0) return { type1: [], type2: [] };

  // 1. Identify target month
  let targetMonthIndex = 9;
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
  } else {
    // Detect primary month from transactions frequency
    const monthCounts: Record<number, number> = {};
    transactions.forEach((t) => {
      const { monthIndex } = extractTransactionDayAndMonth(t);
      monthCounts[monthIndex] = (monthCounts[monthIndex] || 0) + 1;
    });
    let maxCnt = -1;
    for (const [mIdxStr, cnt] of Object.entries(monthCounts)) {
      const mIdx = parseInt(mIdxStr, 10);
      if (cnt > maxCnt) {
        maxCnt = cnt;
        targetMonthIndex = mIdx;
      }
    }
  }

  // 2. Filter transactions strictly to this month
  const monthTransactions = transactions.filter((t) => {
    const { monthIndex } = extractTransactionDayAndMonth(t);
    return monthIndex === targetMonthIndex;
  });

  // If filtered transactions exist for target month, use them; otherwise use provided transactions
  const activeTrx = monthTransactions.length > 0 ? monthTransactions : transactions;

  // Exact total month sales from transactions
  const totalMonthSales = parseFloat(activeTrx.reduce((acc, t) => acc + (t.value || 0), 0).toFixed(2));
  const targetMonthNameLower = (MONTH_NAMES_ES[targetMonthIndex - 1] || 'septiembre').toLowerCase();

  const weekDefs = [
    { id: 'w1', weekName: 'Semana 1', dateRange: `01 al 06 de ${targetMonthNameLower}`, startDay: 1, endDay: 6 },
    { id: 'w2', weekName: 'Semana 2', dateRange: `07 al 13 de ${targetMonthNameLower}`, startDay: 7, endDay: 13 },
    { id: 'w3', weekName: 'Semana 3', dateRange: `14 al 20 de ${targetMonthNameLower}`, startDay: 14, endDay: 20 },
    { id: 'w4', weekName: 'Semana 4', dateRange: `21 al 31 de ${targetMonthNameLower}`, startDay: 21, endDay: 31 },
  ];

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
    else weekIndex = 3; // 21 to 31

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

    // Reconciliation on last week to avoid floating point cent discrepancies
    if (idx === 3 && totalMonthSales > 0) {
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
  const filter: ReportTimeFilter = dataset.reportFilter || { type: 'all' };

  let titleReport1 = 'Reporte Gerencial de Ventas: Distribuidor Upconnect / Distribuidor Connect';
  let titleReport2 = 'REPORTE GERENCIAL VENTAS HASTA EL 20 DE SEPTIEMBRE DE 2026';
  let subtitleDate = dataset.cutoffDate || '20 de Septiembre de 2026';
  let periodLabel = 'Consolidado General 2026';
  let isFiltered = false;

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
    const weekMap: Record<string, { name: string; range: string; start: string; end: string }> = {
      w1: { name: 'Semana 1', range: '01 al 06 de septiembre', start: '2026-09-01', end: '2026-09-06' },
      w2: { name: 'Semana 2', range: '07 al 13 de septiembre', start: '2026-09-07', end: '2026-09-13' },
      w3: { name: 'Semana 3', range: '14 al 20 de septiembre', start: '2026-09-14', end: '2026-09-20' },
      w4: { name: 'Semana 4', range: '21 al 30 de septiembre', start: '2026-09-21', end: '2026-09-30' },
    };
    const targetWeek = weekMap[filter.weekId] || weekMap.w2;

    titleReport1 = `Reporte hasta el 20 de septiembre ${targetWeek.name} (del ${targetWeek.range}): Distribuidor Upconnect / Distribuidor Connect`;
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
    } else {
      // Standard totals from monthly metrics
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
  // REPORTE 2: Calculado EXCLUSIVAMENTE con el archivo de socios (dataset.socios)
  // El archivo de firmas NO debe sobreescribir ni modificar la cartera de socios.
  // =========================================================================
  // filteredSocios permanece como dataset.socios (Reporte 2 nunca se contamina con firmas)
  const sociosTotalSales = filteredSocios.reduce((acc, s) => acc + (s.totalSales || 0), 0);
  let netSales = parseFloat(sociosTotalSales.toFixed(2));
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
  const sistemasMonthlyMetrics = (dataset.sistemasMonthlyMetrics && dataset.sistemasMonthlyMetrics.length > 0)
    ? dataset.sistemasMonthlyMetrics
    : (dataset.sistemasTransactions && dataset.sistemasTransactions.length > 0
        ? deriveMonthlyMetricsFromTransactions(dataset.sistemasTransactions)
        : []);

  // Derive exact partners count breakdown (Distribuidor Upconnect vs Distribuidor Connect) para Reporte 2
  const sociosSummary = computeSociosChannelSummary(filteredSocios);

  // Derive exact weekly breakdown matching filtered transactions and active period/month
  const targetTrxForWeekly = filteredTrx.length > 0 ? filteredTrx : dataset.transactions;
  const weekly = deriveWeeklyBreakdownFromTransactions(
    targetTrxForWeekly,
    activeMonthIndex || activeMonthName
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
  let finalSociosList = filteredSocios.map((s) => {
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
  };
};
