export interface TransactionRecord {
  id: string;
  uniqueId: string;
  date: string; // YYYY-MM-DD or D/M/YYYY
  clientName: string;
  clientLastName: string;
  certificateStatus: 'EMITIDO' | 'EN PROCESO' | 'CANCELADO' | 'PENDIENTE';
  duration: '15 días' | 'Un año' | 'Dos años' | 'Tres años' | 'Cuatro años' | 'Cinco años' | string;
  value: number;
  role: string;
  channel: 'UpConnect' | 'Connectors';
  socio?: string;
  solutionCategory?: 'ERP Contables y Módulos' | 'Planes de Contador' | 'Facturación Electrónica' | string;
  month?: string;       // e.g. 'SEPTIEMBRE'
  monthIndex?: number;  // 1-12
  notes?: string;
  createdAt: string;
  sourceFile?: string;
}

export interface MonthlyMetric {
  month: string;
  monthIndex: number; // 1-12
  upconnectSales: number;
  connectorsSales: number;
  upconnectCount: number;
  connectorsCount: number;
  notes?: string;
}

export interface SocioRecord {
  rank: number;
  name: string;
  totalSales: number;
  group: 'TOP 1-10' | 'TOP 11-20';
  operationsCount?: number;
  averageTicket?: number; // Tiket promedio ($)
  topPlan?: string;       // Plan más vendido
  role?: string;          // Rol que tiene el socio
  channel?: 'UpConnect' | 'Connectors'; // Canal al que pertenece el socio
  note?: string;
}

export interface WeeklySalesBreakdown {
  id: string;
  weekName: string;
  dateRange: string;
  totalAmount: number;
  totalCount: number;
  upconnectAmount: number;
  upconnectCount: number;
  connectorsAmount: number;
  connectorsCount: number;
  highlightText?: string;
}

export interface PortfolioDuration {
  duration: string;
  channel: 'UpConnect' | 'Connectors';
  count: number;
  percentage: number;
  label?: string;
}

export interface SolutionCategory {
  id: string;
  title: string;
  badge: string;
  amount: number;
  percentage: number;
  plansDescription: string;
  iconType?: string;
}

export interface CommunityMetric {
  id: string;
  name: string;
  members: number;
  subtitle: string;
  description: string;
  color: string;
}

export interface GitHubConfig {
  enabled: boolean;
  token: string;
  owner: string;
  repo: string;
  branch: string;
  filePath: string;
  autoSyncOnUpload: boolean;
  lastSyncDate?: string;
  lastCommitSha?: string;
  lastCommitUrl?: string;
}

export interface CommercialCrossPhase {
  commercialTeamSales: number;
  organicSales: number;
  partnersSales: number;
  channel1Name: string;
  channel1Amount: number;
  channel2Name: string;
  channel2Status: string;
}

export interface ReportTimeFilter {
  type: 'all' | 'month' | 'week' | 'range';
  month?: string; // 'ENERO'..'DICIEMBRE' or '09'
  weekId?: string; // 'w1', 'w2', 'w3', 'w4'
  startDate?: string; // YYYY-MM-DD
  endDate?: string; // YYYY-MM-DD
  customLabel?: string; // Generated dynamic title label
}

export interface GlobalDataset {
  transactions: TransactionRecord[];
  monthlyMetrics: MonthlyMetric[];
  socios: SocioRecord[];
  weeklyBreakdownType1: WeeklySalesBreakdown[];
  weeklyBreakdownType2: {
    weekName: string;
    dateRange: string;
    amount: number;
    percentage: number;
    description: string;
    operationsCount?: number;
  }[];
  portfolioDurations: PortfolioDuration[];
  solutionCategories: SolutionCategory[];
  communities: CommunityMetric[];
  commercialCross: CommercialCrossPhase;
  cutoffDate: string;
  reportFilter?: ReportTimeFilter;
  updatedAt: string;
}
