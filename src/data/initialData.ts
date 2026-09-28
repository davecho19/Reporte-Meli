import { GlobalDataset } from '../types';

/**
 * Initial empty dataset.
 * In accordance with user requirements, the platform starts with blank data,
 * and information is populated dynamically when user uploads the templates
 * (Plantilla de Firmas o Plantilla de Socios).
 */
export const initialDataset: GlobalDataset = {
  cutoffDate: 'Sin corte definido',
  updatedAt: new Date().toISOString(),
  
  // Empty until filled by Excel template upload
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
    channel2Status: 'Pendiente de carga',
  },

  reportFilter: {
    type: 'all',
  },
};
