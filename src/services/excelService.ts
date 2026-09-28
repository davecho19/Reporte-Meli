import * as XLSX from 'xlsx';
import { GlobalDataset, MonthlyMetric, SocioRecord, TransactionRecord } from '../types';

export interface ParsedExcelResult {
  sheetNames: string[];
  activeSheet: string;
  recognizedType: 'transactions' | 'monthly_summary' | 'socios' | 'unknown';
  transactions: TransactionRecord[];
  monthlyMetrics?: MonthlyMetric[];
  socios?: SocioRecord[];
  rawRowsCount: number;
  unmappedColumns: string[];
  recognizedColumns: Record<string, string>;
  warnings: string[];
}

// Normalize strings for resilient column matching
const cleanKey = (key: string): string => {
  return String(key || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
};

/**
 * Universal flexible number parser supporting:
 * - Comma as decimal: "29,99", "15,00", "14,50" -> 29.99, 15.00, 14.50
 * - Dot as decimal: "29.99", "15.00", "3.16" -> 29.99, 15.00, 3.16
 * - Spanish thousands + decimals: "1.250,50", "15.000,00" -> 1250.50, 15000.00
 * - US thousands + decimals: "1,250.50", "15,000.00" -> 1250.50, 15000.00
 * - Integers: "30", "1500", "25000" -> 30, 1500, 25000
 * - Currency symbols: "$ 29.99", "$ 29,99", "29,99 USD", "USD 150.00" -> 29.99, 29.99, 29.99, 150.00
 * - Raw numeric values directly from Excel cells
 */
export const parseFlexibleNumber = (raw: any): number => {
  if (raw === null || raw === undefined || raw === '') return 0;
  // If SheetJS cell object
  if (typeof raw === 'object') {
    if ('v' in raw) raw = (raw as any).v;
    else if ('w' in raw) raw = (raw as any).w;
    else if ('value' in raw) raw = (raw as any).value;
    if (raw === null || raw === undefined || raw === '') return 0;
  }
  if (typeof raw === 'number') {
    return isNaN(raw) ? 0 : raw;
  }

  let str = String(raw).trim();
  // Strip currency symbols ($ , USD , EUR , €), non-breaking spaces, text notes
  str = str.replace(/[$€£\s\u00a0a-zA-Z]/g, '').trim();
  // Keep only digits, negative signs, dots, and commas
  str = str.replace(/[^0-9.,-]/g, '').trim();

  if (!str || str === '-') return 0;

  const isNegative = str.startsWith('-');
  if (isNegative) {
    str = str.substring(1);
  }

  const dotCount = (str.match(/\./g) || []).length;
  const commaCount = (str.match(/,/g) || []).length;

  let val = 0;

  if (dotCount === 0 && commaCount === 0) {
    // Pure integer e.g. "1500"
    val = parseFloat(str) || 0;
  } else if (dotCount > 0 && commaCount > 0) {
    // Both separators present: e.g. "1,250.50" (US) or "1.250,50" (ES/LatAm)
    const lastDot = str.lastIndexOf('.');
    const lastComma = str.lastIndexOf(',');
    if (lastDot > lastComma) {
      // US style: "1,250.50" -> dot is decimal
      val = parseFloat(str.replace(/,/g, '')) || 0;
    } else {
      // European/Latin style: "1.250,50" -> comma is decimal
      val = parseFloat(str.replace(/\./g, '').replace(',', '.')) || 0;
    }
  } else if (commaCount > 1) {
    // Multiple commas e.g. "1,000,000" -> thousand separators
    val = parseFloat(str.replace(/,/g, '')) || 0;
  } else if (dotCount > 1) {
    // Multiple dots e.g. "1.000.000" -> thousand separators
    val = parseFloat(str.replace(/\./g, '')) || 0;
  } else if (commaCount === 1) {
    // Single comma e.g. "29,99" or "15,00" or "15,000"
    const parts = str.split(',');
    if (parts[1] === '000' && parts[0].length >= 2) {
      // Explicit thousand e.g. "15,000" -> 15000
      val = parseFloat(parts[0] + '000') || 0;
    } else {
      // Standard comma as decimal e.g. "29,99" -> 29.99
      val = parseFloat(str.replace(',', '.')) || 0;
    }
  } else if (dotCount === 1) {
    // Single dot e.g. "29.99" or "15.00" or "15.000"
    const parts = str.split('.');
    if (parts[1] === '000' && parts[0].length >= 2) {
      // Explicit thousand e.g. "15.000" -> 15000
      val = parseFloat(parts[0] + '000') || 0;
    } else {
      // Standard dot as decimal e.g. "29.99" -> 29.99
      val = parseFloat(str) || 0;
    }
  }

  return isNegative ? -val : val;
};

/**
 * Columna BH en Excel corresponde al índice 59 (0-indexed, 60va columna en la cuadrícula).
 * Requerimiento de usuario: "el valor tomalo de la columna bh, pero hazlo bien, son valores en dolares ecuatorianos, trae la info que es, hazlo bien".
 */
export const EXCEL_COL_BH_INDEX = 59;

export const parseEcuadorianDollarValue = (val: any): number => {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') {
    return isNaN(val) ? 0 : parseFloat(val.toFixed(2));
  }

  if (typeof val === 'object') {
    if (typeof val.v === 'number') return parseFloat(val.v.toFixed(2));
    if (val.v !== undefined && val.v !== null) val = val.v;
    else if (val.w !== undefined && val.w !== null) val = val.w;
  }

  const str = String(val).trim();
  if (!str || str === '-' || str === 'N/A' || str.toLowerCase() === 'gratis') return 0;

  // Clean currency words and symbols: "USD", "usd", "$", spaces, non-breaking spaces
  let clean = str
    .replace(/usd/gi, '')
    .replace(/\$/g, '')
    .replace(/\s+/g, '')
    .replace(/\u00A0/g, '')
    .trim();

  if (!clean) return 0;

  const isNegative = clean.startsWith('-') || (clean.startsWith('(') && clean.endsWith(')'));
  clean = clean.replace(/[()\-]/g, '').trim();

  // Ecuadorian / Spanish format: "1.250,50" -> 1250.50
  if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(clean)) {
    clean = clean.replace(/\./g, '').replace(',', '.');
  }
  // US format: "1,250.50" -> 1250.50
  else if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(clean)) {
    clean = clean.replace(/,/g, '');
  }
  // Single comma format: "15,00" -> 15.00
  else if (clean.includes(',') && !clean.includes('.')) {
    clean = clean.replace(',', '.');
  }
  // Multiple mixed separators: determine decimal separator by the last one
  else if ((clean.match(/[,.]/g) || []).length > 1) {
    const lastComma = clean.lastIndexOf(',');
    const lastDot = clean.lastIndexOf('.');
    if (lastComma > lastDot) {
      clean = clean.replace(/\./g, '').replace(',', '.');
    } else {
      clean = clean.replace(/,/g, '');
    }
  }

  const num = parseFloat(clean);
  if (isNaN(num)) return 0;
  return isNegative ? -parseFloat(num.toFixed(2)) : parseFloat(num.toFixed(2));
};

// Universal flexible date parser
export const parseExcelDate = (rawDate: any): string => {
  if (!rawDate) return '2026-09-01';
  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
    return rawDate.toISOString().split('T')[0];
  }
  // Excel serial number (e.g. 46267)
  if (typeof rawDate === 'number' && rawDate > 20000 && rawDate < 60000) {
    const d = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) {
      return d.toISOString().split('T')[0];
    }
  }

  const str = String(rawDate).trim();
  // If already YYYY-MM-DD or YYYY/MM/DD
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(str)) {
    const parts = str.split(/[-/]/);
    return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
  }
  // If DD/MM/YYYY or D/M/YY
  if (/^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$/.test(str)) {
    const parts = str.split(/[-/]/);
    const d = parts[0].padStart(2, '0');
    const m = parts[1].padStart(2, '0');
    let y = parseInt(parts[2], 10);
    if (y < 100) y += 2000;
    return `${y}-${m}-${d}`;
  }
  return str || '2026-09-01';
};

// Check if header represents quantity / count of operations (not monetary value)
export const isQuantityColumnHeader = (ck: string): boolean => {
  // If explicitly says valor or monto or usd or dinero or precio or importe, it's monetary, NOT quantity!
  if (
    ck.includes('monto') ||
    ck.includes('valor') ||
    ck.includes('precio') ||
    ck.includes('importe') ||
    ck.includes('usd') ||
    ck.includes('dinero') ||
    ck.includes('recaud') ||
    ck.includes('tarifa') ||
    ck.includes('pvp') ||
    ck.includes('costo')
  ) {
    return false;
  }
  if (ck.includes('cantidad') || ck.includes('cantidades') || ck.startsWith('cant')) return true;
  if (ck.includes('numventas') || ck.includes('numeroventas') || ck.includes('unidades') || ck.includes('operaciones') || ck.includes('operacion')) return true;
  if (ck.includes('ventasrealizadas') || ck.includes('ventarealizada')) return true;
  if (ck === 'conteo' || ck === 'firmas' || ck === 'totalfirmas' || ck === 'totaloperaciones') return true;
  return false;
};

// Check if a cleaned column header corresponds to monetary value / amount (monto / valor)
export const isValueColumnHeader = (ck: string): boolean => {
  if (isQuantityColumnHeader(ck)) return false;
  // Direct matches for valor or monto
  if (ck.includes('valor') || ck.includes('monto')) return true;
  if (ck.includes('precio') || ck.includes('importe') || ck.includes('recaud') || ck.includes('totaldinero') || ck.includes('totalusd')) return true;
  if (ck.includes('montototal') || ck.includes('valortotal') || ck.includes('totalventa') || ck.includes('ventatotal') || ck.includes('totalventas')) return true;
  if (ck === 'total' || ck === 'subtotal' || ck === 'pvp' || ck === 'tarifa' || ck === 'costo' || ck === 'venta' || ck === 'ventas') return true;
  if (ck.startsWith('total') && !ck.includes('firma') && !ck.includes('operacion') && !ck.includes('cantidad') && !ck.includes('cliente') && !ck.includes('socio') && !ck.includes('unidades') && !ck.includes('vendedor')) return true;
  return false;
};

// Check if header represents a socio / vendedor
export const isSocioColumnHeader = (ck: string): boolean => {
  if (ck.includes('vendedor') || ck.includes('vendedora') || ck.includes('asesor') || ck.includes('agente') || ck.includes('distribuidor')) return true;
  if (ck.includes('socio') || ck.includes('partner') || ck.includes('franquiciado') || ck.includes('franquicia')) return true;
  if (ck === 'nombredelsocio' || ck === 'nombresocio' || ck === 'nombrevendedor' || ck === 'nombreasesor') return true;
  return false;
};

export const parseExcelFile = async (
  file: File,
  forcedType?: 'firma' | 'socios' | 'monthly'
): Promise<ParsedExcelResult> => {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { cellDates: true, cellNF: true, cellText: true, raw: true });
  
  const sheetNames = workbook.SheetNames;
  if (sheetNames.length === 0) {
    throw new Error('El archivo Excel no contiene ninguna hoja.');
  }

  // Find sheet with the most data rows
  let bestSheetName = sheetNames[0];
  let maxSheetRows = 0;
  for (const name of sheetNames) {
    const ws = workbook.Sheets[name];
    if (!ws) continue;
    const rows: any[] = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });
    if (rows.length > maxSheetRows) {
      maxSheetRows = rows.length;
      bestSheetName = name;
    }
  }

  const activeSheetName = bestSheetName;
  const worksheet = workbook.Sheets[activeSheetName];

  // Recalculate worksheet range to guarantee that Column BH (col 59) and any cells beyond are included
  if (worksheet) {
    let maxR = 0, maxC = 59, minR = 0, minC = 0, hasCells = false;
    for (const cellKey of Object.keys(worksheet)) {
      if (cellKey.startsWith('!')) continue;
      try {
        const decoded = XLSX.utils.decode_cell(cellKey);
        if (!hasCells) {
          minR = maxR = decoded.r;
          minC = maxC = decoded.c;
          hasCells = true;
        } else {
          if (decoded.r > maxR) maxR = decoded.r;
          if (decoded.c > maxC) maxC = decoded.c;
          if (decoded.r < minR) minR = decoded.r;
          if (decoded.c < minC) minC = decoded.c;
        }
      } catch {
        // ignore
      }
    }
    if (hasCells) {
      const origRange = worksheet['!ref']
        ? XLSX.utils.decode_range(worksheet['!ref'])
        : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
      worksheet['!ref'] = XLSX.utils.encode_range({
        s: { r: Math.min(origRange.s.r, minR), c: Math.min(origRange.s.c, minC) },
        e: { r: Math.max(origRange.e.r, maxR), c: Math.max(origRange.e.c, maxC, 59) }
      });
    }
  }

  // Convert to JSON with raw objects
  const rawRows: any[] = XLSX.utils.sheet_to_json(worksheet, { defval: '', header: 1 });

  if (rawRows.length === 0) {
    throw new Error('La hoja seleccionada está vacía.');
  }

  // Robust Header Row Finder:
  // Requires at least 2 non-empty cells to avoid matching single title banners
  let headerRowIndex = 0;
  let maxScore = -1;

  for (let i = 0; i < Math.min(12, rawRows.length); i++) {
    const row = rawRows[i];
    if (!Array.isArray(row)) continue;

    const nonEmptyCells = row.filter((c) => String(c || '').trim() !== '');
    if (nonEmptyCells.length < 2) continue; // Skip title banner lines

    let score = 0;
    let matchCount = 0;

    row.forEach((cell) => {
      const ck = cleanKey(String(cell || ''));
      if (!ck) return;
      if (ck === 'valor' || ck === 'monto' || ck.includes('valor') || ck.includes('monto')) { score += 12; matchCount++; }
      else if (ck.includes('vendedor') || ck.includes('socio') || ck.includes('asesor')) { score += 10; matchCount++; }
      else if (ck.includes('fecha') || ck.includes('date')) { score += 6; matchCount++; }
      else if (ck.includes('nombre') || ck.includes('apellido') || ck.includes('cliente')) { score += 6; matchCount++; }
      else if (ck.includes('id') || ck.includes('codigo') || ck.includes('unico')) { score += 5; matchCount++; }
      else if (ck.includes('estado') || ck.includes('status')) { score += 5; matchCount++; }
      else if (ck.includes('duracion') || ck.includes('vigencia')) { score += 5; matchCount++; }
      else if (ck.includes('tipo') || ck.includes('plan') || ck.includes('solucion') || ck.includes('producto')) { score += 5; matchCount++; }
      else if (ck.includes('cantidad') || ck.includes('ticket') || ck.includes('operacion')) { score += 6; matchCount++; }
      else if (ck.includes('rol') || ck.includes('canal')) { score += 4; matchCount++; }
      else if (ck.includes('mes') && (ck.includes('upconnect') || ck.includes('connectors'))) { score += 10; matchCount++; }
    });

    score += matchCount * 8;

    if (score > maxScore) {
      maxScore = score;
      headerRowIndex = i;
    }
  }

  const rawHeaders: string[] = (rawRows[headerRowIndex] || []).map((h: any) => String(h || '').trim());
  const dataRows = rawRows.slice(headerRowIndex + 1).filter((row) => Array.isArray(row) && row.some((c) => c !== ''));

  // Detect if this is a 9-column format file for Firmas or Socios (forced or detected)
  const is9ColFormat = forcedType === 'firma' || forcedType === 'socios' || 
    (rawHeaders.some((h) => cleanKey(h).includes('vendedor')) ||
     (rawHeaders.some((h) => cleanKey(h).includes('socio')) && !rawHeaders.some((h) => cleanKey(h).includes('idunico'))));

  if (is9ColFormat) {
    // Reference columns:
    // fecha, nombre, apellido, estado, tipo, duracion, valor/monto, vendedor, rol, cantidad, ticket, plan, id
    let fechaCol = -1;
    let nombreCol = -1;
    let apellidoCol = -1;
    let estadoCol = -1;
    let tipoCol = -1;
    let duracionCol = -1;
    let valorCol = -1;
    let vendedorCol = -1;
    let rolCol = -1;
    let canalCol = -1;
    let cantidadCol = -1;
    let ticketCol = -1;
    let planCol = -1;
    let legacyRankCol = -1;
    let idCol = -1;

    // COLUMNA BH (Excel Column 60, 0-indexed 59)
    // Requerimiento de usuario: "el valor tomalo de la columna bh, pero hazlo bien, son valores en dolares ecuatorianos, trae la info que es, hazlo bien"
    const hasColBH = rawHeaders.length > EXCEL_COL_BH_INDEX ||
      dataRows.some((r) => Array.isArray(r) && r.length > EXCEL_COL_BH_INDEX) ||
      Boolean(worksheet && Object.keys(worksheet).some((k) => /^BH\d+$/i.test(k)));

    if (hasColBH || forcedType === 'firma') {
      valorCol = EXCEL_COL_BH_INDEX;
    }

    // STEP 1: Detect valorCol by name only if Columna BH was not found
    if (valorCol === -1) {
      rawHeaders.forEach((h, idx) => {
        const ck = cleanKey(h);
        if (isQuantityColumnHeader(ck)) return;
        if (
          ck === 'valor' ||
          ck === 'monto' ||
          ck.startsWith('valor') ||
          ck.startsWith('monto') ||
          ck.includes('valortotal') ||
          ck.includes('montototal') ||
          ck.includes('valordeventa') ||
          ck.includes('montodeventa') ||
          ck.includes('valormonto') ||
          ck.includes('montovalor') ||
          ck.includes('valorusd') ||
          ck.includes('montousd') ||
          ck.includes('valorneto') ||
          ck.includes('montofacturado') ||
          ck.includes('valorunitario') ||
          ck.includes('valoremision') ||
          ck.includes('valorplan') ||
          ck.includes('valortransaccion') ||
          ck.includes('valorcobrado')
        ) {
          if (valorCol === -1) valorCol = idx;
        }
      });
    }

    if (valorCol === -1) {
      rawHeaders.forEach((h, idx) => {
        const ck = cleanKey(h);
        if (isQuantityColumnHeader(ck)) return;
        if (
          ck.includes('valor') ||
          ck.includes('monto') ||
          ck.includes('precio') ||
          ck.includes('importe') ||
          ck.includes('pvp') ||
          ck.includes('tarifa') ||
          ck.includes('costo') ||
          ck.includes('totalventa') ||
          ck.includes('ventatotal')
        ) {
          if (valorCol === -1) valorCol = idx;
        }
      });
    }

    if (valorCol === -1) {
      rawHeaders.forEach((h, idx) => {
        const ck = cleanKey(h);
        if (isValueColumnHeader(ck)) {
          if (valorCol === -1) valorCol = idx;
        }
      });
    }

    // STEP 2: Detect all remaining columns without overriding valorCol
    rawHeaders.forEach((h, idx) => {
      if (idx === valorCol) return; // Never override valorCol!
      const ck = cleanKey(h);

      if ((ck.includes('idunico') || ck === 'id' || ck === 'codigo' || ck === 'tramite' || ck === 'solicitud' || ck === 'identificador' || ck === 'ruc' || ck === 'cedula' || ck === 'nro') && idCol === -1) {
        idCol = idx;
      } else if (isSocioColumnHeader(ck) && vendedorCol === -1) {
        vendedorCol = idx;
      } else if ((ck === 'canal' || ck.includes('canal') || ck === 'channel' || ck.includes('reddistribucion')) && canalCol === -1) {
        canalCol = idx;
      } else if ((ck === 'rol' || ck === 'cargo' || ck === 'perfil' || ck.includes('rol') || ck.includes('distribuidor')) && rolCol === -1) {
        rolCol = idx;
      } else if ((ck.includes('fecha') || ck.includes('date')) && !ck.includes('venc') && fechaCol === -1) {
        fechaCol = idx;
      } else if (isQuantityColumnHeader(ck) && cantidadCol === -1) {
        cantidadCol = idx;
      } else if ((ck.includes('ticket') || ck.includes('promedio')) && ticketCol === -1) {
        ticketCol = idx;
      } else if ((ck.includes('plan') || ck.includes('solucion') || ck.includes('producto')) && planCol === -1) {
        planCol = idx;
        if (tipoCol === -1) tipoCol = idx;
      } else if ((ck === 'tipo' || ck === 'tipocertificado' || ck === 'tipoplan') && tipoCol === -1) {
        tipoCol = idx;
      } else if ((ck.includes('duracion') || ck.includes('vigencia') || ck.includes('tiempo')) && duracionCol === -1) {
        duracionCol = idx;
      } else if ((ck.includes('estado') || ck.includes('status')) && estadoCol === -1) {
        estadoCol = idx;
      } else if ((ck.includes('apellido') || ck.includes('apellid')) && apellidoCol === -1) {
        apellidoCol = idx;
      } else if ((ck === 'nombre' || ck === 'nombres' || ck === 'nombrecliente' || ck === 'cliente') && nombreCol === -1) {
        nombreCol = idx;
      } else if ((ck === 'ranking' || ck === 'rank' || ck === 'posicion') && legacyRankCol === -1) {
        legacyRankCol = idx;
      }
    });

    // If vendedorCol wasn't found but we have nombreCol in a socios context or without client last name
    if (vendedorCol === -1 && nombreCol !== -1 && (forcedType === 'socios' || (fechaCol === -1 && !apellidoCol))) {
      vendedorCol = nombreCol;
      nombreCol = -1;
    }

    // Smart fallback if valorCol was not matched by header text and Columna BH is not present:
    if (valorCol === -1) {
      let bestCol = -1;
      let maxNumericHits = 0;
      for (let c = 0; c < rawHeaders.length; c++) {
        if (c === fechaCol || c === nombreCol || c === apellidoCol || c === estadoCol || c === tipoCol || c === duracionCol || c === vendedorCol || c === rolCol || c === cantidadCol || c === legacyRankCol || c === idCol) {
          continue;
        }
        let hits = 0;
        for (let r = 0; r < Math.min(10, dataRows.length); r++) {
          const valNum = parseEcuadorianDollarValue(dataRows[r][c]);
          if (valNum > 0) hits++;
        }
        if (hits > maxNumericHits) {
          maxNumericHits = hits;
          bestCol = c;
        }
      }
      if (bestCol !== -1) {
        valorCol = bestCol;
      }
    }

    // Check if this file is a Socios Summary table (has quantity column, ticket, plan, or no date/client)
    // CRITICAL: When forcedType === 'firma', this is ALWAYS individual transactions/emissions, NEVER a summary table!
    const isSummaryTable = forcedType === 'firma'
      ? false
      : forcedType === 'socios' && (cantidadCol !== -1 || (ticketCol !== -1 && valorCol !== -1) || (fechaCol === -1 && apellidoCol === -1));

    const parsedTransactions: TransactionRecord[] = [];
    const socioMap: Record<string, {
      name: string;
      role: string;
      totalSales: number;
      operationsCount: number;
      averageTicket?: number;
      topPlan?: string;
      plans: Record<string, number>;
    }> = {};

    dataRows.forEach((row, idx) => {
      // 1. Vendedor / Socio name
      let socioName = '';
      if (vendedorCol !== -1 && row[vendedorCol]) {
        socioName = String(row[vendedorCol]).trim();
      } else if (nombreCol !== -1 && row[nombreCol]) {
        socioName = String(row[nombreCol]).trim();
      }

      if (socioName.toLowerCase() === 'total' || socioName.toLowerCase().includes('total general')) {
        return;
      }
      if (!socioName) {
        socioName = 'UpConnect Directo';
      }

      // 2. Rol y Canal
      const role = rolCol !== -1 && row[rolCol] ? String(row[rolCol]).trim() : '';
      const canalVal = canalCol !== -1 && row[canalCol] ? String(row[canalCol]).trim() : '';
      const combinedInfo = `${canalVal} ${role} ${socioName}`.toLowerCase();

      // Resolver canal con máxima precisión: UpConnect vs Connectors
      let transactionChannel: 'UpConnect' | 'Connectors' = 'UpConnect';
      if (
        combinedInfo.includes('upconnect') ||
        combinedInfo.includes('up connect') ||
        combinedInfo.includes('directo') ||
        combinedInfo.includes('propio')
      ) {
        transactionChannel = 'UpConnect';
      } else if (
        combinedInfo.includes('connectors') ||
        combinedInfo.includes('connector') ||
        combinedInfo.includes('distribuidor') ||
        combinedInfo.includes('aliado') ||
        combinedInfo.includes('franquicia') ||
        (combinedInfo.includes('connect') && !combinedInfo.includes('up'))
      ) {
        transactionChannel = 'Connectors';
      } else {
        transactionChannel = 'UpConnect';
      }
      const isConnector = transactionChannel === 'Connectors';
      const effectiveRole = role || (isConnector ? 'Distribuidor Connect' : 'Ventas Directas');

      if (isSummaryTable) {
        // SUMMARY ROW: each row is already one socio with total sales and count
        const opCount = cantidadCol !== -1 ? Math.max(1, Math.round(parseFlexibleNumber(row[cantidadCol]))) : 1;
        const totalSales = valorCol !== -1 ? parseFlexibleNumber(row[valorCol]) : 0;
        const avgTicket = ticketCol !== -1 ? parseFlexibleNumber(row[ticketCol]) : (opCount > 0 ? totalSales / opCount : totalSales);
        const topPlan = (planCol !== -1 && row[planCol]) ? String(row[planCol]).trim() : (tipoCol !== -1 && row[tipoCol] ? String(row[tipoCol]).trim() : 'Firma Electrónica (1 año)');

        if (!socioMap[socioName]) {
          socioMap[socioName] = {
            name: socioName,
            role: effectiveRole,
            totalSales: parseFloat(totalSales.toFixed(2)),
            operationsCount: opCount,
            averageTicket: parseFloat(avgTicket.toFixed(2)),
            topPlan,
            plans: { [topPlan]: opCount },
          };
        } else {
          socioMap[socioName].totalSales += parseFloat(totalSales.toFixed(2));
          socioMap[socioName].operationsCount += opCount;
          socioMap[socioName].averageTicket = socioMap[socioName].operationsCount > 0
            ? parseFloat((socioMap[socioName].totalSales / socioMap[socioName].operationsCount).toFixed(2))
            : socioMap[socioName].totalSales;
        }

        // Create transaction representation so charts, monthly totals and channel metrics reflect the exact sales
        parsedTransactions.push({
          id: `socio-trx-${Date.now()}-${idx}`,
          uniqueId: `SOC-${idx + 1}`,
          date: '2026-09-20',
          clientName: socioName,
          clientLastName: '(Socio)',
          certificateStatus: 'EMITIDO',
          duration: 'Un año',
          value: totalSales,
          role: effectiveRole,
          channel: transactionChannel,
          socio: socioName,
          solutionCategory: topPlan,
          createdAt: new Date().toISOString(),
          sourceFile: file.name,
        });
      } else {
        // DETAILED TRANSACTION ROW: 1 row = 1 emission
        const tipo = tipoCol !== -1 && row[tipoCol] ? String(row[tipoCol]).trim() : 'Firma Electrónica';
        const duracion = duracionCol !== -1 && row[duracionCol] ? String(row[duracionCol]).trim() : 'Un año';

        // EXTRACCIÓN DEL VALOR DE LA COLUMNA BH (DÓLARES ECUATORIANOS)
        // User specification: "el valor tomalo de la columna bh, pero hazlo bien, son valores en dolares ecuatorianos, trae la info que es, hazlo bien"
        const excelRowNum = headerRowIndex + 2 + idx;
        let rawValBH: any = undefined;

        // 1. Prioridad: Columna BH (índice 59) del array de la fila
        if (row && row.length > EXCEL_COL_BH_INDEX && row[EXCEL_COL_BH_INDEX] !== undefined && row[EXCEL_COL_BH_INDEX] !== null && String(row[EXCEL_COL_BH_INDEX]).trim() !== '') {
          rawValBH = row[EXCEL_COL_BH_INDEX];
        }

        // 2. Celda directa en worksheet BH{excelRowNum} para soportar archivos con rangos dispersos
        if ((rawValBH === undefined || rawValBH === null || String(rawValBH).trim() === '') && worksheet) {
          const wsCell = worksheet[`BH${excelRowNum}`];
          if (wsCell !== undefined && wsCell !== null) {
            rawValBH = wsCell.v !== undefined && wsCell.v !== null ? wsCell.v : wsCell.w;
          }
        }

        let val = 0;
        if (rawValBH !== undefined && rawValBH !== null && String(rawValBH).trim() !== '') {
          val = parseEcuadorianDollarValue(rawValBH);
        } else if (valorCol !== -1 && valorCol !== EXCEL_COL_BH_INDEX && row && valorCol < row.length) {
          // Solo si el archivo no contenía columna BH, usar la columna de valor detectada
          const fallbackVal = row[valorCol];
          if (fallbackVal !== undefined && fallbackVal !== null && String(fallbackVal).trim() !== '') {
            val = parseEcuadorianDollarValue(fallbackVal);
          }
        }

        // Operaciones / Cantidad
        const opQty = cantidadCol !== -1 && row[cantidadCol]
          ? Math.max(1, Math.round(parseFlexibleNumber(row[cantidadCol])))
          : 1;

        // Fecha
        const dateStr = fechaCol !== -1 ? parseExcelDate(row[fechaCol]) : '2026-09-01';

        // Cliente
        const clientName = nombreCol !== -1 && row[nombreCol] ? String(row[nombreCol]).trim() : 'Cliente';
        const clientLastName = apellidoCol !== -1 && row[apellidoCol] ? String(row[apellidoCol]).trim() : '';

        // ID único
        const rawId = idCol !== -1 && row[idCol] ? String(row[idCol]).trim() : '';
        const uniqueId = rawId || (forcedType === 'firma' ? `FIRM-${idx + 1}` : `TRX-${idx + 1}`);

        // Estado
        const rawStatus = estadoCol !== -1 && row[estadoCol] ? String(row[estadoCol]).trim().toUpperCase() : 'EMITIDO';
        const certificateStatus = rawStatus.includes('CANCEL') ? 'CANCELADO' : rawStatus.includes('PROCES') ? 'EN PROCESO' : 'EMITIDO';

        // Aggregate by Socio (Vendedor)
        if (!socioMap[socioName]) {
          socioMap[socioName] = {
            name: socioName,
            role: effectiveRole,
            totalSales: 0,
            operationsCount: 0,
            plans: {},
          };
        }
        socioMap[socioName].totalSales += val;
        socioMap[socioName].operationsCount += opQty;
        if (effectiveRole && effectiveRole !== 'Distribuidor Connect') {
          socioMap[socioName].role = effectiveRole;
        }

        // Count plan frequency
        const planName = tipo && duracion ? `${tipo} (${duracion})` : (tipo || duracion || 'Plan Estándar');
        socioMap[socioName].plans[planName] = (socioMap[socioName].plans[planName] || 0) + opQty;

        parsedTransactions.push({
          id: `trx-${Date.now()}-${idx}`,
          uniqueId,
          date: dateStr,
          clientName,
          clientLastName,
          certificateStatus,
          duration: duracion,
          value: val,
          role: effectiveRole,
          channel: transactionChannel,
          socio: socioName,
          solutionCategory: tipo,
          createdAt: new Date().toISOString(),
          sourceFile: file.name,
        });
      }
    });

    // Build SocioRecord array with all requested calculations:
    // nombre socio, cantidades ventas realizadas, monto total de venta, ticket promedio, plan mas vendido, rol
    const parsedSocios: SocioRecord[] = Object.values(socioMap).map((agg) => {
      let topPlan = agg.topPlan || 'Firma Electrónica (1 año)';
      if (!agg.topPlan) {
        let maxCount = -1;
        for (const [pName, count] of Object.entries(agg.plans)) {
          if (count > maxCount) {
            maxCount = count;
            topPlan = pName;
          }
        }
      }

      const averageTicket = agg.averageTicket !== undefined
        ? agg.averageTicket
        : (agg.operationsCount > 0 ? (agg.totalSales / agg.operationsCount) : 0);

      return {
        rank: 0,
        name: agg.name,
        totalSales: parseFloat(agg.totalSales.toFixed(2)),
        group: 'TOP 1-10',
        operationsCount: agg.operationsCount,
        averageTicket: parseFloat(averageTicket.toFixed(2)),
        topPlan,
        role: agg.role || 'Distribuidor Connect',
        note: undefined,
      };
    });

    // Sort by sales descending & assign ranks
    parsedSocios.sort((a, b) => b.totalSales - a.totalSales);
    parsedSocios.forEach((s, idx) => {
      s.rank = idx + 1;
      s.group = idx < 10 ? 'TOP 1-10' : 'TOP 11-20';
      if (idx === 0) s.note = 'Líder en Ventas';
    });

    const recognizedCols: Record<string, string> = {};
    if (fechaCol !== -1) recognizedCols['Fecha'] = rawHeaders[fechaCol];
    if (nombreCol !== -1) recognizedCols['Nombre'] = rawHeaders[nombreCol];
    if (apellidoCol !== -1) recognizedCols['Apellido'] = rawHeaders[apellidoCol];
    if (estadoCol !== -1) recognizedCols['Estado'] = rawHeaders[estadoCol];
    if (tipoCol !== -1) recognizedCols['Tipo'] = rawHeaders[tipoCol];
    if (duracionCol !== -1) recognizedCols['Duración'] = rawHeaders[duracionCol];
    if (valorCol !== -1) recognizedCols['Valor / Monto'] = rawHeaders[valorCol];
    if (cantidadCol !== -1) recognizedCols['Cantidades Ventas'] = rawHeaders[cantidadCol];
    if (vendedorCol !== -1) recognizedCols['Vendedor (Socio)'] = rawHeaders[vendedorCol];
    if (rolCol !== -1) recognizedCols['Rol'] = rawHeaders[rolCol];
    if (ticketCol !== -1) recognizedCols['Ticket Promedio'] = rawHeaders[ticketCol];
    if (planCol !== -1) recognizedCols['Plan Más Vendido'] = rawHeaders[planCol];

    return {
      sheetNames,
      activeSheet: activeSheetName,
      recognizedType: forcedType === 'socios' ? 'socios' : 'transactions',
      transactions: parsedTransactions,
      socios: parsedSocios,
      rawRowsCount: dataRows.length,
      unmappedColumns: [],
      recognizedColumns: recognizedCols,
      warnings: [],
    };
  }

  // Detect if this is a Monthly Summary sheet (like Image 3)
  const isMonthlyTable = forcedType === 'monthly' || (rawHeaders.some((h) => cleanKey(h).includes('mes')) &&
    rawHeaders.some((h) => cleanKey(h).includes('upconnect')));

  const recognizedColumns: Record<string, string> = {};
  const unmappedColumns: string[] = [];
  const warnings: string[] = [];

  if (isMonthlyTable) {
    // Process monthly summary format
    const parsedMonthly: MonthlyMetric[] = [];
    
    // Map column positions
    let mesCol = -1;
    let upSalesCol = -1;
    let coSalesCol = -1;
    let upCountCol = -1;
    let coCountCol = -1;

    rawHeaders.forEach((h, idx) => {
      const ck = cleanKey(h);
      if (ck === 'mes') mesCol = idx;
      else if (ck.includes('upconnect') && !ck.includes('numero') && upSalesCol === -1) upSalesCol = idx;
      else if (ck.includes('connectors') && !ck.includes('numero') && coSalesCol === -1) coSalesCol = idx;
      else if ((ck.includes('numero') || ck.includes('num')) && ck.includes('upconnect')) upCountCol = idx;
      else if ((ck.includes('numero') || ck.includes('num')) && ck.includes('connectors')) coCountCol = idx;
      else if (upCountCol === -1 && idx === upSalesCol + 2) upCountCol = idx;
      else if (coCountCol === -1 && idx === upSalesCol + 3) coCountCol = idx;
    });

    const monthMap: Record<string, number> = {
      enero: 1, febrero: 2, marzo: 3, abril: 4, mayo: 5, junio: 6,
      julio: 7, agosto: 8, septiembre: 9, octubre: 10, noviembre: 11, diciembre: 12
    };

    dataRows.forEach((row) => {
      const rawMonth = String(row[mesCol] || '').trim();
      const cleanMonthStr = cleanKey(rawMonth);
      if (!cleanMonthStr || cleanMonthStr === 'total') return;

      const monthIndex = monthMap[cleanMonthStr] || 9;
      const upSales = parseFlexibleNumber(row[upSalesCol]);
      const coSales = parseFlexibleNumber(row[coSalesCol]);
      const upCount = parseInt(String(row[upCountCol] || '0').replace(/[^0-9]/g, ''), 10) || 0;
      const coCount = parseInt(String(row[coCountCol] || '0').replace(/[^0-9]/g, ''), 10) || 0;

      parsedMonthly.push({
        month: rawMonth.toUpperCase(),
        monthIndex,
        upconnectSales: upSales,
        connectorsSales: coSales,
        upconnectCount: upCount,
        connectorsCount: coCount,
      });
    });

    return {
      sheetNames,
      activeSheet: activeSheetName,
      recognizedType: 'monthly_summary',
      transactions: [],
      monthlyMetrics: parsedMonthly,
      rawRowsCount: dataRows.length,
      unmappedColumns: [],
      recognizedColumns: {
        'Mes': rawHeaders[mesCol] || 'MES',
        'Ventas UpConnect': rawHeaders[upSalesCol] || 'UPCONNECT',
        'Ventas Connectors': rawHeaders[coSalesCol] || 'CONNECTORS',
      },
      warnings,
    };
  }

  // Otherwise, treat as detailed transactional report (Image 1 & 2)
  let idCol = -1;
  let dateCol = -1;
  let nameCol = -1;
  let lastNameCol = -1;
  let statusCol = -1;
  let durationCol = -1;
  let valueCol = -1;
  let roleCol = -1;
  let socioCol = -1;
  let categoryCol = -1;

  // COLUMNA BH (Excel Column 60, 0-indexed 59)
  // User specification: "el valor tomalo de la columna bh, pero hazlo bien, son valores en dolares ecuatorianos, trae la info que es, hazlo bien"
  const hasColBH = rawHeaders.length > EXCEL_COL_BH_INDEX ||
    dataRows.some((r) => Array.isArray(r) && r.length > EXCEL_COL_BH_INDEX) ||
    Boolean(worksheet && Object.keys(worksheet).some((k) => /^BH\d+$/i.test(k)));

  if (hasColBH || forcedType === 'firma') {
    valueCol = EXCEL_COL_BH_INDEX;
    recognizedColumns['Valor / Monto (Columna BH)'] = rawHeaders[EXCEL_COL_BH_INDEX] ? `${rawHeaders[EXCEL_COL_BH_INDEX]} (Col. BH)` : 'Columna BH (USD)';
  }

  // STEP 1: High priority for Valor / Monto column if Columna BH is not present
  if (valueCol === -1) {
    rawHeaders.forEach((h, idx) => {
      const ck = cleanKey(h);
      if (isQuantityColumnHeader(ck)) return;
      if (ck === 'valor' || ck === 'monto' || ck.startsWith('valor') || ck.startsWith('monto') ||
          ck.includes('valortotal') || ck.includes('montototal') || ck.includes('valordeventa') || ck.includes('montodeventa') ||
          ck.includes('valormonto') || ck.includes('montovalor') || ck.includes('valorusd') || ck.includes('montousd') ||
          ck.includes('valorneto') || ck.includes('montofacturado')) {
        if (valueCol === -1) {
          valueCol = idx;
          recognizedColumns['Valor / Monto'] = h;
        }
      }
    });
  }

  if (valueCol === -1) {
    rawHeaders.forEach((h, idx) => {
      const ck = cleanKey(h);
      if (isQuantityColumnHeader(ck)) return;
      if (ck.includes('valor') || ck.includes('monto')) {
        if (valueCol === -1) {
          valueCol = idx;
          recognizedColumns['Valor / Monto'] = h;
        }
      }
    });
  }

  if (valueCol === -1) {
    rawHeaders.forEach((h, idx) => {
      const ck = cleanKey(h);
      if (isValueColumnHeader(ck)) {
        if (valueCol === -1) {
          valueCol = idx;
          recognizedColumns['Valor / Monto'] = h;
        }
      }
    });
  }

  // STEP 2: Detect all remaining columns without overriding valueCol
  rawHeaders.forEach((h, idx) => {
    if (idx === valueCol) return;
    const ck = cleanKey(h);
    if ((ck.includes('idunico') || ck === 'id' || ck === 'codigo' || ck === 'identificador' || ck === 'tramite' || ck === 'solicitud' || ck === 'ruc' || ck === 'cedula' || ck === 'nro') && idCol === -1) {
      idCol = idx;
      recognizedColumns['ID único'] = h;
    } else if ((ck.includes('fechasolic') || ck.includes('fecha') || ck === 'datesolic') && dateCol === -1) {
      dateCol = idx;
      recognizedColumns['Fecha Solic'] = h;
    } else if ((ck.includes('nombre') || ck === 'nombres') && nameCol === -1) {
      nameCol = idx;
      recognizedColumns['Nombres'] = h;
    } else if ((ck.includes('apellido') || ck === 'apellidos') && lastNameCol === -1) {
      lastNameCol = idx;
      recognizedColumns['Apellidos'] = h;
    } else if ((ck.includes('certificadoestado') || ck.includes('estado')) && statusCol === -1) {
      statusCol = idx;
      recognizedColumns['Certificado Estado'] = h;
    } else if ((ck.includes('certificadoduracion') || ck.includes('duracion') || ck.includes('vigencia')) && durationCol === -1) {
      durationCol = idx;
      recognizedColumns['Certificado Duración'] = h;
    } else if ((ck.includes('rol') || ck.includes('canal') || ck.includes('distribuidor')) && roleCol === -1) {
      roleCol = idx;
      recognizedColumns['Rol / Canal'] = h;
    } else if ((ck.includes('socio') || ck.includes('vendedor') || ck.includes('agente') || ck.includes('asesor') || ck.includes('operador')) && socioCol === -1) {
      socioCol = idx;
      recognizedColumns['Socio'] = h;
    } else if ((ck.includes('solucion') || ck.includes('categoria') || ck.includes('producto')) && categoryCol === -1) {
      categoryCol = idx;
      recognizedColumns['Categoría'] = h;
    } else {
      unmappedColumns.push(h);
    }
  });

  // Fallback search for valueCol if not found by header and no Columna BH
  if (valueCol === -1) {
    for (let c = 0; c < rawHeaders.length; c++) {
      if (c === idCol || c === dateCol || c === nameCol || c === lastNameCol || c === statusCol || c === durationCol || c === roleCol || c === socioCol || c === categoryCol) continue;
      const hasNumbers = dataRows.slice(0, 5).some((row) => parseEcuadorianDollarValue(row[c]) > 0);
      if (hasNumbers) {
        valueCol = c;
        recognizedColumns['Valor / Monto'] = rawHeaders[c] || 'Monto';
        break;
      }
    }
  }

    const parsedTransactions: TransactionRecord[] = [];
    const socioMap: Record<string, { total: number; count: number; role: string; plans: Record<string, number> }> = {};

    dataRows.forEach((row, rIdx) => {
      // If row is completely empty, skip
      if (!row || row.length === 0) return;

      const rawId = idCol !== -1 ? String(row[idCol] || '').trim() : '';
      const uniqueId = rawId || (forcedType === 'firma' ? `FIRM-${rIdx + 1}` : `gen-${Date.now()}-${rIdx}`);
      
      // Parse Date
      const dateStr = dateCol !== -1 ? parseExcelDate(row[dateCol]) : '2026-09-01';

      const clientName = nameCol !== -1 ? String(row[nameCol] || '').trim() : 'Cliente';
      const clientLastName = lastNameCol !== -1 ? String(row[lastNameCol] || '').trim() : '';
      
      const rawStatus = statusCol !== -1 ? String(row[statusCol] || '').trim().toUpperCase() : 'EMITIDO';
      const certificateStatus = rawStatus.includes('EMIT') ? 'EMITIDO' : 
                                rawStatus.includes('PROCES') ? 'EN PROCESO' : 
                                rawStatus.includes('CANCEL') ? 'CANCELADO' : 'EMITIDO';

      const rawDuration = durationCol !== -1 ? String(row[durationCol] || '').trim() : 'Un año';
      let duration = 'Un año';
      const durLower = rawDuration.toLowerCase();
      if (durLower.includes('15') || durLower.includes('quince')) duration = '15 días';
      else if (durLower.includes('5') || durLower.includes('cinco')) duration = 'Cinco años';
      else if (durLower.includes('4') || durLower.includes('cuatro')) duration = 'Cuatro años';
      else if (durLower.includes('3') || durLower.includes('tres')) duration = 'Tres años';
      else if (durLower.includes('2') || durLower.includes('dos')) duration = 'Dos años';
      else if (durLower.includes('1') || durLower.includes('un')) duration = 'Un año';
      else duration = rawDuration || 'Un año';

      // EXTRACCIÓN DEL VALOR DE LA COLUMNA BH (DÓLARES ECUATORIANOS)
      // User specification: "el valor tomalo de la columna bh, pero hazlo bien, son valores en dolares ecuatorianos, trae la info que es, hazlo bien"
      const excelRowNum = headerRowIndex + 2 + rIdx;
      let rawValBH: any = undefined;

      if (row && row.length > EXCEL_COL_BH_INDEX && row[EXCEL_COL_BH_INDEX] !== undefined && row[EXCEL_COL_BH_INDEX] !== null && String(row[EXCEL_COL_BH_INDEX]).trim() !== '') {
        rawValBH = row[EXCEL_COL_BH_INDEX];
      }

      if ((rawValBH === undefined || rawValBH === null || String(rawValBH).trim() === '') && worksheet) {
        const wsCell = worksheet[`BH${excelRowNum}`];
        if (wsCell !== undefined && wsCell !== null) {
          rawValBH = wsCell.v !== undefined && wsCell.v !== null ? wsCell.v : wsCell.w;
        }
      }

      let val = 0;
      if (rawValBH !== undefined && rawValBH !== null && String(rawValBH).trim() !== '') {
        val = parseEcuadorianDollarValue(rawValBH);
      } else if (valueCol !== -1 && valueCol !== EXCEL_COL_BH_INDEX && row && valueCol < row.length) {
        const fallbackVal = row[valueCol];
        if (fallbackVal !== undefined && fallbackVal !== null && String(fallbackVal).trim() !== '') {
          val = parseEcuadorianDollarValue(fallbackVal);
        }
      }

      // Role & Channel
      const rawRole = roleCol !== -1 ? String(row[roleCol] || '').trim() : 'Distribuidor Connect';
      const isConnector = rawRole.toLowerCase().includes('connect') || rawRole.toLowerCase().includes('distribuidor');
      const channel: 'UpConnect' | 'Connectors' = isConnector ? 'Connectors' : 'UpConnect';

      // Socio
      let socio = socioCol !== -1 ? String(row[socioCol] || '').trim() : '';
      if (!socio) {
        socio = channel === 'Connectors' ? 'Distribuidor Connect' : 'UpConnect Directo';
      }

      // Solution category
      let solutionCategory = categoryCol !== -1 ? String(row[categoryCol] || '').trim() : '';
      if (!solutionCategory) {
        if (val > 50 || duration === 'Cinco años') {
          solutionCategory = 'Planes de Contador';
        } else if (duration === '15 días') {
          solutionCategory = 'Facturación Electrónica';
        } else {
          solutionCategory = 'ERP Contables y Módulos';
        }
      }

      // Aggregate for socios ranking
      if (!socioMap[socio]) {
        socioMap[socio] = { total: 0, count: 0, role: rawRole || (channel === 'Connectors' ? 'Distribuidor Connect' : 'UpConnect Directo'), plans: {} };
      }
      socioMap[socio].total += val;
      socioMap[socio].count += 1;
      const pName = solutionCategory ? `${solutionCategory} (${duration})` : duration;
      socioMap[socio].plans[pName] = (socioMap[socio].plans[pName] || 0) + 1;

      parsedTransactions.push({
        id: `trx-${Date.now()}-${rIdx}`,
        uniqueId,
        date: dateStr,
        clientName,
        clientLastName,
        certificateStatus,
        duration,
        value: val,
        role: rawRole || (channel === 'Connectors' ? 'Distribuidor Connect' : 'UpConnect Directo'),
        channel,
        socio,
        solutionCategory,
        createdAt: new Date().toISOString(),
        sourceFile: file.name,
      });
    });

    const parsedSocios: SocioRecord[] = Object.entries(socioMap)
      .map(([name, data]) => {
        let topPlan = 'Firma Electrónica (1 año)';
        let maxCount = -1;
        for (const [pName, count] of Object.entries(data.plans)) {
          if (count > maxCount) {
            maxCount = count;
            topPlan = pName;
          }
        }
        const avgTicket = data.count > 0 ? (data.total / data.count) : 0;
        return {
          rank: 0,
          name,
          totalSales: parseFloat(data.total.toFixed(2)),
          group: 'TOP 1-10' as const,
          operationsCount: data.count,
          averageTicket: parseFloat(avgTicket.toFixed(2)),
          topPlan,
          role: data.role || 'Distribuidor Connect',
          note: undefined,
        };
      })
      .sort((a, b) => b.totalSales - a.totalSales);

    parsedSocios.forEach((s, idx) => {
      s.rank = idx + 1;
      s.group = idx < 10 ? 'TOP 1-10' : 'TOP 11-20';
      if (idx === 0) s.note = 'Líder en Ventas';
    });

    return {
      sheetNames,
      activeSheet: activeSheetName,
      recognizedType: 'transactions',
      transactions: parsedTransactions,
      socios: parsedSocios,
      rawRowsCount: dataRows.length,
      unmappedColumns,
      recognizedColumns,
      warnings,
    };
};

/**
 * Downloads a rich formatted Excel file with current dataset
 */
export const exportDatasetToExcel = (dataset: GlobalDataset, filename = 'Reporte_Gerencial_UpConnect.xlsx'): void => {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Transactions
  const trxData = dataset.transactions.map((t, idx) => ({
    'N°': idx + 1,
    'ID único': t.uniqueId,
    'Fecha Solic': t.date,
    'Nombres': t.clientName,
    'Apellidos': t.clientLastName,
    'Certificado Estado': t.certificateStatus,
    'Certificado Duración': t.duration,
    'Valor ($)': t.value,
    'Rol': t.role,
    'Canal': t.channel,
    'Socio Asignado': t.socio || '',
    'Categoría Solución': t.solutionCategory || '',
  }));
  const wsTrx = XLSX.utils.json_to_sheet(trxData);
  XLSX.utils.book_append_sheet(wb, wsTrx, 'Transacciones');

  // Sheet 2: Monthly Evolution
  const monthlyData = dataset.monthlyMetrics.map((m) => ({
    'MES': m.month,
    'VENTAS UPCONNECT ($)': m.upconnectSales,
    'VENTAS CONNECTORS ($)': m.connectorsSales,
    'TOTAL VENTAS ($)': m.upconnectSales + m.connectorsSales,
    'FIRMAS UPCONNECT': m.upconnectCount,
    'FIRMAS CONNECTORS': m.connectorsCount,
    'TOTAL FIRMAS': m.upconnectCount + m.connectorsCount,
    'NOTAS': m.notes || '',
  }));
  const wsMonthly = XLSX.utils.json_to_sheet(monthlyData);
  XLSX.utils.book_append_sheet(wb, wsMonthly, 'Consolidado Mensual');

  // Sheet 3: Ranking Socios
  const sociosData = dataset.socios.map((s) => ({
    'Ranking': s.rank,
    'Nombre Socio': s.name,
    'Rol': s.role || 'Distribuidor Connect',
    'Ventas Realizadas': s.operationsCount || 0,
    'Total Ventas ($)': s.totalSales,
    'Ticket Promedio ($)': s.averageTicket ? s.averageTicket.toFixed(2) : (s.totalSales / (s.operationsCount || 1)).toFixed(2),
    'Plan Más Vendido': s.topPlan || 'Firma Electrónica (1 año)',
    'Grupo': s.group,
    'Destacado': s.note || '',
  }));
  const wsSocios = XLSX.utils.json_to_sheet(sociosData);
  XLSX.utils.book_append_sheet(wb, wsSocios, 'Ranking Socios');

  // Trigger download
  XLSX.writeFile(wb, filename);
};

/**
 * Downloads a template Excel specifically for Firmas / Emisiones
 * Contemplates the 9 requested columns: Fecha, Nombre, Apellido, Estado, Tipo, Duración, Valor, Vendedor, Rol
 */
export const downloadTemplateFirmas = (): void => {
  const wb = XLSX.utils.book_new();

  const templateRows = [
    {
      'Fecha': '2026-09-01',
      'Nombre': 'Carlos Andrés',
      'Apellido': 'Mendoza Reyes',
      'Estado': 'EMITIDO',
      'Tipo': 'Firma Electrónica',
      'Duración': '15 días',
      'Valor': 3.16,
      'Vendedor': 'Gabriel Endara',
      'Rol': 'Distribuidor Connect',
    },
    {
      'Fecha': '2026-09-02',
      'Nombre': 'María Elena',
      'Apellido': 'López Castro',
      'Estado': 'EMITIDO',
      'Tipo': 'Firma Electrónica',
      'Duración': 'Un año',
      'Valor': 8.99,
      'Vendedor': 'Katherine Cabrera',
      'Rol': 'Distribuidor Connect',
    },
    {
      'Fecha': '2026-09-03',
      'Nombre': 'Juan José',
      'Apellido': 'Alvarez Vintimilla',
      'Estado': 'EMITIDO',
      'Tipo': 'Firma Electrónica',
      'Duración': 'Cinco años',
      'Valor': 28.17,
      'Vendedor': 'Andrea Burbano',
      'Rol': 'Distribuidor Connect',
    },
    {
      'Fecha': '2026-09-08',
      'Nombre': 'Rosa María',
      'Apellido': 'Ortega Ortiz',
      'Estado': 'EMITIDO',
      'Tipo': 'Firma Electrónica',
      'Duración': 'Dos años',
      'Valor': 14.50,
      'Vendedor': 'Katherine Cabrera',
      'Rol': 'Distribuidor Connect',
    },
    {
      'Fecha': '2026-09-09',
      'Nombre': 'Contabilidad Integral',
      'Apellido': 'S.A.',
      'Estado': 'EMITIDO',
      'Tipo': 'ERP Contables y Módulos',
      'Duración': 'Un año',
      'Valor': 29.99,
      'Vendedor': 'Gabriela González',
      'Rol': 'UpConnect Directo',
    },
  ];

  const ws = XLSX.utils.json_to_sheet(templateRows);
  XLSX.utils.book_append_sheet(wb, ws, 'Plantilla Firmas 9 Columnas');
  XLSX.writeFile(wb, 'Plantilla_Carga_Firmas.xlsx');
};

/**
 * Downloads a template Excel specifically for Socios
 * ONLY contemplates the 9 requested columns:
 * Fecha, Nombre, Apellido, Estado, Tipo, Duración, Valor, Vendedor, Rol
 */
export const downloadTemplateSocios = (): void => {
  const wb = XLSX.utils.book_new();

  const sociosRows = [
    {
      'Fecha': '2026-09-01',
      'Nombre': 'Carlos Andrés',
      'Apellido': 'Mendoza Reyes',
      'Estado': 'EMITIDO',
      'Tipo': 'Firma Electrónica',
      'Duración': 'Un año',
      'Valor': 29.99,
      'Vendedor': 'Katherine Cabrera',
      'Rol': 'Distribuidor Connect',
    },
    {
      'Fecha': '2026-09-02',
      'Nombre': 'María Elena',
      'Apellido': 'López Castro',
      'Estado': 'EMITIDO',
      'Tipo': 'ERP Contable',
      'Duración': 'Dos años',
      'Valor': 149.50,
      'Vendedor': 'Katherine Cabrera',
      'Rol': 'Distribuidor Connect',
    },
    {
      'Fecha': '2026-09-03',
      'Nombre': 'Juan José',
      'Apellido': 'Alvarez Vintimilla',
      'Estado': 'EMITIDO',
      'Tipo': 'Facturación Electrónica',
      'Duración': '15 días',
      'Valor': 15.00,
      'Vendedor': 'Gabriel Endara',
      'Rol': 'Socio Franquicia',
    },
    {
      'Fecha': '2026-09-05',
      'Nombre': 'Rosa María',
      'Apellido': 'Ortega Ortiz',
      'Estado': 'EMITIDO',
      'Tipo': 'Firma Electrónica',
      'Duración': 'Cinco años',
      'Valor': 85.00,
      'Vendedor': 'Andrea Burbano',
      'Rol': 'Contador VIP',
    },
    {
      'Fecha': '2026-09-08',
      'Nombre': 'David Santiago',
      'Apellido': 'Santander Paredes',
      'Estado': 'EMITIDO',
      'Tipo': 'ERP Contable',
      'Duración': 'Un año',
      'Valor': 120.00,
      'Vendedor': 'Katherine Cabrera',
      'Rol': 'Distribuidor Connect',
    },
    {
      'Fecha': '2026-09-10',
      'Nombre': 'Gabriela Patricia',
      'Apellido': 'González Morales',
      'Estado': 'EMITIDO',
      'Tipo': 'Planes de Contador',
      'Duración': 'Un año',
      'Valor': 95.00,
      'Vendedor': 'Gabriel Endara',
      'Rol': 'Socio Franquicia',
    },
    {
      'Fecha': '2026-09-12',
      'Nombre': 'Estefanía Lucía',
      'Apellido': 'Tamay Cárdenas',
      'Estado': 'EMITIDO',
      'Tipo': 'Firma Electrónica',
      'Duración': 'Dos años',
      'Valor': 49.99,
      'Vendedor': 'Andrea Burbano',
      'Rol': 'Contador VIP',
    },
    {
      'Fecha': '2026-09-14',
      'Nombre': 'Lorena Paola',
      'Apellido': 'Vera Villamar',
      'Estado': 'EMITIDO',
      'Tipo': 'Facturación Electrónica',
      'Duración': 'Un año',
      'Valor': 35.00,
      'Vendedor': 'Andrea Perdomo',
      'Rol': 'Distribuidor Connect',
    },
    {
      'Fecha': '2026-09-15',
      'Nombre': 'Janneth Beatriz',
      'Apellido': 'Burbano Erazo',
      'Estado': 'EMITIDO',
      'Tipo': 'Firma Electrónica',
      'Duración': 'Un año',
      'Valor': 29.99,
      'Vendedor': 'Katherine Cabrera',
      'Rol': 'Distribuidor Connect',
    },
    {
      'Fecha': '2026-09-18',
      'Nombre': 'Nancy Carmita',
      'Apellido': 'Duchi Guamán',
      'Estado': 'EMITIDO',
      'Tipo': 'Planes de Contador',
      'Duración': 'Dos años',
      'Valor': 180.00,
      'Vendedor': 'Gabriel Endara',
      'Rol': 'Socio Franquicia',
    },
  ];

  const ws = XLSX.utils.json_to_sheet(sociosRows);
  XLSX.utils.book_append_sheet(wb, ws, 'Plantilla Socios 9 Columnas');
  XLSX.writeFile(wb, 'Plantilla_Carga_Socios.xlsx');
};

/**
 * Downloads a combined template Excel file ready to be filled out by the user or team
 */
export const downloadTemplateExcel = (): void => {
  const wb = XLSX.utils.book_new();

  // Template 1: Detalle de Emisiones / Firmas (Format matching user image 2)
  const templateRows = [
    {
      'N°': 1,
      'ID único': '6a9743c5d5',
      'Fecha Solic': '9/1/2026',
      'Nombres': 'Carlos Andrés',
      'Apellidos': 'Mendoza Reyes',
      'Certificado Estado': 'EMITIDO',
      'Certificado Duración': '15 días',
      'Valor': 3.16,
      'Rol': 'Distribuidor Connect',
      'Socio': 'Gabriel Endara',
      'Categoría': 'Facturación Electrónica',
    },
    {
      'N°': 2,
      'ID único': '6a97494cbe',
      'Fecha Solic': '9/1/2026',
      'Nombres': 'María Elena',
      'Apellidos': 'López Castro',
      'Certificado Estado': 'EMITIDO',
      'Certificado Duración': 'Un año',
      'Valor': 8.99,
      'Rol': 'Distribuidor Connect',
      'Socio': 'Katherine Cabrera',
      'Categoría': 'ERP Contables y Módulos',
    },
    {
      'N°': 3,
      'ID único': '6a97a9dfd5',
      'Fecha Solic': '9/1/2026',
      'Nombres': 'Juan José',
      'Apellidos': 'Alvarez Vintimilla',
      'Certificado Estado': 'EMITIDO',
      'Certificado Duración': 'Cinco años',
      'Valor': 28.17,
      'Rol': 'Distribuidor Connect',
      'Socio': 'Andrea Burbano',
      'Categoría': 'ERP Contables y Módulos',
    },
    {
      'N°': 4,
      'ID único': '6aa01b9aba',
      'Fecha Solic': '9/8/2026',
      'Nombres': 'Rosa María',
      'Apellidos': 'Ortega Ortiz',
      'Certificado Estado': 'EMITIDO',
      'Certificado Duración': 'Dos años',
      'Valor': 14.50,
      'Rol': 'Distribuidor Connect',
      'Socio': 'Katherine Cabrera',
      'Categoría': 'Planes de Contador',
    },
    {
      'N°': 5,
      'ID único': '7bb12a89c1',
      'Fecha Solic': '9/9/2026',
      'Nombres': 'Contabilidad Integral',
      'Apellidos': 'S.A.',
      'Certificado Estado': 'EMITIDO',
      'Certificado Duración': '1 Año',
      'Valor': 29.99,
      'Rol': 'UpConnect Directo',
      'Socio': 'Gabriela González',
      'Categoría': 'ERP Contables y Módulos',
    },
  ];

  const ws = XLSX.utils.json_to_sheet(templateRows);
  XLSX.utils.book_append_sheet(wb, ws, 'Plantilla Emisiones');

  // Sheet 2: Monthly format template (matching user image 3)
  const monthlyRows = [
    { 'MES': 'MARZO', 'UPCONNECT': 26.97, 'CONNECTORS': 0, 'NÚMERO UPCONNECT': 3, 'NÚMERO CONNECTORS': 3 },
    { 'MES': 'ABRIL', 'UPCONNECT': 142.91, 'CONNECTORS': 0, 'NÚMERO UPCONNECT': 10, 'NÚMERO CONNECTORS': 10 },
    { 'MES': 'MAYO', 'UPCONNECT': 228.99, 'CONNECTORS': 0, 'NÚMERO UPCONNECT': 23, 'NÚMERO CONNECTORS': 0 },
    { 'MES': 'JUNIO', 'UPCONNECT': 594.00, 'CONNECTORS': 0, 'NÚMERO UPCONNECT': 51, 'NÚMERO CONNECTORS': 0 },
    { 'MES': 'JULIO', 'UPCONNECT': 1562.57, 'CONNECTORS': 0, 'NÚMERO UPCONNECT': 115, 'NÚMERO CONNECTORS': 0 },
    { 'MES': 'AGOSTO', 'UPCONNECT': 2068.93, 'CONNECTORS': 393.61, 'NÚMERO UPCONNECT': 146, 'NÚMERO CONNECTORS': 38 },
    { 'MES': 'SEPTIEMBRE', 'UPCONNECT': 1159.90, 'CONNECTORS': 305.16, 'NÚMERO UPCONNECT': 97, 'NÚMERO CONNECTORS': 27 },
  ];
  const wsMonth = XLSX.utils.json_to_sheet(monthlyRows);
  XLSX.utils.book_append_sheet(wb, wsMonth, 'Plantilla Resumen Mensual');

  XLSX.writeFile(wb, 'Plantilla_Reporte_UpConnect.xlsx');
};
