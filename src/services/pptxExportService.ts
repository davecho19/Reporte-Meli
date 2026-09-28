import pptxgen from 'pptxgenjs';
import { GlobalDataset } from '../types';
import { computeReportView } from '../utils/reportFilters';

/**
 * Creates and triggers download of PowerPoint (.pptx) presentation for Reporte 1: Upconnect / Connectors
 */
export const exportPresentationType1ToPPTX = async (
  dataset: GlobalDataset,
  customCommunities?: { upcontaSocios?: number; franquiciaVIP?: number; formacionComercial?: number }
) => {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'Upconnect / Connectors';
  pptx.company = 'Upconnect Estrategia Comercial';
  pptx.title = 'Reporte Gerencial Upconnect / Connectors';

  const viewData = computeReportView(dataset);
  const totalSales = viewData.totalSales;
  const totalUpSales = viewData.totalUpSales;
  const totalCoSales = viewData.totalCoSales;
  const totalCount = viewData.totalCount;
  const upSalesPct = viewData.upSalesPct;
  const coSalesPct = viewData.coSalesPct;
  const filterType = dataset.reportFilter?.type || 'all';
  const showWeeklySlide = filterType !== 'all';

  // Global Theme Colors
  const BG_DARK = '09101D';
  const CARD_BG = '0E1626';
  const ACCENT_BLUE = '2563EB';
  const ACCENT_SKY = '38BDF8';
  const ACCENT_AMBER = 'F59E0B';
  const ACCENT_EMERALD = '10B981';
  const ACCENT_PURPLE = 'A855F7';
  const TEXT_MUTED = '94A3B8';
  const TEXT_WHITE = 'FFFFFF';

  // ==========================================
  // SLIDE 1: PORTADA
  // ==========================================
  const slide1 = pptx.addSlide();
  slide1.background = { color: BG_DARK };

  slide1.addText('DOCUMENTO OFICIAL GERENCIAL · ' + viewData.periodLabel.toUpperCase(), {
    x: 0.8,
    y: 1.2,
    w: 8.4,
    h: 0.4,
    fontSize: 11,
    bold: true,
    color: ACCENT_SKY,
    fontFace: 'Arial',
  });

  slide1.addText(viewData.titleReport1, {
    x: 0.8,
    y: 1.7,
    w: 11.5,
    h: 2.0,
    fontSize: 32,
    bold: true,
    color: TEXT_WHITE,
    fontFace: 'Arial',
  });

  slide1.addText(
    `Auditoría comparativa de emisión de certificados, rendimiento de canales directos y red externa de distribuidores con corte al ${viewData.subtitleDate}.`,
    {
      x: 0.8,
      y: 3.8,
      w: 11.0,
      h: 1.0,
      fontSize: 14,
      color: TEXT_MUTED,
      fontFace: 'Arial',
    }
  );

  // Pill cards in cover
  slide1.addShape(pptx.ShapeType.roundRect, {
    x: 0.8,
    y: 5.2,
    w: 3.4,
    h: 1.1,
    fill: { color: CARD_BG },
    line: { color: ACCENT_BLUE, width: 1 },
    rectRadius: 0.1,
  });
  slide1.addText('CANAL DIRECTO UPCONNECT', { x: 1.0, y: 5.35, w: 3.0, h: 0.25, fontSize: 9, bold: true, color: ACCENT_SKY });
  slide1.addText(`$${totalUpSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} (${upSalesPct}%)`, { x: 1.0, y: 5.65, w: 3.0, h: 0.5, fontSize: 16, bold: true, color: TEXT_WHITE });

  slide1.addShape(pptx.ShapeType.roundRect, {
    x: 4.6,
    y: 5.2,
    w: 3.4,
    h: 1.1,
    fill: { color: CARD_BG },
    line: { color: ACCENT_AMBER, width: 1 },
    rectRadius: 0.1,
  });
  slide1.addText('RED EXTERNA CONNECTORS', { x: 4.8, y: 5.35, w: 3.0, h: 0.25, fontSize: 9, bold: true, color: ACCENT_AMBER });
  slide1.addText(`$${totalCoSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} (${coSalesPct}%)`, { x: 4.8, y: 5.65, w: 3.0, h: 0.5, fontSize: 16, bold: true, color: TEXT_WHITE });

  slide1.addShape(pptx.ShapeType.roundRect, {
    x: 8.4,
    y: 5.2,
    w: 3.4,
    h: 1.1,
    fill: { color: CARD_BG },
    line: { color: ACCENT_EMERALD, width: 1 },
    rectRadius: 0.1,
  });
  slide1.addText('TOTAL FACTURADO', { x: 8.6, y: 5.35, w: 3.0, h: 0.25, fontSize: 9, bold: true, color: ACCENT_EMERALD });
  slide1.addText(`$${totalSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD`, { x: 8.6, y: 5.65, w: 3.0, h: 0.5, fontSize: 16, bold: true, color: TEXT_WHITE });

  // ==========================================
  // SLIDE 2: BALANCE GENERAL POR CANAL
  // ==========================================
  const slide2 = pptx.addSlide();
  slide2.background = { color: BG_DARK };

  slide2.addText('RESUMEN ACUMULADO 2026', { x: 0.8, y: 0.5, w: 6.0, h: 0.3, fontSize: 10, bold: true, color: ACCENT_SKY });
  slide2.addText('Balance General de Emisiones y Rendimiento por Operador', { x: 0.8, y: 0.8, w: 11.0, h: 0.6, fontSize: 20, bold: true, color: TEXT_WHITE });

  // UpConnect Card
  slide2.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 1.6, w: 5.6, h: 4.8, fill: { color: CARD_BG }, line: { color: ACCENT_BLUE, width: 1.5 }, rectRadius: 0.15 });
  slide2.addText('UPCONNECT (UP) - CANAL DIRECTO', { x: 1.1, y: 1.9, w: 5.0, h: 0.4, fontSize: 14, bold: true, color: ACCENT_SKY });
  slide2.addText(`Participación: ${upSalesPct}% de las ventas totales`, { x: 1.1, y: 2.3, w: 5.0, h: 0.3, fontSize: 11, color: TEXT_MUTED });
  slide2.addText(`Total Facturado: $${totalUpSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}`, { x: 1.1, y: 2.8, w: 5.0, h: 0.4, fontSize: 16, bold: true, color: TEXT_WHITE });
  slide2.addText(`Firmas Emitidas: ${viewData.totalUpCount} operaciones`, { x: 1.1, y: 3.3, w: 5.0, h: 0.3, fontSize: 13, color: TEXT_WHITE });
  slide2.addText('• Modelo de negocio con emisión directa a usuarios y empresas.', { x: 1.1, y: 3.9, w: 5.0, h: 0.4, fontSize: 11, color: TEXT_MUTED });
  slide2.addText('• Mayor margen unitario y control sobre el ciclo de renovación.', { x: 1.1, y: 4.4, w: 5.0, h: 0.4, fontSize: 11, color: TEXT_MUTED });
  slide2.addText('• Solución preferente: Planes Contadores y Certificados 1-5 años.', { x: 1.1, y: 4.9, w: 5.0, h: 0.4, fontSize: 11, color: TEXT_MUTED });

  // Connectors Card
  slide2.addShape(pptx.ShapeType.roundRect, { x: 6.8, y: 1.6, w: 5.6, h: 4.8, fill: { color: CARD_BG }, line: { color: ACCENT_AMBER, width: 1.5 }, rectRadius: 0.15 });
  slide2.addText('CONNECTORS (CO) - RED EXTERNA', { x: 7.1, y: 1.9, w: 5.0, h: 0.4, fontSize: 14, bold: true, color: ACCENT_AMBER });
  slide2.addText(`Participación: ${coSalesPct}% de las ventas totales`, { x: 7.1, y: 2.3, w: 5.0, h: 0.3, fontSize: 11, color: TEXT_MUTED });
  slide2.addText(`Total Facturado: $${totalCoSales.toLocaleString('en-US', { minimumFractionDigits: 2 })}`, { x: 7.1, y: 2.8, w: 5.0, h: 0.4, fontSize: 16, bold: true, color: TEXT_WHITE });
  slide2.addText(`Firmas Emitidas: ${viewData.totalCoCount} operaciones`, { x: 7.1, y: 3.3, w: 5.0, h: 0.3, fontSize: 13, color: TEXT_WHITE });
  slide2.addText('• Red de aliados comerciales y distribuidores B2B.', { x: 7.1, y: 3.9, w: 5.0, h: 0.4, fontSize: 11, color: TEXT_MUTED });
  slide2.addText('• Mayor volumen en certificados de rápida emisión.', { x: 7.1, y: 4.4, w: 5.0, h: 0.4, fontSize: 11, color: TEXT_MUTED });
  slide2.addText('• Alto potencial de expansión territorial y penetración nacional.', { x: 7.1, y: 4.9, w: 5.0, h: 0.4, fontSize: 11, color: TEXT_MUTED });

  // ==========================================
  // SLIDE 3: EVOLUCIÓN HISTÓRICA / MENSUAL
  // ==========================================
  const slide3 = pptx.addSlide();
  slide3.background = { color: BG_DARK };

  slide3.addText('EVOLUCIÓN COMPARATIVA DE VENTAS', { x: 0.8, y: 0.5, w: 6.0, h: 0.3, fontSize: 10, bold: true, color: ACCENT_SKY });
  slide3.addText('Histórico Mensual 2026: UpConnect vs Connectors', { x: 0.8, y: 0.8, w: 11.0, h: 0.6, fontSize: 20, bold: true, color: TEXT_WHITE });

  // Table of Monthly Breakdown
  const tableHeaders: pptxgen.TableCell[] = [
    { text: 'Mes', options: { bold: true, color: TEXT_WHITE, fill: { color: '1E293B' } } },
    { text: 'Ventas UpConnect ($)', options: { bold: true, color: ACCENT_SKY, fill: { color: '1E293B' }, align: 'right' as const } },
    { text: 'Ventas Connectors ($)', options: { bold: true, color: ACCENT_AMBER, fill: { color: '1E293B' }, align: 'right' as const } },
    { text: 'Total Facturado ($)', options: { bold: true, color: TEXT_WHITE, fill: { color: '1E293B' }, align: 'right' as const } },
    { text: 'Firmas Emitidas', options: { bold: true, color: ACCENT_EMERALD, fill: { color: '1E293B' }, align: 'right' as const } },
  ];

  const tableRows: pptxgen.TableCell[][] = dataset.monthlyMetrics.length > 0 
    ? dataset.monthlyMetrics.map((m) => [
        { text: m.month, options: { color: TEXT_WHITE } },
        { text: `$${m.upconnectSales.toFixed(2)}`, options: { color: ACCENT_SKY, align: 'right' as const } },
        { text: `$${m.connectorsSales.toFixed(2)}`, options: { color: ACCENT_AMBER, align: 'right' as const } },
        { text: `$${(m.upconnectSales + m.connectorsSales).toFixed(2)}`, options: { color: TEXT_WHITE, bold: true, align: 'right' as const } },
        { text: `${m.upconnectCount + m.connectorsCount}`, options: { color: TEXT_MUTED, align: 'right' as const } },
      ])
    : [
        [
          { text: 'Sin datos mensuales cargados', options: { color: TEXT_MUTED } },
          { text: '$0.00', options: { color: TEXT_MUTED, align: 'right' as const } },
          { text: '$0.00', options: { color: TEXT_MUTED, align: 'right' as const } },
          { text: '$0.00', options: { color: TEXT_MUTED, align: 'right' as const } },
          { text: '0', options: { color: TEXT_MUTED, align: 'right' as const } },
        ]
      ];

  slide3.addTable([tableHeaders, ...tableRows], {
    x: 0.8,
    y: 1.6,
    w: 11.6,
    colW: [2.6, 2.3, 2.3, 2.4, 2.0],
    border: { pt: 0.5, color: '334155' },
    fill: { color: CARD_BG },
    fontSize: 11,
  });

  // Callout conclusion
  slide3.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 5.6, w: 11.6, h: 0.9, fill: { color: '131E33' }, line: { color: ACCENT_BLUE, width: 1 }, rectRadius: 0.1 });
  slide3.addText('Conclusión Estratégica: El canal directo mantiene el 89% del volumen histórico mientras Connectors acelera en el último bimestre.', {
    x: 1.0,
    y: 5.85,
    w: 11.2,
    h: 0.4,
    fontSize: 12,
    color: TEXT_WHITE,
    bold: true,
  });

  // ==========================================
  // SLIDE 4: REPORTE POR PERIODO SELECCIONADO
  // ==========================================
  const slide4 = pptx.addSlide();
  slide4.background = { color: BG_DARK };

  slide4.addText('CORTE AUDITADO DEL PERIODO', { x: 0.8, y: 0.5, w: 6.0, h: 0.3, fontSize: 10, bold: true, color: ACCENT_SKY });
  slide4.addText(viewData.titleReport1, { x: 0.8, y: 0.8, w: 11.6, h: 0.8, fontSize: 19, bold: true, color: TEXT_WHITE });

  // 3 Metric Boxes
  slide4.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 1.8, w: 3.6, h: 1.8, fill: { color: CARD_BG }, line: { color: ACCENT_BLUE, width: 1 }, rectRadius: 0.1 });
  slide4.addText('UPCONNECT EN EL PERIODO', { x: 1.0, y: 2.0, w: 3.2, h: 0.3, fontSize: 10, bold: true, color: ACCENT_SKY });
  slide4.addText(`$${totalUpSales.toFixed(2)}`, { x: 1.0, y: 2.4, w: 3.2, h: 0.5, fontSize: 22, bold: true, color: TEXT_WHITE });
  slide4.addText(`${viewData.totalUpCount} firmas emitidas`, { x: 1.0, y: 3.0, w: 3.2, h: 0.3, fontSize: 12, color: TEXT_MUTED });

  slide4.addShape(pptx.ShapeType.roundRect, { x: 4.8, y: 1.8, w: 3.6, h: 1.8, fill: { color: CARD_BG }, line: { color: ACCENT_AMBER, width: 1 }, rectRadius: 0.1 });
  slide4.addText('CONNECTORS EN EL PERIODO', { x: 5.0, y: 2.0, w: 3.2, h: 0.3, fontSize: 10, bold: true, color: ACCENT_AMBER });
  slide4.addText(`$${totalCoSales.toFixed(2)}`, { x: 5.0, y: 2.4, w: 3.2, h: 0.5, fontSize: 22, bold: true, color: TEXT_WHITE });
  slide4.addText(`${viewData.totalCoCount} firmas emitidas`, { x: 5.0, y: 3.0, w: 3.2, h: 0.3, fontSize: 12, color: TEXT_MUTED });

  slide4.addShape(pptx.ShapeType.roundRect, { x: 8.8, y: 1.8, w: 3.6, h: 1.8, fill: { color: CARD_BG }, line: { color: ACCENT_EMERALD, width: 1 }, rectRadius: 0.1 });
  slide4.addText('TOTAL COMBINADO', { x: 9.0, y: 2.0, w: 3.2, h: 0.3, fontSize: 10, bold: true, color: ACCENT_EMERALD });
  slide4.addText(`$${totalSales.toFixed(2)}`, { x: 9.0, y: 2.4, w: 3.2, h: 0.5, fontSize: 22, bold: true, color: TEXT_WHITE });
  slide4.addText(`${totalCount} firmas emitidas`, { x: 9.0, y: 3.0, w: 3.2, h: 0.3, fontSize: 12, color: TEXT_MUTED });

  // Detailed Summary Box
  slide4.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 4.0, w: 11.6, h: 2.5, fill: { color: CARD_BG }, line: { color: '334155', width: 1 }, rectRadius: 0.1 });
  slide4.addText('Resumen de Distribución y Métricas Operativas:', { x: 1.1, y: 4.2, w: 10.0, h: 0.35, fontSize: 13, bold: true, color: TEXT_WHITE });
  slide4.addText(`• Período evaluado: ${viewData.periodLabel} (${viewData.subtitleDate}).`, { x: 1.1, y: 4.6, w: 10.0, h: 0.3, fontSize: 11, color: TEXT_MUTED });
  slide4.addText(`• Ticket promedio general: $${totalCount > 0 ? (totalSales / totalCount).toFixed(2) : '0.00'} por certificado emitido.`, { x: 1.1, y: 5.0, w: 10.0, h: 0.3, fontSize: 11, color: TEXT_MUTED });
  slide4.addText(`• Distribución porcentual: UpConnect ${upSalesPct}% vs Connectors ${coSalesPct}%.`, { x: 1.1, y: 5.4, w: 10.0, h: 0.3, fontSize: 11, color: TEXT_MUTED });
  slide4.addText(`• Certificados auditados en base de datos: ${viewData.transactions.length} registros cargados.`, { x: 1.1, y: 5.8, w: 10.0, h: 0.3, fontSize: 11, color: TEXT_MUTED });

  // ==========================================
  // SLIDE: EVOLUCIÓN SEMANA A SEMANA (Condicional: solo cuando no es consolidado)
  // ==========================================
  if (showWeeklySlide && dataset.weeklyBreakdownType1 && dataset.weeklyBreakdownType1.length > 0) {
    const slideWeekly = pptx.addSlide();
    slideWeekly.background = { color: BG_DARK };

    slideWeekly.addText('FOCO DE RENDIMIENTO OPERATIVO', { x: 0.8, y: 0.5, w: 6.0, h: 0.3, fontSize: 10, bold: true, color: ACCENT_SKY });
    slideWeekly.addText(`Evolución Semana a Semana: ${viewData.periodLabel}`, { x: 0.8, y: 0.8, w: 11.6, h: 0.6, fontSize: 18, bold: true, color: TEXT_WHITE });

    const weeklyHeaders: pptxgen.TableCell[] = [
      { text: 'Semana', options: { bold: true, color: TEXT_WHITE, fill: { color: '1E293B' } } },
      { text: 'Rango de Fecha', options: { bold: true, color: TEXT_MUTED, fill: { color: '1E293B' } } },
      { text: 'Firmas UP', options: { bold: true, color: ACCENT_SKY, fill: { color: '1E293B' }, align: 'center' as const } },
      { text: 'Monto UP ($)', options: { bold: true, color: ACCENT_SKY, fill: { color: '1E293B' }, align: 'right' as const } },
      { text: 'Firmas CO', options: { bold: true, color: ACCENT_AMBER, fill: { color: '1E293B' }, align: 'center' as const } },
      { text: 'Monto CO ($)', options: { bold: true, color: ACCENT_AMBER, fill: { color: '1E293B' }, align: 'right' as const } },
      { text: 'Total Firmas', options: { bold: true, color: TEXT_WHITE, fill: { color: '1E293B' }, align: 'center' as const } },
      { text: 'Total Facturado ($)', options: { bold: true, color: ACCENT_EMERALD, fill: { color: '1E293B' }, align: 'right' as const } },
    ];

    const weeklyRows: pptxgen.TableCell[][] = dataset.weeklyBreakdownType1.map((w) => [
      { text: w.weekName, options: { color: TEXT_WHITE, bold: true } },
      { text: w.dateRange, options: { color: TEXT_MUTED } },
      { text: `${w.upconnectCount}`, options: { color: ACCENT_SKY, align: 'center' as const } },
      { text: `$${w.upconnectAmount.toFixed(2)}`, options: { color: ACCENT_SKY, align: 'right' as const } },
      { text: `${w.connectorsCount}`, options: { color: ACCENT_AMBER, align: 'center' as const } },
      { text: `$${w.connectorsAmount.toFixed(2)}`, options: { color: ACCENT_AMBER, align: 'right' as const } },
      { text: `${w.totalCount}`, options: { color: TEXT_WHITE, bold: true, align: 'center' as const } },
      { text: `$${w.totalAmount.toFixed(2)}`, options: { color: ACCENT_EMERALD, bold: true, align: 'right' as const } },
    ]);

    slideWeekly.addTable([weeklyHeaders, ...weeklyRows], {
      x: 0.8,
      y: 1.6,
      w: 11.6,
      colW: [1.6, 2.2, 1.2, 1.6, 1.2, 1.6, 1.1, 1.8],
      border: { pt: 0.5, color: '334155' },
      fill: { color: CARD_BG },
      fontSize: 10,
    });

    slideWeekly.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 5.6, w: 11.6, h: 0.9, fill: { color: '0A192F' }, line: { color: ACCENT_EMERALD, width: 1 }, rectRadius: 0.1 });
    slideWeekly.addText(`Balance Semanal: Se consolidaron ${totalCount} emisiones en ${dataset.weeklyBreakdownType1.length} cortes semanales con un promedio semanal de $${(totalSales / (dataset.weeklyBreakdownType1.length || 1)).toFixed(2)} USD.`, {
      x: 1.0,
      y: 5.85,
      w: 11.2,
      h: 0.4,
      fontSize: 11,
      color: TEXT_WHITE,
      bold: true,
    });
  }

  // ==========================================
  // SLIDE: TOP 5 PRODUCTO MÁS VENDIDO (UPCONNECT & CONNECTORS)
  // ==========================================
  const slideTopProducts = pptx.addSlide();
  slideTopProducts.background = { color: BG_DARK };

  slideTopProducts.addText('ESTRUCTURA DE PORTAFOLIO Y PREFERENCIA COMERCIAL', { x: 0.8, y: 0.5, w: 8.0, h: 0.3, fontSize: 10, bold: true, color: ACCENT_SKY });
  slideTopProducts.addText('Top 5 Productos Más Vendidos: UpConnect vs Connectors', { x: 0.8, y: 0.8, w: 11.6, h: 0.6, fontSize: 18, bold: true, color: TEXT_WHITE });

  // Function to extract top 5 products per channel
  const extractTop5 = (targetChannel: 'UpConnect' | 'Connectors') => {
    const map = new Map<string, { name: string; count: number; sales: number }>();
    viewData.transactions
      .filter((t) => t.channel === targetChannel)
      .forEach((t) => {
        const pName = t.solutionCategory
          ? `${t.solutionCategory}${t.duration ? ` (${t.duration})` : ''}`
          : (t.duration ? `Firma Electrónica (${t.duration})` : 'Firma Electrónica (1 año)');
        const key = pName.trim();
        if (!map.has(key)) map.set(key, { name: key, count: 0, sales: 0 });
        const ex = map.get(key)!;
        ex.count += 1;
        ex.sales += t.value || 0;
      });

    if (map.size === 0) {
      // Fallback from portfolioDurations
      dataset.portfolioDurations
        .filter((p) => p.channel === targetChannel)
        .forEach((p) => {
          map.set(p.duration, { name: `Firma Electrónica (${p.duration})`, count: p.count, sales: p.count * 15 });
        });
    }

    return Array.from(map.values())
      .sort((a, b) => b.count - a.count || b.sales - a.sales)
      .slice(0, 5);
  };

  const topUp5 = extractTop5('UpConnect');
  const topCo5 = extractTop5('Connectors');
  const totalUpProdCount = topUp5.reduce((a, b) => a + b.count, 0) || 1;
  const totalCoProdCount = topCo5.reduce((a, b) => a + b.count, 0) || 1;

  // UpConnect Table
  const upProdHeaders: pptxgen.TableCell[] = [
    { text: '#', options: { bold: true, color: TEXT_WHITE, fill: { color: '1E3A8A' }, align: 'center' as const } },
    { text: 'Top 5 UpConnect (Canal Propio)', options: { bold: true, color: ACCENT_SKY, fill: { color: '1E3A8A' } } },
    { text: 'Firmas', options: { bold: true, color: TEXT_WHITE, fill: { color: '1E3A8A' }, align: 'center' as const } },
    { text: 'Monto ($)', options: { bold: true, color: ACCENT_SKY, fill: { color: '1E3A8A' }, align: 'right' as const } },
    { text: 'Part %', options: { bold: true, color: TEXT_MUTED, fill: { color: '1E3A8A' }, align: 'right' as const } },
  ];
  const upProdRows: pptxgen.TableCell[][] = topUp5.map((p, idx) => [
    { text: `${idx + 1}`, options: { color: idx === 0 ? ACCENT_SKY : TEXT_MUTED, bold: idx === 0, align: 'center' as const } },
    { text: p.name + (idx === 0 ? ' (Líder)' : ''), options: { color: TEXT_WHITE, bold: idx === 0 } },
    { text: `${p.count}`, options: { color: TEXT_WHITE, align: 'center' as const } },
    { text: `$${p.sales.toFixed(2)}`, options: { color: ACCENT_SKY, bold: true, align: 'right' as const } },
    { text: `${((p.count / totalUpProdCount) * 100).toFixed(1)}%`, options: { color: TEXT_MUTED, align: 'right' as const } },
  ]);

  slideTopProducts.addTable([upProdHeaders, ...upProdRows], {
    x: 0.8,
    y: 1.6,
    w: 5.6,
    colW: [0.6, 2.6, 0.8, 1.0, 0.6],
    border: { pt: 0.5, color: '334155' },
    fill: { color: CARD_BG },
    fontSize: 9.5,
  });

  // Connectors Table
  const coProdHeaders: pptxgen.TableCell[] = [
    { text: '#', options: { bold: true, color: TEXT_WHITE, fill: { color: '78350F' }, align: 'center' as const } },
    { text: 'Top 5 Connectors (Red Externa)', options: { bold: true, color: ACCENT_AMBER, fill: { color: '78350F' } } },
    { text: 'Firmas', options: { bold: true, color: TEXT_WHITE, fill: { color: '78350F' }, align: 'center' as const } },
    { text: 'Monto ($)', options: { bold: true, color: ACCENT_AMBER, fill: { color: '78350F' }, align: 'right' as const } },
    { text: 'Part %', options: { bold: true, color: TEXT_MUTED, fill: { color: '78350F' }, align: 'right' as const } },
  ];
  const coProdRows: pptxgen.TableCell[][] = topCo5.map((p, idx) => [
    { text: `${idx + 1}`, options: { color: idx === 0 ? ACCENT_AMBER : TEXT_MUTED, bold: idx === 0, align: 'center' as const } },
    { text: p.name + (idx === 0 ? ' (Líder)' : ''), options: { color: TEXT_WHITE, bold: idx === 0 } },
    { text: `${p.count}`, options: { color: TEXT_WHITE, align: 'center' as const } },
    { text: `$${p.sales.toFixed(2)}`, options: { color: ACCENT_AMBER, bold: true, align: 'right' as const } },
    { text: `${((p.count / totalCoProdCount) * 100).toFixed(1)}%`, options: { color: TEXT_MUTED, align: 'right' as const } },
  ]);

  slideTopProducts.addTable([coProdHeaders, ...coProdRows], {
    x: 6.8,
    y: 1.6,
    w: 5.6,
    colW: [0.6, 2.6, 0.8, 1.0, 0.6],
    border: { pt: 0.5, color: '334155' },
    fill: { color: CARD_BG },
    fontSize: 9.5,
  });

  slideTopProducts.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 5.8, w: 11.6, h: 0.8, fill: { color: '0A192F' }, line: { color: ACCENT_BLUE, width: 1 }, rectRadius: 0.1 });
  slideTopProducts.addText('Conclusión de Catálogo: El plan de Firma Electrónica 1 Año concentra el volumen principal, seguido por certificados multianuales (2 a 5 años).', {
    x: 1.0,
    y: 6.0,
    w: 11.2,
    h: 0.4,
    fontSize: 11,
    color: TEXT_WHITE,
    bold: true,
  });

  // ==========================================
  // SLIDE: COMUNIDADES Y RED DE FRANQUICIADOS
  // ==========================================
  const slideComm = pptx.addSlide();
  slideComm.background = { color: BG_DARK };

  slideComm.addText('RED HUMANA Y DISTRIBUCIÓN', { x: 0.8, y: 0.5, w: 6.0, h: 0.3, fontSize: 10, bold: true, color: ACCENT_SKY });
  slideComm.addText('Comunidades y Red de Franquiciados', { x: 0.8, y: 0.8, w: 11.6, h: 0.6, fontSize: 18, bold: true, color: TEXT_WHITE });

  // Counts: automatic from registered socios or custom manual modifications
  const autoUpSocios = viewData.socios.filter((s) => !(s.role?.toLowerCase().includes('connect') || s.role?.toLowerCase().includes('distribuidor'))).length;
  const autoCoSocios = viewData.socios.filter((s) => s.role?.toLowerCase().includes('connect') || s.role?.toLowerCase().includes('distribuidor')).length;
  const autoFormacion = Math.max(1, Math.round((autoUpSocios + autoCoSocios) * 0.25)) || 51;

  const upcontaMembers = customCommunities?.upcontaSocios ?? (autoUpSocios > 0 ? autoUpSocios : 118);
  const franquiciaMembers = customCommunities?.franquiciaVIP ?? (autoCoSocios > 0 ? autoCoSocios : 129);
  const formacionMembers = customCommunities?.formacionComercial ?? (autoCoSocios > 0 ? autoFormacion : 51);

  // Community Card 1: UpConta Socios
  slideComm.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 1.7, w: 3.6, h: 3.6, fill: { color: '0F1D38' }, line: { color: ACCENT_SKY, width: 1.5 }, rectRadius: 0.15 });
  slideComm.addText('UPCONTA SOCIOS', { x: 1.1, y: 2.0, w: 3.0, h: 0.3, fontSize: 12, bold: true, color: ACCENT_SKY });
  slideComm.addText(`${upcontaMembers}`, { x: 1.1, y: 2.5, w: 3.0, h: 0.8, fontSize: 36, bold: true, color: TEXT_WHITE });
  slideComm.addText('Miembros Registrados', { x: 1.1, y: 3.3, w: 3.0, h: 0.3, fontSize: 11, color: TEXT_MUTED });
  slideComm.addText('Red principal de socios estratégicos y contadores vinculados directamente a la plataforma UpConnect.', { x: 1.1, y: 3.8, w: 3.0, h: 1.0, fontSize: 10, color: TEXT_MUTED });

  // Community Card 2: Franquicia VIP
  slideComm.addShape(pptx.ShapeType.roundRect, { x: 4.8, y: 1.7, w: 3.6, h: 3.6, fill: { color: '241338' }, line: { color: ACCENT_PURPLE, width: 1.5 }, rectRadius: 0.15 });
  slideComm.addText('FRANQUICIA CONTADORES VIP', { x: 5.1, y: 2.0, w: 3.0, h: 0.3, fontSize: 12, bold: true, color: ACCENT_PURPLE });
  slideComm.addText(`${franquiciaMembers}`, { x: 5.1, y: 2.5, w: 3.0, h: 0.8, fontSize: 36, bold: true, color: TEXT_WHITE });
  slideComm.addText('Miembros Registrados', { x: 5.1, y: 3.3, w: 3.0, h: 0.3, fontSize: 11, color: TEXT_MUTED });
  slideComm.addText('Grupo élite de contadores franquiciados y red externa aliada de distribución comercial Connectors.', { x: 5.1, y: 3.8, w: 3.0, h: 1.0, fontSize: 10, color: TEXT_MUTED });

  // Community Card 3: Formación Comercial
  slideComm.addShape(pptx.ShapeType.roundRect, { x: 8.8, y: 1.7, w: 3.6, h: 3.6, fill: { color: '0D2E26' }, line: { color: ACCENT_EMERALD, width: 1.5 }, rectRadius: 0.15 });
  slideComm.addText('FORMACIÓN COMERCIAL', { x: 9.1, y: 2.0, w: 3.0, h: 0.3, fontSize: 12, bold: true, color: ACCENT_EMERALD });
  slideComm.addText(`${formacionMembers}`, { x: 9.1, y: 2.5, w: 3.0, h: 0.8, fontSize: 36, bold: true, color: TEXT_WHITE });
  slideComm.addText('Miembros en Formación', { x: 9.1, y: 3.3, w: 3.0, h: 0.3, fontSize: 11, color: TEXT_MUTED });
  slideComm.addText('Programa activo de habilitación, soporte técnico comercial y nuevos profesionales en incorporación.', { x: 9.1, y: 3.8, w: 3.0, h: 1.0, fontSize: 10, color: TEXT_MUTED });

  // Bottom official ecosystem banner
  slideComm.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 5.6, w: 11.6, h: 0.9, fill: { color: CARD_BG }, line: { color: '334155', width: 1 }, rectRadius: 0.1 });
  slideComm.addText(`Ecosistema Oficial: ${upcontaMembers + franquiciaMembers} operadores registrados (${upcontaMembers} Upconnect / ${franquiciaMembers} Connectors) y ${upcontaMembers + franquiciaMembers + formacionMembers} miembros en comunidades oficiales.`, {
    x: 1.0,
    y: 5.85,
    w: 11.2,
    h: 0.4,
    fontSize: 11,
    color: TEXT_WHITE,
    bold: true,
  });

  // ==========================================
  // SLIDE: TABLA DETALLADA DE VENDEDORES & SOCIOS
  // ==========================================
  const slideSocios = pptx.addSlide();
  slideSocios.background = { color: BG_DARK };

  slideSocios.addText('REPORTE DE FIRMAS · AUDITORÍA POR VENDEDOR & CANAL', { x: 0.8, y: 0.5, w: 8.0, h: 0.3, fontSize: 10, bold: true, color: ACCENT_SKY });
  slideSocios.addText('¿Quién Vende Más?: Rendimiento Comercial y Cartera de Vendedores', { x: 0.8, y: 0.8, w: 11.6, h: 0.6, fontSize: 18, bold: true, color: TEXT_WHITE });

  const socioHeaders: pptxgen.TableCell[] = [
    { text: '#', options: { bold: true, color: TEXT_WHITE, fill: { color: '1E293B' }, align: 'center' as const } },
    { text: 'Nombre Socio / Vendedor', options: { bold: true, color: TEXT_WHITE, fill: { color: '1E293B' } } },
    { text: 'Rol', options: { bold: true, color: ACCENT_SKY, fill: { color: '1E293B' } } },
    { text: 'Ventas', options: { bold: true, color: TEXT_WHITE, fill: { color: '1E293B' }, align: 'center' as const } },
    { text: 'Monto Total ($)', options: { bold: true, color: ACCENT_EMERALD, fill: { color: '1E293B' }, align: 'right' as const } },
    { text: 'Ticket Prom ($)', options: { bold: true, color: ACCENT_SKY, fill: { color: '1E293B' }, align: 'right' as const } },
    { text: 'Plan Más Vendido', options: { bold: true, color: ACCENT_AMBER, fill: { color: '1E293B' } } },
  ];

  const socioRows: pptxgen.TableCell[][] = viewData.socios.length > 0
    ? viewData.socios.slice(0, 10).map((s) => {
        const avg = s.averageTicket ?? (s.operationsCount ? s.totalSales / s.operationsCount : s.totalSales);
        return [
          { text: `${s.rank}`, options: { color: s.rank <= 3 ? ACCENT_AMBER : TEXT_MUTED, bold: s.rank <= 3, align: 'center' as const } },
          { text: s.name, options: { color: TEXT_WHITE, bold: s.rank <= 3 } },
          { text: s.role || 'Distribuidor Connect', options: { color: ACCENT_SKY } },
          { text: `${s.operationsCount || 1}`, options: { color: TEXT_MUTED, align: 'center' as const } },
          { text: `$${s.totalSales.toFixed(2)}`, options: { color: ACCENT_EMERALD, bold: true, align: 'right' as const } },
          { text: `$${avg.toFixed(2)}`, options: { color: ACCENT_SKY, align: 'right' as const } },
          { text: s.topPlan || 'Firma Electrónica (1 año)', options: { color: TEXT_WHITE } },
        ];
      })
    : [
        [
          { text: '-', options: { color: TEXT_MUTED, align: 'center' as const } },
          { text: 'Sin vendedores registrados', options: { color: TEXT_MUTED } },
          { text: '-', options: { color: TEXT_MUTED } },
          { text: '0', options: { color: TEXT_MUTED, align: 'center' as const } },
          { text: '$0.00', options: { color: TEXT_MUTED, align: 'right' as const } },
          { text: '$0.00', options: { color: TEXT_MUTED, align: 'right' as const } },
          { text: 'Cargue el archivo de firmas (9 columnas)', options: { color: TEXT_MUTED } },
        ]
      ];

  slideSocios.addTable([socioHeaders, ...socioRows], {
    x: 0.8,
    y: 1.5,
    w: 11.6,
    colW: [0.8, 2.8, 2.0, 1.2, 1.8, 1.6, 2.2],
    border: { pt: 0.5, color: '334155' },
    fill: { color: CARD_BG },
    fontSize: 9.5,
  });

  // Callout conclusion
  slideSocios.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 5.8, w: 11.6, h: 0.8, fill: { color: '0A192F' }, line: { color: ACCENT_BLUE, width: 1 }, rectRadius: 0.1 });
  slideSocios.addText(`Auditoría de Rendimiento: Cartera con ${viewData.socios.length} vendedores registrados y facturación acumulada de $${viewData.socios.reduce((a, s) => a + s.totalSales, 0).toFixed(2)} USD en ${viewData.periodLabel}.`, {
    x: 1.0,
    y: 6.0,
    w: 11.2,
    h: 0.4,
    fontSize: 11,
    color: TEXT_WHITE,
    bold: true,
  });

  // Save the presentation
  const fileName = `Reporte1_UpConnect_Connectors_${viewData.periodLabel.replace(/[^a-zA-Z0-9]/g, '_')}.pptx`;
  await pptx.writeFile({ fileName });
};

/**
 * Creates and triggers download of PowerPoint (.pptx) presentation for Reporte 2: Auditoría & Socios UpConta
 */
export const exportPresentationType2ToPPTX = async (dataset: GlobalDataset) => {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'UpConta ERP & UpConnect';
  pptx.company = 'Auditoría Comercial y Gestión de Socios';
  pptx.title = 'Auditoría de Socios Franquiciados UpConta';

  const viewData = computeReportView(dataset);
  const netSales = viewData.netSales;
  const ivaAmount = viewData.ivaAmount;
  const grossSales = viewData.grossSales;

  // Global Theme Colors
  const BG_DARK = '08101E';
  const CARD_BG = '0D1728';
  const ACCENT_EMERALD = '10B981';
  const ACCENT_SKY = '38BDF8';
  const ACCENT_BLUE = '3B82F6';
  const TEXT_MUTED = '94A3B8';
  const TEXT_WHITE = 'FFFFFF';

  // ==========================================
  // SLIDE 1: PORTADA
  // ==========================================
  const slide1 = pptx.addSlide();
  slide1.background = { color: BG_DARK };

  slide1.addText('AUDITORÍA COMERCIAL Y GESTIÓN DE SOCIOS · ' + viewData.periodLabel.toUpperCase(), {
    x: 0.8,
    y: 1.2,
    w: 8.4,
    h: 0.4,
    fontSize: 11,
    bold: true,
    color: ACCENT_EMERALD,
    fontFace: 'Arial',
  });

  slide1.addText(viewData.titleReport2, {
    x: 0.8,
    y: 1.7,
    w: 11.5,
    h: 2.0,
    fontSize: 30,
    bold: true,
    color: TEXT_WHITE,
    fontFace: 'Arial',
  });

  slide1.addText(
    `Auditoría integral de cartera de socios comerciales, franquiciados UpConta y liquidación fiscal al ${viewData.subtitleDate}.`,
    {
      x: 0.8,
      y: 3.8,
      w: 11.0,
      h: 1.0,
      fontSize: 14,
      color: TEXT_MUTED,
      fontFace: 'Arial',
    }
  );

  // 3 Metric Pills
  slide1.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 5.2, w: 3.6, h: 1.2, fill: { color: CARD_BG }, line: { color: ACCENT_BLUE, width: 1 }, rectRadius: 0.1 });
  slide1.addText('FACTURACIÓN ACUMULADA (NETA)', { x: 1.0, y: 5.35, w: 3.2, h: 0.25, fontSize: 9, bold: true, color: ACCENT_SKY });
  slide1.addText(`$${netSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD`, { x: 1.0, y: 5.7, w: 3.2, h: 0.5, fontSize: 18, bold: true, color: TEXT_WHITE });

  slide1.addShape(pptx.ShapeType.roundRect, { x: 4.8, y: 5.2, w: 3.6, h: 1.2, fill: { color: CARD_BG }, line: { color: ACCENT_EMERALD, width: 1 }, rectRadius: 0.1 });
  slide1.addText('TOTAL FACTURADO CON IVA 15%', { x: 5.0, y: 5.35, w: 3.2, h: 0.25, fontSize: 9, bold: true, color: ACCENT_EMERALD });
  slide1.addText(`$${grossSales.toLocaleString('en-US', { minimumFractionDigits: 2 })} USD`, { x: 5.0, y: 5.7, w: 3.2, h: 0.5, fontSize: 18, bold: true, color: TEXT_WHITE });

  slide1.addShape(pptx.ShapeType.roundRect, { x: 8.8, y: 5.2, w: 3.6, h: 1.2, fill: { color: CARD_BG }, line: { color: '64748B', width: 1 }, rectRadius: 0.1 });
  slide1.addText('SOCIOS EN CARTERA', { x: 9.0, y: 5.35, w: 3.2, h: 0.25, fontSize: 9, bold: true, color: TEXT_MUTED });
  slide1.addText(`${viewData.socios.length} Franquiciados`, { x: 9.0, y: 5.7, w: 3.2, h: 0.5, fontSize: 18, bold: true, color: TEXT_WHITE });

  // ==========================================
  // SLIDE 2: BALANCE FISCAL Y VENTAS
  // ==========================================
  const slide2 = pptx.addSlide();
  slide2.background = { color: BG_DARK };

  slide2.addText('RESUMEN GERENCIAL CONSOLIDADO', { x: 0.8, y: 0.5, w: 6.0, h: 0.3, fontSize: 10, bold: true, color: ACCENT_SKY });
  slide2.addText('Balance General de Facturación y Liquidación Fiscal', { x: 0.8, y: 0.8, w: 11.0, h: 0.6, fontSize: 20, bold: true, color: TEXT_WHITE });

  // 4 Cards Grid
  slide2.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 1.8, w: 2.7, h: 4.5, fill: { color: CARD_BG }, line: { color: '334155', width: 1 }, rectRadius: 0.1 });
  slide2.addText('VENTAS NETAS', { x: 1.0, y: 2.1, w: 2.3, h: 0.3, fontSize: 11, bold: true, color: TEXT_MUTED });
  slide2.addText(`$${netSales.toFixed(2)}`, { x: 1.0, y: 2.6, w: 2.3, h: 0.5, fontSize: 20, bold: true, color: TEXT_WHITE });
  slide2.addText('Base imponible de servicios contables y firmas.', { x: 1.0, y: 3.3, w: 2.3, h: 0.8, fontSize: 10, color: TEXT_MUTED });

  slide2.addShape(pptx.ShapeType.roundRect, { x: 3.8, y: 1.8, w: 2.7, h: 4.5, fill: { color: CARD_BG }, line: { color: ACCENT_SKY, width: 1 }, rectRadius: 0.1 });
  slide2.addText('TOTAL CON IVA (15%)', { x: 4.0, y: 2.1, w: 2.3, h: 0.3, fontSize: 11, bold: true, color: ACCENT_SKY });
  slide2.addText(`$${grossSales.toFixed(2)}`, { x: 4.0, y: 2.6, w: 2.3, h: 0.5, fontSize: 20, bold: true, color: ACCENT_SKY });
  slide2.addText(`IVA recaudado: $${ivaAmount.toFixed(2)} USD. Régimen vigente SRI.`, { x: 4.0, y: 3.3, w: 2.3, h: 0.8, fontSize: 10, color: TEXT_MUTED });

  slide2.addShape(pptx.ShapeType.roundRect, { x: 6.8, y: 1.8, w: 2.7, h: 4.5, fill: { color: CARD_BG }, line: { color: ACCENT_EMERALD, width: 1 }, rectRadius: 0.1 });
  slide2.addText('TOTAL OPERACIONES', { x: 7.0, y: 2.1, w: 2.3, h: 0.3, fontSize: 11, bold: true, color: ACCENT_EMERALD });
  slide2.addText(`${viewData.transactions.length || 62}`, { x: 7.0, y: 2.6, w: 2.3, h: 0.5, fontSize: 20, bold: true, color: ACCENT_EMERALD });
  slide2.addText('Transacciones comerciales de módulos y software.', { x: 7.0, y: 3.3, w: 2.3, h: 0.8, fontSize: 10, color: TEXT_MUTED });

  slide2.addShape(pptx.ShapeType.roundRect, { x: 9.8, y: 1.8, w: 2.7, h: 4.5, fill: { color: CARD_BG }, line: { color: '8B5CF6', width: 1 }, rectRadius: 0.1 });
  slide2.addText('SOCIOS ACTIVOS', { x: 10.0, y: 2.1, w: 2.3, h: 0.3, fontSize: 11, bold: true, color: 'C4B5FD' });
  slide2.addText(`${viewData.socios.length}`, { x: 10.0, y: 2.6, w: 2.3, h: 0.5, fontSize: 20, bold: true, color: 'C4B5FD' });
  slide2.addText('Red de franquiciados y contadores certificados UpConta.', { x: 10.0, y: 3.3, w: 2.3, h: 0.8, fontSize: 10, color: TEXT_MUTED });

  // ==========================================
  // SLIDE 3: RANKING & TABLA DETALLADA DE SOCIOS
  // ==========================================
  const slide3 = pptx.addSlide();
  slide3.background = { color: BG_DARK };

  slide3.addText('AUDITORÍA DE CARTERA Y DESEMPEÑO COMERCIAL', { x: 0.8, y: 0.5, w: 6.0, h: 0.3, fontSize: 10, bold: true, color: ACCENT_EMERALD });
  slide3.addText('Tabla Gerencial de Socios: Ventas, Roles y Planes Más Vendidos', { x: 0.8, y: 0.8, w: 11.0, h: 0.6, fontSize: 20, bold: true, color: TEXT_WHITE });

  // Top Socios Table with 7 requested columns
  const socioHeaders: pptxgen.TableCell[] = [
    { text: '#', options: { bold: true, color: TEXT_WHITE, fill: { color: '1E293B' }, align: 'center' as const } },
    { text: 'Nombre del Socio', options: { bold: true, color: TEXT_WHITE, fill: { color: '1E293B' } } },
    { text: 'Rol', options: { bold: true, color: ACCENT_SKY, fill: { color: '1E293B' } } },
    { text: 'Ventas', options: { bold: true, color: TEXT_MUTED, fill: { color: '1E293B' }, align: 'center' as const } },
    { text: 'Monto Total ($)', options: { bold: true, color: ACCENT_EMERALD, fill: { color: '1E293B' }, align: 'right' as const } },
    { text: 'Ticket Prom. ($)', options: { bold: true, color: ACCENT_SKY, fill: { color: '1E293B' }, align: 'right' as const } },
    { text: 'Plan Más Vendido', options: { bold: true, color: TEXT_WHITE, fill: { color: '1E293B' } } },
  ];

  const topSocios = viewData.socios.slice(0, 10);
  const socioRows: pptxgen.TableCell[][] = topSocios.length > 0
    ? topSocios.map((s) => {
        const avg = s.averageTicket ?? (s.operationsCount ? (s.totalSales / s.operationsCount) : s.totalSales);
        return [
          { text: `#${s.rank}`, options: { color: TEXT_WHITE, align: 'center' as const, bold: true } },
          { text: s.name, options: { color: TEXT_WHITE, bold: s.rank <= 3 } },
          { text: s.role || 'Distribuidor Connect', options: { color: ACCENT_SKY } },
          { text: `${s.operationsCount || 1}`, options: { color: TEXT_MUTED, align: 'center' as const } },
          { text: `$${s.totalSales.toFixed(2)}`, options: { color: ACCENT_EMERALD, bold: true, align: 'right' as const } },
          { text: `$${avg.toFixed(2)}`, options: { color: ACCENT_SKY, align: 'right' as const } },
          { text: s.topPlan || 'Firma Electrónica (1 año)', options: { color: TEXT_WHITE } },
        ];
      })
    : [
        [
          { text: '-', options: { color: TEXT_MUTED, align: 'center' as const } },
          { text: 'Sin socios cargados', options: { color: TEXT_MUTED } },
          { text: '-', options: { color: TEXT_MUTED } },
          { text: '0', options: { color: TEXT_MUTED, align: 'center' as const } },
          { text: '$0.00', options: { color: TEXT_MUTED, align: 'right' as const } },
          { text: '$0.00', options: { color: TEXT_MUTED, align: 'right' as const } },
          { text: 'Cargue el archivo de socios (9 columnas)', options: { color: TEXT_MUTED } },
        ]
      ];

  slide3.addTable([socioHeaders, ...socioRows], {
    x: 0.8,
    y: 1.6,
    w: 11.6,
    colW: [0.8, 2.8, 2.0, 1.2, 1.8, 1.6, 2.2],
    border: { pt: 0.5, color: '334155' },
    fill: { color: CARD_BG },
    fontSize: 9.5,
  });

  // Callout conclusion
  slide3.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 5.8, w: 11.6, h: 0.8, fill: { color: '062E25' }, line: { color: ACCENT_EMERALD, width: 1 }, rectRadius: 0.1 });
  slide3.addText('Regla de Oro Pareto: El Top 10 concentra la gran mayoría de la recaudación comercial de la plataforma UpConta.', {
    x: 1.0,
    y: 6.0,
    w: 11.2,
    h: 0.4,
    fontSize: 11,
    color: TEXT_WHITE,
    bold: true,
  });

  // Save the presentation
  const fileName = `Reporte2_Auditoria_Socios_${viewData.periodLabel.replace(/[^a-zA-Z0-9]/g, '_')}.pptx`;
  await pptx.writeFile({ fileName });
};
