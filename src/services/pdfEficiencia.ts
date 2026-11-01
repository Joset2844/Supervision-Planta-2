import { jsPDF } from 'jspdf';
import { CatalogoItem, Maquina, MetaTurno, TurnoMaquina } from '../types';
import { generarFilasTurno } from './excel';
import { formatearFechaTurno, etiquetaTurno } from './reports';

export interface MetricaMaquinaVasos {
  maquina: Maquina;
  meta: number;
  tieneMeta: boolean;
  totalProd: number;
  pct: number;
  diferencia: number;
  estado: 'perfeccion' | 'eficiente' | 'bajo' | 'sin_meta';
  clasificacion: string;
}

export function obtenerMetricasVasos(
  maquinas: Maquina[],
  turno: TurnoMaquina[],
  catalogo: CatalogoItem[]
): {
  items: MetricaMaquinaVasos[];
  totalMaquinas: number;
  conMetaCount: number;
  sinMetaCount: number;
  perfeccionCount: number;
  eficientesCount: number;
  bajoCount: number;
  metaTotal: number;
  prodTotal: number;
  pctGlobal: number;
  diferenciaGlobal: number;
} {
  const filas = generarFilasTurno(maquinas, turno, catalogo);

  // Mapear produccion por máquina
  const prodMap = new Map<string, number>();
  filas.forEach((f) => {
    const k = f.maquina.trim().toLowerCase();
    prodMap.set(k, (prodMap.get(k) || 0) + (f.total || 0));
  });

  // Filtrar exclusivamente maquinas activas del área de Vasos y Envases
  const maquinasVasos = maquinas.filter(
    (m) => m.activa && m.area === 'vasos_envases'
  );

  const items: MetricaMaquinaVasos[] = maquinasVasos.map((m) => {
    const meta = m.meta || 0;
    const tieneMeta = meta > 0;
    const totalProd = prodMap.get(m.nombre.trim().toLowerCase()) || 0;
    const pct = tieneMeta ? (totalProd / meta) * 100 : 0;
    const diferencia = tieneMeta ? totalProd - meta : 0;

    // Regla de negocio:
    // >= 100% -> Perfección
    // 80% a 99.9% -> Eficiente (a partir de 80-85% se considera eficiente)
    // < 80% -> Bajo Meta
    // Meta = 0 -> Sin meta (no se mide)
    let estado: 'perfeccion' | 'eficiente' | 'bajo' | 'sin_meta';
    let clasificacion: string;

    if (!tieneMeta) {
      estado = 'sin_meta';
      clasificacion = 'Sin meta';
    } else if (pct >= 100) {
      estado = 'perfeccion';
      clasificacion = 'Perfección';
    } else if (pct >= 80) {
      estado = 'eficiente';
      clasificacion = 'Eficiente';
    } else {
      estado = 'bajo';
      clasificacion = 'Bajo Meta';
    }

    return {
      maquina: m,
      meta,
      tieneMeta,
      totalProd,
      pct,
      diferencia,
      estado,
      clasificacion,
    };
  });

  const conMeta = items.filter((i) => i.tieneMeta);
  const perfeccionCount = conMeta.filter((i) => i.estado === 'perfeccion').length;
  const eficientesCount = conMeta.filter((i) => i.estado === 'eficiente').length;
  const bajoCount = conMeta.filter((i) => i.estado === 'bajo').length;
  const sinMetaCount = items.length - conMeta.length;

  const metaTotal = conMeta.reduce((acc, i) => acc + i.meta, 0);
  const prodTotal = items.reduce((acc, i) => acc + i.totalProd, 0);
  const prodConMeta = conMeta.reduce((acc, i) => acc + i.totalProd, 0);
  const pctGlobal = metaTotal > 0 ? (prodConMeta / metaTotal) * 100 : 0;
  const diferenciaGlobal = prodConMeta - metaTotal;

  return {
    items,
    totalMaquinas: items.length,
    conMetaCount: conMeta.length,
    sinMetaCount,
    perfeccionCount,
    eficientesCount,
    bajoCount,
    metaTotal,
    prodTotal,
    pctGlobal,
    diferenciaGlobal,
  };
}

export function exportarPdfEficienciaVasos(
  maquinas: Maquina[],
  turno: TurnoMaquina[],
  catalogo: CatalogoItem[],
  metaTurno: MetaTurno
): { success: boolean; nombreArchivo: string } {
  const metricas = obtenerMetricasVasos(maquinas, turno, catalogo);

  const tipoTurno = metaTurno.tipo === 'noche' ? 'Noche' : 'Día';
  const fechaTexto = formatearFechaTurno(metaTurno.fecha);
  const tituloOficial = `Eficiencia de Turno ${tipoTurno} ${fechaTexto}`;
  const nombreArchivo = `${tituloOficial.replace(/\//g, '-')}.pdf`;

  // Crear documento A4 apaisado (landscape)
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4',
  });

  const pageWidth = 297;
  const pageHeight = 210;
  const marginX = 14;
  const contentWidth = pageWidth - marginX * 2; // 269 mm

  // Paleta de colores
  const COLOR_HEADER = [26, 32, 44] as const; // #1A202C
  const COLOR_ACCENT = [242, 169, 59] as const; // #F2A93B (Dorado ámbar)
  const COLOR_BORDER = [45, 55, 72] as const; // #2D3748
  const COLOR_BG_LIGHT = [248, 250, 252] as const; // #F8FAFC
  const COLOR_PERFECCION = [40, 167, 69] as const; // Verde sólido (#28A745)
  const COLOR_EFICIENTE = [56, 161, 105] as const; // Verde medio (#38A169)
  const COLOR_BAJO = [229, 62, 62] as const; // Rojo (#E53E3E)
  const COLOR_MUTED = [113, 128, 150] as const; // Gris (#718096)
  const COLOR_BAR_BG = [226, 232, 240] as const; // Gris claro para fondo de barra

  let y = 12;

  // ================= 1. HEADER BANNER =================
  doc.setFillColor(COLOR_HEADER[0], COLOR_HEADER[1], COLOR_HEADER[2]);
  doc.roundedRect(marginX, y, contentWidth, 22, 2, 2, 'F');

  // Barra de acento lateral
  doc.setFillColor(COLOR_ACCENT[0], COLOR_ACCENT[1], COLOR_ACCENT[2]);
  doc.rect(marginX, y, 4, 22, 'F');

  // Título principal
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(15);
  doc.setTextColor(255, 255, 255);
  doc.text(tituloOficial.toUpperCase(), marginX + 8, y + 8.5);

  // Subtítulo
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.setTextColor(203, 213, 224);
  doc.text(
    `ÁREA EXCLUSIVA: VASOS Y ENVASES  ·  REPORTE OFICIAL DE RENDIMIENTO Y CUMPLIMIENTO`,
    marginX + 8,
    y + 16
  );

  // Badge derecho de turno
  const badgeWidth = 46;
  const badgeX = marginX + contentWidth - badgeWidth - 4;
  doc.setFillColor(58, 46, 27);
  doc.setDrawColor(COLOR_ACCENT[0], COLOR_ACCENT[1], COLOR_ACCENT[2]);
  doc.setLineWidth(0.4);
  doc.roundedRect(badgeX, y + 4, badgeWidth, 14, 2, 2, 'FD');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(COLOR_ACCENT[0], COLOR_ACCENT[1], COLOR_ACCENT[2]);
  doc.text(`TURNO ${tipoTurno.toUpperCase()}`, badgeX + badgeWidth / 2, y + 9.5, { align: 'center' });
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.5);
  doc.setTextColor(255, 255, 255);
  doc.text(fechaTexto, badgeX + badgeWidth / 2, y + 14.5, { align: 'center' });

  y += 25;

  // ================= 2. CRITERIOS DE EVALUACIÓN =================
  doc.setFillColor(241, 245, 249);
  doc.setDrawColor(203, 213, 224);
  doc.setLineWidth(0.3);
  doc.roundedRect(marginX, y, contentWidth, 8, 1.5, 1.5, 'FD');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7.5);
  doc.setTextColor(30, 41, 59);
  doc.text('CRITERIO DE EVALUACIÓN:', marginX + 4, y + 5.2);

  doc.setFont('helvetica', 'normal');
  doc.setTextColor(71, 85, 105);
  doc.text(
    'Eficiente a partir de 80%–85% producido  |  100% es Perfección  |  <80% Bajo Meta  |  Máquinas con meta 0 no se miden producción alguna.',
    marginX + 45,
    y + 5.2
  );

  y += 11;

  // ================= 3. TARJETAS KPI RESUMEN DE VASOS =================
  const kpiCount = 4;
  const kpiGap = 4;
  const kpiWidth = (contentWidth - kpiGap * (kpiCount - 1)) / kpiCount;
  const kpiHeight = 19;

  // KPI 1: Eficiencia Global Vasos
  doc.setFillColor(COLOR_BG_LIGHT[0], COLOR_BG_LIGHT[1], COLOR_BG_LIGHT[2]);
  doc.setDrawColor(COLOR_BORDER[0], COLOR_BORDER[1], COLOR_BORDER[2]);
  doc.setLineWidth(0.3);
  doc.roundedRect(marginX, y, kpiWidth, kpiHeight, 1.5, 1.5, 'FD');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7);
  doc.setTextColor(100, 116, 139);
  doc.text('EFICIENCIA GLOBAL VASOS', marginX + 3.5, y + 4.5);

  const pctGlobalTxt = `${metricas.pctGlobal.toFixed(1)}%`;
  doc.setFontSize(13);
  if (metricas.pctGlobal >= 100) {
    doc.setTextColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
  } else if (metricas.pctGlobal >= 80) {
    doc.setTextColor(COLOR_EFICIENTE[0], COLOR_EFICIENTE[1], COLOR_EFICIENTE[2]);
  } else {
    doc.setTextColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
  }
  doc.text(pctGlobalTxt, marginX + 3.5, y + 11.5);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  doc.setTextColor(71, 85, 105);
  const estadoGlobal =
    metricas.pctGlobal >= 100
      ? '★ PERFECCIÓN'
      : metricas.pctGlobal >= 80
      ? '✓ EFICIENTE'
      : '⚠ BAJO META';
  doc.text(estadoGlobal, marginX + 3.5, y + 16.5);

  // KPI 2: Total Producido vs Meta
  const kpi2X = marginX + kpiWidth + kpiGap;
  doc.setFillColor(COLOR_BG_LIGHT[0], COLOR_BG_LIGHT[1], COLOR_BG_LIGHT[2]);
  doc.roundedRect(kpi2X, y, kpiWidth, kpiHeight, 1.5, 1.5, 'FD');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7);
  doc.setTextColor(100, 116, 139);
  doc.text('PRODUCCIÓN VS META', kpi2X + 3.5, y + 4.5);

  doc.setFontSize(10.5);
  doc.setTextColor(15, 23, 42);
  doc.text(`${metricas.prodTotal.toLocaleString('es-PE')} u`, kpi2X + 3.5, y + 11);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  doc.setTextColor(100, 116, 139);
  doc.text(`Meta total: ${metricas.metaTotal.toLocaleString('es-PE')} u`, kpi2X + 3.5, y + 16.5);

  // KPI 3: Diferencia de Unidades
  const kpi3X = kpi2X + kpiWidth + kpiGap;
  doc.setFillColor(COLOR_BG_LIGHT[0], COLOR_BG_LIGHT[1], COLOR_BG_LIGHT[2]);
  doc.roundedRect(kpi3X, y, kpiWidth, kpiHeight, 1.5, 1.5, 'FD');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7);
  doc.setTextColor(100, 116, 139);
  doc.text('DIFERENCIA NETA', kpi3X + 3.5, y + 4.5);

  const difSign = metricas.diferenciaGlobal >= 0 ? '+' : '';
  const difTxt = `${difSign}${metricas.diferenciaGlobal.toLocaleString('es-PE')} u`;
  doc.setFontSize(10.5);
  if (metricas.diferenciaGlobal >= 0) {
    doc.setTextColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
  } else {
    doc.setTextColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
  }
  doc.text(difTxt, kpi3X + 3.5, y + 11);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  doc.setTextColor(100, 116, 139);
  doc.text('Respecto a meta programada', kpi3X + 3.5, y + 16.5);

  // KPI 4: Conteo de Máquinas
  const kpi4X = kpi3X + kpiWidth + kpiGap;
  doc.setFillColor(COLOR_BG_LIGHT[0], COLOR_BG_LIGHT[1], COLOR_BG_LIGHT[2]);
  doc.roundedRect(kpi4X, y, kpiWidth, kpiHeight, 1.5, 1.5, 'FD');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7);
  doc.setTextColor(100, 116, 139);
  doc.text('ESTADO DE MÁQUINAS VASOS', kpi4X + 3.5, y + 4.5);

  doc.setFontSize(7.5);
  doc.setTextColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
  doc.text(`Perfección (≥100%): ${metricas.perfeccionCount}`, kpi4X + 3.5, y + 10);

  doc.setTextColor(COLOR_EFICIENTE[0], COLOR_EFICIENTE[1], COLOR_EFICIENTE[2]);
  doc.text(`Eficiente (80-99%): ${metricas.eficientesCount}`, kpi4X + 3.5, y + 14.2);

  doc.setTextColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
  doc.text(
    `Bajo meta: ${metricas.bajoCount}  |  Sin meta: ${metricas.sinMetaCount}`,
    kpi4X + 3.5,
    y + 17.5
  );

  y += kpiHeight + 5;

  // ================= 4. GRÁFICA DE EFICIENCIA POR MÁQUINA =================
  // Encabezado de la sección Gráfica
  doc.setFillColor(COLOR_HEADER[0], COLOR_HEADER[1], COLOR_HEADER[2]);
  doc.roundedRect(marginX, y, contentWidth, 7, 1, 1, 'F');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.setTextColor(255, 255, 255);
  doc.text('GRÁFICA DE EFICIENCIA POR MÁQUINA (ÁREA DE VASOS)', marginX + 4, y + 4.8);

  // Guías visuales de la leyenda en el encabezado
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  doc.setTextColor(COLOR_ACCENT[0], COLOR_ACCENT[1], COLOR_ACCENT[2]);
  doc.text(
    'Marcadores:  [ -- 80% Mínimo Eficiente ]   [ | 100% Meta / Perfección ]',
    marginX + contentWidth - 4,
    y + 4.8,
    { align: 'right' }
  );

  y += 9;

  // Parámetros de la gráfica
  const labelColWidth = 52; // Ancho para nombre y código de máquina
  const statsColWidth = 58; // Ancho para texto de producción y meta
  const badgeColWidth = 26; // Ancho para porcentaje y clasificación
  const barStartX = marginX + labelColWidth;
  const barTotalWidth = contentWidth - labelColWidth - statsColWidth - badgeColWidth; // ~133 mm
  const maxDisplayPct = 120; // 100% está en (100 / 120) de la barra

  const rowHeight = 7.8;
  const barHeight = 4.4;

  if (metricas.items.length === 0) {
    doc.setFont('helvetica', 'italic');
    doc.setFontSize(9);
    doc.setTextColor(COLOR_MUTED[0], COLOR_MUTED[1], COLOR_MUTED[2]);
    doc.text('No hay máquinas activas registradas en el área de Vasos y Envases.', marginX + 4, y + 8);
    y += 14;
  } else {
    metricas.items.forEach((item, index) => {
      // Verificar si cabe en la página o agregar nueva
      if (y + rowHeight > pageHeight - 16) {
        doc.addPage('a4', 'landscape');
        y = 14;

        // Encabezado de continuación
        doc.setFillColor(COLOR_HEADER[0], COLOR_HEADER[1], COLOR_HEADER[2]);
        doc.roundedRect(marginX, y, contentWidth, 7, 1, 1, 'F');
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(8);
        doc.setTextColor(255, 255, 255);
        doc.text(
          `GRÁFICA DE EFICIENCIA POR MÁQUINA (CONTINUACIÓN) — ${tituloOficial.toUpperCase()}`,
          marginX + 4,
          y + 4.8
        );
        y += 9;
      }

      // Fila alterna de fondo
      if (index % 2 === 0) {
        doc.setFillColor(250, 250, 250);
        doc.rect(marginX, y, contentWidth, rowHeight, 'F');
      }

      // Nombre y código de máquina
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7.5);
      doc.setTextColor(30, 41, 59);
      const nombreMaq = item.maquina.nombre.length > 20
        ? item.maquina.nombre.slice(0, 19) + '…'
        : item.maquina.nombre;
      doc.text(nombreMaq, marginX + 2, y + 4.8);

      if (item.maquina.codigo) {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(6.5);
        doc.setTextColor(100, 116, 139);
        doc.text(`[${item.maquina.codigo}]`, marginX + 36, y + 4.8);
      }

      // Pista de la barra (fondo gris claro)
      doc.setFillColor(COLOR_BAR_BG[0], COLOR_BAR_BG[1], COLOR_BAR_BG[2]);
      doc.roundedRect(barStartX, y + 1.8, barTotalWidth, barHeight, 0.8, 0.8, 'F');

      // Línea de referencia al 80% (mínimo eficiente)
      const x80 = barStartX + (80 / maxDisplayPct) * barTotalWidth;
      doc.setDrawColor(242, 169, 59); // Dorado
      doc.setLineWidth(0.4);
      doc.line(x80, y + 1.2, x80, y + 1.8 + barHeight + 0.6);

      // Línea de referencia al 100% (perfección)
      const x100 = barStartX + (100 / maxDisplayPct) * barTotalWidth;
      doc.setDrawColor(40, 167, 69); // Verde
      doc.setLineWidth(0.6);
      doc.line(x100, y + 0.8, x100, y + 1.8 + barHeight + 1.0);

      // Barra de progreso según cumplimiento
      if (item.tieneMeta && item.pct > 0) {
        const visualPct = Math.min(item.pct, maxDisplayPct);
        const barFillWidth = Math.max(1.2, (visualPct / maxDisplayPct) * barTotalWidth);

        if (item.estado === 'perfeccion') {
          doc.setFillColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
        } else if (item.estado === 'eficiente') {
          doc.setFillColor(COLOR_EFICIENTE[0], COLOR_EFICIENTE[1], COLOR_EFICIENTE[2]);
        } else {
          doc.setFillColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
        }

        doc.roundedRect(barStartX, y + 1.8, barFillWidth, barHeight, 0.8, 0.8, 'F');
      }

      // Columna de datos de producción y meta
      const statsX = barStartX + barTotalWidth + 4;
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7);
      doc.setTextColor(30, 41, 59);
      doc.text(`${item.totalProd.toLocaleString('es-PE')} u`, statsX, y + 4.8);

      doc.setFont('helvetica', 'normal');
      doc.setFontSize(6.5);
      doc.setTextColor(100, 116, 139);
      if (item.tieneMeta) {
        doc.text(`/ ${item.meta.toLocaleString('es-PE')} meta`, statsX + 22, y + 4.8);
      } else {
        doc.text('(sin meta)', statsX + 22, y + 4.8);
      }

      // Columna de badge y porcentaje
      const badgeXPos = statsX + statsColWidth;
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7.5);

      if (!item.tieneMeta) {
        doc.setTextColor(COLOR_MUTED[0], COLOR_MUTED[1], COLOR_MUTED[2]);
        doc.text('—  Sin meta', badgeXPos, y + 4.8);
      } else {
        const pctFormatted = item.pct >= 100 ? `${Math.round(item.pct)}%` : `${item.pct.toFixed(1)}%`;
        if (item.estado === 'perfeccion') {
          doc.setTextColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
          doc.text(`${pctFormatted} ★ Perfección`, badgeXPos, y + 4.8);
        } else if (item.estado === 'eficiente') {
          doc.setTextColor(COLOR_EFICIENTE[0], COLOR_EFICIENTE[1], COLOR_EFICIENTE[2]);
          doc.text(`${pctFormatted} ✓ Eficiente`, badgeXPos, y + 4.8);
        } else {
          doc.setTextColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
          doc.text(`${pctFormatted} ⚠ Bajo meta`, badgeXPos, y + 4.8);
        }
      }

      y += rowHeight;
    });
  }

  y += 5;

  // ================= 5. TABLA RESUMEN DETALLADA =================
  if (y + 40 > pageHeight - 16) {
    doc.addPage('a4', 'landscape');
    y = 14;
  }

  // Encabezado de tabla
  doc.setFillColor(COLOR_HEADER[0], COLOR_HEADER[1], COLOR_HEADER[2]);
  doc.rect(marginX, y, contentWidth, 6, 'F');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(6.8);
  doc.setTextColor(255, 255, 255);

  const colMaq = marginX + 3;
  const colCod = marginX + 62;
  const colMeta = marginX + 95;
  const colProd = marginX + 130;
  const colDif = marginX + 165;
  const colPct = marginX + 200;
  const colClasif = marginX + 235;

  doc.text('MÁQUINA', colMaq, y + 4.2);
  doc.text('CÓDIGO', colCod, y + 4.2);
  doc.text('META (CANTIDAD)', colMeta, y + 4.2);
  doc.text('PRODUCIDO', colProd, y + 4.2);
  doc.text('DIFERENCIA', colDif, y + 4.2);
  doc.text('EFICIENCIA', colPct, y + 4.2);
  doc.text('CLASIFICACIÓN', colClasif, y + 4.2);

  y += 6;

  // Filas de tabla
  const tableRowH = 5.4;
  metricas.items.forEach((item, index) => {
    if (y + tableRowH > pageHeight - 16) {
      doc.addPage('a4', 'landscape');
      y = 14;

      // Reimprimir encabezado
      doc.setFillColor(COLOR_HEADER[0], COLOR_HEADER[1], COLOR_HEADER[2]);
      doc.rect(marginX, y, contentWidth, 6, 'F');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(6.8);
      doc.setTextColor(255, 255, 255);
      doc.text('MÁQUINA', colMaq, y + 4.2);
      doc.text('CÓDIGO', colCod, y + 4.2);
      doc.text('META (CANTIDAD)', colMeta, y + 4.2);
      doc.text('PRODUCIDO', colProd, y + 4.2);
      doc.text('DIFERENCIA', colDif, y + 4.2);
      doc.text('EFICIENCIA', colPct, y + 4.2);
      doc.text('CLASIFICACIÓN', colClasif, y + 4.2);
      y += 6;
    }

    if (index % 2 === 1) {
      doc.setFillColor(248, 250, 252);
      doc.rect(marginX, y, contentWidth, tableRowH, 'F');
    }

    // Bordes sutiles
    doc.setDrawColor(226, 232, 240);
    doc.setLineWidth(0.2);
    doc.line(marginX, y + tableRowH, marginX + contentWidth, y + tableRowH);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(6.8);
    doc.setTextColor(30, 41, 59);

    doc.text(item.maquina.nombre, colMaq, y + 3.8);
    doc.text(item.maquina.codigo || '—', colCod, y + 3.8);
    doc.text(item.tieneMeta ? `${item.meta.toLocaleString('es-PE')} u` : '0 u (No mide)', colMeta, y + 3.8);
    doc.text(`${item.totalProd.toLocaleString('es-PE')} u`, colProd, y + 3.8);

    // Diferencia
    if (item.tieneMeta) {
      const difTxt = `${item.diferencia >= 0 ? '+' : ''}${item.diferencia.toLocaleString('es-PE')} u`;
      if (item.diferencia >= 0) {
        doc.setTextColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
      } else {
        doc.setTextColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
      }
      doc.text(difTxt, colDif, y + 3.8);
    } else {
      doc.setTextColor(COLOR_MUTED[0], COLOR_MUTED[1], COLOR_MUTED[2]);
      doc.text('—', colDif, y + 3.8);
    }

    // Eficiencia
    if (item.tieneMeta) {
      const pctTxt = item.pct >= 100 ? `${Math.round(item.pct)}%` : `${item.pct.toFixed(1)}%`;
      if (item.estado === 'perfeccion') {
        doc.setTextColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
      } else if (item.estado === 'eficiente') {
        doc.setTextColor(COLOR_EFICIENTE[0], COLOR_EFICIENTE[1], COLOR_EFICIENTE[2]);
      } else {
        doc.setTextColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
      }
      doc.text(pctTxt, colPct, y + 3.8);
    } else {
      doc.setTextColor(COLOR_MUTED[0], COLOR_MUTED[1], COLOR_MUTED[2]);
      doc.text('—', colPct, y + 3.8);
    }

    // Clasificación
    if (item.estado === 'perfeccion') {
      doc.setTextColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
      doc.text('Perfección (≥100%)', colClasif, y + 3.8);
    } else if (item.estado === 'eficiente') {
      doc.setTextColor(COLOR_EFICIENTE[0], COLOR_EFICIENTE[1], COLOR_EFICIENTE[2]);
      doc.text('Eficiente (80-99%)', colClasif, y + 3.8);
    } else if (item.estado === 'bajo') {
      doc.setTextColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
      doc.text('Bajo Meta (<80%)', colClasif, y + 3.8);
    } else {
      doc.setTextColor(COLOR_MUTED[0], COLOR_MUTED[1], COLOR_MUTED[2]);
      doc.text('Sin medición', colClasif, y + 3.8);
    }

    y += tableRowH;
  });

  // Fila de TOTALES
  if (y + tableRowH + 2 > pageHeight - 16) {
    doc.addPage('a4', 'landscape');
    y = 14;
  }

  doc.setFillColor(237, 242, 247);
  doc.rect(marginX, y, contentWidth, tableRowH + 1, 'F');
  doc.setDrawColor(COLOR_BORDER[0], COLOR_BORDER[1], COLOR_BORDER[2]);
  doc.setLineWidth(0.3);
  doc.rect(marginX, y, contentWidth, tableRowH + 1, 'S');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7);
  doc.setTextColor(15, 23, 42);

  doc.text('TOTAL ÁREA VASOS', colMaq, y + 4.2);
  doc.text(`${metricas.items.length} máqs`, colCod, y + 4.2);
  doc.text(`${metricas.metaTotal.toLocaleString('es-PE')} u`, colMeta, y + 4.2);
  doc.text(`${metricas.prodTotal.toLocaleString('es-PE')} u`, colProd, y + 4.2);

  const totalDifTxt = `${metricas.diferenciaGlobal >= 0 ? '+' : ''}${metricas.diferenciaGlobal.toLocaleString('es-PE')} u`;
  if (metricas.diferenciaGlobal >= 0) {
    doc.setTextColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
  } else {
    doc.setTextColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
  }
  doc.text(totalDifTxt, colDif, y + 4.2);

  const totalPctTxt = `${metricas.pctGlobal.toFixed(1)}%`;
  if (metricas.pctGlobal >= 100) {
    doc.setTextColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
  } else if (metricas.pctGlobal >= 80) {
    doc.setTextColor(COLOR_EFICIENTE[0], COLOR_EFICIENTE[1], COLOR_EFICIENTE[2]);
  } else {
    doc.setTextColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
  }
  doc.text(totalPctTxt, colPct, y + 4.2);

  doc.setTextColor(15, 23, 42);
  doc.text(
    metricas.pctGlobal >= 100
      ? '★ PERFECCIÓN'
      : metricas.pctGlobal >= 80
      ? '✓ EFICIENTE'
      : '⚠ BAJO META',
    colClasif,
    y + 4.2
  );

  // ================= 6. PIE DE PÁGINA EN TODAS LAS PÁGINAS =================
  const totalPaginas = (doc as any).internal.getNumberOfPages();
  const fechaGeneracion = new Date().toLocaleString('es-PE', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });

  for (let p = 1; p <= totalPaginas; p++) {
    doc.setPage(p);
    doc.setDrawColor(203, 213, 224);
    doc.setLineWidth(0.2);
    doc.line(marginX, pageHeight - 10, marginX + contentWidth, pageHeight - 10);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(6.5);
    doc.setTextColor(113, 128, 150);
    doc.text(
      `Generado: ${fechaGeneracion}  ·  Supervisión de Planta  ·  ${tituloOficial}`,
      marginX,
      pageHeight - 6.5
    );

    doc.text(`Página ${p} de ${totalPaginas}`, marginX + contentWidth, pageHeight - 6.5, {
      align: 'right',
    });
  }

  // Guardar y descargar archivo PDF directamente
  doc.save(nombreArchivo);

  return { success: true, nombreArchivo };
}
