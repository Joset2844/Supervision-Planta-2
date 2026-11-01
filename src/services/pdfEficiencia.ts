import { jsPDF } from 'jspdf';
import { CatalogoItem, FilaTurnoExcel, Maquina, MetaTurno, TurnoMaquina } from '../types';
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

export interface ResultadoMetricasVasos {
  items: MetricaMaquinaVasos[];
  totalMaquinas: number;
  conMetaCount: number;
  sinMetaCount: number;
  perfeccionCount: number;
  eficientesCount: number;
  bajoCount: number;
  metaTotal: number;
  prodTotal: number;
  prodConMeta: number;
  pctGlobal: number;
  diferenciaGlobal: number;
}

/**
 * Obtiene métricas de eficiencia exclusivamente para el área de Vasos y Envases
 * a partir de un arreglo de filas de producción (FilaTurnoExcel[]).
 */
export function obtenerMetricasVasosDesdeFilas(
  filas: FilaTurnoExcel[],
  maquinas: Maquina[]
): ResultadoMetricasVasos {
  // 1. Mapear la producción total por nombre de máquina
  const prodMap = new Map<string, number>();
  filas.forEach((f) => {
    const k = (f.maquina || '').trim().toLowerCase();
    if (k) {
      prodMap.set(k, (prodMap.get(k) || 0) + (f.total || 0));
    }
  });

  // 2. Filtrar máquinas activas pertenecientes al área de Vasos y Envases
  const maquinasVasos = maquinas.filter(
    (m) => m.activa !== false && m.area === 'vasos_envases'
  );

  // Asegurar que si en las filas hubo alguna máquina de vasos que no esté en la lista activa, se incorpore
  const nombresSet = new Set(maquinasVasos.map((m) => m.nombre.trim().toLowerCase()));
  const maquinasExtras: Maquina[] = [];

  filas.forEach((f) => {
    if (f.area === 'vasos_envases') {
      const k = (f.maquina || '').trim().toLowerCase();
      if (k && !nombresSet.has(k)) {
        nombresSet.add(k);
        const enMaquinas = maquinas.find((m) => m.nombre.trim().toLowerCase() === k);
        maquinasExtras.push(
          enMaquinas || {
            id: `maq_${k}`,
            nombre: f.maquina,
            codigo: f.codigoMaquina || '',
            area: 'vasos_envases',
            activa: true,
            bloqueada: false,
            tipoProducto: '',
            meta: 0,
          }
        );
      }
    }
  });

  const todasLasMaquinas = [...maquinasVasos, ...maquinasExtras];

  // 3. Calcular métricas por máquina
  const items: MetricaMaquinaVasos[] = todasLasMaquinas.map((m) => {
    const meta = m.meta || 0;
    const tieneMeta = meta > 0;
    const totalProd = prodMap.get(m.nombre.trim().toLowerCase()) || 0;
    const pct = tieneMeta ? (totalProd / meta) * 100 : 0;
    const diferencia = tieneMeta ? totalProd - meta : 0;

    // Regla de evaluación solicitada:
    // - Eficiente a partir de 80-85% producido
    // - 100% es Perfección (>= 100%)
    // - <80% Bajo Meta
    // - Meta = 0 no se mide producción alguna (Sin meta)
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

  // 4. Calcular métricas globales
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
    prodConMeta,
    pctGlobal,
    diferenciaGlobal,
  };
}

/**
 * Obtiene métricas de eficiencia para el turno en curso.
 */
export function obtenerMetricasVasos(
  maquinas: Maquina[],
  turno: TurnoMaquina[],
  catalogo: CatalogoItem[]
): ResultadoMetricasVasos {
  const filas = generarFilasTurno(maquinas, turno, catalogo);
  return obtenerMetricasVasosDesdeFilas(filas, maquinas);
}

/**
 * Motor central de generación del reporte PDF con gráficos de alta fidelidad.
 */
function compilarYDescargarPdfVasos(
  metricas: ResultadoMetricasVasos,
  metaTurno: MetaTurno
): { success: boolean; nombreArchivo: string } {
  const tipoTurno = metaTurno.tipo === 'noche' ? 'Noche' : 'Día';
  const fechaTexto = formatearFechaTurno(metaTurno.fecha);
  const tituloOficial = `Eficiencia de Turno ${tipoTurno} ${fechaTexto}`;
  const nombreArchivo = `${tituloOficial.replace(/\//g, '-')}.pdf`;

  // Crear documento A4 en orientación horizontal (landscape)
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4',
  });

  const pageWidth = 297;
  const pageHeight = 210;
  const marginX = 14;
  const contentWidth = pageWidth - marginX * 2; // 269 mm

  // Paleta de colores industriales de alta legibilidad
  const COLOR_HEADER = [26, 32, 44] as const; // #1A202C
  const COLOR_ACCENT = [242, 169, 59] as const; // #F2A93B Dorado ámbar
  const COLOR_BORDER = [45, 55, 72] as const; // #2D3748
  const COLOR_BG_LIGHT = [248, 250, 252] as const; // #F8FAFC
  const COLOR_PERFECCION = [16, 185, 129] as const; // #10B981 Esmeralda
  const COLOR_EFICIENTE = [34, 197, 94] as const; // #22C55E Verde brillante
  const COLOR_BAJO = [239, 68, 68] as const; // #EF4444 Rojo coral
  const COLOR_MUTED = [100, 116, 139] as const; // #64748B Slate gray
  const COLOR_BAR_BG = [226, 232, 240] as const; // #E2E8F0 Pista de la barra

  let y = 12;

  // ================= 1. HEADER BANNER =================
  doc.setFillColor(COLOR_HEADER[0], COLOR_HEADER[1], COLOR_HEADER[2]);
  doc.roundedRect(marginX, y, contentWidth, 22, 2, 2, 'F');

  // Barra de acento lateral dorada
  doc.setFillColor(COLOR_ACCENT[0], COLOR_ACCENT[1], COLOR_ACCENT[2]);
  doc.rect(marginX, y, 4, 22, 'F');

  // Título oficial solicitado
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(15);
  doc.setTextColor(255, 255, 255);
  doc.text(tituloOficial.toUpperCase(), marginX + 8, y + 8.5);

  // Subtítulo de área y contexto
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.setTextColor(203, 213, 224);
  doc.text(
    `ÁREA EXCLUSIVA: VASOS Y ENVASES  ·  GRÁFICA Y REPORTE DE CUMPLIMIENTO POR TURNO`,
    marginX + 8,
    y + 16
  );

  // Badge derecho de turno
  const badgeWidth = 46;
  const badgeX = marginX + contentWidth - badgeWidth - 4;
  doc.setFillColor(35, 42, 56);
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
  doc.roundedRect(marginX, y, contentWidth, 7.5, 1.5, 1.5, 'FD');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7.2);
  doc.setTextColor(30, 41, 59);
  doc.text('CRITERIO OFICIAL:', marginX + 3.5, y + 4.8);

  doc.setFont('helvetica', 'normal');
  doc.setTextColor(71, 85, 105);
  doc.text(
    'Eficiente a partir de 80%–85% producido  |  100% es Perfección  |  <80% Bajo Meta  |  Máquinas con meta 0 no se miden.',
    marginX + 36,
    y + 4.8
  );

  y += 10.5;

  // ================= 3. TARJETAS KPI RESUMEN =================
  const kpiCount = 4;
  const kpiGap = 4;
  const kpiWidth = (contentWidth - kpiGap * (kpiCount - 1)) / kpiCount;
  const kpiHeight = 18.5;

  // KPI 1: Eficiencia Global Vasos
  doc.setFillColor(COLOR_BG_LIGHT[0], COLOR_BG_LIGHT[1], COLOR_BG_LIGHT[2]);
  doc.setDrawColor(COLOR_BORDER[0], COLOR_BORDER[1], COLOR_BORDER[2]);
  doc.setLineWidth(0.3);
  doc.roundedRect(marginX, y, kpiWidth, kpiHeight, 1.5, 1.5, 'FD');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7);
  doc.setTextColor(100, 116, 139);
  doc.text('EFICIENCIA GLOBAL VASOS', marginX + 3.5, y + 4.3);

  const pctGlobalTxt = `${metricas.pctGlobal.toFixed(1)}%`;
  doc.setFontSize(13);
  if (metricas.pctGlobal >= 100) {
    doc.setTextColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
  } else if (metricas.pctGlobal >= 80) {
    doc.setTextColor(COLOR_EFICIENTE[0], COLOR_EFICIENTE[1], COLOR_EFICIENTE[2]);
  } else {
    doc.setTextColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
  }
  doc.text(pctGlobalTxt, marginX + 3.5, y + 11.2);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  doc.setTextColor(71, 85, 105);
  const estadoGlobal =
    metricas.pctGlobal >= 100
      ? '★ PERFECCIÓN (>=100%)'
      : metricas.pctGlobal >= 80
      ? '✓ EFICIENTE (>=80%)'
      : '⚠ BAJO META (<80%)';
  doc.text(estadoGlobal, marginX + 3.5, y + 16);

  // KPI 2: Total Producido vs Meta
  const kpi2X = marginX + kpiWidth + kpiGap;
  doc.setFillColor(COLOR_BG_LIGHT[0], COLOR_BG_LIGHT[1], COLOR_BG_LIGHT[2]);
  doc.roundedRect(kpi2X, y, kpiWidth, kpiHeight, 1.5, 1.5, 'FD');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7);
  doc.setTextColor(100, 116, 139);
  doc.text('PRODUCCIÓN VS META', kpi2X + 3.5, y + 4.3);

  doc.setFontSize(11);
  doc.setTextColor(15, 23, 42);
  doc.text(`${metricas.prodTotal.toLocaleString('es-PE')} u`, kpi2X + 3.5, y + 11);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  doc.setTextColor(100, 116, 139);
  doc.text(`Meta total: ${metricas.metaTotal.toLocaleString('es-PE')} u`, kpi2X + 3.5, y + 16);

  // KPI 3: Diferencia de Unidades
  const kpi3X = kpi2X + kpiWidth + kpiGap;
  doc.setFillColor(COLOR_BG_LIGHT[0], COLOR_BG_LIGHT[1], COLOR_BG_LIGHT[2]);
  doc.roundedRect(kpi3X, y, kpiWidth, kpiHeight, 1.5, 1.5, 'FD');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7);
  doc.setTextColor(100, 116, 139);
  doc.text('DIFERENCIA NETA', kpi3X + 3.5, y + 4.3);

  const difSign = metricas.diferenciaGlobal >= 0 ? '+' : '';
  const difTxt = `${difSign}${metricas.diferenciaGlobal.toLocaleString('es-PE')} u`;
  doc.setFontSize(11);
  if (metricas.diferenciaGlobal >= 0) {
    doc.setTextColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
  } else {
    doc.setTextColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
  }
  doc.text(difTxt, kpi3X + 3.5, y + 11);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  doc.setTextColor(100, 116, 139);
  doc.text('Respecto a meta programada', kpi3X + 3.5, y + 16);

  // KPI 4: Conteo y Distribución de Máquinas
  const kpi4X = kpi3X + kpiWidth + kpiGap;
  doc.setFillColor(COLOR_BG_LIGHT[0], COLOR_BG_LIGHT[1], COLOR_BG_LIGHT[2]);
  doc.roundedRect(kpi4X, y, kpiWidth, kpiHeight, 1.5, 1.5, 'FD');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7);
  doc.setTextColor(100, 116, 139);
  doc.text('DESGLOSE DE MÁQUINAS', kpi4X + 3.5, y + 4.3);

  doc.setFontSize(7.2);
  doc.setTextColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
  doc.text(`★ Perfección (≥100%): ${metricas.perfeccionCount}`, kpi4X + 3.5, y + 9.8);

  doc.setTextColor(COLOR_EFICIENTE[0], COLOR_EFICIENTE[1], COLOR_EFICIENTE[2]);
  doc.text(`✓ Eficiente (80-99%): ${metricas.eficientesCount}`, kpi4X + 3.5, y + 13.5);

  doc.setTextColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
  doc.text(
    `⚠ Bajo meta: ${metricas.bajoCount}   |   Sin meta: ${metricas.sinMetaCount}`,
    kpi4X + 3.5,
    y + 17
  );

  y += kpiHeight + 4.5;

  // ================= 4. BARRA DE DISTRIBUCIÓN GLOBAL (GRÁFICO APILADO) =================
  const totalMedibles = metricas.conMetaCount;
  if (totalMedibles > 0) {
    const distH = 4.2;
    const pctPerfeccion = (metricas.perfeccionCount / totalMedibles) * 100;
    const pctEficientes = (metricas.eficientesCount / totalMedibles) * 100;
    const pctBajo = (metricas.bajoCount / totalMedibles) * 100;

    const wPerf = (pctPerfeccion / 100) * contentWidth;
    const wEfic = (pctEficientes / 100) * contentWidth;
    const wBajo = contentWidth - wPerf - wEfic;

    let currX = marginX;

    // Segmento Perfección
    if (wPerf > 0) {
      doc.setFillColor(COLOR_PERFECCION[0], COLOR_PERFECCION[1], COLOR_PERFECCION[2]);
      doc.rect(currX, y, wPerf, distH, 'F');
      currX += wPerf;
    }
    // Segmento Eficiente
    if (wEfic > 0) {
      doc.setFillColor(COLOR_EFICIENTE[0], COLOR_EFICIENTE[1], COLOR_EFICIENTE[2]);
      doc.rect(currX, y, wEfic, distH, 'F');
      currX += wEfic;
    }
    // Segmento Bajo
    if (wBajo > 0) {
      doc.setFillColor(COLOR_BAJO[0], COLOR_BAJO[1], COLOR_BAJO[2]);
      doc.rect(currX, y, wBajo, distH, 'F');
    }

    // Borde exterior de la barra
    doc.setDrawColor(COLOR_BORDER[0], COLOR_BORDER[1], COLOR_BORDER[2]);
    doc.setLineWidth(0.3);
    doc.rect(marginX, y, contentWidth, distH, 'S');

    y += distH + 4;
  }

  // ================= 5. GRÁFICA DE EFICIENCIA POR MÁQUINA =================
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

  y += 8.5;

  // Parámetros y dimensiones de la gráfica horizontal
  const labelColWidth = 52; // Nombre y código de máquina
  const statsColWidth = 56; // Producción y meta
  const badgeColWidth = 28; // Badge de eficiencia y estado
  const barStartX = marginX + labelColWidth;
  const barTotalWidth = contentWidth - labelColWidth - statsColWidth - badgeColWidth; // ~133 mm
  const maxDisplayPct = 120; // 100% se sitúa en (100 / 120) del ancho total de la barra

  const rowHeight = 7.6;
  const barHeight = 4.2;

  if (metricas.items.length === 0) {
    doc.setFont('helvetica', 'italic');
    doc.setFontSize(8.5);
    doc.setTextColor(COLOR_MUTED[0], COLOR_MUTED[1], COLOR_MUTED[2]);
    doc.text('No hay máquinas registradas en el área de Vasos y Envases.', marginX + 4, y + 8);
    y += 14;
  } else {
    metricas.items.forEach((item, index) => {
      // Paginación si excede el límite vertical
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
        y += 8.5;
      }

      // Fila alterna con sombreado sutil
      if (index % 2 === 0) {
        doc.setFillColor(250, 250, 250);
        doc.rect(marginX, y, contentWidth, rowHeight, 'F');
      }

      // Nombre y código de máquina
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7.5);
      doc.setTextColor(30, 41, 59);
      const nombreMaq =
        item.maquina.nombre.length > 20
          ? item.maquina.nombre.slice(0, 19) + '…'
          : item.maquina.nombre;
      doc.text(nombreMaq, marginX + 2, y + 4.8);

      if (item.maquina.codigo) {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(6.5);
        doc.setTextColor(100, 116, 139);
        doc.text(`[${item.maquina.codigo}]`, marginX + 36, y + 4.8);
      }

      // Pista de la barra (fondo gris claro con borde)
      doc.setFillColor(COLOR_BAR_BG[0], COLOR_BAR_BG[1], COLOR_BAR_BG[2]);
      doc.roundedRect(barStartX, y + 1.7, barTotalWidth, barHeight, 0.8, 0.8, 'F');
      doc.setDrawColor(203, 213, 224);
      doc.setLineWidth(0.2);
      doc.roundedRect(barStartX, y + 1.7, barTotalWidth, barHeight, 0.8, 0.8, 'S');

      // Línea de referencia al 80% (mínimo eficiente)
      const x80 = barStartX + (80 / maxDisplayPct) * barTotalWidth;
      doc.setDrawColor(242, 169, 59); // Dorado
      doc.setLineWidth(0.5);
      doc.line(x80, y + 1.1, x80, y + 1.7 + barHeight + 0.6);

      // Línea de referencia al 100% (perfección)
      const x100 = barStartX + (100 / maxDisplayPct) * barTotalWidth;
      doc.setDrawColor(16, 185, 129); // Verde esmeralda
      doc.setLineWidth(0.7);
      doc.line(x100, y + 0.7, x100, y + 1.7 + barHeight + 1.0);

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

        doc.roundedRect(barStartX, y + 1.7, barFillWidth, barHeight, 0.8, 0.8, 'F');
      }

      // Columna de datos numéricos (Producción / Meta)
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
      doc.setFontSize(7.2);

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

  y += 4.5;

  // ================= 6. TABLA RESUMEN DETALLADA =================
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
  const tableRowH = 5.3;
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

  // ================= 7. PIE DE PÁGINA EN TODAS LAS PÁGINAS =================
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

/**
 * Exporta el PDF de eficiencia de Vasos para el turno activo en producción.
 */
export function exportarPdfEficienciaVasos(
  maquinas: Maquina[],
  turno: TurnoMaquina[],
  catalogo: CatalogoItem[],
  metaTurno: MetaTurno
): { success: boolean; nombreArchivo: string } {
  const metricas = obtenerMetricasVasos(maquinas, turno, catalogo);
  return compilarYDescargarPdfVasos(metricas, metaTurno);
}

/**
 * Exporta el PDF de eficiencia de Vasos directamente desde filas (para turnos del historial).
 */
export function exportarPdfEficienciaVasosDesdeFilas(
  filas: FilaTurnoExcel[],
  maquinas: Maquina[],
  metaTurno: MetaTurno
): { success: boolean; nombreArchivo: string } {
  const metricas = obtenerMetricasVasosDesdeFilas(filas, maquinas);
  return compilarYDescargarPdfVasos(metricas, metaTurno);
}
