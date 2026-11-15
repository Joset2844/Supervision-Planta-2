import { jsPDF } from 'jspdf';
import { CatalogoItem } from '../types';
import { formatearFechaTurno } from './reports';
import { ConsultaOpResultado, TurnoOpResumen } from './opHistory';

interface ModeloResumenPdf {
  nombre: string;
  sap: string;
  cantidadOrden: number;
  ptTotal: number;
  ptMerma: number;
  ppTotal: number;
  ppMerma: number;
  sinTipoTotal: number;
  sinTipoMerma: number;
  maquinas: string[];
  turnosTexto: string;
}

function limpiarPliegoSufijo(texto: string): string {
  return (texto || '')
    .replace(/\s*-\s*(?:Pliego\s*\d+|Tira|Retira)$/i, '')
    .trim();
}

function construirModelosResumen(
  catalogMatch: CatalogoItem | null,
  todosLosTurnos: TurnoOpResumen[]
): ModeloResumenPdf[] {
  interface AcumModelo {
    nombre: string;
    modeloCorto: string;
    sap: string;
    cantidadOrden: number;
    ptTotal: number;
    ptMerma: number;
    ppTotal: number;
    ppMerma: number;
    sinTipoTotal: number;
    sinTipoMerma: number;
    maquinasSet: Set<string>;
    turnosSet: Map<string, string>;
  }

  const lista: AcumModelo[] = [];

  // 1. Sembrar modelos del catálogo (incluyendo los que aún estén en 0)
  if (catalogMatch && catalogMatch.modelos && catalogMatch.modelos.length > 0) {
    catalogMatch.modelos.forEach((mo, idx) => {
      lista.push({
        nombre: mo.descripcion || mo.modelo || `Modelo ${idx + 1}`,
        modeloCorto: mo.modelo || '',
        sap: (mo.sap || '').trim(),
        cantidadOrden: Number(mo.cantidadTotal) || 0,
        ptTotal: 0,
        ptMerma: 0,
        ppTotal: 0,
        ppMerma: 0,
        sinTipoTotal: 0,
        sinTipoMerma: 0,
        maquinasSet: new Set<string>(),
        turnosSet: new Map<string, string>(),
      });
    });
  }

  const buscarOCrearDestino = (productoRaw: string, sapRaw: string): AcumModelo => {
    const sapClean = (sapRaw || '').trim().toLowerCase();
    const prodClean = limpiarPliegoSufijo(productoRaw);
    const prodLower = prodClean.toLowerCase();

    // Match por SAP exacto si existe
    if (sapClean) {
      const bySap = lista.find((m) => m.sap && m.sap.toLowerCase() === sapClean);
      if (bySap) return bySap;
    }

    // Match por descripción o nombre corto del modelo
    if (prodLower) {
      const byExact = lista.find(
        (m) =>
          m.nombre.toLowerCase() === prodLower ||
          (m.modeloCorto && m.modeloCorto.toLowerCase() === prodLower)
      );
      if (byExact) return byExact;

      const byPartial = lista.find(
        (m) =>
          (m.modeloCorto && prodLower.includes(m.modeloCorto.toLowerCase())) ||
          (m.nombre && (prodLower.includes(m.nombre.toLowerCase()) || m.nombre.toLowerCase().includes(prodLower)))
      );
      if (byPartial) return byPartial;
    }

    // Si el catálogo solo tiene 1 modelo y el registro vino sin producto especificado
    if (lista.length === 1 && !prodLower) {
      return lista[0];
    }

    // Crear entrada nueva para variante no listada en catálogo
    const nuevoNombre = prodClean || 'Sin modelo especificado';
    const existente = lista.find(
      (m) => m.nombre.toLowerCase() === nuevoNombre.toLowerCase() && m.sap.toLowerCase() === sapClean
    );
    if (existente) return existente;

    const nuevo: AcumModelo = {
      nombre: nuevoNombre,
      modeloCorto: '',
      sap: (sapRaw || '').trim(),
      cantidadOrden: catalogMatch && lista.length === 0 ? Number(catalogMatch.cantidadTotal) || 0 : 0,
      ptTotal: 0,
      ptMerma: 0,
      ppTotal: 0,
      ppMerma: 0,
      sinTipoTotal: 0,
      sinTipoMerma: 0,
      maquinasSet: new Set<string>(),
      turnosSet: new Map<string, string>(),
    };
    lista.push(nuevo);
    return nuevo;
  };

  // 2. Acumular registros reales de todos los turnos
  todosLosTurnos.forEach((t) => {
    const fShort = formatearFechaTurno(t.fecha).slice(0, 5); // dd/mm
    const tShort = t.tipo === 'noche' ? 'Noche' : t.tipo === 'dia' ? 'Día' : '';
    const turnoLabel = t.fuente === 'actual' ? `${fShort} ${tShort} (Act.)`.trim() : `${fShort} ${tShort}`.trim();

    t.filas.forEach((f) => {
      const dest = buscarOCrearDestino(f.producto, f.sap);
      if (f.tipoProducto === 'pp') {
        dest.ppTotal += f.total;
        dest.ppMerma += f.merma;
      } else if (f.tipoProducto === 'pt') {
        dest.ptTotal += f.total;
        dest.ptMerma += f.merma;
      } else {
        dest.sinTipoTotal += f.total;
        dest.sinTipoMerma += f.merma;
      }
      if (f.maquina) dest.maquinasSet.add(f.maquina);
      dest.turnosSet.set(t.key, turnoLabel);
    });
  });

  return lista.map((m) => ({
    nombre: m.nombre,
    sap: m.sap,
    cantidadOrden: m.cantidadOrden,
    ptTotal: m.ptTotal,
    ptMerma: m.ptMerma,
    ppTotal: m.ppTotal,
    ppMerma: m.ppMerma,
    sinTipoTotal: m.sinTipoTotal,
    sinTipoMerma: m.sinTipoMerma,
    maquinas: Array.from(m.maquinasSet),
    turnosTexto: Array.from(m.turnosSet.values()).join(', '),
  }));
}

export function exportarPdfReporteOp(params: {
  op: string;
  catalogMatch: CatalogoItem | null;
  resultado: ConsultaOpResultado;
}): { nombreArchivo: string } {
  const { op, catalogMatch, resultado } = params;
  const { totales, porMaquina, turnosAnteriores, turnoActualResumen } = resultado;

  const todosLosTurnos: TurnoOpResumen[] = [];
  if (turnoActualResumen) todosLosTurnos.push(turnoActualResumen);
  todosLosTurnos.push(...turnosAnteriores);

  const modelos = construirModelosResumen(catalogMatch, todosLosTurnos);

  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4',
  });

  const pageWidth = 297;
  const pageHeight = 210;
  const marginX = 12;
  const contentWidth = pageWidth - marginX * 2; // 273 mm
  const bottomLimit = pageHeight - 14;

  let y = 11;

  const ensureSpace = (needed: number, onNewPage?: () => void) => {
    if (y + needed > bottomLimit) {
      doc.addPage('a4', 'landscape');
      y = 11;
      if (onNewPage) onNewPage();
      return true;
    }
    return false;
  };

  // ================= 1. MAIN BANNER =================
  doc.setFillColor(18, 22, 31); // #12161F
  doc.rect(marginX, y, contentWidth, 18, 'F');
  doc.setFillColor(242, 169, 59); // #F2A93B
  doc.rect(marginX, y, 3.5, 18, 'F');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(13.5);
  doc.setTextColor(255, 255, 255);
  doc.text(`REPORTE DE PRODUCCIÓN — OP ${op}`, marginX + 7, y + 6.8);

  const cantidadOrden = Number(catalogMatch?.cantidadTotal || 0);
  const descBase = catalogMatch
    ? catalogMatch.descripcionBase || catalogMatch.descripcion || ''
    : 'Registro histórico de producción';
  const clienteTxt = catalogMatch?.cliente ? `Cliente: ${catalogMatch.cliente}` : '';
  const ordenTxt = cantidadOrden > 0 ? `Orden: ${cantidadOrden.toLocaleString('es-PE')} u` : '';
  const subParts = [descBase, clienteTxt, ordenTxt].filter(Boolean).join('  ·  ');

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8.2);
  doc.setTextColor(203, 213, 224);
  const subLines = doc.splitTextToSize(subParts, 175) as string[];
  doc.text(subLines[0] || '', marginX + 7, y + 12.8);

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(9.5);
  doc.setTextColor(242, 169, 59);
  doc.text(
    `PT: ${totales.acumPT.toLocaleString('es-PE')} u   |   PP: ${totales.acumPP.toLocaleString('es-PE')} u`,
    marginX + contentWidth - 4,
    y + 7.2,
    { align: 'right' }
  );

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.8);
  doc.setTextColor(203, 213, 224);
  doc.text(
    `${todosLosTurnos.length} turno(s)  ·  ${porMaquina.length} máquina(s)  ·  ${modelos.length} modelo(s)`,
    marginX + contentWidth - 4,
    y + 13.0,
    { align: 'right' }
  );

  y += 21.5;

  // ================= 2. KPI SUMMARY CARDS (PT & PP) =================
  const cardGap = 5;
  const cardW = (contentWidth - cardGap) / 2;
  const cardH = 22;

  // PT Card
  doc.setFillColor(239, 246, 255);
  doc.setDrawColor(147, 197, 253);
  doc.setLineWidth(0.3);
  doc.rect(marginX, y, cardW, cardH, 'FD');
  doc.setFillColor(37, 99, 235);
  doc.rect(marginX, y, 2.2, cardH, 'F');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.setTextColor(30, 64, 175);
  doc.text('PT — PRODUCTO TERMINADO', marginX + 5, y + 5.2);

  const pctPT = cantidadOrden > 0 ? (totales.acumPT / cantidadOrden) * 100 : 0;
  if (cantidadOrden > 0) {
    doc.text(`${pctPT.toFixed(1)}% de orden`, marginX + cardW - 4, y + 5.2, { align: 'right' });
  }

  doc.setFontSize(14);
  doc.setTextColor(17, 24, 39);
  doc.text(`${totales.acumPT.toLocaleString('es-PE')} u`, marginX + 5, y + 12.2);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.5);
  doc.setTextColor(71, 85, 105);
  const ptSub = [
    `Turnos ant.: ${totales.antPT.toLocaleString('es-PE')} u`,
    totales.actualPT > 0 ? `Turno actual: +${totales.actualPT.toLocaleString('es-PE')} u` : '',
    totales.mermaAcumPT > 0 ? `Merma: -${totales.mermaAcumPT.toLocaleString('es-PE')} u` : '',
  ]
    .filter(Boolean)
    .join('  ·  ');
  doc.text(ptSub, marginX + 5, y + 17.0);

  if (cantidadOrden > 0) {
    const difPT = cantidadOrden - totales.acumPT;
    const difTxt =
      difPT > 0
        ? `Faltan ${difPT.toLocaleString('es-PE')} u`
        : difPT < 0
        ? `+${Math.abs(difPT).toLocaleString('es-PE')} u extra`
        : 'Orden completa';
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(difPT <= 0 ? 21 : 30, difPT <= 0 ? 128 : 64, difPT <= 0 ? 61 : 175);
    doc.text(difTxt, marginX + cardW - 4, y + 17.0, { align: 'right' });
  }

  // PP Card
  const ppX = marginX + cardW + cardGap;
  doc.setFillColor(255, 251, 235);
  doc.setDrawColor(252, 211, 77);
  doc.setLineWidth(0.3);
  doc.rect(ppX, y, cardW, cardH, 'FD');
  doc.setFillColor(217, 119, 6);
  doc.rect(ppX, y, 2.2, cardH, 'F');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.setTextColor(180, 83, 9);
  doc.text('PP — PRODUCTO EN PROCESO', ppX + 5, y + 5.2);

  doc.setFontSize(14);
  doc.setTextColor(17, 24, 39);
  doc.text(`${totales.acumPP.toLocaleString('es-PE')} u`, ppX + 5, y + 12.2);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.5);
  doc.setTextColor(71, 85, 105);
  const ppSub = [
    `Turnos ant.: ${totales.antPP.toLocaleString('es-PE')} u`,
    totales.actualPP > 0 ? `Turno actual: +${totales.actualPP.toLocaleString('es-PE')} u` : '',
    totales.mermaAcumPP > 0 ? `Merma: -${totales.mermaAcumPP.toLocaleString('es-PE')} u` : '',
  ]
    .filter(Boolean)
    .join('  ·  ');
  doc.text(ppSub, ppX + 5, y + 17.0);

  doc.setFontSize(6.8);
  doc.setTextColor(120, 113, 108);
  doc.text('Contabilización independiente de proceso / pliegos', ppX + 5, y + 20.5);

  y += cardH + 4;

  if (totales.acumSinTipo > 0) {
    doc.setFillColor(243, 244, 246);
    doc.setDrawColor(209, 213, 219);
    doc.rect(marginX, y, contentWidth, 6, 'FD');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.5);
    doc.setTextColor(55, 65, 81);
    doc.text(
      `REGISTROS SIN CLASIFICAR (PP/PT): ${totales.acumSinTipo.toLocaleString('es-PE')} u`,
      marginX + 3,
      y + 4.1
    );
    y += 8;
  }

  // Helper to draw a section banner
  const drawSectionBanner = (label: string, fillRgb: [number, number, number] = [217, 119, 6]) => {
    doc.setFillColor(fillRgb[0], fillRgb[1], fillRgb[2]);
    doc.setDrawColor(183, 183, 183);
    doc.setLineWidth(0.2);
    doc.rect(marginX, y, contentWidth, 6.2, 'FD');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.3);
    doc.setTextColor(255, 255, 255);
    doc.text(label, marginX + 3, y + 4.3);
    y += 6.2;
  };

  // ================= 3. INFORMACIÓN POR MODELOS =================
  if (modelos.length > 0) {
    const modCols = [68, 21, 24, 23, 15, 23, 17, 40, 42]; // Sum = 273
    const modHeaders = [
      'MODELO / PRODUCTO',
      'SAP',
      'CANT. ORDEN',
      'PROD. PT',
      '% PT',
      'PROD. PP',
      'MERMA',
      'MÁQUINAS',
      'FECHAS / TURNOS',
    ];
    const modColX: number[] = [];
    let mx = marginX;
    modCols.forEach((w) => {
      modColX.push(mx);
      mx += w;
    });

    const drawModelosHeader = (cont = false) => {
      drawSectionBanner(
        cont
          ? `1. INFORMACIÓN Y PRODUCCIÓN POR MODELOS (CONTINUACIÓN)`
          : `1. INFORMACIÓN Y PRODUCCIÓN POR MODELOS (${modelos.length})`,
        [30, 58, 138]
      );
      doc.setFillColor(242, 242, 242);
      doc.setDrawColor(183, 183, 183);
      doc.rect(marginX, y, contentWidth, 6, 'FD');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(6.9);
      doc.setTextColor(51, 51, 51);
      modHeaders.forEach((h, i) => {
        doc.line(modColX[i], y, modColX[i], y + 6);
        doc.text(h, modColX[i] + modCols[i] / 2, y + 4.0, { align: 'center' });
      });
      doc.line(marginX + contentWidth, y, marginX + contentWidth, y + 6);
      y += 6;
    };

    ensureSpace(20);
    drawModelosHeader(false);

    modelos.forEach((mo, idx) => {
      const mermaTot = mo.ptMerma + mo.ppMerma + mo.sinTipoMerma;
      const pctMo = mo.cantidadOrden > 0 ? `${((mo.ptTotal / mo.cantidadOrden) * 100).toFixed(1)}%` : '—';
      const vals = [
        mo.nombre,
        mo.sap || '—',
        mo.cantidadOrden > 0 ? `${mo.cantidadOrden.toLocaleString('es-PE')}` : '—',
        `${mo.ptTotal.toLocaleString('es-PE')}`,
        pctMo,
        `${mo.ppTotal.toLocaleString('es-PE')}`,
        mermaTot > 0 ? `${mermaTot.toLocaleString('es-PE')}` : '0',
        mo.maquinas.length > 0 ? mo.maquinas.join(', ') : 'Sin registros',
        mo.turnosTexto || '—',
      ];

      doc.setFontSize(6.9);
      const wrapped = vals.map((v, cIdx) => doc.splitTextToSize(v, modCols[cIdx] - 2.4) as string[]);
      const maxLines = Math.max(...wrapped.map((l) => l.length), 1);
      const rowH = Math.max(5.5, maxLines * 3.1 + 2.2);

      ensureSpace(rowH, () => drawModelosHeader(true));

      if (idx % 2 === 1) {
        doc.setFillColor(250, 250, 250);
        doc.rect(marginX, y, contentWidth, rowH, 'F');
      }

      doc.setDrawColor(183, 183, 183);
      doc.setLineWidth(0.2);
      doc.rect(marginX, y, contentWidth, rowH, 'S');
      modColX.forEach((cx) => doc.line(cx, y, cx, y + rowH));

      wrapped.forEach((lines, cIdx) => {
        if (cIdx === 0 || cIdx === 3 || cIdx === 5) doc.setFont('helvetica', 'bold');
        else doc.setFont('helvetica', 'normal');

        if (cIdx === 3 && mo.ptTotal > 0) doc.setTextColor(30, 64, 175);
        else if (cIdx === 5 && mo.ppTotal > 0) doc.setTextColor(180, 83, 9);
        else if (cIdx === 6 && mermaTot > 0) doc.setTextColor(192, 57, 43);
        else doc.setTextColor(17, 24, 39);

        let ty = y + (rowH - lines.length * 3.1) / 2 + 2.4;
        lines.forEach((line) => {
          if (cIdx === 1 || cIdx === 4) {
            doc.text(line, modColX[cIdx] + modCols[cIdx] / 2, ty, { align: 'center' });
          } else if (cIdx === 2 || cIdx === 3 || cIdx === 5 || cIdx === 6) {
            doc.text(line, modColX[cIdx] + modCols[cIdx] - 1.5, ty, { align: 'right' });
          } else {
            doc.text(line, modColX[cIdx] + 1.3, ty);
          }
          ty += 3.1;
        });
      });

      y += rowH;
    });

    y += 5;
  }

  // ================= 4. TOTALES CONSOLIDADOS POR MÁQUINA =================
  if (porMaquina.length > 0) {
    const mqCols = [58, 28, 18, 20, 38, 38, 43, 30]; // Sum = 273
    const mqHeaders = [
      'MÁQUINA',
      'CÓDIGO',
      'TIPO',
      'TURNOS',
      'GUARDADO',
      'TURNO ACTUAL',
      'TOTAL PRODUCIDO',
      'MERMA',
    ];
    const mqColX: number[] = [];
    let qx = marginX;
    mqCols.forEach((w) => {
      mqColX.push(qx);
      qx += w;
    });

    const drawMaqHeader = (cont = false) => {
      drawSectionBanner(
        cont
          ? `2. TOTALES CONSOLIDADOS POR MÁQUINA (CONTINUACIÓN)`
          : `2. TOTALES CONSOLIDADOS POR MÁQUINA (${porMaquina.length})`,
        [217, 119, 6]
      );
      doc.setFillColor(242, 242, 242);
      doc.setDrawColor(183, 183, 183);
      doc.rect(marginX, y, contentWidth, 6, 'FD');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7);
      doc.setTextColor(51, 51, 51);
      mqHeaders.forEach((h, i) => {
        doc.line(mqColX[i], y, mqColX[i], y + 6);
        doc.text(h, mqColX[i] + mqCols[i] / 2, y + 4.0, { align: 'center' });
      });
      doc.line(marginX + contentWidth, y, marginX + contentWidth, y + 6);
      y += 6;
    };

    ensureSpace(20);
    drawMaqHeader(false);

    porMaquina.forEach((mq, idx) => {
      const rowH = 5.6;
      ensureSpace(rowH, () => drawMaqHeader(true));

      if (idx % 2 === 1) {
        doc.setFillColor(250, 250, 250);
        doc.rect(marginX, y, contentWidth, rowH, 'F');
      }
      doc.setDrawColor(183, 183, 183);
      doc.setLineWidth(0.2);
      doc.rect(marginX, y, contentWidth, rowH, 'S');
      mqColX.forEach((cx) => doc.line(cx, y, cx, y + rowH));

      const vals = [
        mq.maquina,
        mq.codigoMaquina || '—',
        mq.tipoProducto ? mq.tipoProducto.toUpperCase() : '—',
        `${mq.turnosCount}`,
        `${mq.totalAnterior.toLocaleString('es-PE')} u`,
        mq.totalActual > 0 ? `+${mq.totalActual.toLocaleString('es-PE')} u` : '0 u',
        `${mq.totalAcumulado.toLocaleString('es-PE')} u`,
        mq.mermaAcumulada > 0 ? `${mq.mermaAcumulada.toLocaleString('es-PE')} u` : '0',
      ];

      doc.setFontSize(7.1);
      vals.forEach((v, cIdx) => {
        if (cIdx === 0 || cIdx === 2 || cIdx === 6) doc.setFont('helvetica', 'bold');
        else doc.setFont('helvetica', 'normal');

        if (cIdx === 2 && mq.tipoProducto === 'pt') doc.setTextColor(30, 64, 175);
        else if (cIdx === 2 && mq.tipoProducto === 'pp') doc.setTextColor(180, 83, 9);
        else if (cIdx === 6 && mq.tipoProducto === 'pt') doc.setTextColor(30, 64, 175);
        else if (cIdx === 6 && mq.tipoProducto === 'pp') doc.setTextColor(180, 83, 9);
        else if (cIdx === 7 && mq.mermaAcumulada > 0) doc.setTextColor(192, 57, 43);
        else doc.setTextColor(17, 24, 39);

        if (cIdx === 1 || cIdx === 2 || cIdx === 3) {
          doc.text(v, mqColX[cIdx] + mqCols[cIdx] / 2, y + 3.9, { align: 'center' });
        } else if (cIdx >= 4) {
          doc.text(v, mqColX[cIdx] + mqCols[cIdx] - 1.6, y + 3.9, { align: 'right' });
        } else {
          doc.text(v, mqColX[cIdx] + 1.5, y + 3.9);
        }
      });

      y += rowH;
    });

    // Subtotals separated by PT and PP
    const subtotalRows: { label: string; ant: number; act: number; acum: number; merma: number; color: [number, number, number] }[] = [];
    if (totales.acumPT > 0 || porMaquina.some((m) => m.tipoProducto === 'pt')) {
      subtotalRows.push({
        label: 'TOTAL PT (PRODUCTO TERMINADO)',
        ant: totales.antPT,
        act: totales.actualPT,
        acum: totales.acumPT,
        merma: totales.mermaAcumPT,
        color: [219, 234, 254],
      });
    }
    if (totales.acumPP > 0 || porMaquina.some((m) => m.tipoProducto === 'pp')) {
      subtotalRows.push({
        label: 'TOTAL PP (PRODUCTO EN PROCESO)',
        ant: totales.antPP,
        act: totales.actualPP,
        acum: totales.acumPP,
        merma: totales.mermaAcumPP,
        color: [253, 235, 208],
      });
    }

    subtotalRows.forEach((sr) => {
      ensureSpace(6);
      doc.setFillColor(sr.color[0], sr.color[1], sr.color[2]);
      doc.setDrawColor(183, 183, 183);
      doc.rect(marginX, y, contentWidth, 5.8, 'FD');
      mqColX.slice(4).forEach((cx) => doc.line(cx, y, cx, y + 5.8));

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7.1);
      doc.setTextColor(17, 24, 39);
      doc.text(sr.label, marginX + 2, y + 4.0);
      doc.text(`${sr.ant.toLocaleString('es-PE')} u`, mqColX[4] + mqCols[4] - 1.6, y + 4.0, { align: 'right' });
      doc.text(sr.act > 0 ? `+${sr.act.toLocaleString('es-PE')} u` : '0 u', mqColX[5] + mqCols[5] - 1.6, y + 4.0, { align: 'right' });
      doc.text(`${sr.acum.toLocaleString('es-PE')} u`, mqColX[6] + mqCols[6] - 1.6, y + 4.0, { align: 'right' });
      doc.setTextColor(192, 57, 43);
      doc.text(`${sr.merma.toLocaleString('es-PE')} u`, mqColX[7] + mqCols[7] - 1.6, y + 4.0, { align: 'right' });
      y += 5.8;
    });

    y += 5;
  }

  // ================= 5. DETALLE POR FECHAS, TURNOS Y MÁQUINAS =================
  if (todosLosTurnos.length > 0) {
    const detCols = [21, 18, 32, 18, 12, 20, 72, 24, 18, 38]; // Sum = 273
    const detHeaders = [
      'FECHA',
      'TURNO',
      'MÁQUINA',
      'CÓD.',
      'TIPO',
      'SAP',
      'MODELO / PRODUCTO',
      'CANTIDAD',
      'MERMA',
      'OBSERVACIONES',
    ];
    const detColX: number[] = [];
    let dx = marginX;
    detCols.forEach((w) => {
      detColX.push(dx);
      dx += w;
    });

    const drawDetalleHeader = (cont = false) => {
      drawSectionBanner(
        cont
          ? `3. DETALLE POR FECHAS, TURNOS Y MÁQUINAS (CONTINUACIÓN)`
          : `3. DETALLE POR FECHAS, TURNOS Y MÁQUINAS (${todosLosTurnos.length} TURNO(S))`,
        [18, 22, 31]
      );
      doc.setFillColor(242, 242, 242);
      doc.setDrawColor(183, 183, 183);
      doc.rect(marginX, y, contentWidth, 6, 'FD');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(6.8);
      doc.setTextColor(51, 51, 51);
      detHeaders.forEach((h, i) => {
        doc.line(detColX[i], y, detColX[i], y + 6);
        doc.text(h, detColX[i] + detCols[i] / 2, y + 4.0, { align: 'center' });
      });
      doc.line(marginX + contentWidth, y, marginX + contentWidth, y + 6);
      y += 6;
    };

    ensureSpace(20);
    drawDetalleHeader(false);

    todosLosTurnos.forEach((t) => {
      const fechaTxt = formatearFechaTurno(t.fecha);
      const turnoTxt =
        t.fuente === 'actual'
          ? `${t.tipo === 'noche' ? 'Noche' : 'Día'} (Act.)`
          : t.tipo === 'noche'
          ? 'Noche'
          : t.tipo === 'dia'
          ? 'Día'
          : '—';

      t.filas.forEach((f, idx) => {
        const isFirstOfShift = idx === 0;
        const isLastOfShift = idx === t.filas.length - 1;

        const values = [
          isFirstOfShift ? fechaTxt : '',
          isFirstOfShift ? turnoTxt : '',
          f.maquina,
          f.codigoMaquina || '—',
          f.tipoProducto ? f.tipoProducto.toUpperCase() : '—',
          f.sap || '—',
          f.producto || '—',
          `${f.total.toLocaleString('es-PE')} u`,
          f.merma > 0 ? `${f.merma.toLocaleString('es-PE')}` : '0',
          f.observaciones || '',
        ];

        doc.setFontSize(6.8);
        const wrapped = values.map((v, cIdx) => {
          if (!v) return [''];
          return doc.splitTextToSize(v, detCols[cIdx] - 2.2) as string[];
        });
        const maxLines = Math.max(...wrapped.map((l) => l.length), 1);
        const rowH = Math.max(5.3, maxLines * 3.0 + 2.0);

        const broke = ensureSpace(rowH, () => {
          drawDetalleHeader(true);
          wrapped[0] = [fechaTxt];
          wrapped[1] = [turnoTxt];
        });
        void broke;

        if (t.fuente === 'actual') {
          doc.setFillColor(240, 253, 244);
          doc.rect(marginX, y, contentWidth, rowH, 'F');
        } else if (idx % 2 === 1) {
          doc.setFillColor(250, 250, 250);
          doc.rect(marginX, y, contentWidth, rowH, 'F');
        }

        doc.setDrawColor(183, 183, 183);
        doc.setLineWidth(0.2);
        detColX.forEach((cx) => doc.line(cx, y, cx, y + rowH));
        doc.line(marginX + contentWidth, y, marginX + contentWidth, y + rowH);

        if (isLastOfShift) {
          doc.line(marginX, y + rowH, marginX + contentWidth, y + rowH);
        } else {
          doc.line(detColX[2], y + rowH, marginX + contentWidth, y + rowH);
        }

        wrapped.forEach((lines, cIdx) => {
          if (lines.length === 1 && lines[0] === '') return;

          if (cIdx === 0 || cIdx === 2 || cIdx === 4 || cIdx === 7) {
            doc.setFont('helvetica', 'bold');
          } else {
            doc.setFont('helvetica', 'normal');
          }

          if ((cIdx === 4 || cIdx === 7) && f.tipoProducto === 'pt') doc.setTextColor(30, 64, 175);
          else if ((cIdx === 4 || cIdx === 7) && f.tipoProducto === 'pp') doc.setTextColor(180, 83, 9);
          else if (cIdx === 8 && f.merma > 0) doc.setTextColor(192, 57, 43);
          else doc.setTextColor(17, 24, 39);

          let ty = y + (rowH - lines.length * 3.0) / 2 + 2.3;
          lines.forEach((line) => {
            if (cIdx === 0 || cIdx === 1 || cIdx === 3 || cIdx === 4 || cIdx === 5) {
              doc.text(line, detColX[cIdx] + detCols[cIdx] / 2, ty, { align: 'center' });
            } else if (cIdx === 7 || cIdx === 8) {
              doc.text(line, detColX[cIdx] + detCols[cIdx] - 1.4, ty, { align: 'right' });
            } else {
              doc.text(line, detColX[cIdx] + 1.2, ty);
            }
            ty += 3.0;
          });
        });

        y += rowH;
      });

      // Subtotal bar per shift/date
      ensureSpace(5.6, () => drawDetalleHeader(true));
      doc.setFillColor(253, 235, 208);
      doc.setDrawColor(183, 183, 183);
      doc.rect(marginX, y, contentWidth, 5.2, 'FD');

      const partsShift: string[] = [];
      if (t.totalPT > 0 || t.filas.some((f) => f.tipoProducto === 'pt')) {
        partsShift.push(`PT: ${t.totalPT.toLocaleString('es-PE')} u`);
      }
      if (t.totalPP > 0 || t.filas.some((f) => f.tipoProducto === 'pp')) {
        partsShift.push(`PP: ${t.totalPP.toLocaleString('es-PE')} u`);
      }
      if (t.totalSinTipo > 0) {
        partsShift.push(`S/T: ${t.totalSinTipo.toLocaleString('es-PE')} u`);
      }
      const mermaTurnoTot = t.mermaPT + t.mermaPP + t.mermaSinTipo;
      if (mermaTurnoTot > 0) {
        partsShift.push(`Merma: -${mermaTurnoTot.toLocaleString('es-PE')} u`);
      }

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(6.8);
      doc.setTextColor(55, 65, 81);
      doc.text(`Subtotal ${fechaTxt} · ${turnoTxt}:`, marginX + 2.5, y + 3.6);
      doc.setTextColor(17, 24, 39);
      doc.text(partsShift.join('   |   '), marginX + contentWidth - 2.5, y + 3.6, { align: 'right' });
      y += 5.2;
    });
  }

  // ================= 6. FOOTER ON ALL PAGES =================
  const totalPages = (doc as any).internal.getNumberOfPages();
  const generado = new Date().toLocaleString('es-PE', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });

  for (let p = 1; p <= totalPages; p++) {
    doc.setPage(p);
    doc.setDrawColor(203, 213, 224);
    doc.setLineWidth(0.2);
    doc.line(marginX, pageHeight - 10, marginX + contentWidth, pageHeight - 10);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(100, 116, 139);
    doc.text(
      `Generado: ${generado}  ·  Reporte de Producción OP ${op}`,
      marginX,
      pageHeight - 6.2
    );
    doc.text(`Página ${p} de ${totalPages}`, marginX + contentWidth, pageHeight - 6.2, {
      align: 'right',
    });
  }

  const safeOp = op.replace(/[^a-zA-Z0-9_-]/g, '_');
  const hoy = new Date().toISOString().slice(0, 10);
  const nombreArchivo = `reporte_OP_${safeOp}_${hoy}.pdf`;
  doc.save(nombreArchivo);

  return { nombreArchivo };
}
