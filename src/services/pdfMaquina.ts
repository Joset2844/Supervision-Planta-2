import { jsPDF } from 'jspdf';
import { Maquina } from '../types';
import { formatearFechaTurno } from './reports';
import {
  ConsultaMaquinaResultado,
  MaquinaOpConsolidado,
  TurnoMaquinaResumen,
} from './opHistory';

export interface ExportarPdfMaquinaParams {
  maquina: Maquina;
  turnosFiltrados: TurnoMaquinaResumen[];
  porOpFiltrado: MaquinaOpConsolidado[];
  resumenFiltrado: {
    totalAnterior: number;
    mermaAnterior: number;
    totalActual: number;
    mermaActual: number;
    totalProducido: number;
    totalMerma: number;
    promedioPorTurno: number;
    turnosDia: number;
    turnosNoche: number;
  };
  fechaDesde?: string;
  fechaHasta?: string;
  tipoTurnoFiltro?: '' | 'dia' | 'noche';
  resultadoOriginal?: ConsultaMaquinaResultado;
}

export function exportarPdfReporteMaquina(params: ExportarPdfMaquinaParams): {
  nombreArchivo: string;
} {
  const {
    maquina,
    turnosFiltrados,
    porOpFiltrado,
    resumenFiltrado,
    fechaDesde,
    fechaHasta,
    tipoTurnoFiltro,
  } = params;

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

  const areaTxt =
    maquina.area === 'vasos_envases'
      ? 'Vasos y Envases'
      : maquina.area === 'papel_carton'
      ? 'Papel y Cartón'
      : '';
  const tipoTxt = maquina.tipoProducto ? maquina.tipoProducto.toUpperCase() : '';
  const meta = Number(maquina.meta) || 0;

  let rangoTxt = 'Todo el historial';
  if (fechaDesde && fechaHasta) {
    rangoTxt = `Del ${formatearFechaTurno(fechaDesde)} al ${formatearFechaTurno(fechaHasta)}`;
  } else if (fechaDesde) {
    rangoTxt = `Desde ${formatearFechaTurno(fechaDesde)}`;
  } else if (fechaHasta) {
    rangoTxt = `Hasta ${formatearFechaTurno(fechaHasta)}`;
  }

  const turnoFiltroTxt =
    tipoTurnoFiltro === 'dia'
      ? 'Solo Turno Día'
      : tipoTurnoFiltro === 'noche'
      ? 'Solo Turno Noche'
      : 'Todos los turnos';

  // ================= 1. MAIN BANNER =================
  doc.setFillColor(18, 22, 31); // #12161F
  doc.rect(marginX, y, contentWidth, 18, 'F');
  doc.setFillColor(242, 169, 59); // #F2A93B
  doc.rect(marginX, y, 3.5, 18, 'F');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(13.5);
  doc.setTextColor(255, 255, 255);
  doc.text(`REPORTE DE PRODUCCIÓN — MÁQUINA ${maquina.nombre}`, marginX + 7, y + 6.8);

  const subParts = [
    maquina.codigo ? `Código: ${maquina.codigo}` : '',
    tipoTxt ? `Tipo: ${tipoTxt}` : '',
    areaTxt ? `Área: ${areaTxt}` : '',
    meta > 0 ? `Meta/turno: ${meta.toLocaleString('es-PE')} u` : '',
    `${rangoTxt} (${turnoFiltroTxt})`,
  ]
    .filter(Boolean)
    .join('  ·  ');

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8.0);
  doc.setTextColor(203, 213, 224);
  const subLines = doc.splitTextToSize(subParts, 185) as string[];
  doc.text(subLines[0] || '', marginX + 7, y + 12.8);

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(9.5);
  doc.setTextColor(242, 169, 59);
  doc.text(
    `Total: ${resumenFiltrado.totalProducido.toLocaleString('es-PE')} u`,
    marginX + contentWidth - 4,
    y + 7.2,
    { align: 'right' }
  );

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.8);
  doc.setTextColor(203, 213, 224);
  doc.text(
    `${turnosFiltrados.length} turno(s)  ·  ${porOpFiltrado.length} OP/reg.  ·  Prom: ${resumenFiltrado.promedioPorTurno.toLocaleString('es-PE')} u/turno`,
    marginX + contentWidth - 4,
    y + 13.0,
    { align: 'right' }
  );

  y += 21.5;

  // ================= 2. KPI SUMMARY CARDS (4 CARDS) =================
  const gap = 4;
  const cardW = (contentWidth - gap * 3) / 4;
  const cardH = 19;

  const pctMetaProm =
    meta > 0 && turnosFiltrados.length > 0
      ? (resumenFiltrado.promedioPorTurno / meta) * 100
      : 0;
  const turnosMetaOk =
    meta > 0 ? turnosFiltrados.filter((t) => t.totalProducido >= meta).length : 0;

  const distinctOpsCount = new Set(
    porOpFiltrado.map((o) => (o.op || '').trim().toLowerCase()).filter(Boolean)
  ).size;

  const kpiCards = [
    {
      label: 'TOTAL PRODUCIDO',
      value: `${resumenFiltrado.totalProducido.toLocaleString('es-PE')} u`,
      sub: [
        `Ant: ${resumenFiltrado.totalAnterior.toLocaleString('es-PE')}`,
        resumenFiltrado.totalActual > 0
          ? `Act: +${resumenFiltrado.totalActual.toLocaleString('es-PE')}`
          : '',
        resumenFiltrado.totalMerma > 0
          ? `Merma: -${resumenFiltrado.totalMerma.toLocaleString('es-PE')}`
          : '',
      ]
        .filter(Boolean)
        .join(' · '),
      accent: [217, 119, 6] as [number, number, number],
      bg: [255, 251, 235] as [number, number, number],
      border: [252, 211, 77] as [number, number, number],
    },
    {
      label: 'TURNOS REGISTRADOS',
      value: `${turnosFiltrados.length}`,
      sub: `Día: ${resumenFiltrado.turnosDia}  ·  Noche: ${resumenFiltrado.turnosNoche}`,
      accent: [37, 99, 235] as [number, number, number],
      bg: [239, 246, 255] as [number, number, number],
      border: [147, 197, 253] as [number, number, number],
    },
    {
      label: 'PROMEDIO / TURNO',
      value: `${resumenFiltrado.promedioPorTurno.toLocaleString('es-PE')} u`,
      sub:
        meta > 0 && turnosFiltrados.length > 0
          ? `${pctMetaProm.toFixed(1)}% meta · ${turnosMetaOk}/${turnosFiltrados.length} en meta`
          : meta > 0
          ? `Meta: ${meta.toLocaleString('es-PE')} u/turno`
          : 'Sin meta asignada',
      accent: [21, 128, 61] as [number, number, number],
      bg: [240, 253, 244] as [number, number, number],
      border: [134, 239, 172] as [number, number, number],
    },
    {
      label: 'OPs TRABAJADAS',
      value: `${distinctOpsCount || porOpFiltrado.length}`,
      sub: `${porOpFiltrado.length} ítem(s) consolidado(s)`,
      accent: [71, 85, 105] as [number, number, number],
      bg: [248, 250, 252] as [number, number, number],
      border: [203, 213, 225] as [number, number, number],
    },
  ];

  kpiCards.forEach((c, idx) => {
    const cx = marginX + idx * (cardW + gap);
    doc.setFillColor(c.bg[0], c.bg[1], c.bg[2]);
    doc.setDrawColor(c.border[0], c.border[1], c.border[2]);
    doc.setLineWidth(0.28);
    doc.rect(cx, y, cardW, cardH, 'FD');
    doc.setFillColor(c.accent[0], c.accent[1], c.accent[2]);
    doc.rect(cx, y, 2.0, cardH, 'F');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.2);
    doc.setTextColor(c.accent[0], c.accent[1], c.accent[2]);
    doc.text(c.label, cx + 4.2, y + 4.8);

    doc.setFontSize(12.5);
    doc.setTextColor(17, 24, 39);
    doc.text(c.value, cx + 4.2, y + 11.2);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(6.6);
    doc.setTextColor(71, 85, 105);
    const lines = doc.splitTextToSize(c.sub, cardW - 6) as string[];
    doc.text(lines[0] || '', cx + 4.2, y + 15.8);
  });

  y += cardH + 5;

  const drawSectionBanner = (label: string, fillRgb: [number, number, number]) => {
    doc.setFillColor(fillRgb[0], fillRgb[1], fillRgb[2]);
    doc.setDrawColor(183, 183, 183);
    doc.setLineWidth(0.2);
    doc.rect(marginX, y, contentWidth, 6.2, 'FD');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.2);
    doc.setTextColor(255, 255, 255);
    doc.text(label, marginX + 3, y + 4.3);
    y += 6.2;
  };

  // ================= 3. CONSOLIDADO POR OP / PRODUCTO =================
  if (porOpFiltrado.length > 0) {
    const opCols = [22, 22, 74, 42, 16, 26, 24, 27, 20]; // Sum = 273
    const opHeaders = [
      'OP',
      'SAP',
      'MODELO / PRODUCTO',
      'CLIENTE',
      'TURNOS',
      'GUARDADO',
      'TURNO ACT.',
      'TOTAL PROD.',
      'MERMA',
    ];
    const opColX: number[] = [];
    let ox = marginX;
    opCols.forEach((w) => {
      opColX.push(ox);
      ox += w;
    });

    const drawOpHeader = (cont = false) => {
      drawSectionBanner(
        cont
          ? `1. CONSOLIDADO POR OP / PRODUCTO (CONTINUACIÓN)`
          : `1. CONSOLIDADO POR OP / PRODUCTO (${porOpFiltrado.length})`,
        [30, 58, 138]
      );
      doc.setFillColor(242, 242, 242);
      doc.setDrawColor(183, 183, 183);
      doc.rect(marginX, y, contentWidth, 6, 'FD');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(6.8);
      doc.setTextColor(51, 51, 51);
      opHeaders.forEach((h, i) => {
        doc.line(opColX[i], y, opColX[i], y + 6);
        doc.text(h, opColX[i] + opCols[i] / 2, y + 4.0, { align: 'center' });
      });
      doc.line(marginX + contentWidth, y, marginX + contentWidth, y + 6);
      y += 6;
    };

    ensureSpace(20);
    drawOpHeader(false);

    porOpFiltrado.forEach((item, idx) => {
      const vals = [
        item.op || 'Sin OP',
        item.sap || '—',
        item.producto || '—',
        item.cliente || '—',
        `${item.turnosCount}`,
        `${item.totalAnterior.toLocaleString('es-PE')} u`,
        item.totalActual > 0 ? `+${item.totalActual.toLocaleString('es-PE')} u` : '0 u',
        `${item.totalAcumulado.toLocaleString('es-PE')} u`,
        item.mermaAcumulada > 0 ? `${item.mermaAcumulada.toLocaleString('es-PE')}` : '0',
      ];

      doc.setFontSize(6.8);
      const wrapped = vals.map((v, cIdx) => doc.splitTextToSize(v, opCols[cIdx] - 2.2) as string[]);
      const maxLines = Math.max(...wrapped.map((l) => l.length), 1);
      const rowH = Math.max(5.4, maxLines * 3.0 + 2.1);

      ensureSpace(rowH, () => drawOpHeader(true));

      if (idx % 2 === 1) {
        doc.setFillColor(250, 250, 250);
        doc.rect(marginX, y, contentWidth, rowH, 'F');
      }
      doc.setDrawColor(183, 183, 183);
      doc.setLineWidth(0.2);
      doc.rect(marginX, y, contentWidth, rowH, 'S');
      opColX.forEach((cx) => doc.line(cx, y, cx, y + rowH));

      wrapped.forEach((lines, cIdx) => {
        if (cIdx === 0 || cIdx === 7) doc.setFont('helvetica', 'bold');
        else doc.setFont('helvetica', 'normal');

        if (cIdx === 0 && item.op) doc.setTextColor(180, 83, 9);
        else if (cIdx === 7) doc.setTextColor(30, 64, 175);
        else if (cIdx === 8 && item.mermaAcumulada > 0) doc.setTextColor(192, 57, 43);
        else doc.setTextColor(17, 24, 39);

        let ty = y + (rowH - lines.length * 3.0) / 2 + 2.3;
        lines.forEach((line) => {
          if (cIdx === 0 || cIdx === 1 || cIdx === 4) {
            doc.text(line, opColX[cIdx] + opCols[cIdx] / 2, ty, { align: 'center' });
          } else if (cIdx >= 5) {
            doc.text(line, opColX[cIdx] + opCols[cIdx] - 1.5, ty, { align: 'right' });
          } else {
            doc.text(line, opColX[cIdx] + 1.3, ty);
          }
          ty += 3.0;
        });
      });

      y += rowH;
    });

    // Total Row
    ensureSpace(6.2);
    doc.setFillColor(253, 235, 208);
    doc.setDrawColor(183, 183, 183);
    doc.rect(marginX, y, contentWidth, 5.8, 'FD');
    opColX.slice(5).forEach((cx) => doc.line(cx, y, cx, y + 5.8));

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.1);
    doc.setTextColor(17, 24, 39);
    doc.text('TOTAL CONSOLIDADO DEL PERÍODO', marginX + 2.5, y + 4.0);
    doc.text(
      `${resumenFiltrado.totalAnterior.toLocaleString('es-PE')} u`,
      opColX[5] + opCols[5] - 1.5,
      y + 4.0,
      { align: 'right' }
    );
    doc.text(
      resumenFiltrado.totalActual > 0
        ? `+${resumenFiltrado.totalActual.toLocaleString('es-PE')} u`
        : '0 u',
      opColX[6] + opCols[6] - 1.5,
      y + 4.0,
      { align: 'right' }
    );
    doc.text(
      `${resumenFiltrado.totalProducido.toLocaleString('es-PE')} u`,
      opColX[7] + opCols[7] - 1.5,
      y + 4.0,
      { align: 'right' }
    );
    doc.setTextColor(192, 57, 43);
    doc.text(
      `${resumenFiltrado.totalMerma.toLocaleString('es-PE')}`,
      opColX[8] + opCols[8] - 1.5,
      y + 4.0,
      { align: 'right' }
    );
    y += 10.5;
  }

  // ================= 4. DETALLE POR FECHAS Y TURNOS =================
  if (turnosFiltrados.length > 0) {
    const detCols = [22, 20, 22, 22, 84, 26, 19, 18, 40]; // Sum = 273
    const detHeaders = [
      'FECHA',
      'TURNO',
      'OP',
      'SAP',
      'MODELO / PRODUCTO',
      'CANTIDAD',
      'MERMA',
      '% META',
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
          ? `2. DETALLE CRONOLÓGICO POR FECHAS Y TURNOS (CONTINUACIÓN)`
          : `2. DETALLE CRONOLÓGICO POR FECHAS Y TURNOS (${turnosFiltrados.length} TURNO(S))`,
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

    turnosFiltrados.forEach((t) => {
      const fechaTxt = formatearFechaTurno(t.fecha);
      const turnoTxt =
        t.fuente === 'actual'
          ? `${t.tipo === 'noche' ? 'Noche' : 'Día'} (Act.)`
          : t.tipo === 'noche'
          ? 'Noche'
          : t.tipo === 'dia'
          ? 'Día'
          : '—';
      const pctTurnoMeta = meta > 0 ? `${((t.totalProducido / meta) * 100).toFixed(0)}%` : '—';

      t.filas.forEach((f, idx) => {
        const isFirst = idx === 0;
        const isLast = idx === t.filas.length - 1;

        const values = [
          isFirst ? fechaTxt : '',
          isFirst ? turnoTxt : '',
          f.op || 'Sin OP',
          f.sap || '—',
          f.producto || '—',
          `${f.total.toLocaleString('es-PE')} u`,
          f.merma > 0 ? `${f.merma.toLocaleString('es-PE')}` : '0',
          isFirst ? pctTurnoMeta : '',
          f.observaciones || '',
        ];

        doc.setFontSize(6.8);
        const wrapped = values.map((v, cIdx) => {
          if (!v) return [''];
          return doc.splitTextToSize(v, detCols[cIdx] - 2.2) as string[];
        });
        const maxLines = Math.max(...wrapped.map((l) => l.length), 1);
        const rowH = Math.max(5.3, maxLines * 3.0 + 2.0);

        ensureSpace(rowH, () => {
          drawDetalleHeader(true);
          wrapped[0] = [fechaTxt];
          wrapped[1] = [turnoTxt];
          wrapped[7] = [pctTurnoMeta];
        });

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

        if (isLast) {
          doc.line(marginX, y + rowH, marginX + contentWidth, y + rowH);
        } else {
          doc.line(detColX[2], y + rowH, detColX[7], y + rowH);
          doc.line(detColX[8], y + rowH, marginX + contentWidth, y + rowH);
        }

        wrapped.forEach((lines, cIdx) => {
          if (lines.length === 1 && lines[0] === '') return;

          if (cIdx === 0 || cIdx === 2 || cIdx === 5 || cIdx === 7) {
            doc.setFont('helvetica', 'bold');
          } else {
            doc.setFont('helvetica', 'normal');
          }

          if (cIdx === 2 && f.op) doc.setTextColor(180, 83, 9);
          else if (cIdx === 5) doc.setTextColor(21, 128, 61);
          else if (cIdx === 6 && f.merma > 0) doc.setTextColor(192, 57, 43);
          else if (cIdx === 7 && meta > 0) {
            const p = (t.totalProducido / meta) * 100;
            doc.setTextColor(p >= 100 ? 21 : p >= 80 ? 180 : 192, p >= 100 ? 128 : p >= 80 ? 83 : 57, p >= 100 ? 61 : p >= 80 ? 9 : 43);
          } else {
            doc.setTextColor(17, 24, 39);
          }

          let ty = y + (rowH - lines.length * 3.0) / 2 + 2.3;
          lines.forEach((line) => {
            if (cIdx <= 3 || cIdx === 7) {
              doc.text(line, detColX[cIdx] + detCols[cIdx] / 2, ty, { align: 'center' });
            } else if (cIdx === 5 || cIdx === 6) {
              doc.text(line, detColX[cIdx] + detCols[cIdx] - 1.4, ty, { align: 'right' });
            } else {
              doc.text(line, detColX[cIdx] + 1.2, ty);
            }
            ty += 3.0;
          });
        });

        y += rowH;
      });

      // Subtotal bar per shift
      ensureSpace(5.6, () => drawDetalleHeader(true));
      doc.setFillColor(253, 235, 208);
      doc.setDrawColor(183, 183, 183);
      doc.rect(marginX, y, contentWidth, 5.2, 'FD');

      const partsShift = [`Producido: ${t.totalProducido.toLocaleString('es-PE')} u`];
      if (t.totalMerma > 0) {
        partsShift.push(`Merma: -${t.totalMerma.toLocaleString('es-PE')} u`);
      }
      if (meta > 0) {
        partsShift.push(`Cumplimiento meta: ${((t.totalProducido / meta) * 100).toFixed(1)}%`);
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

  // ================= 5. FOOTER ON ALL PAGES =================
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
      `Generado: ${generado}  ·  Reporte de Máquina ${maquina.nombre}`,
      marginX,
      pageHeight - 6.2
    );
    doc.text(`Página ${p} de ${totalPages}`, marginX + contentWidth, pageHeight - 6.2, {
      align: 'right',
    });
  }

  const safeMaq = (maquina.nombre || 'maquina').replace(/[^a-zA-Z0-9_-]/g, '_');
  const hoy = new Date().toISOString().slice(0, 10);
  const nombreArchivo = `reporte_MAQUINA_${safeMaq}_${hoy}.pdf`;
  doc.save(nombreArchivo);

  return { nombreArchivo };
}
