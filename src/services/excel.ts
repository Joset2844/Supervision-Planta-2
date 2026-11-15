import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { jsPDF } from 'jspdf';
import { CatalogoItem, CatalogoModelo, FilaTurnoExcel, ListaItem, Maquina, TurnoMaquina } from '../types';

export const AREAS_ORDEN = [
  { key: 'vasos_envases', label: 'VASOS Y ENVASES' },
  { key: 'papel_carton', label: 'PAPEL Y CARTON' },
  { key: '', label: 'SIN AREA ASIGNADA' },
];

export const ACT_HEADERS = [
  'MAQUINA',
  'CODIGO',
  'ESTADO',
  'MOTIVO DE PARADA',
  'OP',
  'PRODUCTO',
  'PERSONAL EN MAQUINA',
  'OBSERVACIONES',
];

export function normalizar(h: any): string {
  return String(h || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

export function escapeHtml(s: any): string {
  return String(s || '').replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] || c;
  });
}

export function etiquetaPliego(p: any): string {
  if (p === '' || p === null || typeof p === 'undefined') return '';
  if (p === 'Tira' || p === 'Retira') return p;
  return 'Pliego ' + p;
}

export function buscarEnCatalogo(catalogo: CatalogoItem[], op: string): CatalogoItem | null {
  if (!op) return null;
  const q = op.trim().toLowerCase();
  if (!q) return null;
  const exact = catalogo.find((c) => c.op.toLowerCase() === q);
  if (exact) return exact;
  const digits = q.replace(/\D/g, '');
  if (digits) {
    const byDigits = catalogo.find((c) => c.op.replace(/\D/g, '') === digits);
    if (byDigits) return byDigits;
  }
  return null;
}

export function resolverModelo(match: CatalogoItem | null, modeloIdx: number): { descripcion: string; sap: string } {
  if (!match || !match.modelos || match.modelos.length === 0) {
    return { descripcion: match ? match.descripcion || '' : '', sap: match ? match.sap || '' : '' };
  }
  if (match.modelos.length === 1) return { descripcion: match.modelos[0].descripcion, sap: match.modelos[0].sap || '' };
  if (typeof modeloIdx === 'number' && modeloIdx >= 0 && match.modelos[modeloIdx]) {
    return { descripcion: match.modelos[modeloIdx].descripcion, sap: match.modelos[modeloIdx].sap || '' };
  }
  return { descripcion: match.descripcionBase || match.modelos[0].descripcion || '', sap: '' };
}

export function descripcionCorta(match: CatalogoItem | null): string {
  if (!match || !match.modelos || match.modelos.length === 0) return match ? match.descripcion || '' : '';
  if (match.modelos.length === 1) return match.modelos[0].descripcion;
  return match.modelos.length + ' modelos disponibles';
}

// ----------------- PARSING CATALOGO -----------------
export function parsearCantidad(raw: any): number {
  if (raw === null || raw === undefined || raw === '') return 0;
  if (typeof raw === 'number') {
    return isNaN(raw) ? 0 : Math.round(raw);
  }
  const s = String(raw).trim();
  if (!s) return 0;
  // Clean spaces and thousands separators (e.g. 15,000 or 15.000 or 15 000)
  let clean = s.replace(/[^0-9.,-]/g, '').trim();
  if (!clean) return 0;

  if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(clean)) {
    clean = clean.replace(/,/g, '');
  } else if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(clean)) {
    clean = clean.replace(/\./g, '').replace(/,/g, '.');
  } else if (clean.includes(',') && !clean.includes('.')) {
    clean = clean.replace(/,/g, '');
  }

  const n = parseFloat(clean);
  return isNaN(n) ? 0 : Math.round(n);
}

export function esColumnaCantidad(h: string): boolean {
  if (!h) return false;
  const s = h.toLowerCase().trim();
  return (
    s.includes('cantidad') ||
    s === 'cant' ||
    s.startsWith('cant') ||
    s.includes('cant.') ||
    s.includes('cant ') ||
    s.includes('total') ||
    s.includes('unidade') ||
    s === 'und' ||
    s === 'unds' ||
    s === 'unid' ||
    s.includes('pedido') ||
    s.includes('programad') ||
    s === 'meta' ||
    s === 'objetivo' ||
    s === 'ctd' ||
    s === 'ctdad'
  );
}

function buscarBase(mapa: Record<string, any>, op: string) {
  if (mapa[op]) return mapa[op];
  const k = Object.keys(mapa).find((x) => x.trim().toLowerCase() === op.trim().toLowerCase());
  return k ? mapa[k] : null;
}

function parsearHojaCatalogo(sheet: any) {
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (rows.length < 2) return null;
  const headers = rows[0].map(normalizar);
  const idxOp = headers.findIndex((h) => h === 'op' || h.indexOf('op') !== -1);
  const idxDesc = headers.findIndex((h) => h.indexOf('descripcion') !== -1 || h.indexOf('producto') !== -1 || h.indexOf('articulo') !== -1);
  const idxCliente = headers.findIndex((h) => h.indexOf('cliente') !== -1 || h.indexOf('razon') !== -1);
  const idxCantidad = headers.findIndex(esColumnaCantidad);
  const idxSap = headers.findIndex((h) => h.indexOf('sap') !== -1 || h.indexOf('codigo') !== -1);
  if (idxOp === -1 || idxDesc === -1 || idxCliente === -1) return null;

  const filasValidas = rows.slice(1).filter((r) => String(r[idxOp]).trim() !== '');
  const mapaOps: Record<string, any> = {};
  const ordenOps: string[] = [];

  filasValidas.forEach((r) => {
    const op = String(r[idxOp]).trim();
    const cliente = String(r[idxCliente]).trim();
    const cantNum = idxCantidad !== -1 ? parsearCantidad(r[idxCantidad]) : 0;
    const desc = String(r[idxDesc]).trim();
    const sap = idxSap !== -1 ? String(r[idxSap]).trim() : '';
    const modelo: CatalogoModelo = {
      descripcion: desc,
      sap: sap,
      cantidadTotal: cantNum,
    };
    if (!mapaOps[op]) {
      mapaOps[op] = { op, cliente, cantidadTotal: 0, modelos: [] };
      ordenOps.push(op);
    }
    const yaExiste = mapaOps[op].modelos.some((mo: any) => {
      return mo.descripcion === modelo.descripcion && (mo.sap || '') === (modelo.sap || '');
    });
    if (!yaExiste) {
      mapaOps[op].modelos.push(modelo);
    } else {
      const idxExistente = mapaOps[op].modelos.findIndex(
        (mo: any) => mo.descripcion === modelo.descripcion && (mo.sap || '') === (modelo.sap || '')
      );
      if (idxExistente !== -1 && cantNum > 0) {
        const cantPrev = parsearCantidad(mapaOps[op].modelos[idxExistente].cantidadTotal);
        mapaOps[op].modelos[idxExistente].cantidadTotal = cantPrev + cantNum;
      }
    }
    mapaOps[op].cantidadTotal += cantNum;
  });
  return { mapaOps, ordenOps };
}

function parsearHojaModelos(sheet: any, mapaBase: Record<string, any>) {
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (rows.length < 2) return null;
  const headers = rows[0].map(normalizar);
  const idxOp = headers.findIndex((h) => h === 'op' || h.indexOf('op') !== -1);
  const idxModelo = headers.findIndex((h) => h.indexOf('modelo') !== -1 || h.indexOf('descripcion') !== -1 || h.indexOf('producto') !== -1);
  const idxCantidad = headers.findIndex(esColumnaCantidad);
  const idxSap = headers.findIndex((h) => h.indexOf('sap') !== -1 || h.indexOf('codigo') !== -1);
  if (idxOp === -1 || idxModelo === -1) return null;

  const filasValidas = rows.slice(1).filter((r) => String(r[idxOp]).trim() !== '');
  const mapaOps: Record<string, any> = {};
  const ordenOps: string[] = [];

  filasValidas.forEach((r) => {
    const op = String(r[idxOp]).trim();
    const modeloTxt = String(r[idxModelo]).trim();
    const cantNum = idxCantidad !== -1 ? parsearCantidad(r[idxCantidad]) : 0;
    const base = buscarBase(mapaBase, op);
    const descBase = base ? (base.modelos && base.modelos[0] ? base.modelos[0].descripcion : (base.descripcion || '')) : '';
    const descripcionFinal =
      descBase && !modeloTxt.toLowerCase().includes(descBase.toLowerCase())
        ? `${descBase} - ${modeloTxt}`
        : modeloTxt;
    const modelo: CatalogoModelo = {
      modelo: modeloTxt,
      descripcion: descripcionFinal,
      sap: idxSap !== -1 ? String(r[idxSap]).trim() : '',
      cantidadTotal: cantNum,
    };
    if (!mapaOps[op]) {
      mapaOps[op] = { op, cliente: base ? base.cliente : '', cantidadTotal: 0, modelos: [] };
      ordenOps.push(op);
    }
    const yaExisteModelo = mapaOps[op].modelos.some((mo: any) => {
      return (mo.modelo === modeloTxt || mo.descripcion === descripcionFinal) && mo.sap === modelo.sap;
    });
    if (!yaExisteModelo) {
      mapaOps[op].modelos.push(modelo);
    } else {
      const idxM = mapaOps[op].modelos.findIndex(
        (mo: any) => (mo.modelo === modeloTxt || mo.descripcion === descripcionFinal) && mo.sap === modelo.sap
      );
      if (idxM !== -1 && cantNum > 0) {
        const cantPrev = parsearCantidad(mapaOps[op].modelos[idxM].cantidadTotal);
        mapaOps[op].modelos[idxM].cantidadTotal = cantPrev + cantNum;
      }
    }
    mapaOps[op].cantidadTotal += cantNum;
  });
  return { mapaOps, ordenOps };
}

export function construirCatalogo(sheetPrincipal: any, sheetModelos: any): CatalogoItem[] | null {
  const principal = parsearHojaCatalogo(sheetPrincipal);
  if (!principal) return null;

  const mapaFinal = principal.mapaOps;
  const ordenFinal = principal.ordenOps.slice();

  if (sheetModelos) {
    const modelos = parsearHojaModelos(sheetModelos, principal.mapaOps);
    if (modelos) {
      modelos.ordenOps.forEach((op) => {
        const base = buscarBase(principal.mapaOps, op);
        const modelosData = modelos.mapaOps[op];

        if (!mapaFinal[op]) {
          ordenFinal.push(op);
          mapaFinal[op] = modelosData;
        } else {
          const descBase = base && base.modelos[0] ? base.modelos[0].descripcion : (base?.descripcion || '');
          // If models sheet has quantities, use sum of models, otherwise preserve base quantity
          const cantModelos = modelosData.cantidadTotal || 0;
          const cantBase = base ? (base.cantidadTotal || 0) : 0;
          const cantFinal = cantModelos > 0 ? cantModelos : cantBase;

          mapaFinal[op] = {
            ...modelosData,
            cliente: base ? base.cliente : modelosData.cliente,
            cantidadTotal: cantFinal,
            descripcionBase: descBase,
          };
        }
      });
    }
  }

  const nuevo: CatalogoItem[] = ordenFinal.map((op) => {
    const item = mapaFinal[op];
    const tieneMultiples = item.modelos && item.modelos.length > 1;

    // Calculate total from models sum if available, or item.cantidadTotal
    let totalOp = parsearCantidad(item.cantidadTotal);
    if (item.modelos && item.modelos.length > 0) {
      const sumModelos = item.modelos.reduce(
        (acc: number, m: any) => acc + parsearCantidad(m.cantidadTotal),
        0
      );
      if (sumModelos > 0) {
        totalOp = sumModelos;
      }
    }

    return {
      op: item.op,
      cliente: item.cliente || '',
      cantidadTotal: totalOp,
      descripcion: tieneMultiples
        ? 'Varios modelos (' + item.modelos.length + ')'
        : item.modelos[0]?.descripcion || item.descripcion || '',
      sap: tieneMultiples ? '' : item.modelos[0]?.sap || item.sap || '',
      descripcionBase: item.descripcionBase || item.modelos[0]?.descripcion || item.descripcion || '',
      modelos: item.modelos || [],
    };
  });

  return nuevo;
}

export function procesarWorkbookCatalogo(wb: XLSX.WorkBook): CatalogoItem[] | null {
  const sheetPrincipal = wb.Sheets[wb.SheetNames[0]];
  const sheetModelos = wb.SheetNames.length > 1 ? wb.Sheets[wb.SheetNames[1]] : null;
  return construirCatalogo(sheetPrincipal, sheetModelos);
}

// ----------------- PARSING MAQUINAS -----------------
export function parsearCodigosMaquinas(wb: XLSX.WorkBook, maquinasActuales: Maquina[]): { actualizadas: number; creadas: number; maquinas: Maquina[] } {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (rows.length < 2) {
    throw new Error('El archivo no tiene datos');
  }
  const headers = rows[0].map(normalizar);
  const idxMaquina = headers.findIndex((h) => h.indexOf('maquina') !== -1);
  const idxCodigo = headers.findIndex((h) => h.indexOf('codigo') !== -1);
  if (idxMaquina === -1 || idxCodigo === -1) {
    throw new Error('Debe tener columnas MAQUINA y CODIGO');
  }
  const idxTipo = headers.findIndex((h) => h === 'tipo' || h.indexOf('tipo') === 0);
  const idxArea = headers.findIndex((h) => h === 'area' || h.indexOf('area') === 0);
  const idxCantidad = headers.findIndex((h) => h === 'cantidad' || h.indexOf('cantidad') === 0 || h === 'meta' || h.indexOf('meta') === 0 || h === 'objetivo');
  const idxUndMin = headers.findIndex((h) => h.indexOf('undxmin') !== -1 || h.indexOf('und/min') !== -1 || (h.indexOf('und') !== -1 && h.indexOf('min') !== -1));

  function leerTipo(r: any[]) {
    if (idxTipo === -1) return null;
    const v = String(r[idxTipo]).trim().toLowerCase();
    return v === 'pp' || v === 'pt' ? (v as 'pp' | 'pt') : null;
  }
  function leerArea(r: any[]) {
    if (idxArea === -1) return null;
    const v = String(r[idxArea]).trim().toLowerCase();
    if (v === 'v') return 'vasos_envases' as const;
    if (v === 'p') return 'papel_carton' as const;
    return null;
  }
  function leerMeta(r: any[]): number {
    if (idxCantidad === -1) return 0;
    const raw = String(r[idxCantidad] ?? '').trim().replace(/,/g, '').replace(/\./g, '');
    const n = parseInt(raw, 10);
    return isNaN(n) ? 0 : Math.max(0, n);
  }
  function leerUndMin(r: any[]): number {
    if (idxUndMin === -1) return 0;
    const raw = String(r[idxUndMin] ?? '').trim().replace(/,/g, '');
    const n = parseInt(raw, 10);
    return isNaN(n) ? 0 : Math.max(0, n);
  }

  const copiaMaquinas = [...maquinasActuales];
  let actualizadas = 0;
  let creadas = 0;

  rows.slice(1).forEach((r) => {
    const nombre = String(r[idxMaquina]).trim();
    const codigo = String(r[idxCodigo]).trim();
    const tipoNuevo = leerTipo(r);
    const areaNueva = leerArea(r);
    const metaNueva = leerMeta(r);
    const undMinNueva = leerUndMin(r);
    if (!nombre) return;
    const existente = copiaMaquinas.find((m) => m.nombre.trim().toLowerCase() === nombre.toLowerCase());
    if (existente) {
      existente.codigo = codigo;
      existente.bloqueada = true;
      if (tipoNuevo) existente.tipoProducto = tipoNuevo;
      if (areaNueva) existente.area = areaNueva;
      if (idxCantidad !== -1) existente.meta = metaNueva;
      if (idxUndMin !== -1) existente.undPorMin = undMinNueva;
      actualizadas++;
    } else {
      copiaMaquinas.push({
        id: Date.now() + '-' + Math.random().toString(36).slice(2, 7) + '-' + Math.random().toString(36).slice(2, 5),
        nombre,
        activa: false,
        area: areaNueva || '',
        codigo,
        bloqueada: true,
        tipoProducto: tipoNuevo || '',
        meta: metaNueva || 0,
        undPorMin: undMinNueva || 0,
      });
      creadas++;
    }
  });

  return { actualizadas, creadas, maquinas: copiaMaquinas };
}

// ----------------- EXPORT HELPERS -----------------
export function descargarBlob(blob: Blob, nombre: string) {
  if (navigator.canShare && window.File) {
    try {
      const archivo = new File([blob], nombre, { type: blob.type });
      if (navigator.canShare({ files: [archivo] })) {
        navigator.share({ files: [archivo], title: nombre }).catch((err) => {
          if (err && err.name === 'AbortError') return;
          descargaClasica(blob, nombre);
        });
        return;
      }
    } catch {}
  }
  descargaClasica(blob, nombre);
}

function descargaClasica(blob: Blob, nombre: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = nombre;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

// ----------------- EXPORT INVENTARIO LISTA -----------------
export function agruparParaExportar(items: ListaItem[]) {
  const mapa: Record<string, any> = {};
  const orden: string[] = [];
  items.forEach((item) => {
    const key = `${item.op}||${item.descripcion}||${item.tipoEmpaque}||${item.unidadesXCaja}`;
    if (!mapa[key]) {
      mapa[key] = {
        op: item.op,
        descripcion: item.descripcion,
        sap: item.sap || '',
        cliente: item.cliente,
        tipoEmpaque: item.tipoEmpaque,
        unidadesXCaja: item.unidadesXCaja,
        cantidadCajas: 0,
        totalUnidades: 0,
        observaciones: [] as string[],
      };
      orden.push(key);
    }
    mapa[key].cantidadCajas += item.cantidadCajas;
    mapa[key].totalUnidades += item.totalUnidades;
    if (item.observacion && mapa[key].observaciones.indexOf(item.observacion) === -1) {
      mapa[key].observaciones.push(item.observacion);
    }
  });
  return orden.map((k) => {
    const g = mapa[k];
    g.observacion = g.observaciones.join(' / ');
    return g;
  });
}

export function prepararFilasInventario(lista: ListaItem[], catalogo: CatalogoItem[]) {
  // Mantener cada línea individual en orden cronológico de registro (sin agrupar ni sumar líneas)
  const ordenadas = [...(lista || [])].reverse();
  return ordenadas.map((item, idx) => {
    const match = buscarEnCatalogo(catalogo, item.op);
    const empaqueTxt = item.tipoEmpaque === 'paquete' ? 'Paquete' : 'Caja';
    const sap = item.sap || (match && match.sap ? match.sap : '');
    const unidadesXEmp = Number(item.unidadesXCaja) || 0;
    const cantEmp = Number(item.cantidadCajas) || 0;
    const totalUnid = Number(item.totalUnidades) || unidadesXEmp * cantEmp;
    return {
      indice: idx + 1,
      op: item.op || '',
      sap,
      descripcion: item.descripcion || '',
      cliente: item.cliente || '',
      empaque: empaqueTxt,
      unidadesXEmp,
      cantEmp,
      totalUnid,
      observacion: item.observacion || '',
    };
  });
}

export async function exportarInventarioExcel(lista: ListaItem[], catalogo: CatalogoItem[]) {
  const filas = prepararFilasInventario(lista, catalogo);
  if (filas.length === 0) return;

  const totalUnidades = filas.reduce((acc, f) => acc + f.totalUnid, 0);
  const totalEmpaques = filas.reduce((acc, f) => acc + f.cantEmp, 0);
  const ahora = new Date();
  const hoyIso = ahora.toISOString().slice(0, 10);
  const fechaCorta = `${hoyIso.slice(8, 10)}/${hoyIso.slice(5, 7)}/${hoyIso.slice(0, 4)}`;
  const fechaHora = ahora.toLocaleString('es-PE', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });
  const nombre = `inventario_${hoyIso}.xlsx`;

  try {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Inventario');
    ws.columns = [
      { width: 6 },  // #
      { width: 13 }, // OP
      { width: 15 }, // SAP
      { width: 42 }, // DESCRIPCIÓN
      { width: 24 }, // CLIENTE
      { width: 12 }, // EMPAQUE
      { width: 15 }, // UNID. X EMP.
      { width: 14 }, // CANT. EMP.
      { width: 17 }, // TOTAL UNIDADES
      { width: 30 }, // OBSERVACIÓN
    ];

    // Row 1: Main Title Banner
    ws.mergeCells(1, 1, 1, 10);
    const rTitulo = ws.getRow(1);
    rTitulo.height = 24;
    rTitulo.getCell(1).value = `INVENTARIO DE PRODUCCIÓN — ${fechaCorta}`;
    rTitulo.getCell(1).font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    rTitulo.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF12161F' } };
    rTitulo.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    // Row 2: Sub-banner with summary
    ws.mergeCells(2, 1, 2, 10);
    const rSub = ws.getRow(2);
    rSub.height = 18;
    rSub.getCell(1).value = `Generado: ${fechaHora}   |   Líneas registradas: ${filas.length}   |   Total empaques: ${totalEmpaques.toLocaleString('es-PE')}   |   Total unidades: ${totalUnidades.toLocaleString('es-PE')} u`;
    rSub.getCell(1).font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    rSub.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD97706' } };
    rSub.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    // Row 3: Column Headers
    const headers = [
      '#',
      'OP',
      'SAP',
      'DESCRIPCIÓN / MODELO',
      'CLIENTE',
      'EMPAQUE',
      'UNID. X EMP.',
      'CANT. EMP.',
      'TOTAL UNIDADES',
      'OBSERVACIÓN',
    ];
    const rHead = ws.getRow(3);
    rHead.height = 22;
    headers.forEach((h, i) => {
      const c = rHead.getCell(i + 1);
      c.value = h;
      c.font = { bold: true, size: 10, color: { argb: 'FF333333' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      c.border = EXCELJS_BORDES;
    });

    // Data rows (sin agrupar ni sumar líneas entre sí)
    let rowIdx = 4;
    filas.forEach((f, idx) => {
      const r = ws.getRow(rowIdx);
      const vals = [
        f.indice,
        f.op,
        f.sap,
        f.descripcion,
        f.cliente,
        f.empaque,
        f.unidadesXEmp,
        f.cantEmp,
        f.totalUnid,
        f.observacion,
      ];
      const esPar = idx % 2 === 1;
      vals.forEach((v, colIdx) => {
        const c = r.getCell(colIdx + 1);
        c.value = v;
        c.border = EXCELJS_BORDES;
        if (esPar) {
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFAFAFA' } };
        }
        if (colIdx === 0 || colIdx === 1 || colIdx === 2 || colIdx === 5) {
          c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          if (colIdx === 1) c.font = { bold: true, color: { argb: 'FF12161F' } };
        } else if (colIdx === 6 || colIdx === 7 || colIdx === 8) {
          c.alignment = { horizontal: 'right', vertical: 'middle' };
          c.numFmt = '#,##0';
          if (colIdx === 8) c.font = { bold: true, color: { argb: 'FF1E40AF' } };
        } else {
          c.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        }
      });
      rowIdx++;
    });

    // Total Row
    ws.mergeCells(rowIdx, 1, rowIdx, 7);
    const rTot = ws.getRow(rowIdx);
    rTot.height = 20;
    rTot.getCell(1).value = `TOTAL GENERAL (${filas.length} LÍNEA${filas.length !== 1 ? 'S' : ''})`;
    for (let i = 1; i <= 10; i++) {
      const cell = rTot.getCell(i);
      cell.font = { bold: true, size: 10.5 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDEBD0' } };
      cell.border = EXCELJS_BORDES;
      cell.alignment = { vertical: 'middle' };
    }
    rTot.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
    rTot.getCell(8).value = totalEmpaques;
    rTot.getCell(8).numFmt = '#,##0';
    rTot.getCell(8).alignment = { horizontal: 'right', vertical: 'middle' };
    rTot.getCell(9).value = totalUnidades;
    rTot.getCell(9).numFmt = '#,##0';
    rTot.getCell(9).alignment = { horizontal: 'right', vertical: 'middle' };
    rTot.getCell(9).font = { bold: true, size: 11, color: { argb: 'FFB45309' } };

    ws.views = [{ state: 'frozen', ySplit: 3 }];
    ws.pageSetup = {
      orientation: 'landscape',
      paperSize: 9,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
    };

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    descargarBlob(blob, nombre);
  } catch {
    // Fallback to SheetJS (sin agrupar ni sumar líneas)
    const filasSimple = filas.map((f) => ({
      '#': f.indice,
      OP: f.op,
      SAP: f.sap,
      Descripción: f.descripcion,
      Cliente: f.cliente,
      Empaque: f.empaque,
      'Unidades x empaque': f.unidadesXEmp,
      'Cantidad de empaques': f.cantEmp,
      'Total unidades': f.totalUnid,
      Observación: f.observacion,
    }));
    const ws = XLSX.utils.json_to_sheet(filasSimple);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 13 },
      { wch: 14 },
      { wch: 38 },
      { wch: 22 },
      { wch: 11 },
      { wch: 16 },
      { wch: 16 },
      { wch: 16 },
      { wch: 30 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Inventario');
    const wbArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbArray], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    descargarBlob(blob, nombre);
  }
}

export function imprimirInventarioFormatoExcel(lista: ListaItem[], catalogo: CatalogoItem[]) {
  const filas = prepararFilasInventario(lista, catalogo);
  if (filas.length === 0) return;

  const totalUnidades = filas.reduce((acc, f) => acc + f.totalUnid, 0);
  const totalEmpaques = filas.reduce((acc, f) => acc + f.cantEmp, 0);
  const ahora = new Date();
  const fechaCorta = ahora.toLocaleDateString('es-PE', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  });
  const fechaHora = ahora.toLocaleString('es-PE', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });

  const titulo = `INVENTARIO DE PRODUCCIÓN — ${fechaCorta}`;
  const subtitulo = `Generado: ${fechaHora} · Detalle por línea registrada (${filas.length} registro${filas.length !== 1 ? 's' : ''})`;

  const tbodyHtml = filas
    .map(
      (f, idx) => `
      <tr class="${idx % 2 === 1 ? 'alt-row' : ''}">
        <td class="idx-cell">${f.indice}</td>
        <td class="op-cell">${escapeHtml(f.op)}</td>
        <td class="sap-cell">${escapeHtml(f.sap || '—')}</td>
        <td class="prod-cell">${escapeHtml(f.descripcion)}</td>
        <td class="cli-cell">${escapeHtml(f.cliente || '—')}</td>
        <td class="emp-cell">${escapeHtml(f.empaque)}</td>
        <td class="num">${f.unidadesXEmp.toLocaleString('es-PE')}</td>
        <td class="num">${f.cantEmp.toLocaleString('es-PE')}</td>
        <td class="num total-cell">${f.totalUnid.toLocaleString('es-PE')}</td>
        <td class="obs-cell">${escapeHtml(f.observacion)}</td>
      </tr>`
    )
    .join('');

  const html = `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>${escapeHtml(titulo)}</title>
  <style>
    @page {
      size: A4 landscape;
      margin: 5mm;
    }
    * {
      box-sizing: border-box;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    html, body {
      margin: 0;
      padding: 0;
      width: 285mm;
      max-width: 285mm;
      background: #ffffff;
      font-family: Arial, Helvetica, sans-serif;
      color: #111827;
      font-size: 9.5px;
    }
    @media print {
      html, body {
        width: 285mm;
        height: auto;
        max-height: 192mm;
        overflow: hidden !important;
      }
    }
    #print-viewport {
      width: 285mm;
      overflow: hidden;
      page-break-inside: avoid;
      break-inside: avoid;
      page-break-after: avoid;
      break-after: avoid;
    }
    #sheet-content {
      width: 285mm;
      transform-origin: top left;
    }
    .title-banner {
      background: #12161F;
      color: #FFFFFF;
      padding: 6px 10px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      border-bottom: 3px solid #F2A93B;
      margin-bottom: 6px;
    }
    .title-main {
      font-size: 13px;
      font-weight: 700;
      letter-spacing: 0.4px;
      text-transform: uppercase;
    }
    .title-sub {
      font-size: 9px;
      color: #CBD5E0;
      margin-top: 1px;
    }
    .title-stats {
      text-align: right;
      font-size: 10px;
      white-space: nowrap;
    }
    .title-stats strong {
      color: #F2A93B;
      font-size: 11px;
    }
    .area-banner {
      background: #D97706;
      color: #FFFFFF;
      font-weight: 700;
      font-size: 10px;
      padding: 3px 7px;
      letter-spacing: 0.4px;
      border: 1px solid #B7B7B7;
      border-bottom: none;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
    }
    th, td {
      border: 1px solid #B7B7B7;
      padding: 2.5px 4.5px;
      vertical-align: middle;
      word-wrap: break-word;
      overflow-wrap: anywhere;
      line-height: 1.2;
    }
    th {
      background: #F2F2F2;
      color: #333333;
      font-weight: 700;
      font-size: 8.5px;
      text-align: center;
    }
    .alt-row td {
      background: #FAFAFA;
    }
    .idx-cell, .op-cell, .sap-cell, .emp-cell {
      font-family: 'Courier New', Courier, monospace;
      text-align: center;
      font-size: 9px;
    }
    .op-cell {
      font-weight: 700;
      color: #12161F;
    }
    .prod-cell, .cli-cell, .obs-cell {
      text-align: left;
    }
    .num {
      text-align: right;
      font-family: 'Courier New', Courier, monospace;
      font-weight: 600;
    }
    .total-cell {
      font-weight: 700;
      color: #1E40AF;
    }
    .total-row td {
      background: #FDEBD0;
      font-weight: 700;
      font-size: 9.5px;
    }
    .grand-total {
      margin-top: 5px;
      background: #12161F;
      color: #FFFFFF;
      padding: 5px 9px;
      display: flex;
      justify-content: space-between;
      font-size: 10px;
      font-weight: 700;
      border-radius: 2px;
    }
    .grand-total span.val {
      color: #F2A93B;
      font-family: 'Courier New', Courier, monospace;
    }
    .grand-total span.sub-val {
      color: #93C5FD;
      font-family: 'Courier New', Courier, monospace;
      margin-right: 12px;
    }

    #sheet-content.compact .title-banner { padding: 4px 8px; margin-bottom: 4px; }
    #sheet-content.compact .title-main { font-size: 12px; }
    #sheet-content.compact .area-banner { padding: 2px 6px; font-size: 9.5px; }
    #sheet-content.compact th,
    #sheet-content.compact td { padding: 1.8px 3.5px; font-size: 8.8px; line-height: 1.15; }
    #sheet-content.compact th { font-size: 8px; }
    #sheet-content.compact .grand-total { padding: 4px 8px; font-size: 9px; margin-top: 3px; }

    #sheet-content.ultra-compact .title-banner { padding: 3px 7px; margin-bottom: 3px; }
    #sheet-content.ultra-compact .title-main { font-size: 11px; }
    #sheet-content.ultra-compact .area-banner { padding: 1.5px 5px; font-size: 9px; }
    #sheet-content.ultra-compact th,
    #sheet-content.ultra-compact td { padding: 1.2px 3px; font-size: 8.2px; line-height: 1.12; }
    #sheet-content.ultra-compact th { font-size: 7.6px; }
    #sheet-content.ultra-compact .grand-total { padding: 3px 7px; font-size: 8.5px; margin-top: 2px; }
  </style>
</head>
<body>
  <div id="mm-ruler" style="position:absolute;top:-9999px;left:0;width:100mm;height:100mm;visibility:hidden;pointer-events:none;"></div>
  <div id="print-viewport">
    <div id="sheet-content">
      <div class="title-banner">
        <div>
          <div class="title-main">${escapeHtml(titulo)}</div>
          <div class="title-sub">${escapeHtml(subtitulo)}</div>
        </div>
        <div class="title-stats">
          Total: <strong>${totalUnidades.toLocaleString('es-PE')} u</strong>
          <span style="color:#A0AEC0;margin-left:8px;">(${filas.length} líneas)</span>
        </div>
      </div>

      <div class="area-banner">DETALLE DE PIEZAS Y EMPAQUES REGISTRADOS</div>
      <table>
        <thead>
          <tr>
            <th style="width:4%">#</th>
            <th style="width:8%">OP</th>
            <th style="width:9%">SAP</th>
            <th style="width:24%">DESCRIPCIÓN / MODELO</th>
            <th style="width:15%">CLIENTE</th>
            <th style="width:7%">EMPAQUE</th>
            <th style="width:8%">UNID. × EMP.</th>
            <th style="width:7%">CANT. EMP.</th>
            <th style="width:9%">TOTAL UNID.</th>
            <th style="width:9%">OBSERVACIÓN</th>
          </tr>
        </thead>
        <tbody>
          ${tbodyHtml}
          <tr class="total-row">
            <td colspan="7">TOTAL GENERAL (${filas.length} LÍNEA${filas.length !== 1 ? 'S' : ''})</td>
            <td class="num">${totalEmpaques.toLocaleString('es-PE')}</td>
            <td class="num total-cell">${totalUnidades.toLocaleString('es-PE')}</td>
            <td></td>
          </tr>
        </tbody>
      </table>

      <div class="grand-total">
        <span>RESUMEN DE INVENTARIO (${filas.length} LÍNEAS)</span>
        <div>
          EMPAQUES: <span class="sub-val">${totalEmpaques.toLocaleString('es-PE')}</span>
          TOTAL UNIDADES: <span class="val">${totalUnidades.toLocaleString('es-PE')} u</span>
        </div>
      </div>
    </div>
  </div>
</body>
</html>`;

  const iframe = document.createElement('iframe');
  iframe.style.position = 'fixed';
  iframe.style.left = '-12000px';
  iframe.style.top = '0';
  iframe.style.width = '1200px';
  iframe.style.height = '900px';
  iframe.style.border = '0';
  iframe.style.opacity = '0';
  iframe.style.pointerEvents = 'none';
  iframe.setAttribute('aria-hidden', 'true');
  document.body.appendChild(iframe);

  const doc = iframe.contentWindow?.document;
  if (!doc || !iframe.contentWindow) {
    document.body.removeChild(iframe);
    return;
  }

  doc.open();
  doc.write(html);
  doc.close();

  const fitSingleLandscapePage = () => {
    const ruler = doc.getElementById('mm-ruler');
    const viewport = doc.getElementById('print-viewport');
    const content = doc.getElementById('sheet-content');
    if (!ruler || !viewport || !content) return;

    const rulerRect = ruler.getBoundingClientRect();
    const pxPerMm = rulerRect.height > 0 ? rulerRect.height / 100 : 3.7795275591;
    const widthMm = 285;
    const maxHeightMm = 190;
    const maxH = maxHeightMm * pxPerMm;

    content.style.width = `${widthMm}mm`;

    if (content.scrollHeight > maxH) {
      content.classList.add('compact');
    }
    if (content.scrollHeight > maxH * 1.18) {
      content.classList.add('ultra-compact');
    }

    let bestScale = 1;
    if (content.scrollHeight > maxH) {
      let low = 0.25;
      let high = 1.0;
      bestScale = 0.25;

      for (let step = 0; step < 14; step++) {
        const mid = (low + high) / 2;
        content.style.width = `${(widthMm / mid).toFixed(3)}mm`;
        const h = content.scrollHeight;
        if (h * mid <= maxH * 0.985) {
          bestScale = mid;
          low = mid;
        } else {
          high = mid;
        }
      }

      content.style.width = `${(widthMm / bestScale).toFixed(3)}mm`;
      content.style.transform = `scale(${bestScale.toFixed(5)})`;
      content.style.transformOrigin = 'top left';
    }

    const scaledHeightMm = Math.min(maxHeightMm, (content.scrollHeight / pxPerMm) * bestScale + 0.5);
    viewport.style.width = `${widthMm}mm`;
    viewport.style.height = `${scaledHeightMm.toFixed(2)}mm`;
    viewport.style.overflow = 'hidden';

    ruler.remove();
  };

  const cleanup = () => {
    if (iframe.parentNode) {
      iframe.parentNode.removeChild(iframe);
    }
  };

  iframe.contentWindow.addEventListener('afterprint', cleanup, { once: true });
  setTimeout(() => {
    try {
      fitSingleLandscapePage();
      iframe.contentWindow?.focus();
      iframe.contentWindow?.print();
    } finally {
      setTimeout(cleanup, 60000);
    }
  }, 180);
}

// ----------------- TURNO INSTANCIAS & MULTI-OP -----------------
export function turnosDeMaquina(m: Maquina, turno: TurnoMaquina[]): TurnoMaquina[] {
  if (!m.multiOp) {
    const t = turno.find((x) => x.maquinaId === m.id && (x.instancia || 1) === 1);
    return [
      t || {
        maquinaId: m.id,
        instancia: 1,
        op: '',
        objetivo: null,
        motivoParada: '',
        observaciones: '',
        personal: '',
        modeloIdx: -1,
        pliego: '',
        registros: [],
      },
    ];
  }
  let lista = turno.filter((x) => x.maquinaId === m.id);
  if (!lista.length) {
    lista = [
      {
        maquinaId: m.id,
        instancia: 1,
        op: '',
        objetivo: null,
        motivoParada: '',
        observaciones: '',
        personal: '',
        modeloIdx: -1,
        pliego: '',
        registros: [],
      },
    ];
  }
  return lista.slice().sort((a, b) => (a.instancia || 1) - (b.instancia || 1));
}

export function claveTarjeta(m: Maquina, inst = 1): string {
  return m.multiOp ? `${m.id}#${inst}` : m.id;
}

export function nombreTarjeta(m: Maquina, inst = 1): string {
  return m.multiOp ? `${m.nombre}(${inst})` : m.nombre;
}

export function opsDeTurno(t: TurnoMaquina): { op: string; modeloIdx: number; pliego: string | number; personal: string }[] {
  return t.op
    ? [
        {
          op: t.op,
          modeloIdx: typeof t.modeloIdx === 'number' ? t.modeloIdx : -1,
          pliego: t.pliego || '',
          personal: t.personal || '',
        },
      ]
    : [];
}

// ----------------- TURNO EXCEL & FORMATTING -----------------
export function generarFilasTurno(maquinas: Maquina[], turno: TurnoMaquina[], catalogo: CatalogoItem[]): FilaTurnoExcel[] {
  const filas: FilaTurnoExcel[] = [];
  const activas = maquinas.filter((m) => m.activa).sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' }));

  activas.forEach((m) => {
    turnosDeMaquina(m, turno).forEach((t) => {
      function textoObservaciones(tn: TurnoMaquina) {
        const partes: string[] = [];
        if (tn.motivoParada) partes.push('Detenida: ' + tn.motivoParada);
        if (tn.observaciones) partes.push(tn.observaciones);
        return partes.join(' | ');
      }

      function textoObservacionesGrupo(tn: TurnoMaquina, g: any) {
        const partes: string[] = [];
        if (g.etq) partes.push(g.etq);
        if (tn.motivoParada) partes.push('Detenida: ' + tn.motivoParada);
        if (g.obs && g.obs.length) partes.push(g.obs.join(' / '));
        return partes.join(' | ');
      }

      if (!t.registros || t.registros.length === 0) {
        const opsVacias = opsDeTurno(t);
        const listaOps = opsVacias.length ? opsVacias : [{ op: '', modeloIdx: -1, pliego: '', personal: '' }];
        listaOps.forEach((o) => {
          const match = buscarEnCatalogo(catalogo, o.op);
          const modeloVacio = resolverModelo(match, o.modeloIdx);
          filas.push({
            maquina: m.nombre,
            codigoMaquina: m.codigo || '',
            tipoProducto: m.tipoProducto || '',
            area: m.area || '',
            op: o.op || '',
            sap: match ? modeloVacio.sap || '' : '',
            producto: match ? modeloVacio.descripcion : '',
            total: 0,
            merma: 0,
            observaciones: textoObservaciones(t),
          });
        });
      } else {
        const grupos: Record<string, any> = {};
        const ordenGrupos: string[] = [];
        const opsConGrupo: string[] = [];

        t.registros.forEach((r) => {
          const opKey = r.op || t.op || '';
          const descKey = r.descripcion || '';
          const pliegoKey = r.pliego || '';
          const claveGrupo = `${opKey}||${descKey}||${pliegoKey}`;
          if (!grupos[claveGrupo]) {
            const match = buscarEnCatalogo(catalogo, opKey);
            const descBase = descKey || (match ? resolverModelo(match, -1).descripcion : '');
            grupos[claveGrupo] = {
              op: opKey,
              descripcion: descBase + (pliegoKey ? ' - ' + etiquetaPliego(pliegoKey) : ''),
              sap: r.sap || (match ? match.sap || '' : ''),
              total: 0,
              merma: 0,
              etq: etiquetaPliego(pliegoKey),
              obs: [] as string[],
            };
            ordenGrupos.push(claveGrupo);
            if (opsConGrupo.indexOf(opKey) === -1) opsConGrupo.push(opKey);
          }
          grupos[claveGrupo].total += r.cantidad || 0;
          grupos[claveGrupo].merma += r.merma || 0;
          const o = typeof r.observacion === 'undefined' ? t.observaciones || '' : r.observacion;
          if (o && grupos[claveGrupo].obs.indexOf(o) === -1) grupos[claveGrupo].obs.push(o);
        });

        ordenGrupos.forEach((claveGrupo) => {
          const g = grupos[claveGrupo];
          const match = buscarEnCatalogo(catalogo, g.op);
          const prodDesc = g.descripcion || (match ? resolverModelo(match, -1).descripcion : '');
          filas.push({
            maquina: m.nombre,
            codigoMaquina: m.codigo || '',
            tipoProducto: m.tipoProducto || '',
            area: m.area || '',
            op: g.op,
            sap: g.sap || '',
            producto: prodDesc,
            total: g.total,
            merma: g.merma,
            observaciones: textoObservacionesGrupo(t, g),
          });
        });

        opsDeTurno(t).forEach((o) => {
          if (!o.op || opsConGrupo.indexOf(o.op) !== -1) return;
          const matchActual = buscarEnCatalogo(catalogo, o.op);
          const modeloActual = resolverModelo(matchActual, o.modeloIdx);
          filas.push({
            maquina: m.nombre,
            codigoMaquina: m.codigo || '',
            tipoProducto: m.tipoProducto || '',
            area: m.area || '',
            op: o.op,
            sap: matchActual ? modeloActual.sap || '' : '',
            producto: matchActual ? modeloActual.descripcion : '',
            total: 0,
            merma: 0,
            observaciones: textoObservaciones(t),
          });
        });
      }
    });
  });

  return filas;
}

const EXCELJS_BORDE_FINO = { style: 'thin' as const, color: { argb: 'FFB7B7B7' } };
const EXCELJS_BORDES = { top: EXCELJS_BORDE_FINO, bottom: EXCELJS_BORDE_FINO, left: EXCELJS_BORDE_FINO, right: EXCELJS_BORDE_FINO };

export async function exportarTurnoExcel(filas: FilaTurnoExcel[], titulo: string, nombreBase: string) {
  try {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Turno');
    ws.columns = [{ width: 18 }, { width: 15 }, { width: 13 }, { width: 14 }, { width: 42 }, { width: 17 }, { width: 10 }, { width: 32 }];

    ws.mergeCells(1, 1, 1, 8);
    const rTitulo = ws.getRow(1);
    rTitulo.height = 22;
    rTitulo.getCell(1).value = titulo;
    rTitulo.getCell(1).font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    rTitulo.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF12161F' } };
    rTitulo.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    let fila = 2;
    AREAS_ORDEN.forEach((areaDef) => {
      const filasDeArea = filas.filter((f) => (f.area || '') === areaDef.key);
      if (filasDeArea.length === 0) return;

      ws.mergeCells(fila, 1, fila, 8);
      const rArea = ws.getRow(fila);
      rArea.getCell(1).value = areaDef.label;
      rArea.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      rArea.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD97706' } };
      fila++;

      const headers = ['MAQUINA', 'CODIGO MAQUINA', "OP'S", 'SAP', 'PRODUCTO', 'TOTAL PRODUCCION', 'MERMA', 'OBSERVACIONES'];
      const rHead = ws.getRow(fila);
      headers.forEach((h, i) => {
        const c = rHead.getCell(i + 1);
        c.value = h;
        c.font = { bold: true, color: { argb: 'FF333333' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
        c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        c.border = EXCELJS_BORDES;
      });
      fila++;

      let totalArea = 0;
      let mermaArea = 0;
      let maquinaAnterior: string | null = null;
      let inicioRun = fila;

      filasDeArea.forEach((f, idx) => {
        totalArea += f.total || 0;
        mermaArea += f.merma || 0;
        const r = ws.getRow(fila);
        const vals = [f.maquina, f.codigoMaquina || '', f.op, f.sap || '', f.producto, f.total || 0, f.merma || 0, f.observaciones || ''];
        vals.forEach((v, i) => {
          const c = r.getCell(i + 1);
          c.value = v;
          c.border = EXCELJS_BORDES;
          c.alignment = { vertical: 'middle', wrapText: true };
          if (i === 5) c.alignment = { horizontal: 'right', vertical: 'middle' };
          if (i === 6) {
            c.alignment = { horizontal: 'right', vertical: 'middle' };
            c.font = { color: { argb: 'FFC0392B' } };
          }
        });
        if (f.maquina !== maquinaAnterior) {
          if (maquinaAnterior !== null && fila - inicioRun > 1) {
            ws.mergeCells(inicioRun, 1, fila - 1, 1);
            ws.mergeCells(inicioRun, 2, fila - 1, 2);
          }
          inicioRun = fila;
          maquinaAnterior = f.maquina;
        }
        fila++;
        if (idx === filasDeArea.length - 1 && fila - inicioRun > 1) {
          ws.mergeCells(inicioRun, 1, fila - 1, 1);
          ws.mergeCells(inicioRun, 2, fila - 1, 2);
        }
      });

      ws.mergeCells(fila, 1, fila, 5);
      const rTot = ws.getRow(fila);
      rTot.getCell(1).value = 'TOTAL';
      for (let i = 1; i <= 8; i++) {
        rTot.getCell(i).font = { bold: true };
        rTot.getCell(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDEBD0' } };
        rTot.getCell(i).border = EXCELJS_BORDES;
      }
      rTot.getCell(6).value = totalArea;
      rTot.getCell(6).alignment = { horizontal: 'right' };
      rTot.getCell(7).value = mermaArea;
      rTot.getCell(7).alignment = { horizontal: 'right' };
      rTot.getCell(7).font = { bold: true, color: { argb: 'FFC0392B' } };
      fila++;
    });

    ws.views = [{ state: 'frozen', ySplit: 1 }];
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    descargarBlob(blob, `${nombreBase}.xlsx`);
  } catch (err) {
    // Fallback to SheetJS
    const aoa: any[][] = [[titulo], []];
    AREAS_ORDEN.forEach((areaDef) => {
      const filasDeArea = filas.filter((f) => (f.area || '') === areaDef.key);
      if (filasDeArea.length === 0) return;
      aoa.push([areaDef.label]);
      aoa.push(['MAQUINA', 'CODIGO MAQUINA', "OP'S", 'SAP', 'PRODUCTO', 'TOTAL PRODUCCION', 'MERMA', 'OBSERVACIONES']);
      let totalArea = 0;
      let mermaArea = 0;
      filasDeArea.forEach((f) => {
        totalArea += f.total || 0;
        mermaArea += f.merma || 0;
        aoa.push([f.maquina, f.codigoMaquina || '', f.op, f.sap || '', f.producto, f.total || 0, f.merma || 0, f.observaciones || '']);
      });
      aoa.push(['TOTAL', '', '', '', '', totalArea, mermaArea, '']);
      aoa.push([]);
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 18 }, { wch: 15 }, { wch: 13 }, { wch: 14 }, { wch: 42 }, { wch: 17 }, { wch: 10 }, { wch: 32 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Turno');
    const wbArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    descargarBlob(blob, `${nombreBase}.xlsx`);
  }
}

export function imprimirTurnoFormatoExcel(
  filasOriginales: FilaTurnoExcel[],
  titulo: string,
  subtitulo: string,
  maquinasCatalogo: Maquina[] = []
) {
  // Enrich rows that might miss area/codigoMaquina in legacy saved shifts
  const maqMap = new Map<string, Maquina>();
  maquinasCatalogo.forEach((m) => {
    maqMap.set(m.nombre.trim().toLowerCase(), m);
  });

  const filas = (filasOriginales || []).map((f) => {
    const match = maqMap.get((f.maquina || '').trim().toLowerCase());
    return {
      ...f,
      area: f.area || match?.area || '',
      codigoMaquina: f.codigoMaquina || match?.codigo || '',
    };
  });

  const totalGen = filas.reduce((s, f) => s + (f.total || 0), 0);
  const mermaGen = filas.reduce((s, f) => s + (f.merma || 0), 0);

  const bloquesAreas = AREAS_ORDEN.map((areaDef) => {
    const filasDeArea = filas.filter((f) => (f.area || '') === areaDef.key);
    if (filasDeArea.length === 0) return '';

    let totalArea = 0;
    let mermaArea = 0;
    filasDeArea.forEach((f) => {
      totalArea += f.total || 0;
      mermaArea += f.merma || 0;
    });

    // Build rows with vertical rowspan merging for consecutive identical machine names (matching Excel export)
    let tbodyHtml = '';
    let i = 0;
    while (i < filasDeArea.length) {
      let j = i;
      const maqKey = (filasDeArea[i].maquina || '').trim().toLowerCase();
      while (
        j + 1 < filasDeArea.length &&
        (filasDeArea[j + 1].maquina || '').trim().toLowerCase() === maqKey
      ) {
        j++;
      }
      const span = j - i + 1;
      const rs = span > 1 ? ` rowspan="${span}"` : '';

      for (let k = i; k <= j; k++) {
        const f = filasDeArea[k];
        tbodyHtml += '<tr>';
        if (k === i) {
          tbodyHtml += `<td${rs} class="maq-cell">${escapeHtml(f.maquina || '')}</td>`;
          tbodyHtml += `<td${rs} class="cod-cell">${escapeHtml(f.codigoMaquina || '')}</td>`;
        }
        tbodyHtml += `<td class="op-cell">${escapeHtml(f.op || '')}</td>`;
        tbodyHtml += `<td class="sap-cell">${escapeHtml(f.sap || '')}</td>`;
        tbodyHtml += `<td class="prod-cell">${escapeHtml(f.producto || '')}</td>`;
        tbodyHtml += `<td class="num">${(f.total || 0).toLocaleString('es-PE')}</td>`;
        tbodyHtml += `<td class="num merma">${(f.merma || 0).toLocaleString('es-PE')}</td>`;
        tbodyHtml += `<td class="obs-cell">${escapeHtml(f.observaciones || '')}</td>`;
        tbodyHtml += '</tr>';
      }
      i = j + 1;
    }

    return `
      <div class="area-block">
        <div class="area-banner">${escapeHtml(areaDef.label)}</div>
        <table>
          <thead>
            <tr>
              <th style="width:13%">MÁQUINA</th>
              <th style="width:9%">CÓDIGO MÁQ.</th>
              <th style="width:9%">OP'S</th>
              <th style="width:9%">SAP</th>
              <th style="width:24%">PRODUCTO</th>
              <th style="width:11%">TOTAL PROD.</th>
              <th style="width:8%">MERMA</th>
              <th style="width:17%">OBSERVACIONES</th>
            </tr>
          </thead>
          <tbody>
            ${tbodyHtml}
            <tr class="total-row">
              <td colspan="5">TOTAL ${escapeHtml(areaDef.label)}</td>
              <td class="num">${totalArea.toLocaleString('es-PE')}</td>
              <td class="num merma">${mermaArea.toLocaleString('es-PE')}</td>
              <td></td>
            </tr>
          </tbody>
        </table>
      </div>
    `;
  })
    .filter(Boolean)
    .join('');

  const html = `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>${escapeHtml(titulo)}</title>
  <style>
    @page {
      size: A4 landscape;
      margin: 5mm;
    }
    * {
      box-sizing: border-box;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    html, body {
      margin: 0;
      padding: 0;
      width: 285mm;
      max-width: 285mm;
      background: #ffffff;
      font-family: Arial, Helvetica, sans-serif;
      color: #111827;
      font-size: 9.5px;
    }
    @media print {
      html, body {
        width: 285mm;
        height: auto;
        max-height: 192mm;
        overflow: hidden !important;
      }
    }
    #print-viewport {
      width: 285mm;
      overflow: hidden;
      page-break-inside: avoid;
      break-inside: avoid;
      page-break-after: avoid;
      break-after: avoid;
    }
    #sheet-content {
      width: 285mm;
      transform-origin: top left;
    }
    .title-banner {
      background: #12161F;
      color: #FFFFFF;
      padding: 6px 10px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      border-bottom: 3px solid #F2A93B;
      margin-bottom: 6px;
    }
    .title-main {
      font-size: 13px;
      font-weight: 700;
      letter-spacing: 0.4px;
      text-transform: uppercase;
    }
    .title-sub {
      font-size: 9px;
      color: #CBD5E0;
      margin-top: 1px;
    }
    .title-stats {
      text-align: right;
      font-size: 10px;
      white-space: nowrap;
    }
    .title-stats strong {
      color: #F2A93B;
      font-size: 11px;
    }
    .title-stats .merma-badge {
      color: #FC8181;
      font-weight: 700;
      margin-left: 8px;
    }
    .area-block {
      margin-bottom: 6px;
    }
    .area-banner {
      background: #D97706;
      color: #FFFFFF;
      font-weight: 700;
      font-size: 10px;
      padding: 3px 7px;
      letter-spacing: 0.4px;
      border: 1px solid #B7B7B7;
      border-bottom: none;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
    }
    th, td {
      border: 1px solid #B7B7B7;
      padding: 2.5px 4.5px;
      vertical-align: middle;
      word-wrap: break-word;
      overflow-wrap: anywhere;
      line-height: 1.2;
    }
    th {
      background: #F2F2F2;
      color: #333333;
      font-weight: 700;
      font-size: 8.5px;
      text-align: center;
    }
    .maq-cell {
      font-weight: 700;
      background: #FAFAFA;
    }
    .cod-cell, .op-cell, .sap-cell {
      font-family: 'Courier New', Courier, monospace;
      text-align: center;
      font-size: 9px;
    }
    .prod-cell, .obs-cell {
      text-align: left;
    }
    .num {
      text-align: right;
      font-family: 'Courier New', Courier, monospace;
      font-weight: 600;
    }
    .merma {
      color: #C0392B;
    }
    .total-row td {
      background: #FDEBD0;
      font-weight: 700;
      font-size: 9.5px;
    }
    .grand-total {
      margin-top: 4px;
      background: #12161F;
      color: #FFFFFF;
      padding: 5px 9px;
      display: flex;
      justify-content: space-between;
      font-size: 10px;
      font-weight: 700;
      border-radius: 2px;
    }
    .grand-total span.val {
      color: #F2A93B;
      font-family: 'Courier New', Courier, monospace;
    }
    .grand-total span.merma-val {
      color: #FC8181;
      font-family: 'Courier New', Courier, monospace;
      margin-left: 12px;
    }

    /* Compact density tiers applied automatically before scaling when needed */
    #sheet-content.compact .title-banner { padding: 4px 8px; margin-bottom: 4px; }
    #sheet-content.compact .title-main { font-size: 12px; }
    #sheet-content.compact .area-block { margin-bottom: 4px; }
    #sheet-content.compact .area-banner { padding: 2px 6px; font-size: 9.5px; }
    #sheet-content.compact th,
    #sheet-content.compact td { padding: 1.8px 3.5px; font-size: 8.8px; line-height: 1.15; }
    #sheet-content.compact th { font-size: 8px; }
    #sheet-content.compact .grand-total { padding: 4px 8px; font-size: 9px; margin-top: 3px; }

    #sheet-content.ultra-compact .title-banner { padding: 3px 7px; margin-bottom: 3px; }
    #sheet-content.ultra-compact .title-main { font-size: 11px; }
    #sheet-content.ultra-compact .area-block { margin-bottom: 3px; }
    #sheet-content.ultra-compact .area-banner { padding: 1.5px 5px; font-size: 9px; }
    #sheet-content.ultra-compact th,
    #sheet-content.ultra-compact td { padding: 1.2px 3px; font-size: 8.2px; line-height: 1.12; }
    #sheet-content.ultra-compact th { font-size: 7.6px; }
    #sheet-content.ultra-compact .grand-total { padding: 3px 7px; font-size: 8.5px; margin-top: 2px; }
  </style>
</head>
<body>
  <div id="mm-ruler" style="position:absolute;top:-9999px;left:0;width:100mm;height:100mm;visibility:hidden;pointer-events:none;"></div>
  <div id="print-viewport">
    <div id="sheet-content">
      <div class="title-banner">
        <div>
          <div class="title-main">${escapeHtml(titulo)}</div>
          <div class="title-sub">${escapeHtml(subtitulo)}</div>
        </div>
        <div class="title-stats">
          Total: <strong>${totalGen.toLocaleString('es-PE')} u</strong>
          ${mermaGen > 0 ? `<span class="merma-badge">Merma: ${mermaGen.toLocaleString('es-PE')} u</span>` : ''}
          <span style="color:#A0AEC0;margin-left:8px;">(${filas.length} filas)</span>
        </div>
      </div>
      ${bloquesAreas}
      <div class="grand-total">
        <span>RESUMEN GENERAL DEL TURNO (${filas.length} REGISTROS)</span>
        <div>
          TOTAL PRODUCIDO: <span class="val">${totalGen.toLocaleString('es-PE')} u</span>
          <span class="merma-val">MERMA TOTAL: ${mermaGen.toLocaleString('es-PE')} u</span>
        </div>
      </div>
    </div>
  </div>
</body>
</html>`;

  // Render inside an off-screen A4-sized iframe so DOM height is accurately measured and scaled to 1 page
  const iframe = document.createElement('iframe');
  iframe.style.position = 'fixed';
  iframe.style.left = '-12000px';
  iframe.style.top = '0';
  iframe.style.width = '1200px';
  iframe.style.height = '900px';
  iframe.style.border = '0';
  iframe.style.opacity = '0';
  iframe.style.pointerEvents = 'none';
  iframe.setAttribute('aria-hidden', 'true');
  document.body.appendChild(iframe);

  const doc = iframe.contentWindow?.document;
  if (!doc || !iframe.contentWindow) {
    document.body.removeChild(iframe);
    return;
  }

  doc.open();
  doc.write(html);
  doc.close();

  const fitSingleLandscapePage = () => {
    const ruler = doc.getElementById('mm-ruler');
    const viewport = doc.getElementById('print-viewport');
    const content = doc.getElementById('sheet-content');
    if (!ruler || !viewport || !content) return;

    const rulerRect = ruler.getBoundingClientRect();
    const pxPerMm = rulerRect.height > 0 ? rulerRect.height / 100 : 3.7795275591;
    const widthMm = 285;
    const maxHeightMm = 190; // Safe single-page budget inside 210mm A4 height with 5mm top/bottom margins
    const maxH = maxHeightMm * pxPerMm;

    content.style.width = `${widthMm}mm`;

    // 1) Try compact density tiers before applying geometric transform scaling
    if (content.scrollHeight > maxH) {
      content.classList.add('compact');
    }
    if (content.scrollHeight > maxH * 1.18) {
      content.classList.add('ultra-compact');
    }

    // 2) If content still exceeds single-page height, solve optimal scale via binary search:
    // widening virtual width to (285 / scale)mm prevents narrow wrapping while uniform scale keeps full width & <=190mm height
    let bestScale = 1;
    if (content.scrollHeight > maxH) {
      let low = 0.25;
      let high = 1.0;
      bestScale = 0.25;

      for (let step = 0; step < 14; step++) {
        const mid = (low + high) / 2;
        content.style.width = `${(widthMm / mid).toFixed(3)}mm`;
        const h = content.scrollHeight;
        if (h * mid <= maxH * 0.985) {
          bestScale = mid;
          low = mid;
        } else {
          high = mid;
        }
      }

      content.style.width = `${(widthMm / bestScale).toFixed(3)}mm`;
      content.style.transform = `scale(${bestScale.toFixed(5)})`;
      content.style.transformOrigin = 'top left';
    }

    const scaledHeightMm = Math.min(maxHeightMm, (content.scrollHeight / pxPerMm) * bestScale + 0.5);
    viewport.style.width = `${widthMm}mm`;
    viewport.style.height = `${scaledHeightMm.toFixed(2)}mm`;
    viewport.style.overflow = 'hidden';

    ruler.remove();
  };

  const cleanup = () => {
    if (iframe.parentNode) {
      iframe.parentNode.removeChild(iframe);
    }
  };

  iframe.contentWindow.addEventListener('afterprint', cleanup, { once: true });
  setTimeout(() => {
    try {
      fitSingleLandscapePage();
      iframe.contentWindow?.focus();
      iframe.contentWindow?.print();
    } finally {
      setTimeout(cleanup, 60000);
    }
  }, 180);
}

// ----------------- ACTIVACION EXCEL & PDF -----------------
export function filasActivacion(maquinas: Maquina[], turno: TurnoMaquina[], catalogo: CatalogoItem[]): any[] {
  const filas: any[] = [];
  const getR = (area: string) => (area === 'vasos_envases' ? 0 : area === 'papel_carton' ? 1 : 2);
  const activas = maquinas
    .filter((m) => m.activa)
    .sort((a, b) => {
      const ra = getR(a.area);
      const rb = getR(b.area);
      if (ra !== rb) return ra - rb;
      return a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' });
    });

  activas.forEach((m) => {
    turnosDeMaquina(m, turno).forEach((t) => {
      const o = opsDeTurno(t)[0] || { op: '', modeloIdx: -1, pliego: '' };
      const mt = buscarEnCatalogo(catalogo, o.op);
      const d = mt ? resolverModelo(mt, o.modeloIdx).descripcion : '';
      filas.push({
        maqId: m.id,
        uid: `${m.id}#${t.instancia || 1}`,
        maquina: m.nombre,
        codigo: m.codigo || '',
        area: m.area || '',
        parada: !!t.motivoParada,
        motivo: t.motivoParada || '',
        op: o.op,
        producto: [d, etiquetaPliego(o.pliego)].filter(Boolean).join(' - '),
        personal: t.personal || '',
        observaciones: t.observaciones || '',
      });
    });
  });

  return filas;
}

export function contarMaquinasAct(filas: any[]): { total: number; paradas: number } {
  const ids: Record<string, boolean> = {};
  const pa: Record<string, boolean> = {};
  filas.forEach((f) => {
    ids[f.uid] = true;
    if (f.parada) pa[f.uid] = true;
  });
  return { total: Object.keys(ids).length, paradas: Object.keys(pa).length };
}

export async function exportarActivacionExcel(filas: any[], titulo: string, nombreBase: string) {
  try {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Activacion');
    ws.columns = [{ width: 18 }, { width: 13 }, { width: 12 }, { width: 20 }, { width: 14 }, { width: 40 }, { width: 28 }, { width: 30 }];
    ws.mergeCells(1, 1, 1, 8);
    const rt = ws.getRow(1);
    rt.height = 22;
    rt.getCell(1).value = titulo;
    rt.getCell(1).font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    rt.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF12161F' } };
    rt.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    let fila = 2;
    AREAS_ORDEN.forEach((areaDef) => {
      const de = filas.filter((f) => (f.area || '') === areaDef.key);
      if (de.length === 0) return;
      ws.mergeCells(fila, 1, fila, 8);
      const a = ws.getRow(fila).getCell(1);
      a.value = areaDef.label;
      a.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      a.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD97706' } };
      a.alignment = { horizontal: 'left', vertical: 'middle' };
      fila++;

      const rh = ws.getRow(fila);
      ACT_HEADERS.forEach((h, i) => {
        const c = rh.getCell(i + 1);
        c.value = h;
        c.font = { bold: true, color: { argb: 'FF333333' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
        c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        c.border = EXCELJS_BORDES;
      });
      fila++;

      let previo: string | null = null;
      let inicioRun = fila;
      const cerrarRun = (desde: number, hasta: number) => {
        if (hasta > desde) {
          [1, 2].forEach((col) => {
            ws.mergeCells(desde, col, hasta, col);
          });
        }
      };

      de.forEach((f) => {
        if (previo !== null && f.maqId !== previo) {
          cerrarRun(inicioRun, fila - 1);
          inicioRun = fila;
        }
        previo = f.maqId;
        const r = ws.getRow(fila);
        [f.maquina, f.codigo, f.parada ? 'PARADA' : 'ACTIVA', f.motivo, f.op || '—', f.producto, f.personal, f.observaciones].forEach((v, i) => {
          const c = r.getCell(i + 1);
          c.value = v;
          c.border = EXCELJS_BORDES;
          c.alignment = { vertical: 'middle', wrapText: true };
          if (i === 0 || i === 2) c.font = { bold: true };
          if (f.parada) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF176' } };
        });
        fila++;
      });
      cerrarRun(inicioRun, fila - 1);
    });

    const cnt = contarMaquinasAct(filas);
    ws.mergeCells(fila, 1, fila, 8);
    const rTot = ws.getRow(fila).getCell(1);
    rTot.value = `TOTAL: ${cnt.total} máquina(s) · ${cnt.paradas} parada(s)`;
    rTot.font = { bold: true };
    rTot.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDEBD0' } };
    rTot.border = EXCELJS_BORDES;
    ws.views = [{ state: 'frozen', ySplit: 1 }];

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    descargarBlob(blob, `${nombreBase}.xlsx`);
  } catch (e) {
    const aoa: any[][] = [[titulo], []];
    AREAS_ORDEN.forEach((areaDef) => {
      const de = filas.filter((f) => (f.area || '') === areaDef.key);
      if (de.length === 0) return;
      aoa.push([areaDef.label]);
      aoa.push(ACT_HEADERS);
      de.forEach((f) => {
        aoa.push([f.maquina, f.codigo, f.parada ? 'PARADA' : 'ACTIVA', f.motivo, f.op || '—', f.producto, f.personal, f.observaciones]);
      });
      aoa.push([]);
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 18 }, { wch: 13 }, { wch: 12 }, { wch: 20 }, { wch: 14 }, { wch: 40 }, { wch: 28 }, { wch: 30 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Activacion');
    const wbArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    descargarBlob(blob, `${nombreBase}.xlsx`);
  }
}

export function imprimirActivacion(filas: any[], fechaTxt: string, turnoTxt: string, nombreArchivo: string) {
  const cntAct = contarMaquinasAct(filas);

  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4',
  });

  const pageWidth = 297;
  const pageHeight = 210;
  const marginX = 12;
  const contentWidth = pageWidth - marginX * 2; // 273 mm

  // Column widths summing strictly to 273 mm
  const colWidths = [32, 21, 18, 35, 20, 63, 40, 44];
  const colX: number[] = [];
  let accX = marginX;
  colWidths.forEach((w) => {
    colX.push(accX);
    accX += w;
  });

  let y = 12;

  const drawMainHeader = () => {
    doc.setFillColor(18, 22, 31); // #12161F
    doc.rect(marginX, y, contentWidth, 16, 'F');
    doc.setFillColor(242, 169, 59); // #F2A93B gold accent
    doc.rect(marginX, y, 3.5, 16, 'F');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(13.5);
    doc.setTextColor(255, 255, 255);
    doc.text('ACTIVACIÓN DE MÁQUINAS', marginX + 7, y + 6.8);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.5);
    doc.setTextColor(203, 213, 224);
    const sub = `${fechaTxt}${turnoTxt ? '  ·  ' + turnoTxt : ''}`;
    doc.text(sub, marginX + 7, y + 12.5);

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9.5);
    doc.setTextColor(242, 169, 59);
    doc.text(
      `${cntAct.total} máquina(s)  ·  ${cntAct.paradas} parada(s)`,
      marginX + contentWidth - 4,
      y + 9.5,
      { align: 'right' }
    );

    y += 19;
  };

  const drawAreaAndTableHeaders = (areaLabel: string, isContinuation = false) => {
    // Area header banner (#D97706)
    doc.setFillColor(217, 119, 6);
    doc.setDrawColor(183, 183, 183);
    doc.setLineWidth(0.2);
    doc.rect(marginX, y, contentWidth, 6.5, 'FD');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.5);
    doc.setTextColor(255, 255, 255);
    doc.text(
      isContinuation ? `${areaLabel} (CONTINUACIÓN)` : areaLabel,
      marginX + 3,
      y + 4.4
    );
    y += 6.5;

    // Column headers (#F2F2F2)
    doc.setFillColor(242, 242, 242);
    doc.rect(marginX, y, contentWidth, 6.2, 'FD');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.2);
    doc.setTextColor(51, 51, 51);
    ACT_HEADERS.forEach((h, i) => {
      doc.line(colX[i], y, colX[i], y + 6.2);
      doc.text(h, colX[i] + colWidths[i] / 2, y + 4.2, { align: 'center' });
    });
    doc.line(marginX + contentWidth, y, marginX + contentWidth, y + 6.2);
    y += 6.2;
  };

  drawMainHeader();

  AREAS_ORDEN.forEach((areaDef) => {
    const de = filas.filter((f) => (f.area || '') === areaDef.key);
    if (de.length === 0) return;

    // Ensure enough room for area header + table header + at least 1 row
    if (y + 22 > pageHeight - 14) {
      doc.addPage('a4', 'landscape');
      y = 12;
      drawMainHeader();
    }

    drawAreaAndTableHeaders(areaDef.label, false);

    let prevMaqKey: string | null = null;

    de.forEach((f, idx) => {
      const maqKey = String(f.maqId || f.maquina || '');
      const isSameAsPrev = prevMaqKey !== null && prevMaqKey === maqKey;
      const nextIsSame = idx + 1 < de.length && String(de[idx + 1].maqId || de[idx + 1].maquina || '') === maqKey;

      const values = [
        isSameAsPrev ? '' : String(f.maquina || ''),
        isSameAsPrev ? '' : String(f.codigo || ''),
        f.parada ? 'PARADA' : 'ACTIVA',
        String(f.motivo || ''),
        String(f.op || '—'),
        String(f.producto || ''),
        String(f.personal || ''),
        String(f.observaciones || ''),
      ];

      doc.setFontSize(7.2);
      const wrappedLines = values.map((val, cIdx) => {
        if (!val) return [''];
        return doc.splitTextToSize(val, colWidths[cIdx] - 2.6) as string[];
      });

      const maxLines = Math.max(...wrappedLines.map((l) => l.length), 1);
      const rowH = Math.max(5.8, maxLines * 3.3 + 2.4);

      if (y + rowH > pageHeight - 14) {
        doc.addPage('a4', 'landscape');
        y = 12;
        drawAreaAndTableHeaders(areaDef.label, true);
        // After page break, reprint machine name/code for clarity
        wrappedLines[0] = f.maquina ? (doc.splitTextToSize(String(f.maquina), colWidths[0] - 2.6) as string[]) : [''];
        wrappedLines[1] = f.codigo ? (doc.splitTextToSize(String(f.codigo), colWidths[1] - 2.6) as string[]) : [''];
      }

      // Row background
      if (f.parada) {
        doc.setFillColor(255, 241, 118); // #FFF176 yellow highlight
        doc.rect(marginX, y, contentWidth, rowH, 'F');
      } else if (idx % 2 === 1) {
        doc.setFillColor(250, 250, 250);
        doc.rect(marginX, y, contentWidth, rowH, 'F');
      }

      // Draw borders cleanly (omit top line in cols 0 & 1 when merged with previous row on same page)
      doc.setDrawColor(183, 183, 183);
      doc.setLineWidth(0.2);

      // Vertical borders
      colX.forEach((cx) => {
        doc.line(cx, y, cx, y + rowH);
      });
      doc.line(marginX + contentWidth, y, marginX + contentWidth, y + rowH);

      // Bottom border (for cols 0..1 omit if next row on same machine, else draw full)
      if (!nextIsSame) {
        doc.line(marginX, y + rowH, marginX + contentWidth, y + rowH);
      } else {
        doc.line(colX[2], y + rowH, marginX + contentWidth, y + rowH);
      }

      // Cell text
      wrappedLines.forEach((lines, cIdx) => {
        if (lines.length === 1 && lines[0] === '') return;

        if (cIdx === 0 || cIdx === 2) {
          doc.setFont('helvetica', 'bold');
        } else {
          doc.setFont('helvetica', 'normal');
        }

        if (cIdx === 2 && f.parada) {
          doc.setTextColor(185, 28, 28); // Red bold for PARADA
        } else if (cIdx === 2) {
          doc.setTextColor(21, 128, 61); // Green bold for ACTIVA
        } else {
          doc.setTextColor(17, 24, 39);
        }

        const totalTextH = lines.length * 3.3;
        let startTextY = y + (rowH - totalTextH) / 2 + 2.6;

        lines.forEach((line) => {
          if (cIdx === 1 || cIdx === 2 || cIdx === 4) {
            doc.text(line, colX[cIdx] + colWidths[cIdx] / 2, startTextY, { align: 'center' });
          } else {
            doc.text(line, colX[cIdx] + 1.4, startTextY);
          }
          startTextY += 3.3;
        });
      });

      prevMaqKey = maqKey;
      y += rowH;
    });

    y += 4;
  });

  // Totals footer bar (#FDEBD0)
  if (y + 9 > pageHeight - 14) {
    doc.addPage('a4', 'landscape');
    y = 14;
  }

  doc.setFillColor(253, 235, 208);
  doc.setDrawColor(183, 183, 183);
  doc.setLineWidth(0.25);
  doc.rect(marginX, y, contentWidth, 7, 'FD');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(17, 24, 39);
  doc.text(
    `TOTAL: ${cntAct.total} máquina(s)  ·  ${cntAct.paradas} parada(s)`,
    marginX + 3,
    y + 4.7
  );

  // Page footer on all pages
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
      `Generado: ${generado}  ·  Activación de Máquinas`,
      marginX,
      pageHeight - 6.2
    );
    doc.text(`Página ${p} de ${totalPages}`, marginX + contentWidth, pageHeight - 6.2, {
      align: 'right',
    });
  }

  doc.save(`${nombreArchivo}.pdf`);
}

// ----------------- MIGRACION EXCEL -----------------
export function exportarMigracion(todasLasFilas: FilaTurnoExcel[], fecha: string, tipo: string) {
  const filas = todasLasFilas.filter((f) => f.tipoProducto === 'pt' && f.total > 0);
  if (filas.length === 0) {
    throw new Error('No hay producción PT para migrar');
  }
  const aoa = [['ODTCOD', 'CODPTSAP', 'CODMAQ', 'ODTCANT', 'ODTOBS']];
  filas.forEach((f) => {
    aoa.push([f.op || '', f.sap || '', f.codigoMaquina || '', String(f.total), f.observaciones || '']);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 30 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Hoja1');
  const wbArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  descargarBlob(blob, `migracion_${fecha}_${tipo}.xlsx`);
}
