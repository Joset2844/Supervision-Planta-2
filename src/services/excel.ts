import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
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
function buscarBase(mapa: Record<string, any>, op: string) {
  if (mapa[op]) return mapa[op];
  const k = Object.keys(mapa).find((x) => x.trim().toLowerCase() === op.trim().toLowerCase());
  return k ? mapa[k] : null;
}

function parsearHojaCatalogo(sheet: any) {
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (rows.length < 2) return null;
  const headers = rows[0].map(normalizar);
  const idxOp = headers.findIndex((h) => h.indexOf('op') !== -1);
  const idxDesc = headers.findIndex((h) => h.indexOf('descripcion') !== -1);
  const idxCliente = headers.findIndex((h) => h.indexOf('cliente') !== -1);
  const idxCantidad = headers.findIndex((h) => h.indexOf('cantidad') !== -1 || h.indexOf('total') !== -1);
  const idxSap = headers.findIndex((h) => h.indexOf('sap') !== -1);
  if (idxOp === -1 || idxDesc === -1 || idxCliente === -1) return null;

  const filasValidas = rows.slice(1).filter((r) => String(r[idxOp]).trim() !== '');
  const mapaOps: Record<string, any> = {};
  const ordenOps: string[] = [];

  filasValidas.forEach((r) => {
    const op = String(r[idxOp]).trim();
    const cliente = String(r[idxCliente]).trim();
    const cantidad = idxCantidad !== -1 ? String(r[idxCantidad]).trim() : '';
    const modelo: CatalogoModelo = {
      descripcion: String(r[idxDesc]).trim(),
      sap: idxSap !== -1 ? String(r[idxSap]).trim() : '',
      cantidadTotal: cantidad,
    };
    if (!mapaOps[op]) {
      mapaOps[op] = { op, cliente, cantidadTotal: 0, modelos: [] };
      ordenOps.push(op);
    }
    const yaExiste = mapaOps[op].modelos.some((mo: any) => {
      return mo.descripcion === modelo.descripcion && (mo.sap || '') === (modelo.sap || '');
    });
    if (!yaExiste) mapaOps[op].modelos.push(modelo);
    const n = parseInt(cantidad, 10);
    if (!isNaN(n)) mapaOps[op].cantidadTotal += n;
  });
  return { mapaOps, ordenOps };
}

function parsearHojaModelos(sheet: any, mapaBase: Record<string, any>) {
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (rows.length < 2) return null;
  const headers = rows[0].map(normalizar);
  const idxOp = headers.findIndex((h) => h.indexOf('op') !== -1);
  const idxModelo = headers.findIndex((h) => h.indexOf('modelo') !== -1);
  const idxCantidad = headers.findIndex((h) => h.indexOf('cantidad') !== -1 || h.indexOf('total') !== -1);
  const idxSap = headers.findIndex((h) => h.indexOf('sap') !== -1);
  if (idxOp === -1 || idxModelo === -1) return null;

  const filasValidas = rows.slice(1).filter((r) => String(r[idxOp]).trim() !== '');
  const mapaOps: Record<string, any> = {};
  const ordenOps: string[] = [];

  filasValidas.forEach((r) => {
    const op = String(r[idxOp]).trim();
    const modeloTxt = String(r[idxModelo]).trim();
    const cantidad = idxCantidad !== -1 ? String(r[idxCantidad]).trim() : '';
    const base = buscarBase(mapaBase, op);
    const descBase = base ? (base.modelos && base.modelos[0] ? base.modelos[0].descripcion : (base.descripcion || '')) : '';
    const descripcionFinal = descBase ? `${descBase} - ${modeloTxt}` : modeloTxt;
    const modelo: CatalogoModelo = {
      modelo: modeloTxt,
      descripcion: descripcionFinal,
      sap: idxSap !== -1 ? String(r[idxSap]).trim() : '',
      cantidadTotal: cantidad,
    };
    if (!mapaOps[op]) {
      mapaOps[op] = { op, cliente: base ? base.cliente : '', cantidadTotal: 0, modelos: [] };
      ordenOps.push(op);
    }
    const yaExisteModelo = mapaOps[op].modelos.some((mo: any) => {
      return mo.modelo === modeloTxt && mo.sap === modelo.sap;
    });
    if (!yaExisteModelo) mapaOps[op].modelos.push(modelo);
    const n = parseInt(cantidad, 10);
    if (!isNaN(n)) mapaOps[op].cantidadTotal += n;
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
        if (!mapaFinal[op]) ordenFinal.push(op);
        const base = buscarBase(principal.mapaOps, op);
        const descripcionBase = base && base.modelos[0] ? base.modelos[0].descripcion : '';
        mapaFinal[op] = modelos.mapaOps[op];
        mapaFinal[op].descripcionBase = descripcionBase;
      });
    }
  }

  const nuevo: CatalogoItem[] = ordenFinal.map((op) => {
    const item = mapaFinal[op];
    if (item.modelos.length > 1) {
      item.descripcion = 'Varios modelos (' + item.modelos.length + ')';
      item.sap = '';
      if (!item.descripcionBase) item.descripcionBase = item.modelos[0].descripcion;
    } else {
      item.descripcion = item.modelos[0].descripcion;
      item.sap = item.modelos[0].sap;
      item.descripcionBase = item.modelos[0].descripcion;
    }
    return item;
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

  const copiaMaquinas = [...maquinasActuales];
  let actualizadas = 0;
  let creadas = 0;

  rows.slice(1).forEach((r) => {
    const nombre = String(r[idxMaquina]).trim();
    const codigo = String(r[idxCodigo]).trim();
    const tipoNuevo = leerTipo(r);
    const areaNueva = leerArea(r);
    if (!nombre) return;
    const existente = copiaMaquinas.find((m) => m.nombre.trim().toLowerCase() === nombre.toLowerCase());
    if (existente) {
      existente.codigo = codigo;
      existente.bloqueada = true;
      if (tipoNuevo) existente.tipoProducto = tipoNuevo;
      if (areaNueva) existente.area = areaNueva;
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

export function exportarInventarioExcel(lista: ListaItem[], catalogo: CatalogoItem[]) {
  const agrupado = agruparParaExportar(lista);
  const filas = agrupado.map((i: any) => {
    const match = buscarEnCatalogo(catalogo, i.op);
    const empaqueTxt = i.tipoEmpaque === 'paquete' ? 'Paquete' : 'Caja';
    return {
      OP: i.op,
      Descripción: i.descripcion,
      Cliente: i.cliente,
      SAP: i.sap || (match && match.sap ? match.sap : ''),
      Empaque: empaqueTxt,
      'Unidades x empaque': i.unidadesXCaja,
      'Cantidad de empaques': i.cantidadCajas,
      'Total unidades': i.totalUnidades,
      Observación: i.observacion || '',
    };
  });
  const ws = XLSX.utils.json_to_sheet(filas);
  ws['!cols'] = [{ wch: 14 }, { wch: 36 }, { wch: 22 }, { wch: 12 }, { wch: 10 }, { wch: 16 }, { wch: 18 }, { wch: 14 }, { wch: 30 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Inventario');
  const hoy = new Date().toISOString().slice(0, 10);
  const nombre = `inventario_${hoy}.xlsx`;
  const wbArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  descargarBlob(blob, nombre);
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
  function bloque(areaDef: { key: string; label: string }) {
    const de = filas.filter((f) => (f.area || '') === areaDef.key);
    if (de.length === 0) return '';
    let h = `<h2>${areaDef.label}</h2><table><thead><tr>` + ACT_HEADERS.map((x) => `<th>${x}</th>`).join('') + `</tr></thead><tbody>`;
    let i = 0;
    while (i < de.length) {
      let j = i;
      while (j + 1 < de.length && de[j + 1].maqId === de[i].maqId) j++;
      const n = j - i + 1;
      const f0 = de[i];
      const rs = n > 1 ? ` rowspan='${n}'` : '';
      for (let k = i; k <= j; k++) {
        const f = de[k];
        h += `<tr class='${f.parada ? 'parada' : ''}'>`;
        if (k === i) h += `<td${rs}><b>${escapeHtml(f0.maquina)}</b></td><td${rs}>${escapeHtml(f0.codigo)}</td>`;
        h += `<td><b>${f.parada ? 'PARADA' : 'ACTIVA'}</b></td><td>${escapeHtml(f.motivo)}</td><td>${escapeHtml(f.op || '—')}</td><td>${escapeHtml(
          f.producto
        )}</td><td>${escapeHtml(f.personal)}</td><td>${escapeHtml(f.observaciones)}</td></tr>`;
      }
      i = j + 1;
    }
    return h + `</tbody></table>`;
  }

  const cntAct = contarMaquinasAct(filas);
  const html =
    `<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'><title>${nombreArchivo}</title><style>` +
    `@page{size:A4 landscape;margin:10mm;}body{font-family:Arial,Helvetica,sans-serif;color:#111;font-size:11px;-webkit-print-color-adjust:exact;print-color-adjust:exact;}` +
    `h1{font-size:16px;margin:0 0 2px;}h2{font-size:13px;background:#d97706;color:#fff;padding:4px 6px;margin:12px 0 4px;}` +
    `table{width:100%;border-collapse:collapse;margin-bottom:4px;}th,td{border:1px solid #999;padding:3px 5px;text-align:left;vertical-align:middle;}` +
    `th{background:#f2f2f2;text-align:center;}tr.parada td{background:#fff176;}` +
    `.top{display:flex;justify-content:space-between;align-items:baseline;border-bottom:2px solid #333;padding-bottom:6px;margin-bottom:4px;}` +
    `</style></head><body><div class='top'><div><h1>Activación de máquinas</h1><div>${fechaTxt}${
      turnoTxt ? ' · ' + turnoTxt : ''
    }</div></div><div><b>${cntAct.total}</b> máquina(s) · <b>${cntAct.paradas}</b> parada(s)</div></div>${AREAS_ORDEN.map(bloque).join('')}</body></html>`;

  const w = window.open('', '_blank');
  if (!w) {
    alert('El navegador bloqueó la ventana de impresión');
    return;
  }
  w.document.open();
  w.document.write(html);
  w.document.close();
  w.onload = function () {
    w.focus();
    w.print();
  };
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
