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

function opDigits(op: any): string {
  return String(op || '').replace(/\D/g, '');
}

function normalizarClaveOp(op: any): string {
  let s = String(op || '').trim().toUpperCase();
  if (!s) return '';
  s = s.replace(/^(?:O\.?\s*P\.?|ODT|N[°º.]?|#)\s*[-:#.]?\s*/i, '').trim();
  if (/^\d+$/.test(s)) {
    s = s.replace(/^0+/, '') || '0';
  }
  return s;
}

function opSoloPrefijoYDigitos(op: any): boolean {
  const s = String(op || '').trim().replace(/^(?:O\.?\s*P\.?|ODT|N[°º.]?|#)\s*[-:#.]?\s*/i, '').trim();
  return /^\d+$/.test(s);
}

function opDigitsSinCeros(op: any): string {
  const d = opDigits(op).replace(/^0+/, '');
  return d || opDigits(op);
}

export function normalizarCatalogoItem(raw: any): CatalogoItem | null {
  if (!raw || typeof raw !== 'object') return null;
  const opStr = String(raw.op ?? '').trim();
  if (!opStr) return null;

  const rawModelos = Array.isArray(raw.modelos) ? raw.modelos : [];
  const modelos: CatalogoModelo[] = rawModelos
    .filter((m: any) => m && typeof m === 'object')
    .map((m: any) => {
      const cant = parsearCantidad(m.cantidadTotal);
      return {
        ...(m.modelo ? { modelo: String(m.modelo).trim() } : {}),
        descripcion: String(m.descripcion || m.modelo || '').trim(),
        sap: String(m.sap || '').trim(),
        cantidadTotal: cant,
      };
    });

  if (modelos.length === 0) {
    modelos.push({
      descripcion: String(raw.descripcion || '').trim(),
      sap: String(raw.sap || '').trim(),
      cantidadTotal: parsearCantidad(raw.cantidadTotal),
    });
  }

  let totalOp = parsearCantidad(raw.cantidadTotal);
  const sumModelos = modelos.reduce((acc, m) => acc + parsearCantidad(m.cantidadTotal), 0);
  if (sumModelos > 0) {
    totalOp = sumModelos;
  }

  const tieneMultiples = modelos.length > 1;
  const descripcionBase =
    String(raw.descripcionBase || '').trim() ||
    modelos[0]?.descripcion ||
    String(raw.descripcion || '').trim();

  return {
    op: opStr,
    cliente: String(raw.cliente || '').trim(),
    cantidadTotal: totalOp,
    descripcion: tieneMultiples
      ? `Varios modelos (${modelos.length})`
      : modelos[0]?.descripcion || String(raw.descripcion || '').trim(),
    sap: tieneMultiples ? '' : modelos[0]?.sap || String(raw.sap || '').trim(),
    descripcionBase,
    modelos,
  };
}

export function buscarEnCatalogo(catalogo: CatalogoItem[], op: string): CatalogoItem | null {
  if (!op || !Array.isArray(catalogo) || catalogo.length === 0) return null;
  const q = String(op).trim().toLowerCase();
  if (!q) return null;

  const exactMatches = catalogo.filter((c) => String(c.op || '').trim().toLowerCase() === q);
  if (exactMatches.length === 1) return exactMatches[0];
  if (exactMatches.length > 1) {
    return exactMatches.reduce((best, cur) =>
      (cur.modelos?.length || 0) > (best.modelos?.length || 0) ? cur : best
    );
  }

  const qNorm = normalizarClaveOp(q);
  if (qNorm) {
    const byNorm = catalogo.filter((c) => normalizarClaveOp(c.op) === qNorm);
    if (byNorm.length > 0) {
      return byNorm.reduce((best, cur) =>
        (cur.modelos?.length || 0) > (best.modelos?.length || 0) ? cur : best
      );
    }
  }

  if (opSoloPrefijoYDigitos(q)) {
    const digitsNoZero = opDigitsSinCeros(q);
    if (digitsNoZero) {
      const byDigits = catalogo.filter(
        (c) => opSoloPrefijoYDigitos(c.op) && opDigitsSinCeros(c.op) === digitsNoZero
      );
      if (byDigits.length > 0) {
        return byDigits.reduce((best, cur) =>
          (cur.modelos?.length || 0) > (best.modelos?.length || 0) ? cur : best
        );
      }
    }
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
  if (
    s.includes('modelo') ||
    s.includes('caja') ||
    s.includes('paquete') ||
    s.includes('min') ||
    s.includes('merma')
  ) {
    return false;
  }
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
    s === 'ctdad' ||
    s === 'odtcant'
  );
}

function encontrarColumnaOp(headers: string[]): number {
  const exactos = ['op', 'ops', "op's", 'odtcod', 'odt', 'nro op', 'n op', 'num op', 'numero op', 'codigo op', 'cod op'];
  let idx = headers.findIndex((h) => exactos.includes(h));
  if (idx !== -1) return idx;
  idx = headers.findIndex((h) => /\b(op|odt)\b/.test(h) && !h.includes('operad') && !h.includes('operar') && !h.includes('opcion') && !h.includes('copia'));
  if (idx !== -1) return idx;
  return headers.findIndex(
    (h) =>
      h.includes('op') &&
      !h.includes('operad') &&
      !h.includes('operar') &&
      !h.includes('opcion') &&
      !h.includes('copia') &&
      !h.includes('soporte') &&
      !h.includes('propied')
  );
}

function encontrarColumnaSap(headers: string[]): number {
  let idx = headers.findIndex((h) => h === 'sap' || h === 'codptsap' || h.includes('sap'));
  if (idx !== -1) return idx;
  idx = headers.findIndex(
    (h) =>
      (h === 'codigo' || h.includes('codigo producto') || h.includes('cod producto') || h.includes('codigo pt') || h.includes('cod pt')) &&
      !h.includes('op') &&
      !h.includes('cliente') &&
      !h.includes('maq')
  );
  if (idx !== -1) return idx;
  return headers.findIndex((h) => h.includes('codigo') && !h.includes('op') && !h.includes('cliente') && !h.includes('maq'));
}

function encontrarColumnaCantidad(headers: string[]): number {
  const exactos = ['cantidad', 'cantidad total', 'cant total', 'cant', 'cant.', 'total', 'unidades', 'odtcant'];
  const idx = headers.findIndex((h) => exactos.includes(h));
  if (idx !== -1) return idx;
  return headers.findIndex(esColumnaCantidad);
}

function encontrarFilaEncabezados(rows: any[][], validar: (headers: string[]) => boolean): { rowIdx: number; headers: string[] } | null {
  const maxScan = Math.min(rows.length - 1, 8);
  for (let i = 0; i < maxScan; i++) {
    const headers = (rows[i] || []).map(normalizar);
    if (validar(headers)) {
      return { rowIdx: i, headers };
    }
  }
  return null;
}

function buscarClaveBase(mapa: Record<string, any>, op: string): string | null {
  if (mapa[op]) return op;
  const qLower = op.trim().toLowerCase();
  const exactKey = Object.keys(mapa).find((x) => x.trim().toLowerCase() === qLower);
  if (exactKey) return exactKey;

  const qNorm = normalizarClaveOp(op);
  if (qNorm) {
    const normKey = Object.keys(mapa).find((x) => normalizarClaveOp(x) === qNorm);
    if (normKey) return normKey;
  }

  if (opSoloPrefijoYDigitos(op)) {
    const qDigits = opDigitsSinCeros(op);
    if (qDigits) {
      const digitKey = Object.keys(mapa).find(
        (x) => opSoloPrefijoYDigitos(x) && opDigitsSinCeros(x) === qDigits
      );
      if (digitKey) return digitKey;
    }
  }
  return null;
}

function buscarBase(mapa: Record<string, any>, op: string) {
  const k = buscarClaveBase(mapa, op);
  return k ? mapa[k] : null;
}

function parsearHojaCatalogo(sheet: any) {
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (rows.length < 2) return null;

  const found = encontrarFilaEncabezados(rows, (headers) => {
    const idxOp = encontrarColumnaOp(headers);
    const idxDesc = headers.findIndex((h) => h.includes('descripcion') || h.includes('producto') || h.includes('articulo'));
    const idxCliente = headers.findIndex((h) => h.includes('cliente') || h.includes('razon') || h.includes('empresa'));
    return idxOp !== -1 && idxDesc !== -1 && idxCliente !== -1;
  });
  if (!found) return null;

  const { rowIdx, headers } = found;
  const idxOp = encontrarColumnaOp(headers);
  const idxDesc =
    headers.findIndex((h) => h === 'descripcion' || h === 'producto') !== -1
      ? headers.findIndex((h) => h === 'descripcion' || h === 'producto')
      : headers.findIndex((h) => h.includes('descripcion') || h.includes('producto') || h.includes('articulo'));
  const idxModelo =
    headers.findIndex((h) => h === 'modelo') !== -1
      ? headers.findIndex((h) => h === 'modelo')
      : headers.findIndex((h) => h.includes('modelo'));
  const idxCliente = headers.findIndex((h) => h.includes('cliente') || h.includes('razon') || h.includes('empresa'));
  const idxCantidad = encontrarColumnaCantidad(headers);
  const idxSap = encontrarColumnaSap(headers);

  const filasValidas = rows.slice(rowIdx + 1).filter((r) => String(r[idxOp] ?? '').trim() !== '');
  const mapaOps: Record<string, any> = {};
  const ordenOps: string[] = [];

  filasValidas.forEach((r) => {
    const opRaw = String(r[idxOp] ?? '').trim();
    const claveExistente = buscarClaveBase(mapaOps, opRaw);
    const opKey = claveExistente || opRaw;
    const cliente = String(r[idxCliente] ?? '').trim();
    const cantNum = idxCantidad !== -1 ? parsearCantidad(r[idxCantidad]) : 0;
    const descBase = String(r[idxDesc] ?? '').trim();
    const modeloTxt = idxModelo !== -1 ? String(r[idxModelo] ?? '').trim() : '';
    const desc =
      modeloTxt && !descBase.toLowerCase().includes(modeloTxt.toLowerCase())
        ? `${descBase} - ${modeloTxt}`
        : descBase || modeloTxt;
    const sap = idxSap !== -1 ? String(r[idxSap] ?? '').trim() : '';
    const modelo: CatalogoModelo = {
      ...(modeloTxt ? { modelo: modeloTxt } : {}),
      descripcion: desc,
      sap: sap,
      cantidadTotal: cantNum,
    };
    if (!mapaOps[opKey]) {
      mapaOps[opKey] = { op: opKey, cliente, cantidadTotal: 0, descripcionBase: descBase, modelos: [] };
      ordenOps.push(opKey);
    } else if (!mapaOps[opKey].cliente && cliente) {
      mapaOps[opKey].cliente = cliente;
    }
    const idxExistente = mapaOps[opKey].modelos.findIndex((mo: any) => {
      const sameDesc = mo.descripcion === modelo.descripcion;
      const sameMod = (mo.modelo || '') === (modelo.modelo || '');
      const sameSap = (mo.sap || '') === (modelo.sap || '');
      return sameDesc && sameMod && sameSap;
    });
    if (idxExistente === -1) {
      mapaOps[opKey].modelos.push(modelo);
    } else if (cantNum > 0) {
      const cantPrev = parsearCantidad(mapaOps[opKey].modelos[idxExistente].cantidadTotal);
      mapaOps[opKey].modelos[idxExistente].cantidadTotal = cantPrev + cantNum;
    }
    mapaOps[opKey].cantidadTotal += cantNum;
  });
  return { mapaOps, ordenOps };
}

function parsearHojaModelos(sheet: any, mapaBase: Record<string, any>) {
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (rows.length < 2) return null;

  const found = encontrarFilaEncabezados(rows, (headers) => {
    const idxOp = encontrarColumnaOp(headers);
    const hasModelo = headers.some(
      (h) => h === 'modelo' || h.includes('modelo') || h.includes('descripcion') || h.includes('producto') || h.includes('variante') || h.includes('sabor')
    );
    return idxOp !== -1 && hasModelo;
  });
  if (!found) return null;

  const { rowIdx, headers } = found;
  const idxOp = encontrarColumnaOp(headers);
  let idxModelo = headers.findIndex((h) => h === 'modelo' || h.includes('modelo'));
  const idxDescAlt = headers.findIndex((h) => h.includes('descripcion') || h.includes('producto') || h.includes('variante') || h.includes('sabor'));
  if (idxModelo === -1) {
    idxModelo = idxDescAlt;
  }
  const idxCliente = headers.findIndex((h) => h.includes('cliente') || h.includes('razon') || h.includes('empresa'));
  const idxCantidad = encontrarColumnaCantidad(headers);
  const idxSap = encontrarColumnaSap(headers);
  if (idxOp === -1 || idxModelo === -1) return null;

  const filasValidas = rows.slice(rowIdx + 1).filter((r) => String(r[idxOp] ?? '').trim() !== '');
  const mapaOps: Record<string, any> = {};
  const ordenOps: string[] = [];

  filasValidas.forEach((r) => {
    const opRaw = String(r[idxOp] ?? '').trim();
    const base = buscarBase(mapaBase, opRaw);
    const opKey = base ? base.op : buscarClaveBase(mapaOps, opRaw) || opRaw;
    let modeloTxt = String(r[idxModelo] ?? '').trim();
    if (!modeloTxt && idxDescAlt !== -1) {
      modeloTxt = String(r[idxDescAlt] ?? '').trim();
    }
    if (!modeloTxt) return;

    const cantNum = idxCantidad !== -1 ? parsearCantidad(r[idxCantidad]) : 0;
    const descBase = base
      ? base.descripcionBase || (base.modelos && base.modelos[0] ? base.modelos[0].descripcion : base.descripcion || '')
      : '';
    const descripcionFinal =
      descBase && !modeloTxt.toLowerCase().includes(descBase.toLowerCase())
        ? `${descBase} - ${modeloTxt}`
        : modeloTxt;
    const sap = idxSap !== -1 ? String(r[idxSap] ?? '').trim() : '';
    const clienteMod = idxCliente !== -1 ? String(r[idxCliente] ?? '').trim() : '';
    const modelo: CatalogoModelo = {
      modelo: modeloTxt,
      descripcion: descripcionFinal,
      sap,
      cantidadTotal: cantNum,
    };
    if (!mapaOps[opKey]) {
      mapaOps[opKey] = {
        op: opKey,
        cliente: base ? base.cliente : clienteMod,
        cantidadTotal: 0,
        modelos: [],
      };
      ordenOps.push(opKey);
    } else if (!mapaOps[opKey].cliente && clienteMod) {
      mapaOps[opKey].cliente = clienteMod;
    }
    const idxM = mapaOps[opKey].modelos.findIndex(
      (mo: any) =>
        (mo.modelo === modeloTxt || mo.descripcion === descripcionFinal) &&
        (mo.sap || '') === (modelo.sap || '')
    );
    if (idxM === -1) {
      mapaOps[opKey].modelos.push(modelo);
    } else if (cantNum > 0) {
      const cantPrev = parsearCantidad(mapaOps[opKey].modelos[idxM].cantidadTotal);
      mapaOps[opKey].modelos[idxM].cantidadTotal = cantPrev + cantNum;
    }
    mapaOps[opKey].cantidadTotal += cantNum;
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
        const claveBase = buscarClaveBase(principal.mapaOps, op);
        const base = claveBase ? principal.mapaOps[claveBase] : null;
        const targetOp = claveBase || op;
        const modelosData = modelos.mapaOps[op];

        if (!mapaFinal[targetOp]) {
          ordenFinal.push(targetOp);
          mapaFinal[targetOp] = { ...modelosData, op: targetOp };
        } else {
          const descBase =
            base && base.modelos[0]
              ? base.descripcionBase || base.modelos[0].descripcion
              : base?.descripcion || '';
          const cantModelos = modelosData.cantidadTotal || 0;
          const cantBase = base ? base.cantidadTotal || 0 : 0;
          const cantFinal = cantModelos > 0 ? cantModelos : cantBase;

          mapaFinal[targetOp] = {
            ...modelosData,
            op: targetOp,
            cliente: (base && base.cliente) || modelosData.cliente || '',
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
  if (!wb || !wb.SheetNames || wb.SheetNames.length === 0) return null;

  let principalName = wb.SheetNames[0];
  if (!parsearHojaCatalogo(wb.Sheets[principalName])) {
    const otraPrincipal = wb.SheetNames.find((n) => parsearHojaCatalogo(wb.Sheets[n]) !== null);
    if (otraPrincipal) principalName = otraPrincipal;
  }

  const sheetPrincipal = wb.Sheets[principalName];
  let sheetModelos: any = null;
  const otrasHojas = wb.SheetNames.filter((n) => n !== principalName);
  if (otrasHojas.length > 0) {
    const nombreModelo = otrasHojas.find((n) => {
      const norm = normalizar(n);
      return norm.includes('modelo') || norm.includes('detalle') || norm.includes('variante') || norm.includes('sabor');
    });
    const candidata = nombreModelo || otrasHojas.find((n) => parsearHojaModelos(wb.Sheets[n], {}) !== null);
    if (candidata) sheetModelos = wb.Sheets[candidata];
  }
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
        lineas: 0,
        observaciones: [] as string[],
      };
      orden.push(key);
    }
    mapa[key].cantidadCajas += item.cantidadCajas;
    mapa[key].totalUnidades += item.totalUnidades;
    mapa[key].lineas += 1;
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

export function prepararFilasConsolidadoInventario(lista: ListaItem[], catalogo: CatalogoItem[]) {
  const ordenadas = [...(lista || [])].reverse();
  const grupos = agruparParaExportar(ordenadas);
  return grupos.map((g, idx) => {
    const match = buscarEnCatalogo(catalogo, g.op);
    const empaqueTxt = g.tipoEmpaque === 'paquete' ? 'Paquete' : 'Caja';
    const sap = g.sap || (match && match.sap ? match.sap : '');
    const unidadesXEmp = Number(g.unidadesXCaja) || 0;
    const cantEmp = Number(g.cantidadCajas) || 0;
    const totalUnid = Number(g.totalUnidades) || unidadesXEmp * cantEmp;
    return {
      indice: idx + 1,
      op: g.op || '',
      sap,
      descripcion: g.descripcion || '',
      cliente: g.cliente || '',
      empaque: empaqueTxt,
      unidadesXEmp,
      cantEmp,
      totalUnid,
      lineas: Number(g.lineas) || 1,
      observacion: g.observacion || '',
    };
  });
}

export async function exportarInventarioExcel(lista: ListaItem[], catalogo: CatalogoItem[]) {
  const filas = prepararFilasInventario(lista, catalogo);
  if (filas.length === 0) return;
  const filasConsolidado = prepararFilasConsolidadoInventario(lista, catalogo);

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

    // ==================== HOJA 1: DETALLE POR LÍNEA (SIN SUMAR) ====================
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

    ws.mergeCells(1, 1, 1, 10);
    const rTitulo = ws.getRow(1);
    rTitulo.height = 24;
    rTitulo.getCell(1).value = `INVENTARIO DE PRODUCCIÓN — ${fechaCorta}`;
    rTitulo.getCell(1).font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    rTitulo.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF12161F' } };
    rTitulo.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    ws.mergeCells(2, 1, 2, 10);
    const rSub = ws.getRow(2);
    rSub.height = 18;
    rSub.getCell(1).value = `Generado: ${fechaHora}   |   Líneas registradas: ${filas.length}   |   Total empaques: ${totalEmpaques.toLocaleString('es-PE')}   |   Total unidades: ${totalUnidades.toLocaleString('es-PE')} u`;
    rSub.getCell(1).font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    rSub.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD97706' } };
    rSub.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

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

    // ==================== HOJA 2: CONSOLIDADO AGRUPADO POR OP / MODELO ====================
    const wsRes = wb.addWorksheet('Consolidado por OP');
    wsRes.columns = [
      { width: 6 },  // #
      { width: 13 }, // OP
      { width: 15 }, // SAP
      { width: 42 }, // DESCRIPCIÓN
      { width: 24 }, // CLIENTE
      { width: 12 }, // EMPAQUE
      { width: 15 }, // UNID. X EMP.
      { width: 14 }, // CANT. EMP.
      { width: 17 }, // TOTAL UNIDADES
      { width: 10 }, // LÍNEAS
      { width: 30 }, // OBSERVACIONES
    ];

    wsRes.mergeCells(1, 1, 1, 11);
    const rTitulo2 = wsRes.getRow(1);
    rTitulo2.height = 24;
    rTitulo2.getCell(1).value = `CONSOLIDADO DE INVENTARIO POR OP — ${fechaCorta}`;
    rTitulo2.getCell(1).font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    rTitulo2.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF12161F' } };
    rTitulo2.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    wsRes.mergeCells(2, 1, 2, 11);
    const rSub2 = wsRes.getRow(2);
    rSub2.height = 18;
    rSub2.getCell(1).value = `Generado: ${fechaHora}   |   Grupos consolidados: ${filasConsolidado.length}   |   Líneas origen: ${filas.length}   |   Total unidades: ${totalUnidades.toLocaleString('es-PE')} u`;
    rSub2.getCell(1).font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    rSub2.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD97706' } };
    rSub2.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    const headers2 = [
      '#',
      'OP',
      'SAP',
      'DESCRIPCIÓN / MODELO',
      'CLIENTE',
      'EMPAQUE',
      'UNID. X EMP.',
      'CANT. EMP.',
      'TOTAL UNIDADES',
      'LÍNEAS',
      'OBSERVACIONES',
    ];
    const rHead2 = wsRes.getRow(3);
    rHead2.height = 22;
    headers2.forEach((h, i) => {
      const c = rHead2.getCell(i + 1);
      c.value = h;
      c.font = { bold: true, size: 10, color: { argb: 'FF333333' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      c.border = EXCELJS_BORDES;
    });

    let rowIdx2 = 4;
    filasConsolidado.forEach((f, idx) => {
      const r = wsRes.getRow(rowIdx2);
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
        f.lineas,
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
        if (colIdx === 0 || colIdx === 1 || colIdx === 2 || colIdx === 5 || colIdx === 9) {
          c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          if (colIdx === 1) c.font = { bold: true, color: { argb: 'FF12161F' } };
          if (colIdx === 9) c.numFmt = '#,##0';
        } else if (colIdx === 6 || colIdx === 7 || colIdx === 8) {
          c.alignment = { horizontal: 'right', vertical: 'middle' };
          c.numFmt = '#,##0';
          if (colIdx === 8) c.font = { bold: true, color: { argb: 'FF1E40AF' } };
        } else {
          c.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        }
      });
      rowIdx2++;
    });

    wsRes.mergeCells(rowIdx2, 1, rowIdx2, 7);
    const rTot2 = wsRes.getRow(rowIdx2);
    rTot2.height = 20;
    rTot2.getCell(1).value = `TOTAL CONSOLIDADO (${filasConsolidado.length} GRUPO${filasConsolidado.length !== 1 ? 'S' : ''})`;
    for (let i = 1; i <= 11; i++) {
      const cell = rTot2.getCell(i);
      cell.font = { bold: true, size: 10.5 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDEBD0' } };
      cell.border = EXCELJS_BORDES;
      cell.alignment = { vertical: 'middle' };
    }
    rTot2.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
    rTot2.getCell(8).value = totalEmpaques;
    rTot2.getCell(8).numFmt = '#,##0';
    rTot2.getCell(8).alignment = { horizontal: 'right', vertical: 'middle' };
    rTot2.getCell(9).value = totalUnidades;
    rTot2.getCell(9).numFmt = '#,##0';
    rTot2.getCell(9).alignment = { horizontal: 'right', vertical: 'middle' };
    rTot2.getCell(9).font = { bold: true, size: 11, color: { argb: 'FFB45309' } };
    rTot2.getCell(10).value = filas.length;
    rTot2.getCell(10).numFmt = '#,##0';
    rTot2.getCell(10).alignment = { horizontal: 'center', vertical: 'middle' };

    wsRes.views = [{ state: 'frozen', ySplit: 3 }];
    wsRes.pageSetup = {
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
    // Fallback to SheetJS (Hoja 1 sin agrupar + Hoja 2 consolidado)
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

    const filasSimple2 = filasConsolidado.map((f) => ({
      '#': f.indice,
      OP: f.op,
      SAP: f.sap,
      Descripción: f.descripcion,
      Cliente: f.cliente,
      Empaque: f.empaque,
      'Unidades x empaque': f.unidadesXEmp,
      'Cantidad de empaques': f.cantEmp,
      'Total unidades': f.totalUnid,
      Líneas: f.lineas,
      Observaciones: f.observacion,
    }));
    const ws2 = XLSX.utils.json_to_sheet(filasSimple2);
    ws2['!cols'] = [
      { wch: 6 },
      { wch: 13 },
      { wch: 14 },
      { wch: 38 },
      { wch: 22 },
      { wch: 11 },
      { wch: 16 },
      { wch: 16 },
      { wch: 16 },
      { wch: 10 },
      { wch: 30 },
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Inventario');
    XLSX.utils.book_append_sheet(wb, ws2, 'Consolidado por OP');
    const wbArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbArray], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    descargarBlob(blob, nombre);
  }
}

export function imprimirInventarioFormatoExcel(lista: ListaItem[], catalogo: CatalogoItem[]): string {
  const filas = prepararFilasInventario(lista, catalogo);
  if (filas.length === 0) return '';

  const totalUnidades = filas.reduce((acc, f) => acc + f.totalUnid, 0);
  const totalEmpaques = filas.reduce((acc, f) => acc + f.cantEmp, 0);
  const ahora = new Date();
  const hoyIso = ahora.toISOString().slice(0, 10);
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
  const subtitulo = `Generado: ${fechaHora}  ·  Detalle por línea registrada (${filas.length} registro${filas.length !== 1 ? 's' : ''})`;
  const nombreArchivo = `inventario_${hoyIso}.pdf`;

  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4',
  });

  const pageWidth = 297;
  const pageHeight = 210;
  const marginX = 10;
  const contentWidth = pageWidth - marginX * 2; // 277 mm

  // Column widths summing strictly to 277 mm
  const colWidths = [10, 20, 22, 76, 44, 20, 21, 19, 23, 22];
  const colX: number[] = [];
  let accX = marginX;
  colWidths.forEach((w) => {
    colX.push(accX);
    accX += w;
  });

  const headers = [
    '#',
    'OP',
    'SAP',
    'DESCRIPCIÓN / MODELO',
    'CLIENTE',
    'EMPAQUE',
    'UNID. × EMP.',
    'CANT. EMP.',
    'TOTAL UNID.',
    'OBSERVACIÓN',
  ];

  let y = 11;

  const drawMainHeader = () => {
    doc.setFillColor(18, 22, 31); // #12161F
    doc.rect(marginX, y, contentWidth, 15.5, 'F');
    doc.setFillColor(242, 169, 59); // #F2A93B
    doc.rect(marginX, y, 3.5, 15.5, 'F');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(13);
    doc.setTextColor(255, 255, 255);
    doc.text(titulo, marginX + 6.5, y + 6.5);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.2);
    doc.setTextColor(203, 213, 224);
    doc.text(subtitulo, marginX + 6.5, y + 12);

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9.5);
    doc.setTextColor(242, 169, 59);
    doc.text(
      `Total: ${totalUnidades.toLocaleString('es-PE')} u  (${filas.length} líneas)`,
      marginX + contentWidth - 4,
      y + 9.2,
      { align: 'right' }
    );

    y += 18;
  };

  const drawSectionAndTableHeaders = (isContinuation = false) => {
    doc.setFillColor(217, 119, 6); // #D97706
    doc.setDrawColor(183, 183, 183);
    doc.setLineWidth(0.2);
    doc.rect(marginX, y, contentWidth, 6.2, 'FD');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.5);
    doc.setTextColor(255, 255, 255);
    doc.text(
      isContinuation
        ? 'DETALLE DE PIEZAS Y EMPAQUES REGISTRADOS (CONTINUACIÓN)'
        : 'DETALLE DE PIEZAS Y EMPAQUES REGISTRADOS',
      marginX + 3,
      y + 4.2
    );
    y += 6.2;

    doc.setFillColor(242, 242, 242); // #F2F2F2
    doc.rect(marginX, y, contentWidth, 6.2, 'FD');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.1);
    doc.setTextColor(51, 51, 51);
    headers.forEach((h, i) => {
      doc.line(colX[i], y, colX[i], y + 6.2);
      doc.text(h, colX[i] + colWidths[i] / 2, y + 4.2, { align: 'center' });
    });
    doc.line(marginX + contentWidth, y, marginX + contentWidth, y + 6.2);
    y += 6.2;
  };

  drawMainHeader();
  drawSectionAndTableHeaders(false);

  filas.forEach((f, idx) => {
    const values = [
      String(f.indice),
      String(f.op || ''),
      String(f.sap || '—'),
      String(f.descripcion || ''),
      String(f.cliente || '—'),
      String(f.empaque || ''),
      f.unidadesXEmp.toLocaleString('es-PE'),
      f.cantEmp.toLocaleString('es-PE'),
      f.totalUnid.toLocaleString('es-PE'),
      String(f.observacion || ''),
    ];

    doc.setFontSize(7.2);
    const wrappedLines = values.map((val, cIdx) => {
      if (!val) return [''];
      return doc.splitTextToSize(val, colWidths[cIdx] - 2.4) as string[];
    });

    const maxLines = Math.max(...wrappedLines.map((l) => l.length), 1);
    const rowH = Math.max(5.5, maxLines * 3.2 + 2.2);

    if (y + rowH > pageHeight - 14) {
      doc.addPage('a4', 'landscape');
      y = 11;
      drawMainHeader();
      drawSectionAndTableHeaders(true);
    }

    if (idx % 2 === 1) {
      doc.setFillColor(250, 250, 250);
      doc.rect(marginX, y, contentWidth, rowH, 'F');
    }

    doc.setDrawColor(183, 183, 183);
    doc.setLineWidth(0.2);
    colX.forEach((cx) => {
      doc.line(cx, y, cx, y + rowH);
    });
    doc.line(marginX + contentWidth, y, marginX + contentWidth, y + rowH);
    doc.line(marginX, y + rowH, marginX + contentWidth, y + rowH);

    wrappedLines.forEach((lines, cIdx) => {
      if (lines.length === 1 && lines[0] === '') return;

      if (cIdx === 1 || cIdx === 8) {
        doc.setFont('helvetica', 'bold');
      } else {
        doc.setFont('helvetica', 'normal');
      }

      if (cIdx === 8) {
        doc.setTextColor(30, 64, 175); // #1E40AF
      } else {
        doc.setTextColor(17, 24, 39);
      }

      const totalTextH = lines.length * 3.2;
      let startTextY = y + (rowH - totalTextH) / 2 + 2.5;

      lines.forEach((line) => {
        if (cIdx === 0 || cIdx === 1 || cIdx === 2 || cIdx === 5) {
          doc.text(line, colX[cIdx] + colWidths[cIdx] / 2, startTextY, { align: 'center' });
        } else if (cIdx === 6 || cIdx === 7 || cIdx === 8) {
          doc.text(line, colX[cIdx] + colWidths[cIdx] - 1.4, startTextY, { align: 'right' });
        } else {
          doc.text(line, colX[cIdx] + 1.4, startTextY);
        }
        startTextY += 3.2;
      });
    });

    y += rowH;
  });

  // Total Row (#FDEBD0) + Grand Total Banner (#12161F)
  if (y + 18 > pageHeight - 13) {
    doc.addPage('a4', 'landscape');
    y = 12;
  }

  doc.setFillColor(253, 235, 208);
  doc.setDrawColor(183, 183, 183);
  doc.setLineWidth(0.22);
  doc.rect(marginX, y, contentWidth, 6.4, 'FD');

  colX.forEach((cx, cIdx) => {
    if (cIdx >= 7) {
      doc.line(cx, y, cx, y + 6.4);
    }
  });

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.setTextColor(17, 24, 39);
  doc.text(`TOTAL GENERAL (${filas.length} LÍNEA${filas.length !== 1 ? 'S' : ''})`, marginX + 2.5, y + 4.3);
  doc.text(totalEmpaques.toLocaleString('es-PE'), colX[7] + colWidths[7] - 1.4, y + 4.3, { align: 'right' });
  doc.setTextColor(30, 64, 175);
  doc.text(totalUnidades.toLocaleString('es-PE'), colX[8] + colWidths[8] - 1.4, y + 4.3, { align: 'right' });
  y += 8.5;

  doc.setFillColor(18, 22, 31);
  doc.rect(marginX, y, contentWidth, 7.5, 'F');
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(255, 255, 255);
  doc.text(`RESUMEN DE INVENTARIO (${filas.length} LÍNEAS)`, marginX + 3, y + 5);

  doc.setTextColor(242, 169, 59);
  doc.text(
    `EMPAQUES: ${totalEmpaques.toLocaleString('es-PE')}   ·   TOTAL UNIDADES: ${totalUnidades.toLocaleString('es-PE')} u`,
    marginX + contentWidth - 3,
    y + 5,
    { align: 'right' }
  );

  // Page footer on all pages
  const totalPages = (doc as any).internal.getNumberOfPages();
  for (let p = 1; p <= totalPages; p++) {
    doc.setPage(p);
    doc.setDrawColor(203, 213, 224);
    doc.setLineWidth(0.2);
    doc.line(marginX, pageHeight - 10, marginX + contentWidth, pageHeight - 10);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(100, 116, 139);
    doc.text(`Generado: ${fechaHora}  ·  Inventario de Producción`, marginX, pageHeight - 6.2);
    doc.text(`Página ${p} de ${totalPages}`, marginX + contentWidth, pageHeight - 6.2, {
      align: 'right',
    });
  }

  doc.save(nombreArchivo);
  return nombreArchivo;
}

// ----------------- TURNO INSTANCIAS & MULTI-OP -----------------
export function turnosDeMaquina(m: Maquina, turno: TurnoMaquina[]): TurnoMaquina[] {
  if (!m.multiOp) {
    const t = turno.find((x) => x.maquinaId === m.id && (x.instancia || 1) === 1);
    const extrasConRegistros = turno
      .filter((x) => x.maquinaId === m.id && (x.instancia || 1) !== 1 && (x.registros || []).length > 0)
      .sort((a, b) => (a.instancia || 1) - (b.instancia || 1));
    const principal: TurnoMaquina = t || {
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
    };
    return [principal, ...extrasConRegistros];
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
          personal: String(t.personal || '').toLocaleUpperCase('es-PE'),
        },
      ]
    : [];
}

// ----------------- TURNO EXCEL & FORMATTING -----------------
export function generarFilasTurno(maquinas: Maquina[], turno: TurnoMaquina[], catalogo: CatalogoItem[]): FilaTurnoExcel[] {
  const filas: FilaTurnoExcel[] = [];
  const activas = maquinas
    .filter(
      (m) =>
        m.activa ||
        turno.some(
          (t) =>
            t.maquinaId === m.id &&
            (t.registros || []).some((r) => (r.cantidad || 0) > 0 || (r.merma || 0) > 0)
        )
    )
    .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' }));

  activas.forEach((m) => {
    const tipoMaqNorm = (m.tipoProducto || '').trim().toLowerCase();
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
            tipoProducto: tipoMaqNorm,
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
            let sapGrupo = r.sap || '';
            if (!sapGrupo && match) {
              if (match.modelos && match.modelos.length > 1 && descKey) {
                const moMatch = match.modelos.find(
                  (mo) =>
                    mo.descripcion === descKey ||
                    (mo.modelo && descKey.toLowerCase().includes(mo.modelo.toLowerCase()))
                );
                sapGrupo = moMatch?.sap || match.sap || '';
              } else if (match.modelos && match.modelos.length === 1) {
                sapGrupo = match.modelos[0].sap || match.sap || '';
              } else {
                sapGrupo = match.sap || '';
              }
            }
            grupos[claveGrupo] = {
              op: opKey,
              descripcion: descBase + (pliegoKey ? ' - ' + etiquetaPliego(pliegoKey) : ''),
              sap: sapGrupo,
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
            tipoProducto: tipoMaqNorm,
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
            tipoProducto: tipoMaqNorm,
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
  maquinasCatalogo: Maquina[] = [],
  nombreArchivoBase?: string
): string {
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

  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4',
  });

  const pageWidth = 297;
  const pageHeight = 210;
  const marginX = 10;
  const contentWidth = pageWidth - marginX * 2; // 277 mm

  // Column widths summing strictly to 277 mm
  const colWidths = [34, 22, 22, 22, 76, 28, 20, 53];
  const colX: number[] = [];
  let accX = marginX;
  colWidths.forEach((w) => {
    colX.push(accX);
    accX += w;
  });

  const headers = [
    'MÁQUINA',
    'CÓDIGO MÁQ.',
    "OP'S",
    'SAP',
    'PRODUCTO',
    'TOTAL PROD.',
    'MERMA',
    'OBSERVACIONES',
  ];

  let y = 11;

  const drawMainHeader = () => {
    doc.setFillColor(18, 22, 31); // #12161F
    doc.rect(marginX, y, contentWidth, 15.5, 'F');
    doc.setFillColor(242, 169, 59); // #F2A93B
    doc.rect(marginX, y, 3.5, 15.5, 'F');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(13);
    doc.setTextColor(255, 255, 255);
    doc.text(titulo, marginX + 6.5, y + 6.5);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.2);
    doc.setTextColor(203, 213, 224);
    doc.text(subtitulo, marginX + 6.5, y + 12);

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9.5);
    doc.setTextColor(242, 169, 59);
    const statsParts = [`Total: ${totalGen.toLocaleString('es-PE')} u`];
    if (mermaGen > 0) statsParts.push(`Merma: ${mermaGen.toLocaleString('es-PE')} u`);
    statsParts.push(`(${filas.length} filas)`);
    doc.text(statsParts.join('  ·  '), marginX + contentWidth - 4, y + 9.2, {
      align: 'right',
    });

    y += 18;
  };

  const drawAreaAndTableHeaders = (areaLabel: string, isContinuation = false) => {
    doc.setFillColor(217, 119, 6); // #D97706
    doc.setDrawColor(183, 183, 183);
    doc.setLineWidth(0.2);
    doc.rect(marginX, y, contentWidth, 6.2, 'FD');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.5);
    doc.setTextColor(255, 255, 255);
    doc.text(
      isContinuation ? `${areaLabel} (CONTINUACIÓN)` : areaLabel,
      marginX + 3,
      y + 4.2
    );
    y += 6.2;

    doc.setFillColor(242, 242, 242); // #F2F2F2
    doc.rect(marginX, y, contentWidth, 6.2, 'FD');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.1);
    doc.setTextColor(51, 51, 51);
    headers.forEach((h, i) => {
      doc.line(colX[i], y, colX[i], y + 6.2);
      doc.text(h, colX[i] + colWidths[i] / 2, y + 4.2, { align: 'center' });
    });
    doc.line(marginX + contentWidth, y, marginX + contentWidth, y + 6.2);
    y += 6.2;
  };

  drawMainHeader();

  AREAS_ORDEN.forEach((areaDef) => {
    const filasDeArea = filas.filter((f) => (f.area || '') === areaDef.key);
    if (filasDeArea.length === 0) return;

    const totalArea = filasDeArea.reduce((s, f) => s + (f.total || 0), 0);
    const mermaArea = filasDeArea.reduce((s, f) => s + (f.merma || 0), 0);

    if (y + 20 > pageHeight - 14) {
      doc.addPage('a4', 'landscape');
      y = 11;
      drawMainHeader();
    }

    drawAreaAndTableHeaders(areaDef.label, false);

    let prevMaqKey: string | null = null;

    filasDeArea.forEach((f, idx) => {
      const maqKey = (f.maquina || '').trim().toLowerCase();
      const isSameAsPrev = prevMaqKey !== null && prevMaqKey === maqKey;
      const nextIsSame =
        idx + 1 < filasDeArea.length &&
        (filasDeArea[idx + 1].maquina || '').trim().toLowerCase() === maqKey;

      const values = [
        isSameAsPrev ? '' : String(f.maquina || ''),
        isSameAsPrev ? '' : String(f.codigoMaquina || ''),
        String(f.op || ''),
        String(f.sap || ''),
        String(f.producto || ''),
        (f.total || 0).toLocaleString('es-PE'),
        (f.merma || 0).toLocaleString('es-PE'),
        String(f.observaciones || ''),
      ];

      doc.setFontSize(7.2);
      const wrappedLines = values.map((val, cIdx) => {
        if (!val) return [''];
        return doc.splitTextToSize(val, colWidths[cIdx] - 2.4) as string[];
      });

      const maxLines = Math.max(...wrappedLines.map((l) => l.length), 1);
      const rowH = Math.max(5.5, maxLines * 3.2 + 2.2);

      if (y + rowH > pageHeight - 14) {
        doc.addPage('a4', 'landscape');
        y = 11;
        drawMainHeader();
        drawAreaAndTableHeaders(areaDef.label, true);
        // Reprint machine & code after page break
        wrappedLines[0] = f.maquina
          ? (doc.splitTextToSize(String(f.maquina), colWidths[0] - 2.4) as string[])
          : [''];
        wrappedLines[1] = f.codigoMaquina
          ? (doc.splitTextToSize(String(f.codigoMaquina), colWidths[1] - 2.4) as string[])
          : [''];
      }

      if (idx % 2 === 1) {
        doc.setFillColor(250, 250, 250);
        doc.rect(marginX, y, contentWidth, rowH, 'F');
      }

      doc.setDrawColor(183, 183, 183);
      doc.setLineWidth(0.2);
      colX.forEach((cx) => {
        doc.line(cx, y, cx, y + rowH);
      });
      doc.line(marginX + contentWidth, y, marginX + contentWidth, y + rowH);

      if (!nextIsSame) {
        doc.line(marginX, y + rowH, marginX + contentWidth, y + rowH);
      } else {
        doc.line(colX[2], y + rowH, marginX + contentWidth, y + rowH);
      }

      wrappedLines.forEach((lines, cIdx) => {
        if (lines.length === 1 && lines[0] === '') return;

        if (cIdx === 0 || cIdx === 5) {
          doc.setFont('helvetica', 'bold');
        } else {
          doc.setFont('helvetica', 'normal');
        }

        if (cIdx === 6) {
          doc.setTextColor(192, 57, 43); // #C0392B merma
        } else if (cIdx === 5) {
          doc.setTextColor(30, 64, 175); // #1E40AF total
        } else {
          doc.setTextColor(17, 24, 39);
        }

        const totalTextH = lines.length * 3.2;
        let startTextY = y + (rowH - totalTextH) / 2 + 2.5;

        lines.forEach((line) => {
          if (cIdx === 1 || cIdx === 2 || cIdx === 3) {
            doc.text(line, colX[cIdx] + colWidths[cIdx] / 2, startTextY, { align: 'center' });
          } else if (cIdx === 5 || cIdx === 6) {
            doc.text(line, colX[cIdx] + colWidths[cIdx] - 1.4, startTextY, { align: 'right' });
          } else {
            doc.text(line, colX[cIdx] + 1.4, startTextY);
          }
          startTextY += 3.2;
        });
      });

      prevMaqKey = maqKey;
      y += rowH;
    });

    // Area subtotal row (#FDEBD0)
    if (y + 7 > pageHeight - 14) {
      doc.addPage('a4', 'landscape');
      y = 12;
    }

    doc.setFillColor(253, 235, 208);
    doc.setDrawColor(183, 183, 183);
    doc.setLineWidth(0.22);
    doc.rect(marginX, y, contentWidth, 6.2, 'FD');
    doc.line(colX[5], y, colX[5], y + 6.2);
    doc.line(colX[6], y, colX[6], y + 6.2);
    doc.line(colX[7], y, colX[7], y + 6.2);

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.setTextColor(17, 24, 39);
    doc.text(`TOTAL ${areaDef.label}`, marginX + 2.5, y + 4.2);

    doc.setTextColor(30, 64, 175);
    doc.text(totalArea.toLocaleString('es-PE'), colX[5] + colWidths[5] - 1.4, y + 4.2, {
      align: 'right',
    });

    doc.setTextColor(192, 57, 43);
    doc.text(mermaArea.toLocaleString('es-PE'), colX[6] + colWidths[6] - 1.4, y + 4.2, {
      align: 'right',
    });

    y += 9;
  });

  // Grand Total Summary Banner (#12161F)
  if (y + 10 > pageHeight - 13) {
    doc.addPage('a4', 'landscape');
    y = 12;
  }

  doc.setFillColor(18, 22, 31);
  doc.rect(marginX, y, contentWidth, 7.5, 'F');
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(255, 255, 255);
  doc.text(`RESUMEN GENERAL DEL TURNO (${filas.length} REGISTROS)`, marginX + 3, y + 5);

  doc.setTextColor(242, 169, 59);
  const grandRight = `TOTAL PRODUCIDO: ${totalGen.toLocaleString('es-PE')} u   ·   MERMA TOTAL: ${mermaGen.toLocaleString('es-PE')} u`;
  doc.text(grandRight, marginX + contentWidth - 3, y + 5, { align: 'right' });

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
    doc.text(`Generado: ${generado}  ·  Reporte de Producción`, marginX, pageHeight - 6.2);
    doc.text(`Página ${p} de ${totalPages}`, marginX + contentWidth, pageHeight - 6.2, {
      align: 'right',
    });
  }

  const baseClean = (
    nombreArchivoBase ||
    titulo
      .toLowerCase()
      .replace(/[^a-z0-9]+/gi, '_')
      .replace(/^_+|_+$/g, '') ||
    'reporte_turno'
  ).replace(/\.pdf$/i, '');
  const nombreArchivo = `${baseClean}.pdf`;

  doc.save(nombreArchivo);
  return nombreArchivo;
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
        personal: String(t.personal || '').toLocaleUpperCase('es-PE'),
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
export function exportarMigracion(
  todasLasFilas: FilaTurnoExcel[],
  fecha: string,
  tipo: string,
  maquinas?: Maquina[],
  catalogo?: CatalogoItem[]
) {
  const maqByName = new Map<string, Maquina>();
  const maqByNorm = new Map<string, Maquina>();
  const maqByCode = new Map<string, Maquina>();

  const normKey = (s: string) =>
    String(s || '')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '');

  (maquinas || []).forEach((m) => {
    const nameL = (m.nombre || '').trim().toLowerCase();
    if (nameL) {
      maqByName.set(nameL, m);
      maqByNorm.set(normKey(nameL), m);
    }
    const codL = (m.codigo || '').trim().toLowerCase();
    if (codL) {
      maqByCode.set(codL, m);
    }
  });

  const buscarMaquinaFila = (f: FilaTurnoExcel): Maquina | undefined => {
    const nL = (f.maquina || '').trim().toLowerCase();
    const cL = (f.codigoMaquina || '').trim().toLowerCase();
    return (
      (nL ? maqByName.get(nL) : undefined) ||
      (cL ? maqByCode.get(cL) : undefined) ||
      (nL ? maqByNorm.get(normKey(nL)) : undefined)
    );
  };

  const filasEnriquecidas = (todasLasFilas || [])
    .map((f) => {
      const maq = buscarMaquinaFila(f);
      const tipoRaw = (f.tipoProducto || '').trim().toLowerCase();
      const tipoMaq = (maq?.tipoProducto || '').trim().toLowerCase();
      const tipoResuelto = tipoRaw === 'pt' || tipoRaw === 'pp' ? tipoRaw : tipoMaq;
      const totalNum = typeof f.total === 'number' ? f.total : parsearCantidad(f.total);

      let sapResuelto = (f.sap || '').trim();
      if (!sapResuelto && catalogo && catalogo.length > 0 && f.op) {
        const match = buscarEnCatalogo(catalogo, f.op);
        if (match) {
          if (match.modelos && match.modelos.length > 1 && f.producto) {
            const prodL = f.producto.toLowerCase();
            const moMatch = match.modelos.find(
              (mo) =>
                mo.descripcion.toLowerCase() === prodL ||
                (mo.modelo && prodL.includes(mo.modelo.toLowerCase()))
            );
            sapResuelto = moMatch?.sap || match.sap || '';
          } else if (match.modelos && match.modelos.length === 1) {
            sapResuelto = match.modelos[0].sap || match.sap || '';
          } else {
            sapResuelto = match.sap || '';
          }
        }
      }

      return {
        ...f,
        tipoResuelto,
        totalNum,
        codigoResuelto: (f.codigoMaquina || '').trim() || maq?.codigo || '',
        sapResuelto,
      };
    })
    .filter((f) => f.tipoResuelto === 'pt' && f.totalNum > 0);

  if (filasEnriquecidas.length === 0) {
    throw new Error('No hay producción PT para migrar');
  }
  const aoa = [['ODTCOD', 'CODPTSAP', 'CODMAQ', 'ODTCANT', 'ODTOBS']];
  filasEnriquecidas.forEach((f) => {
    aoa.push([f.op || '', f.sapResuelto || '', f.codigoResuelto || '', String(f.totalNum), f.observaciones || '']);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 30 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Hoja1');
  const wbArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  descargarBlob(blob, `migracion_${fecha}_${tipo}.xlsx`);
}
