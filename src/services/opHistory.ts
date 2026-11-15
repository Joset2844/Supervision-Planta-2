import {
  CatalogoItem,
  HistorialTurnoLocal,
  Maquina,
  MetaTurno,
  SupabaseSession,
  TurnoDBRow,
  TurnoMaquina,
} from '../types';
import { getSupabaseConfig, sbFetch } from './supabase';
import { buscarEnCatalogo } from './excel';

export interface MaquinaRegistroOp {
  maquina: string;
  codigoMaquina: string;
  tipoProducto: 'pp' | 'pt' | '';
  sap: string;
  producto: string;
  total: number;
  merma: number;
  observaciones: string;
}

export interface MaquinaPPSum {
  maquina: string;
  codigoMaquina: string;
  total: number;
  merma: number;
}

export interface TurnoOpResumen {
  key: string;
  fecha: string;
  tipo: string;
  fuente: 'nube' | 'local' | 'actual';
  totalPP: number;
  mermaPP: number;
  ppPorMaquina: MaquinaPPSum[];
  totalPT: number;
  mermaPT: number;
  totalSinTipo: number;
  mermaSinTipo: number;
  filas: MaquinaRegistroOp[];
}

export interface MaquinaTotalOp {
  key: string;
  maquina: string;
  codigoMaquina: string;
  tipoProducto: 'pp' | 'pt' | '';
  totalAnterior: number;
  mermaAnterior: number;
  totalActual: number;
  mermaActual: number;
  totalAcumulado: number;
  mermaAcumulada: number;
  turnosCount: number;
}

export interface TotalesOpSeparados {
  antPP: number;
  mermaAntPP: number;
  actualPP: number;
  mermaActualPP: number;
  acumPP: number;
  mermaAcumPP: number;
  antPT: number;
  mermaAntPT: number;
  actualPT: number;
  mermaActualPT: number;
  acumPT: number;
  mermaAcumPT: number;
  antSinTipo: number;
  mermaAntSinTipo: number;
  actualSinTipo: number;
  mermaActualSinTipo: number;
  acumSinTipo: number;
  mermaAcumSinTipo: number;
}

export interface ConsultaOpResultado {
  turnosAnteriores: TurnoOpResumen[];
  turnoActualResumen: TurnoOpResumen | null;
  porMaquina: MaquinaTotalOp[];
  totales: TotalesOpSeparados;
  nubeDisponible: boolean;
  nubeMensaje: string;
}

export function opCoincide(opA: string, opB: string): boolean {
  const a = (opA || '').trim().toLowerCase();
  const b = (opB || '').trim().toLowerCase();
  if (!a || !b) return false;
  if (a === b) return true;
  const da = a.replace(/\D/g, '');
  const db = b.replace(/\D/g, '');
  if (da && db && da === db) return true;
  return false;
}

function buildMaqTipoMap(maquinas: Maquina[]): Map<string, 'pp' | 'pt' | ''> {
  const map = new Map<string, 'pp' | 'pt' | ''>();
  maquinas.forEach((m) => {
    if (m.nombre) map.set(m.nombre.trim().toLowerCase(), m.tipoProducto || '');
    if (m.codigo) map.set(`cod:${m.codigo.trim().toLowerCase()}`, m.tipoProducto || '');
  });
  return map;
}

function resolverTipo(
  maqTipoMap: Map<string, 'pp' | 'pt' | ''>,
  rawTipo?: string,
  nombreMaq?: string,
  codMaq?: string
): 'pp' | 'pt' | '' {
  const t = (rawTipo || '').trim().toLowerCase();
  if (t === 'pp' || t === 'pt') return t;
  const byName = nombreMaq ? maqTipoMap.get(nombreMaq.trim().toLowerCase()) : '';
  if (byName === 'pp' || byName === 'pt') return byName;
  const byCode = codMaq ? maqTipoMap.get(`cod:${codMaq.trim().toLowerCase()}`) : '';
  if (byCode === 'pp' || byCode === 'pt') return byCode;
  return '';
}

function construirTurnoOpResumen(
  key: string,
  fecha: string,
  tipo: string,
  fuente: 'nube' | 'local' | 'actual',
  filasRaw: MaquinaRegistroOp[]
): TurnoOpResumen {
  let totalPP = 0;
  let mermaPP = 0;
  let totalPT = 0;
  let mermaPT = 0;
  let totalSinTipo = 0;
  let mermaSinTipo = 0;

  const ppMap = new Map<string, MaquinaPPSum>();

  filasRaw.forEach((r) => {
    if (r.tipoProducto === 'pp') {
      totalPP += r.total;
      mermaPP += r.merma;
      const mKey = (r.maquina || 'Sin máquina').trim().toLowerCase();
      if (!ppMap.has(mKey)) {
        ppMap.set(mKey, {
          maquina: r.maquina || 'Sin máquina',
          codigoMaquina: r.codigoMaquina || '',
          total: 0,
          merma: 0,
        });
      }
      const entry = ppMap.get(mKey)!;
      if (!entry.codigoMaquina && r.codigoMaquina) entry.codigoMaquina = r.codigoMaquina;
      entry.total += r.total;
      entry.merma += r.merma;
    } else if (r.tipoProducto === 'pt') {
      totalPT += r.total;
      mermaPT += r.merma;
    } else {
      totalSinTipo += r.total;
      mermaSinTipo += r.merma;
    }
  });

  const peso = (tp: string) => (tp === 'pt' ? 0 : tp === 'pp' ? 1 : 2);
  const filas = [...filasRaw].sort((a, b) => peso(a.tipoProducto) - peso(b.tipoProducto));
  const ppPorMaquina = Array.from(ppMap.values());

  return {
    key,
    fecha,
    tipo,
    fuente,
    totalPP,
    mermaPP,
    ppPorMaquina,
    totalPT,
    mermaPT,
    totalSinTipo,
    mermaSinTipo,
    filas,
  };
}

export async function consultarHistorialOp(params: {
  op: string;
  catalogo?: CatalogoItem[];
  historial: HistorialTurnoLocal[];
  maquinas: Maquina[];
  turno: TurnoMaquina[];
  metaTurno: MetaTurno;
  session: SupabaseSession | null;
}): Promise<ConsultaOpResultado> {
  const { op, historial, maquinas, turno, metaTurno, session } = params;
  const maqTipoMap = buildMaqTipoMap(maquinas);

  let dbRows: TurnoDBRow[] = [];
  let nubeDisponible = false;
  let nubeMensaje = '';

  const cfg = getSupabaseConfig(false);
  if (cfg) {
    try {
      const res = await sbFetch<TurnoDBRow[]>(
        cfg,
        'turnos?select=fecha,tipo,filas&order=fecha.desc,tipo.asc&limit=500',
        undefined,
        session
      );
      if (Array.isArray(res)) {
        dbRows = res;
        nubeDisponible = true;
      }
    } catch {
      nubeDisponible = false;
      nubeMensaje = session
        ? 'No se pudo consultar la base de datos en la nube. Mostrando registros locales.'
        : 'Inicia sesión en "Turnos Guardados" para consultar también los turnos guardados en la nube.';
    }
  } else {
    nubeMensaje = 'Mostrando solo registros guardados en este dispositivo.';
  }

  const mapaTurnos = new Map<string, TurnoOpResumen>();
  const currentFecha = (metaTurno.fecha || '').trim();
  const currentTipo = (metaTurno.tipo || '').trim().toLowerCase();

  // 1. Procesar turnos de Supabase
  dbRows.forEach((rec) => {
    const fFecha = (rec.fecha || '').trim();
    const fTipo = (rec.tipo || '').trim().toLowerCase();
    if (currentFecha && currentTipo && fFecha === currentFecha && fTipo === currentTipo) {
      return;
    }

    const coincidentes = (rec.filas || []).filter((f) => opCoincide(f.op, op));
    if (coincidentes.length === 0) return;

    const filasMap: MaquinaRegistroOp[] = coincidentes.map((f) => ({
      maquina: f.maquina || 'Sin máquina',
      codigoMaquina: f.codigoMaquina || '',
      tipoProducto: resolverTipo(maqTipoMap, f.tipoProducto, f.maquina, f.codigoMaquina),
      sap: f.sap || '',
      producto: f.producto || '',
      total: Number(f.total) || 0,
      merma: Number(f.merma) || 0,
      observaciones: f.observaciones || '',
    }));

    const key = `${fFecha}|${fTipo}`;
    mapaTurnos.set(key, construirTurnoOpResumen(key, fFecha, fTipo, 'nube', filasMap));
  });

  // 2. Complementar con historial local
  (historial || []).forEach((h) => {
    const fFecha = (h.fechaTurno || (h.fecha ? h.fecha.slice(0, 10) : '')).trim();
    const fTipo = (h.tipoTurno || '').trim().toLowerCase();
    if (!fFecha) return;
    if (currentFecha && currentTipo && fFecha === currentFecha && fTipo === currentTipo) {
      return;
    }

    const key = `${fFecha}|${fTipo}`;
    if (mapaTurnos.has(key)) return;

    const coincidentes = (h.filas || []).filter((f) => opCoincide(f.op, op));
    if (coincidentes.length === 0) return;

    const filasMap: MaquinaRegistroOp[] = coincidentes.map((f) => ({
      maquina: f.maquina || 'Sin máquina',
      codigoMaquina: f.codigoMaquina || '',
      tipoProducto: resolverTipo(maqTipoMap, f.tipoProducto, f.maquina, f.codigoMaquina),
      sap: f.sap || '',
      producto: f.producto || '',
      total: Number(f.total) || 0,
      merma: Number(f.merma) || 0,
      observaciones: f.observaciones || '',
    }));

    mapaTurnos.set(key, construirTurnoOpResumen(key, fFecha, fTipo, 'local', filasMap));
  });

  const turnosAnteriores = Array.from(mapaTurnos.values()).sort((a, b) => {
    const cmpFecha = b.fecha.localeCompare(a.fecha);
    if (cmpFecha !== 0) return cmpFecha;
    const pesoTipo = (t: string) => (t === 'noche' ? 2 : t === 'dia' ? 1 : 0);
    return pesoTipo(b.tipo) - pesoTipo(a.tipo);
  });

  // 3. Procesar turno activo actual (agrupado por máquina + producto + sap)
  const filasActuales: MaquinaRegistroOp[] = [];
  (turno || []).forEach((t) => {
    const m = maquinas.find((x) => x.id === t.maquinaId);
    const regs = (t.registros || []).filter((r) => opCoincide(r.op || t.op, op));
    const gruposActual = new Map<string, MaquinaRegistroOp>();
    regs.forEach((r) => {
      const cant = Number(r.cantidad) || 0;
      const merma = Number(r.merma) || 0;
      if (cant <= 0 && merma <= 0) return;
      const prod = r.descripcion || '';
      const sap = r.sap || '';
      const gKey = `${prod}||${sap}`;
      if (!gruposActual.has(gKey)) {
        gruposActual.set(gKey, {
          maquina: m ? m.nombre : 'Máquina',
          codigoMaquina: m?.codigo || '',
          tipoProducto: m?.tipoProducto || '',
          sap,
          producto: prod,
          total: 0,
          merma: 0,
          observaciones: t.observaciones || '',
        });
      }
      const g = gruposActual.get(gKey)!;
      g.total += cant;
      g.merma += merma;
    });
    gruposActual.forEach((g) => filasActuales.push(g));
  });

  const turnoActualResumen =
    filasActuales.length > 0
      ? construirTurnoOpResumen(
          `actual|${currentFecha}|${currentTipo}`,
          currentFecha || new Date().toISOString().slice(0, 10),
          currentTipo || 'dia',
          'actual',
          filasActuales
        )
      : null;

  // 4. Totales acumulados separados por PT y PP
  const totales: TotalesOpSeparados = {
    antPP: 0,
    mermaAntPP: 0,
    actualPP: turnoActualResumen?.totalPP || 0,
    mermaActualPP: turnoActualResumen?.mermaPP || 0,
    acumPP: 0,
    mermaAcumPP: 0,
    antPT: 0,
    mermaAntPT: 0,
    actualPT: turnoActualResumen?.totalPT || 0,
    mermaActualPT: turnoActualResumen?.mermaPT || 0,
    acumPT: 0,
    mermaAcumPT: 0,
    antSinTipo: 0,
    mermaAntSinTipo: 0,
    actualSinTipo: turnoActualResumen?.totalSinTipo || 0,
    mermaActualSinTipo: turnoActualResumen?.mermaSinTipo || 0,
    acumSinTipo: 0,
    mermaAcumSinTipo: 0,
  };

  turnosAnteriores.forEach((t) => {
    totales.antPP += t.totalPP;
    totales.mermaAntPP += t.mermaPP;
    totales.antPT += t.totalPT;
    totales.mermaAntPT += t.mermaPT;
    totales.antSinTipo += t.totalSinTipo;
    totales.mermaAntSinTipo += t.mermaSinTipo;
  });

  totales.acumPP = totales.antPP + totales.actualPP;
  totales.mermaAcumPP = totales.mermaAntPP + totales.mermaActualPP;
  totales.acumPT = totales.antPT + totales.actualPT;
  totales.mermaAcumPT = totales.mermaAntPT + totales.mermaActualPT;
  totales.acumSinTipo = totales.antSinTipo + totales.actualSinTipo;
  totales.mermaAcumSinTipo = totales.mermaAntSinTipo + totales.mermaActualSinTipo;

  // 5. Agrupar totales por máquina + tipoProducto
  const maqMap = new Map<string, MaquinaTotalOp & { _turnosSet: Set<string> }>();

  const acumularFilasDeTurno = (t: TurnoOpResumen, esActual: boolean) => {
    t.filas.forEach((f) => {
      const mKey = `${f.maquina.trim().toLowerCase()}||${f.tipoProducto}`;
      if (!maqMap.has(mKey)) {
        maqMap.set(mKey, {
          key: mKey,
          maquina: f.maquina,
          codigoMaquina: f.codigoMaquina,
          tipoProducto: f.tipoProducto,
          totalAnterior: 0,
          mermaAnterior: 0,
          totalActual: 0,
          mermaActual: 0,
          totalAcumulado: 0,
          mermaAcumulada: 0,
          turnosCount: 0,
          _turnosSet: new Set<string>(),
        });
      }
      const entry = maqMap.get(mKey)!;
      if (!entry.codigoMaquina && f.codigoMaquina) {
        entry.codigoMaquina = f.codigoMaquina;
      }
      if (esActual) {
        entry.totalActual += f.total;
        entry.mermaActual += f.merma;
      } else {
        entry.totalAnterior += f.total;
        entry.mermaAnterior += f.merma;
      }
      entry.totalAcumulado = entry.totalAnterior + entry.totalActual;
      entry.mermaAcumulada = entry.mermaAnterior + entry.mermaActual;
      entry._turnosSet.add(t.key);
      entry.turnosCount = entry._turnosSet.size;
    });
  };

  turnosAnteriores.forEach((t) => acumularFilasDeTurno(t, false));
  if (turnoActualResumen) acumularFilasDeTurno(turnoActualResumen, true);

  const pesoTipo = (tp: string) => (tp === 'pt' ? 0 : tp === 'pp' ? 1 : 2);
  const porMaquina: MaquinaTotalOp[] = Array.from(maqMap.values())
    .map(({ _turnosSet, ...rest }) => rest)
    .sort((a, b) => {
      const pt = pesoTipo(a.tipoProducto) - pesoTipo(b.tipoProducto);
      if (pt !== 0) return pt;
      return b.totalAcumulado - a.totalAcumulado;
    });

  return {
    turnosAnteriores,
    turnoActualResumen,
    porMaquina,
    totales,
    nubeDisponible,
    nubeMensaje,
  };
}

export interface MaquinaTurnoFila {
  op: string;
  sap: string;
  producto: string;
  total: number;
  merma: number;
  observaciones: string;
}

export interface TurnoMaquinaResumen {
  key: string;
  fecha: string;
  tipo: string;
  fuente: 'nube' | 'local' | 'actual';
  totalProducido: number;
  totalMerma: number;
  filas: MaquinaTurnoFila[];
}

export interface MaquinaOpConsolidado {
  key: string;
  op: string;
  sap: string;
  producto: string;
  cliente: string;
  totalAnterior: number;
  mermaAnterior: number;
  totalActual: number;
  mermaActual: number;
  totalAcumulado: number;
  mermaAcumulada: number;
  turnosCount: number;
}

export interface ConsultaMaquinaResultado {
  turnosAnteriores: TurnoMaquinaResumen[];
  turnoActualResumen: TurnoMaquinaResumen | null;
  porOp: MaquinaOpConsolidado[];
  totalAnterior: number;
  mermaAnterior: number;
  totalActual: number;
  mermaActual: number;
  totalProducido: number;
  totalMerma: number;
  promedioPorTurno: number;
  nubeDisponible: boolean;
  nubeMensaje: string;
}

function maquinaCoincide(
  fila: { maquina?: string; codigoMaquina?: string },
  maquina: Maquina
): boolean {
  const fNombre = (fila.maquina || '').trim().toLowerCase();
  const mNombre = (maquina.nombre || '').trim().toLowerCase();
  if (fNombre && mNombre && fNombre === mNombre) return true;

  const fCod = (fila.codigoMaquina || '').trim().toLowerCase();
  const mCod = (maquina.codigo || '').trim().toLowerCase();
  if (fCod && mCod && fCod === mCod) return true;

  return false;
}

function fechaEnRango(fecha: string, desde?: string, hasta?: string): boolean {
  const f = (fecha || '').trim();
  if (!f) return false;
  if (desde && f < desde) return false;
  if (hasta && f > hasta) return false;
  return true;
}

export async function consultarHistorialMaquina(params: {
  maquina: Maquina;
  historial: HistorialTurnoLocal[];
  metaTurno: MetaTurno;
  session: SupabaseSession | null;
  turno?: TurnoMaquina[];
  catalogo?: CatalogoItem[];
  fechaDesde?: string;
  fechaHasta?: string;
}): Promise<ConsultaMaquinaResultado> {
  const { maquina, historial, metaTurno, session, turno, catalogo = [], fechaDesde, fechaHasta } = params;

  let dbRows: TurnoDBRow[] = [];
  let nubeDisponible = false;
  let nubeMensaje = '';

  const cfg = getSupabaseConfig(false);
  if (cfg) {
    try {
      let q = 'turnos?select=fecha,tipo,filas&order=fecha.desc,tipo.asc&limit=500';
      if (fechaDesde) q += `&fecha=gte.${encodeURIComponent(fechaDesde)}`;
      if (fechaHasta) q += `&fecha=lte.${encodeURIComponent(fechaHasta)}`;
      const res = await sbFetch<TurnoDBRow[]>(cfg, q, undefined, session);
      if (Array.isArray(res)) {
        dbRows = res;
        nubeDisponible = true;
      }
    } catch {
      nubeDisponible = false;
      nubeMensaje = session
        ? 'No se pudo consultar la base de datos en la nube. Mostrando registros locales.'
        : 'Inicia sesión en "Turnos Guardados" para consultar también los turnos guardados en la nube.';
    }
  } else {
    nubeMensaje = 'Mostrando solo registros guardados en este dispositivo.';
  }

  const mapaTurnos = new Map<string, TurnoMaquinaResumen>();
  const currentFecha = (metaTurno.fecha || '').trim();
  const currentTipo = (metaTurno.tipo || '').trim().toLowerCase();

  // 1. Procesar turnos desde Supabase
  dbRows.forEach((rec) => {
    const fFecha = (rec.fecha || '').trim();
    const fTipo = (rec.tipo || '').trim().toLowerCase();
    if (!fechaEnRango(fFecha, fechaDesde, fechaHasta)) return;
    if (currentFecha && currentTipo && fFecha === currentFecha && fTipo === currentTipo) {
      return;
    }

    const coincidentes = (rec.filas || []).filter((f) => maquinaCoincide(f, maquina));
    if (coincidentes.length === 0) return;

    const filasMap: MaquinaTurnoFila[] = coincidentes.map((f) => ({
      op: f.op || '',
      sap: f.sap || '',
      producto: f.producto || '',
      total: Number(f.total) || 0,
      merma: Number(f.merma) || 0,
      observaciones: f.observaciones || '',
    }));

    const key = `${fFecha}|${fTipo}`;
    mapaTurnos.set(key, {
      key,
      fecha: fFecha,
      tipo: fTipo,
      fuente: 'nube',
      totalProducido: filasMap.reduce((s, r) => s + r.total, 0),
      totalMerma: filasMap.reduce((s, r) => s + r.merma, 0),
      filas: filasMap,
    });
  });

  // 2. Complementar con historial local
  (historial || []).forEach((h) => {
    const fFecha = (h.fechaTurno || (h.fecha ? h.fecha.slice(0, 10) : '')).trim();
    const fTipo = (h.tipoTurno || '').trim().toLowerCase();
    if (!fechaEnRango(fFecha, fechaDesde, fechaHasta)) return;
    if (currentFecha && currentTipo && fFecha === currentFecha && fTipo === currentTipo) {
      return;
    }

    const key = `${fFecha}|${fTipo}`;
    if (mapaTurnos.has(key)) return;

    const coincidentes = (h.filas || []).filter((f) => maquinaCoincide(f, maquina));
    if (coincidentes.length === 0) return;

    const filasMap: MaquinaTurnoFila[] = coincidentes.map((f) => ({
      op: f.op || '',
      sap: f.sap || '',
      producto: f.producto || '',
      total: Number(f.total) || 0,
      merma: Number(f.merma) || 0,
      observaciones: f.observaciones || '',
    }));

    mapaTurnos.set(key, {
      key,
      fecha: fFecha,
      tipo: fTipo,
      fuente: 'local',
      totalProducido: filasMap.reduce((s, r) => s + r.total, 0),
      totalMerma: filasMap.reduce((s, r) => s + r.merma, 0),
      filas: filasMap,
    });
  });

  const turnosAnteriores = Array.from(mapaTurnos.values()).sort((a, b) => {
    const cmpFecha = b.fecha.localeCompare(a.fecha);
    if (cmpFecha !== 0) return cmpFecha;
    const pesoTipo = (t: string) => (t === 'noche' ? 2 : t === 'dia' ? 1 : 0);
    return pesoTipo(b.tipo) - pesoTipo(a.tipo);
  });

  // 3. Procesar turno activo actual si se pasó `turno`
  let turnoActualResumen: TurnoMaquinaResumen | null = null;
  const activeFecha = currentFecha || new Date().toISOString().slice(0, 10);
  if (turno && fechaEnRango(activeFecha, fechaDesde, fechaHasta)) {
    const gruposActual = new Map<string, MaquinaTurnoFila>();
    turno
      .filter((t) => t.maquinaId === maquina.id)
      .forEach((t) => {
        (t.registros || []).forEach((r) => {
          const cant = Number(r.cantidad) || 0;
          const merma = Number(r.merma) || 0;
          if (cant <= 0 && merma <= 0) return;
          const opVal = (r.op || t.op || '').trim();
          const sapVal = (r.sap || '').trim();
          const prodVal = (r.descripcion || '').trim();
          const gKey = `${opVal.toLowerCase()}||${sapVal.toLowerCase()}||${prodVal.toLowerCase()}`;
          if (!gruposActual.has(gKey)) {
            gruposActual.set(gKey, {
              op: opVal,
              sap: sapVal,
              producto: prodVal,
              total: 0,
              merma: 0,
              observaciones: t.observaciones || '',
            });
          }
          const g = gruposActual.get(gKey)!;
          g.total += cant;
          g.merma += merma;
        });
      });

    const filasActuales = Array.from(gruposActual.values());
    if (filasActuales.length > 0) {
      turnoActualResumen = {
        key: `actual|${activeFecha}|${currentTipo || 'dia'}`,
        fecha: activeFecha,
        tipo: currentTipo || 'dia',
        fuente: 'actual',
        totalProducido: filasActuales.reduce((s, r) => s + r.total, 0),
        totalMerma: filasActuales.reduce((s, r) => s + r.merma, 0),
        filas: filasActuales,
      };
    }
  }

  // 4. Consolidar por OP / SAP / Producto
  const opMap = new Map<string, MaquinaOpConsolidado & { _turnosSet: Set<string> }>();

  const acumularPorOp = (t: TurnoMaquinaResumen, esActual: boolean) => {
    t.filas.forEach((f) => {
      const opKey = (f.op || '').trim();
      const sapKey = (f.sap || '').trim();
      const prodKey = (f.producto || '').trim();
      const mapKey = `${opKey.toLowerCase()}||${sapKey.toLowerCase()}||${prodKey.toLowerCase()}`;

      if (!opMap.has(mapKey)) {
        const cat = opKey ? buscarEnCatalogo(catalogo, opKey) : null;
        opMap.set(mapKey, {
          key: mapKey,
          op: opKey,
          sap: sapKey,
          producto: prodKey || cat?.descripcionBase || cat?.descripcion || '',
          cliente: cat?.cliente || '',
          totalAnterior: 0,
          mermaAnterior: 0,
          totalActual: 0,
          mermaActual: 0,
          totalAcumulado: 0,
          mermaAcumulada: 0,
          turnosCount: 0,
          _turnosSet: new Set<string>(),
        });
      }

      const entry = opMap.get(mapKey)!;
      if (esActual) {
        entry.totalActual += f.total;
        entry.mermaActual += f.merma;
      } else {
        entry.totalAnterior += f.total;
        entry.mermaAnterior += f.merma;
      }
      entry.totalAcumulado = entry.totalAnterior + entry.totalActual;
      entry.mermaAcumulada = entry.mermaAnterior + entry.mermaActual;
      entry._turnosSet.add(t.key);
      entry.turnosCount = entry._turnosSet.size;
    });
  };

  turnosAnteriores.forEach((t) => acumularPorOp(t, false));
  if (turnoActualResumen) acumularPorOp(turnoActualResumen, true);

  const porOp: MaquinaOpConsolidado[] = Array.from(opMap.values())
    .map(({ _turnosSet, ...rest }) => rest)
    .sort((a, b) => b.totalAcumulado - a.totalAcumulado);

  const totalAnterior = turnosAnteriores.reduce((s, t) => s + t.totalProducido, 0);
  const mermaAnterior = turnosAnteriores.reduce((s, t) => s + t.totalMerma, 0);
  const totalActual = turnoActualResumen?.totalProducido || 0;
  const mermaActual = turnoActualResumen?.totalMerma || 0;
  const totalProducido = totalAnterior + totalActual;
  const totalMerma = mermaAnterior + mermaActual;
  const totalTurnos = turnosAnteriores.length + (turnoActualResumen ? 1 : 0);
  const promedioPorTurno = totalTurnos > 0 ? Math.round(totalProducido / totalTurnos) : 0;

  return {
    turnosAnteriores,
    turnoActualResumen,
    porOp,
    totalAnterior,
    mermaAnterior,
    totalActual,
    mermaActual,
    totalProducido,
    totalMerma,
    promedioPorTurno,
    nubeDisponible,
    nubeMensaje,
  };
}

export interface AvanceOpResult {
  anterior: number;
  actual: number;
  acumulado: number;
  meta: number;
  pct: number;
  restante: number;
  esPorModelo: boolean;
  alcance: 'pt' | 'pp' | 'sin_tipo';
  nubeUsada: boolean;
}

export async function cargarTurnosNubeSnapshot(
  session: SupabaseSession | null
): Promise<{ dbRows: TurnoDBRow[]; nubeDisponible: boolean }> {
  const cfg = getSupabaseConfig(false);
  if (!cfg) return { dbRows: [], nubeDisponible: false };
  try {
    const res = await sbFetch<TurnoDBRow[]>(
      cfg,
      'turnos?select=fecha,tipo,filas&order=fecha.desc,tipo.asc&limit=500',
      undefined,
      session
    );
    if (Array.isArray(res)) {
      return { dbRows: res, nubeDisponible: true };
    }
  } catch {
    // Fallback gracefully to local history only
  }
  return { dbRows: [], nubeDisponible: false };
}

function limpiarSufijoPliego(texto: string): string {
  return (texto || '')
    .replace(/\s*-\s*(?:Pliego\s*\d+|Tira|Retira)$/i, '')
    .trim();
}

export function calcularAvanceOpEnTarjeta(params: {
  op: string;
  modeloIdx: number;
  maquina: Maquina;
  catalogMatch: CatalogoItem | null;
  dbRows: TurnoDBRow[];
  nubeDisponible: boolean;
  historial: HistorialTurnoLocal[];
  turno: TurnoMaquina[];
  metaTurno: MetaTurno;
  maquinas: Maquina[];
}): AvanceOpResult | null {
  const {
    op,
    modeloIdx,
    maquina,
    catalogMatch,
    dbRows,
    nubeDisponible,
    historial,
    turno,
    metaTurno,
    maquinas,
  } = params;

  const opClean = (op || '').trim();
  if (!opClean || !catalogMatch || maquina.tipoProducto === 'pp') return null;

  const mod =
    modeloIdx >= 0 && catalogMatch.modelos && catalogMatch.modelos[modeloIdx]
      ? catalogMatch.modelos[modeloIdx]
      : null;
  const metaModelo = Number(mod?.cantidadTotal || 0);
  const metaOrden = Number(catalogMatch.cantidadTotal || 0);
  const esPorModelo = metaModelo > 0;
  const meta = esPorModelo ? metaModelo : metaOrden;

  if (!meta || meta <= 0) return null;

  const alcance: 'pt' | 'sin_tipo' =
    maquina.tipoProducto === 'pt' ? 'pt' : 'sin_tipo';

  const maqTipoMap = buildMaqTipoMap(maquinas);
  const currentFecha = (metaTurno.fecha || '').trim();
  const currentTipo = (metaTurno.tipo || '').trim().toLowerCase();

  const coincideModelo = (sapRaw?: string, prodRaw?: string): boolean => {
    if (!esPorModelo || !mod) return true;
    const modSap = (mod.sap || '').trim().toLowerCase();
    const rSap = (sapRaw || '').trim().toLowerCase();
    if (modSap && rSap && modSap === rSap) return true;

    const modDesc = (mod.descripcion || '').trim().toLowerCase();
    const modCorto = (mod.modelo || '').trim().toLowerCase();
    const rProd = limpiarSufijoPliego(prodRaw || '').toLowerCase();
    if (!rProd) return false;
    if (modDesc && (rProd === modDesc || rProd.includes(modDesc) || modDesc.includes(rProd))) {
      return true;
    }
    if (modCorto && (rProd === modCorto || rProd.includes(modCorto))) {
      return true;
    }
    return false;
  };

  const coincideAlcanceFila = (
    rawTipo?: string,
    nombreMaq?: string,
    codMaq?: string
  ): boolean => {
    const tipoRes = resolverTipo(maqTipoMap, rawTipo, nombreMaq, codMaq);
    if (alcance === 'pt') {
      // PT suma únicamente producción terminada (PT)
      return tipoRes === 'pt';
    }
    // Sin tipo definido en la máquina: suma solo registros de esta misma máquina
    return maquinaCoincide({ maquina: nombreMaq, codigoMaquina: codMaq }, maquina);
  };

  // 1. Sumar turnos anteriores deduplicados (nube primero, luego local)
  const vistos = new Set<string>();
  let anterior = 0;

  dbRows.forEach((rec) => {
    const fFecha = (rec.fecha || '').trim();
    const fTipo = (rec.tipo || '').trim().toLowerCase();
    if (currentFecha && currentTipo && fFecha === currentFecha && fTipo === currentTipo) {
      return;
    }
    const key = `${fFecha}|${fTipo}`;
    vistos.add(key);

    (rec.filas || []).forEach((f) => {
      if (!opCoincide(f.op, opClean)) return;
      if (!coincideAlcanceFila(f.tipoProducto, f.maquina, f.codigoMaquina)) return;
      if (!coincideModelo(f.sap, f.producto)) return;
      anterior += Number(f.total) || 0;
    });
  });

  (historial || []).forEach((h) => {
    const fFecha = (h.fechaTurno || (h.fecha ? h.fecha.slice(0, 10) : '')).trim();
    const fTipo = (h.tipoTurno || '').trim().toLowerCase();
    if (!fFecha) return;
    if (currentFecha && currentTipo && fFecha === currentFecha && fTipo === currentTipo) {
      return;
    }
    const key = `${fFecha}|${fTipo}`;
    if (vistos.has(key)) return;
    vistos.add(key);

    (h.filas || []).forEach((f) => {
      if (!opCoincide(f.op, opClean)) return;
      if (!coincideAlcanceFila(f.tipoProducto, f.maquina, f.codigoMaquina)) return;
      if (!coincideModelo(f.sap, f.producto)) return;
      anterior += Number(f.total) || 0;
    });
  });

  // 2. Sumar producción en vivo del turno actual
  let actual = 0;
  (turno || []).forEach((t) => {
    const m = maquinas.find((x) => x.id === t.maquinaId);
    if (!m) return;
    if (alcance === 'pt') {
      if (m.tipoProducto !== 'pt') return;
    } else {
      if (m.id !== maquina.id && !maquinaCoincide({ maquina: m.nombre, codigoMaquina: m.codigo }, maquina)) {
        return;
      }
    }

    (t.registros || []).forEach((r) => {
      const regOp = (r.op || t.op || '').trim();
      if (!opCoincide(regOp, opClean)) return;
      if (!coincideModelo(r.sap, r.descripcion)) return;
      actual += Number(r.cantidad) || 0;
    });
  });

  const acumulado = anterior + actual;
  const pct = meta > 0 ? (acumulado / meta) * 100 : 0;
  const restante = meta - acumulado;

  return {
    anterior,
    actual,
    acumulado,
    meta,
    pct,
    restante,
    esPorModelo,
    alcance,
    nubeUsada: nubeDisponible,
  };
}


