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

export interface MaquinaRegistroOp {
  maquina: string;
  codigoMaquina: string;
  tipoProducto: 'pp' | 'pt' | '';
  producto: string;
  total: number;
  merma: number;
  observaciones: string;
}

export interface TurnoOpResumen {
  key: string;
  fecha: string;
  tipo: string;
  fuente: 'nube' | 'local' | 'actual';
  totalPP: number;
  mermaPP: number;
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

  filasRaw.forEach((r) => {
    if (r.tipoProducto === 'pp') {
      totalPP += r.total;
      mermaPP += r.merma;
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

  return {
    key,
    fecha,
    tipo,
    fuente,
    totalPP,
    mermaPP,
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

  // 3. Procesar turno activo actual
  const filasActuales: MaquinaRegistroOp[] = [];
  (turno || []).forEach((t) => {
    const m = maquinas.find((x) => x.id === t.maquinaId);
    const regs = (t.registros || []).filter((r) => opCoincide(r.op || t.op, op));
    const sumaMaq = regs.reduce((s, r) => s + (Number(r.cantidad) || 0), 0);
    const mermaMaq = regs.reduce((s, r) => s + (Number(r.merma) || 0), 0);
    if (sumaMaq > 0 || mermaMaq > 0) {
      const desc = regs.find((r) => r.descripcion)?.descripcion || '';
      filasActuales.push({
        maquina: m ? m.nombre : 'Máquina',
        codigoMaquina: m?.codigo || '',
        tipoProducto: m?.tipoProducto || '',
        producto: desc,
        total: sumaMaq,
        merma: mermaMaq,
        observaciones: t.observaciones || '',
      });
    }
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
  fuente: 'nube' | 'local';
  totalProducido: number;
  totalMerma: number;
  filas: MaquinaTurnoFila[];
}

export interface ConsultaMaquinaResultado {
  turnosAnteriores: TurnoMaquinaResumen[];
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

export async function consultarHistorialMaquina(params: {
  maquina: Maquina;
  historial: HistorialTurnoLocal[];
  metaTurno: MetaTurno;
  session: SupabaseSession | null;
}): Promise<ConsultaMaquinaResultado> {
  const { maquina, historial, metaTurno, session } = params;

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
        : 'Inicia sesión en la pestaña Historial para consultar también los turnos guardados en la nube.';
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
    if (!fFecha) return;
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

  const totalProducido = turnosAnteriores.reduce((s, t) => s + t.totalProducido, 0);
  const totalMerma = turnosAnteriores.reduce((s, t) => s + t.totalMerma, 0);
  const promedioPorTurno =
    turnosAnteriores.length > 0 ? Math.round(totalProducido / turnosAnteriores.length) : 0;

  return {
    turnosAnteriores,
    totalProducido,
    totalMerma,
    promedioPorTurno,
    nubeDisponible,
    nubeMensaje,
  };
}

