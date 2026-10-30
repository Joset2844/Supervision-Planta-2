import { CatalogoItem, ListaItem, Maquina, TurnoMaquina, HistorialTurnoLocal, PlantillaItem, MetaTurno } from '../types';

export const STORAGE_KEYS = {
  CATALOGO: 'inventario_catalogo_op',
  LISTA: 'inventario_lista_pendiente',
  MAQUINAS: 'inventario_maquinas',
  MAQUINAS_BAK: 'inventario_maquinas_bak',
  TURNO: 'inventario_turno_maquinas',
  HISTORIAL: 'inventario_historial_turnos',
  PLANTILLA: 'inventario_plantilla_turno',
  META_TURNO: 'inventario_meta_turno',
  CAT_CURSOR: 'inventario_cat_cursor',
  URL_SHEETS_MAQUINAS: 'inventario_url_sheets_maquinas',
  SB_URL: 'inventario_sb_url',
  SB_KEY: 'inventario_sb_key',
  SB_SESION: 'inventario_sb_sesion',
};

export function getTodayDateString(): string {
  const hoy = new Date();
  return `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, '0')}-${String(hoy.getDate()).padStart(2, '0')}`;
}

export function loadInitialData() {
  let catalogo: CatalogoItem[] = [];
  try {
    const c = localStorage.getItem(STORAGE_KEYS.CATALOGO);
    if (c) catalogo = JSON.parse(c);
  } catch {}

  let migroCatalogo = false;
  catalogo.forEach((item: any) => {
    if (!item.modelos) {
      item.modelos = [{ descripcion: item.descripcion || '', sap: item.sap || '', cantidadTotal: item.cantidadTotal || '' }];
      migroCatalogo = true;
    }
    if (!item.descripcionBase) {
      item.descripcionBase = item.modelos[0] ? item.modelos[0].descripcion : (item.descripcion || '');
      migroCatalogo = true;
    }
  });
  if (migroCatalogo) {
    try { localStorage.setItem(STORAGE_KEYS.CATALOGO, JSON.stringify(catalogo)); } catch {}
  }

  let lista: ListaItem[] = [];
  try {
    const l = localStorage.getItem(STORAGE_KEYS.LISTA);
    if (l) lista = JSON.parse(l);
  } catch {}

  let migroLista = false;
  lista.forEach((item: any) => {
    if (!item.tipoEmpaque) {
      item.tipoEmpaque = 'caja';
      migroLista = true;
    }
  });
  if (migroLista) {
    try { localStorage.setItem(STORAGE_KEYS.LISTA, JSON.stringify(lista)); } catch {}
  }

  let maquinas: Maquina[] = [];
  try {
    const m = localStorage.getItem(STORAGE_KEYS.MAQUINAS);
    if (m) maquinas = JSON.parse(m);
    if (!maquinas.length) {
      const mb = localStorage.getItem(STORAGE_KEYS.MAQUINAS_BAK);
      if (mb) maquinas = JSON.parse(mb);
    }
  } catch {}

  let turno: TurnoMaquina[] = [];
  try {
    const t = localStorage.getItem(STORAGE_KEYS.TURNO);
    if (t) turno = JSON.parse(t);
  } catch {}

  let migroMaquinas = false;
  maquinas.forEach((m: any) => {
    if (typeof m.activa === 'undefined') {
      const tExistente = turno.find((x) => x.maquinaId === m.id);
      m.activa = tExistente ? !!(tExistente as any).activa : false;
      migroMaquinas = true;
    }
    if (typeof m.area === 'undefined') { m.area = ''; migroMaquinas = true; }
    if (typeof m.codigo === 'undefined') { m.codigo = ''; migroMaquinas = true; }
    if (typeof m.bloqueada === 'undefined') { m.bloqueada = !!m.codigoBloqueado; migroMaquinas = true; }
    if (typeof m.tipoProducto === 'undefined') { m.tipoProducto = ''; migroMaquinas = true; }
  });
  if (migroMaquinas) {
    try { localStorage.setItem(STORAGE_KEYS.MAQUINAS, JSON.stringify(maquinas)); } catch {}
  }

  let migroTurno = false;
  turno.forEach((t: any) => {
    if (typeof t.objetivo === 'undefined') {
      const match = catalogo.find((c) => t.op && c.op.toLowerCase() === t.op.trim().toLowerCase());
      t.objetivo = (match && match.cantidadTotal) ? (parseInt(String(match.cantidadTotal), 10) || null) : null;
      migroTurno = true;
    }
    if (typeof t.motivoParada === 'undefined') { t.motivoParada = ''; migroTurno = true; }
    if (typeof t.modeloIdx === 'undefined') { t.modeloIdx = -1; migroTurno = true; }
    if (typeof t.observaciones === 'undefined') { t.observaciones = ''; migroTurno = true; }
    if (typeof t.pliego === 'undefined') { t.pliego = ''; migroTurno = true; }
    if (!t.registros) t.registros = [];
    t.registros.forEach((r: any) => {
      if (typeof r.merma === 'undefined') { r.merma = 0; migroTurno = true; }
    });
  });
  if (migroTurno) {
    try { localStorage.setItem(STORAGE_KEYS.TURNO, JSON.stringify(turno)); } catch {}
  }

  let historial: HistorialTurnoLocal[] = [];
  try {
    const h = localStorage.getItem(STORAGE_KEYS.HISTORIAL);
    if (h) historial = JSON.parse(h);
  } catch {}

  let plantilla: PlantillaItem[] = [];
  try {
    const pl = localStorage.getItem(STORAGE_KEYS.PLANTILLA);
    if (pl) plantilla = JSON.parse(pl);
  } catch {}

  let metaTurno: MetaTurno = { tipo: 'dia', fecha: getTodayDateString() };
  try {
    const mt = localStorage.getItem(STORAGE_KEYS.META_TURNO);
    if (mt) metaTurno = JSON.parse(mt);
  } catch {}
  if (!metaTurno.fecha) {
    metaTurno.fecha = getTodayDateString();
  }

  return { catalogo, lista, maquinas, turno, historial, plantilla, metaTurno };
}

export function saveStorageItem(key: string, data: any) {
  try {
    localStorage.setItem(key, JSON.stringify(data));
  } catch (e) {
    console.error('Error saving storage item', key, e);
  }
}
