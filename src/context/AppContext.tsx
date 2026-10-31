import React, { createContext, useContext, useState, useEffect, useCallback } from 'react';
import {
  CatalogoItem,
  ListaItem,
  Maquina,
  TurnoMaquina,
  HistorialTurnoLocal,
  PlantillaItem,
  MetaTurno,
  SupabaseSession,
  FilaTurnoExcel,
} from '../types';
import {
  STORAGE_KEYS,
  loadInitialData,
  saveStorageItem,
  getTodayDateString,
} from '../services/storage';
import {
  getSupabaseConfig,
  getStoredSession,
  sbFetch,
  saveSession,
} from '../services/supabase';
import { generarFilasTurno } from '../services/excel';
import { generarReporteProduccion } from '../services/reports';

export interface ToastMessage {
  id: string;
  type: 'ok' | 'error';
  text: string;
}

interface AppContextType {
  activeTab: 'inventario' | 'maquinas' | 'produccion' | 'resumen' | 'historialdb';
  setActiveTab: (tab: 'inventario' | 'maquinas' | 'produccion' | 'resumen' | 'historialdb') => void;
  subTabInventario: 'agregar' | 'consultar';
  setSubTabInventario: (tab: 'agregar' | 'consultar') => void;

  catalogo: CatalogoItem[];
  setCatalogo: React.Dispatch<React.SetStateAction<CatalogoItem[]>>;
  lista: ListaItem[];
  setLista: React.Dispatch<React.SetStateAction<ListaItem[]>>;
  maquinas: Maquina[];
  setMaquinas: React.Dispatch<React.SetStateAction<Maquina[]>>;
  turno: TurnoMaquina[];
  setTurno: React.Dispatch<React.SetStateAction<TurnoMaquina[]>>;
  historial: HistorialTurnoLocal[];
  setHistorial: React.Dispatch<React.SetStateAction<HistorialTurnoLocal[]>>;
  plantilla: PlantillaItem[];
  setPlantilla: React.Dispatch<React.SetStateAction<PlantillaItem[]>>;
  metaTurno: MetaTurno;
  setMetaTurno: React.Dispatch<React.SetStateAction<MetaTurno>>;
  session: SupabaseSession | null;
  setSession: (s: SupabaseSession | null) => void;

  toasts: ToastMessage[];
  showToast: (type: 'ok' | 'error', text: string) => void;
  removeToast: (id: string) => void;

  sincronizarCatalogoSupabase: () => Promise<number>;
  subirCatalogoASupabase: (nuevo: CatalogoItem[]) => Promise<void>;
  reiniciarTurno: () => void;
  guardarTurnoDB: (filasPersonalizadas?: FilaTurnoExcel[], forzarReemplazo?: boolean) => Promise<{ requiereConfirmacion?: boolean; existenteInfo?: any } | void>;
}

const AppContext = createContext<AppContextType | null>(null);

function canon(o: any): string {
  if (Array.isArray(o)) return '[' + o.map(canon).join(',') + ']';
  if (o && typeof o === 'object') {
    return (
      '{' +
      Object.keys(o)
        .sort()
        .filter((k) => o[k] !== undefined)
        .map((k) => JSON.stringify(k) + ':' + canon(o[k]))
        .join(',') +
      '}'
    );
  }
  return JSON.stringify(o);
}

export const AppProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [activeTab, setActiveTab] = useState<'inventario' | 'maquinas' | 'produccion' | 'resumen' | 'historialdb'>('inventario');
  const [subTabInventario, setSubTabInventario] = useState<'agregar' | 'consultar'>('agregar');

  const [catalogo, setCatalogoState] = useState<CatalogoItem[]>([]);
  const [lista, setListaState] = useState<ListaItem[]>([]);
  const [maquinas, setMaquinasState] = useState<Maquina[]>([]);
  const [turno, setTurnoState] = useState<TurnoMaquina[]>([]);
  const [historial, setHistorialState] = useState<HistorialTurnoLocal[]>([]);
  const [plantilla, setPlantillaState] = useState<PlantillaItem[]>([]);
  const [metaTurno, setMetaTurnoState] = useState<MetaTurno>({ tipo: 'dia', fecha: getTodayDateString() });
  const [session, setSessionState] = useState<SupabaseSession | null>(null);
  const [toasts, setToasts] = useState<ToastMessage[]>([]);
  const lastToastRef = React.useRef<{ text: string; time: number }>({ text: '', time: 0 });

  const showToast = useCallback((type: 'ok' | 'error', text: string) => {
    const now = Date.now();
    if (lastToastRef.current.text === text && now - lastToastRef.current.time < 800) {
      return;
    }
    lastToastRef.current = { text, time: now };
    const id = now + '-' + Math.random().toString(36).slice(2, 6);
    setToasts((prev) => [...prev, { id, type, text }]);
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 2800);
  }, []);

  const removeToast = useCallback((id: string) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  }, []);

  // Initialize
  useEffect(() => {
    const initial = loadInitialData();
    setCatalogoState(initial.catalogo);
    setListaState(initial.lista);
    setMaquinasState(initial.maquinas);
    setTurnoState(initial.turno);
    setHistorialState(initial.historial);
    setPlantillaState(initial.plantilla);
    setMetaTurnoState(initial.metaTurno);
    setSessionState(getStoredSession());

    // Request persistent storage
    try {
      if (navigator.storage && navigator.storage.persist) {
        navigator.storage.persist();
      }
    } catch {}
  }, []);

  // Save wrappers
  const setCatalogo: React.Dispatch<React.SetStateAction<CatalogoItem[]>> = useCallback((valOrFn) => {
    setCatalogoState((prev) => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      saveStorageItem(STORAGE_KEYS.CATALOGO, next);
      return next;
    });
  }, []);

  const setLista: React.Dispatch<React.SetStateAction<ListaItem[]>> = useCallback((valOrFn) => {
    setListaState((prev) => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      saveStorageItem(STORAGE_KEYS.LISTA, next);
      return next;
    });
  }, []);

  const setMaquinas: React.Dispatch<React.SetStateAction<Maquina[]>> = useCallback((valOrFn) => {
    setMaquinasState((prev) => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      saveStorageItem(STORAGE_KEYS.MAQUINAS, next);
      if (next.length) saveStorageItem(STORAGE_KEYS.MAQUINAS_BAK, next);
      return next;
    });
  }, []);

  const setTurno: React.Dispatch<React.SetStateAction<TurnoMaquina[]>> = useCallback((valOrFn) => {
    setTurnoState((prev) => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      saveStorageItem(STORAGE_KEYS.TURNO, next);
      return next;
    });
  }, []);

  const setHistorial: React.Dispatch<React.SetStateAction<HistorialTurnoLocal[]>> = useCallback((valOrFn) => {
    setHistorialState((prev) => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      saveStorageItem(STORAGE_KEYS.HISTORIAL, next);
      return next;
    });
  }, []);

  const setPlantilla: React.Dispatch<React.SetStateAction<PlantillaItem[]>> = useCallback((valOrFn) => {
    setPlantillaState((prev) => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      saveStorageItem(STORAGE_KEYS.PLANTILLA, next);
      return next;
    });
  }, []);

  const setMetaTurno: React.Dispatch<React.SetStateAction<MetaTurno>> = useCallback((valOrFn) => {
    setMetaTurnoState((prev) => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      saveStorageItem(STORAGE_KEYS.META_TURNO, next);
      return next;
    });
  }, []);

  const setSession = useCallback((s: SupabaseSession | null) => {
    setSessionState(s);
    saveSession(s);
  }, []);

  // Supabase Sync Catalogo
  const sincronizarCatalogoSupabase = useCallback(async (): Promise<number> => {
    const cfg = getSupabaseConfig(false);
    if (!cfg) return 0;
    let cursor = localStorage.getItem(STORAGE_KEYS.CAT_CURSOR) || '';
    if (!catalogo.length) cursor = '';
    const filtro = cursor ? '&actualizado_en=gt.' + encodeURIComponent(cursor) : '';
    const PAG = 1000;
    let pagina = 0;
    const recibidas: CatalogoItem[] = [];
    let maxTs = cursor;
    let cambiadas = 0;

    async function pedir(): Promise<void> {
      const rows: any[] = await sbFetch(
        cfg!,
        `catalogo_op?select=datos,actualizado_en&order=actualizado_en.asc,op.asc&limit=${PAG}&offset=${pagina * PAG}${filtro}`,
        undefined,
        session
      );
      if (rows && rows.length) {
        rows.forEach((r) => {
          if (r.datos && r.datos.op) recibidas.push(r.datos);
          if (r.actualizado_en && r.actualizado_en > maxTs) maxTs = r.actualizado_en;
        });
        if (rows.length === PAG) {
          pagina++;
          return pedir();
        }
      }
    }

    await pedir();
    if (!recibidas.length) return 0;

    setCatalogoState((currentCatalogo) => {
      let nuevoCatalogo = [...currentCatalogo];
      if (!cursor) {
        nuevoCatalogo = recibidas;
        cambiadas = recibidas.length;
      } else {
        const indice: Record<string, number> = {};
        nuevoCatalogo.forEach((c, k) => {
          indice[c.op] = k;
        });
        recibidas.forEach((it) => {
          if (indice[it.op] !== undefined) {
            if (canon(nuevoCatalogo[indice[it.op]]) !== canon(it)) {
              nuevoCatalogo[indice[it.op]] = it;
              cambiadas++;
            }
          } else {
            indice[it.op] = nuevoCatalogo.length;
            nuevoCatalogo.push(it);
            cambiadas++;
          }
        });
      }
      try {
        localStorage.setItem(STORAGE_KEYS.CAT_CURSOR, maxTs);
      } catch {}
      saveStorageItem(STORAGE_KEYS.CATALOGO, nuevoCatalogo);
      return nuevoCatalogo;
    });

    return cambiadas;
  }, [catalogo.length, session]);

  // Upload catalog to Supabase
  const subirCatalogoASupabase = useCallback(
    async (nuevo: CatalogoItem[]) => {
      if (!session) {
        showToast('error', 'Inicia sesión en la pestaña Historial para cargar el catálogo para todos');
        return;
      }
      const cfg = getSupabaseConfig(false);
      if (!cfg) return;

      await sincronizarCatalogoSupabase();
      const cursor = localStorage.getItem(STORAGE_KEYS.CAT_CURSOR);
      const base = cursor ? catalogo : [];
      const conocidas: Record<string, string> = {};
      base.forEach((c) => {
        conocidas[c.op] = canon(c);
      });

      const cambios: { op: string; datos: CatalogoItem }[] = [];
      let nuevas = 0;
      let modificadas = 0;
      let iguales = 0;

      nuevo.forEach((it) => {
        const previo = conocidas[it.op];
        if (previo === undefined) {
          nuevas++;
          cambios.push({ op: it.op, datos: it });
        } else if (previo !== canon(it)) {
          modificadas++;
          cambios.push({ op: it.op, datos: it });
        } else {
          iguales++;
        }
      });

      if (!cambios.length) {
        showToast('ok', `Sin cambios: las ${iguales} OP ya están en el catálogo`);
        return;
      }

      const confirmed = window.confirm(
        `Catálogo del Excel: ${nuevo.length} OP\n\n• Nuevas: ${nuevas}\n• Modificadas: ${modificadas}\n• Sin cambios: ${iguales}\n\nSe subirán ${cambios.length} OP para todos (no se borra ninguna).\n¿Confirmas?`
      );
      if (!confirmed) return;

      const LOTE = 500;
      let k = 0;
      async function subirLote(): Promise<void> {
        if (k >= cambios.length) return;
        const lote = cambios.slice(k, k + LOTE);
        k += LOTE;
        await sbFetch(
          cfg!,
          'catalogo_op?on_conflict=op',
          { method: 'POST', json: lote, prefer: 'resolution=merge-duplicates,return=minimal' },
          session
        );
        return subirLote();
      }

      await subirLote();
      await sincronizarCatalogoSupabase();
      showToast('ok', `Catálogo actualizado para todos: ${nuevas} nuevas, ${modificadas} modificadas`);
    },
    [session, catalogo, sincronizarCatalogoSupabase, showToast]
  );

  // Auto sync on mount if online
  useEffect(() => {
    if (navigator.onLine) {
      setTimeout(() => {
        sincronizarCatalogoSupabase().catch(() => {});
      }, 600);
    }
  }, [sincronizarCatalogoSupabase]);

  // Restart shift
  const reiniciarTurno = useCallback(() => {
    if (turno.length === 0) return;
    const totalActual = turno.reduce(
      (acc, t) => acc + (t.registros || []).reduce((rAcc, r) => rAcc + (r.cantidad || 0), 0),
      0
    );
    const confirmed = window.confirm(
      `¿Reiniciar el turno?\n\nEsto borra lo registrado en pantalla (${totalActual} u en total).\n\nSe guardará una copia en Historial (dentro de esta pestaña) antes de borrar, pero no se sube a Supabase ni a Sheets automáticamente.\n\n¿Confirmas?`
    );
    if (!confirmed) return;

    const filas = generarFilasTurno(maquinas, turno, catalogo);
    const nuevoHistorial = [
      {
        id: Date.now() + '-' + Math.random().toString(36).slice(2, 7),
        fecha: new Date().toISOString(),
        fechaTurno: metaTurno.fecha,
        tipoTurno: metaTurno.tipo,
        resumen: generarReporteProduccion(maquinas, turno, catalogo),
        filas,
      },
      ...historial,
    ].slice(0, 30);

    setHistorial(nuevoHistorial);
    setTurno([]);
    const hRe = new Date();
    const hoyStr = `${hRe.getFullYear()}-${String(hRe.getMonth() + 1).padStart(2, '0')}-${String(hRe.getDate()).padStart(2, '0')}`;
    setMetaTurno({ fecha: hoyStr, tipo: '' });
    showToast('ok', 'Turno reiniciado. Elige fecha y turno (Día/Noche)');
  }, [turno, maquinas, catalogo, historial, setHistorial, setTurno, setMetaTurno, showToast]);

  // Save active shift to Supabase
  const guardarTurnoDB = useCallback(
    async (filasPersonalizadas?: FilaTurnoExcel[], forzarReemplazo = false): Promise<{ requiereConfirmacion?: boolean; existenteInfo?: any } | void> => {
      if (!session) {
        showToast('error', 'Inicia sesión para subir los datos a la nube');
        return;
      }
      const fecha = metaTurno.fecha || getTodayDateString();
      const tipo = metaTurno.tipo || 'dia';

      const cfg = getSupabaseConfig(false);
      if (!cfg) {
        showToast('error', 'Falta configuración de Supabase');
        return;
      }

      const filas = filasPersonalizadas || generarFilasTurno(maquinas, turno, catalogo);
      if (filas.length === 0) {
        showToast('error', 'No hay máquinas activas para guardar');
        return;
      }

      const totalFilas = (fs: FilaTurnoExcel[]) => fs.reduce((a, f) => a + (f.total || 0), 0);
      const formatearFecha = (d: string) => {
        const p = d.split('-');
        return p.length === 3 ? `${p[2]}/${p[1]}/${p[0]}` : d;
      };
      const etiqueta = `${formatearFecha(fecha)} - ${tipo === 'noche' ? 'Turno Noche' : 'Turno Día'}`;
      const ahora = new Date().toISOString();

      try {
        const ex: any[] = await sbFetch(cfg, `turnos?select=*&fecha=eq.${fecha}&tipo=eq.${tipo}`, undefined, session);
        const existente = ex && ex[0];
        if (existente && !forzarReemplazo) {
          const esAdmin = session.usuario === 'admin';
          const duenio = !!existente.created_by_email && existente.created_by_email === session.email;
          if (!esAdmin && !duenio) {
            showToast('error', 'Ese turno pertenece a otro usuario; solo el admin puede reemplazarlo.');
            return;
          }
          const horas = (Date.now() - new Date(existente.creado_en || existente.updated_at || 0).getTime()) / 3600000;
          if (!esAdmin && (isNaN(horas) || horas > 24)) {
            showToast('error', `Pasaron más de 24 h desde que se guardó ${etiqueta}. Puedes editarlo en Historial.`);
            return;
          }

          // Return notice so UI displays custom modal instead of failing native window.confirm
          return {
            requiereConfirmacion: true,
            existenteInfo: {
              fecha,
              tipo,
              etiqueta,
              totalExistente: totalFilas(existente.filas || []),
              totalNuevo: totalFilas(filas),
              filasCount: filas.length,
            },
          };
        }

        if (existente && forzarReemplazo) {
          await sbFetch(
            cfg,
            `turnos?fecha=eq.${fecha}&tipo=eq.${tipo}`,
            { method: 'PATCH', json: { filas, turno, updated_at: ahora }, prefer: 'return=minimal' },
            session
          );

          const copiaLocal: HistorialTurnoLocal = {
            id: Date.now() + '-' + Math.random().toString(36).slice(2, 7),
            fecha: ahora,
            fechaTurno: metaTurno.fecha,
            tipoTurno: metaTurno.tipo,
            resumen: generarReporteProduccion(maquinas, turno, catalogo),
            filas,
          };
          setHistorial((prev) => [copiaLocal, ...prev].slice(0, 30));

          showToast('ok', `Turno actualizado en BD: ${etiqueta} (${totalFilas(filas)} u)`);
          return;
        }

        // New record insert
        await sbFetch(
          cfg,
          'turnos',
          {
            method: 'POST',
            json: [
              {
                fecha,
                tipo,
                filas,
                turno,
                created_by: session.user_id,
                created_by_email: session.email,
                updated_at: ahora,
              },
            ],
            prefer: 'return=minimal',
          },
          session
        );

        // Also add to local history so it is immediately visible
        const copiaLocal: HistorialTurnoLocal = {
          id: Date.now() + '-' + Math.random().toString(36).slice(2, 7),
          fecha: ahora,
          fechaTurno: metaTurno.fecha,
          tipoTurno: metaTurno.tipo,
          resumen: generarReporteProduccion(maquinas, turno, catalogo),
          filas,
        };
        setHistorial((prev) => [copiaLocal, ...prev].slice(0, 30));

        showToast('ok', `Guardado en base de datos: ${etiqueta} (${totalFilas(filas)} u)`);
      } catch (err: any) {
        console.error('Error al guardar en Supabase:', err);
        if (err.message && (err.message.includes('401') || err.message.includes('Sesión vencida') || err.message.includes('JWT') || err.message.includes('row-level security'))) {
          setSession(null);
          showToast('error', 'Sesión vencida o no autorizada. Por favor inicia sesión nuevamente.');
          return;
        }
        const msg = /HTTP 409|duplicate|23505/.test(err.message)
          ? 'Ese turno ya existe; no se sobrescribió'
          : 'No se pudo guardar: ' + (err.message || '').slice(0, 80);
        showToast('error', msg);
      }
    },
    [session, metaTurno, maquinas, turno, catalogo, setHistorial, showToast, setSession]
  );

  return (
    <AppContext.Provider
      value={{
        activeTab,
        setActiveTab,
        subTabInventario,
        setSubTabInventario,
        catalogo,
        setCatalogo,
        lista,
        setLista,
        maquinas,
        setMaquinas,
        turno,
        setTurno,
        historial,
        setHistorial,
        plantilla,
        setPlantilla,
        metaTurno,
        setMetaTurno,
        session,
        setSession,
        toasts,
        showToast,
        removeToast,
        sincronizarCatalogoSupabase,
        subirCatalogoASupabase,
        reiniciarTurno,
        guardarTurnoDB,
      }}
    >
      {children}
    </AppContext.Provider>
  );
};

export const useApp = () => {
  const context = useContext(AppContext);
  if (!context) throw new Error('useApp must be used within an AppProvider');
  return context;
};
