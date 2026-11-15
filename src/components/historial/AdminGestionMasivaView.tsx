import React, { useState, useEffect, useCallback, useMemo } from 'react';
import {
  ArrowLeft,
  Search,
  Trash2,
  Save,
  CheckSquare,
  Square,
  RefreshCw,
  AlertTriangle,
  Sun,
  Moon,
  ArrowRightLeft,
  RotateCcw,
  Filter,
  ShieldCheck,
} from 'lucide-react';
import {
  FilaTurnoExcel,
  Maquina,
  SupabaseSession,
  TurnoDBRow,
} from '../../types';
import { getSupabaseConfig, sbFetch } from '../../services/supabase';

interface AdminGestionMasivaViewProps {
  session: SupabaseSession;
  maquinas: Maquina[];
  onClose: () => void;
  onDataChanged: () => void;
  showToast: (type: 'ok' | 'error' | 'info', msg: string) => void;
}

interface FlatRow {
  rowKey: string;
  shiftKey: string;
  fecha: string;
  tipo: string;
  rowIndex: number;
  autor: string;
  updatedAt?: string;
  fila: FilaTurnoExcel;
}

const fechaCorta = (iso: string): string => {
  if (!iso) return '';
  const p = iso.split('-');
  return p.length === 3 ? `${p[2]}/${p[1]}/${p[0]}` : iso;
};

export const AdminGestionMasivaView: React.FC<AdminGestionMasivaViewProps> = ({
  session,
  maquinas,
  onClose,
  onDataChanged,
  showToast,
}) => {
  // Query filters for Supabase
  const [desde, setDesde] = useState('');
  const [hasta, setHasta] = useState('');
  const [filtroTurno, setFiltroTurno] = useState('');
  const [limite, setLimite] = useState<number>(300);

  // Loaded shifts from Supabase
  const [turnos, setTurnos] = useState<TurnoDBRow[]>([]);
  const [cargando, setCargando] = useState(false);

  // In-memory audit filters
  const [filtroMaquina, setFiltroMaquina] = useState('');
  const [filtroOps, setFiltroOps] = useState('');
  const [filtroTexto, setFiltroTexto] = useState('');
  const [soloConProduccion, setSoloConProduccion] = useState(false);

  // Selection & staged edits
  const [selectedKeys, setSelectedKeys] = useState<Set<string>>(new Set());
  const [draftEdits, setDraftEdits] = useState<Record<string, Partial<FilaTurnoExcel>>>({});

  // Bulk apply helpers
  const [maquinaDestinoId, setMaquinaDestinoId] = useState('');
  const [opLote, setOpLote] = useState('');
  const [obsLote, setObsLote] = useState('');

  // Save / Delete states
  const [guardandoCambios, setGuardandoCambios] = useState(false);
  const [modalEliminar, setModalEliminar] = useState<{
    rows: FlatRow[];
    turnosAfectados: number;
    totalUnidades: number;
    opsResumen: string[];
    maquinasResumen: string[];
  } | null>(null);
  const [eliminarTurnoSiQuedaVacio, setEliminarTurnoSiQuedaVacio] = useState(false);
  const [eliminandoMasivo, setEliminandoMasivo] = useState(false);

  const cargarTurnosAuditoria = useCallback(async () => {
    const cfg = getSupabaseConfig(false);
    if (!cfg) {
      showToast('error', 'Falta configuración de Supabase');
      return;
    }
    setCargando(true);
    let q = `turnos?select=fecha,tipo,filas,created_by,created_by_email,updated_at&order=fecha.desc,tipo.asc&limit=${limite}`;
    if (desde) q += '&fecha=gte.' + encodeURIComponent(desde);
    if (hasta) q += '&fecha=lte.' + encodeURIComponent(hasta);
    if (filtroTurno) q += '&tipo=eq.' + encodeURIComponent(filtroTurno);

    try {
      const rows: TurnoDBRow[] = await sbFetch(cfg, q, undefined, session);
      setTurnos(rows || []);
      setSelectedKeys(new Set());
      setDraftEdits({});
    } catch (err: any) {
      showToast('error', 'Error al cargar registros: ' + (err.message || ''));
    } finally {
      setCargando(false);
    }
  }, [desde, hasta, filtroTurno, limite, session, showToast]);

  useEffect(() => {
    cargarTurnosAuditoria();
  }, [cargarTurnosAuditoria]);

  // Flatten all shift rows
  const allRows: FlatRow[] = useMemo(() => {
    const out: FlatRow[] = [];
    for (const t of turnos) {
      const filas = t.filas || [];
      const autor = t.created_by_email ? t.created_by_email.split('@')[0] : 'desconocido';
      const shiftKey = `${t.fecha}::${t.tipo}`;
      for (let i = 0; i < filas.length; i++) {
        out.push({
          rowKey: `${shiftKey}::${i}`,
          shiftKey,
          fecha: t.fecha,
          tipo: t.tipo,
          rowIndex: i,
          autor,
          updatedAt: t.updated_at,
          fila: filas[i],
        });
      }
    }
    return out;
  }, [turnos]);

  // Effective row values (original merged with draftEdits)
  const getEffectiveFila = useCallback(
    (row: FlatRow): FilaTurnoExcel => {
      const patch = draftEdits[row.rowKey];
      return patch ? { ...row.fila, ...patch } : row.fila;
    },
    [draftEdits]
  );

  // Distinct machine options present in loaded rows + catalog count
  const maquinasConConteo = useMemo(() => {
    const map = new Map<string, { nombre: string; count: number; ops: Set<string> }>();
    for (const r of allRows) {
      const f = getEffectiveFila(r);
      const nombre = (f.maquina || '').trim();
      if (!nombre) continue;
      if (!map.has(nombre)) {
        map.set(nombre, { nombre, count: 0, ops: new Set() });
      }
      const entry = map.get(nombre)!;
      entry.count += 1;
      const op = (f.op || '').trim();
      if (op && op !== '-') entry.ops.add(op);
    }
    for (const m of maquinas) {
      if (!map.has(m.nombre)) {
        map.set(m.nombre, { nombre: m.nombre, count: 0, ops: new Set() });
      }
    }
    return Array.from(map.values()).sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
  }, [allRows, maquinas, getEffectiveFila]);

  // Rows filtered by machine/date/text/zero (before OP chip quick-summary)
  const rowsFilteredBeforeOp = useMemo(() => {
    const textLower = filtroTexto.trim().toLowerCase();
    return allRows.filter((r) => {
      const f = getEffectiveFila(r);
      if (filtroMaquina && (f.maquina || '').trim() !== filtroMaquina) return false;
      if (soloConProduccion && Number(f.total || 0) <= 0) return false;
      if (textLower) {
        const hay = `${f.producto || ''} ${f.sap || ''} ${f.observaciones || ''} ${f.codigoMaquina || ''} ${r.autor}`.toLowerCase();
        if (!hay.includes(textLower)) return false;
      }
      return true;
    });
  }, [allRows, getEffectiveFila, filtroMaquina, soloConProduccion, filtroTexto]);

  // Parse OP filter tokens (supports comma/space/newline separated OPs)
  const opTokens = useMemo(() => {
    return filtroOps
      .split(/[\s,;]+/)
      .map((s) => s.trim().toLowerCase())
      .filter(Boolean);
  }, [filtroOps]);

  // Visible rows after OP filter
  const visibleRows = useMemo(() => {
    if (opTokens.length === 0) return rowsFilteredBeforeOp;
    return rowsFilteredBeforeOp.filter((r) => {
      const f = getEffectiveFila(r);
      const opStr = String(f.op || '').trim().toLowerCase();
      return opTokens.some((tok) => opStr.includes(tok));
    });
  }, [rowsFilteredBeforeOp, opTokens, getEffectiveFila]);

  // OP summary chips on current machine/date filter (helps spot wrong OPs immediately!)
  const opsSummary = useMemo(() => {
    const map = new Map<
      string,
      { op: string; rowKeys: string[]; count: number; totalUnidades: number; productoSample: string }
    >();
    for (const r of rowsFilteredBeforeOp) {
      const f = getEffectiveFila(r);
      const opKey = (f.op || '').trim() || '(Sin OP)';
      if (!map.has(opKey)) {
        map.set(opKey, {
          op: opKey,
          rowKeys: [],
          count: 0,
          totalUnidades: 0,
          productoSample: f.producto || '',
        });
      }
      const item = map.get(opKey)!;
      item.rowKeys.push(r.rowKey);
      item.count += 1;
      item.totalUnidades += Number(f.total || 0);
      if (!item.productoSample && f.producto) item.productoSample = f.producto;
    }
    return Array.from(map.values()).sort((a, b) => b.count - a.count || a.op.localeCompare(b.op));
  }, [rowsFilteredBeforeOp, getEffectiveFila]);

  const dirtyKeys = useMemo(() => Object.keys(draftEdits), [draftEdits]);

  const selectedVisibleCount = useMemo(() => {
    let c = 0;
    for (const r of visibleRows) {
      if (selectedKeys.has(r.rowKey)) c++;
    }
    return c;
  }, [visibleRows, selectedKeys]);

  const selectedStats = useMemo(() => {
    const selectedRows = allRows.filter((r) => selectedKeys.has(r.rowKey));
    const shifts = new Set(selectedRows.map((r) => r.shiftKey));
    const units = selectedRows.reduce((acc, r) => acc + Number(getEffectiveFila(r).total || 0), 0);
    return {
      rows: selectedRows,
      count: selectedRows.length,
      shiftsCount: shifts.size,
      totalUnits: units,
    };
  }, [allRows, selectedKeys, getEffectiveFila]);

  // Handlers
  const toggleRowSelection = (rowKey: string) => {
    setSelectedKeys((prev) => {
      const next = new Set(prev);
      if (next.has(rowKey)) next.delete(rowKey);
      else next.add(rowKey);
      return next;
    });
  };

  const seleccionarTodosVisibles = () => {
    setSelectedKeys((prev) => {
      const next = new Set(prev);
      for (const r of visibleRows) next.add(r.rowKey);
      return next;
    });
  };

  const deseleccionarTodos = () => {
    setSelectedKeys(new Set());
  };

  const toggleSeleccionPorOp = (opKeys: string[]) => {
    setSelectedKeys((prev) => {
      const next = new Set(prev);
      const allSelected = opKeys.every((k) => next.has(k));
      if (allSelected) {
        for (const k of opKeys) next.delete(k);
      } else {
        for (const k of opKeys) next.add(k);
      }
      return next;
    });
  };

  const stageRowPatch = (row: FlatRow, patch: Partial<FilaTurnoExcel>) => {
    setDraftEdits((prev) => {
      const merged = { ...(prev[row.rowKey] || {}), ...patch };
      // Check if everything matches original
      const orig = row.fila;
      const keys = Object.keys(merged) as (keyof FilaTurnoExcel)[];
      const differs = keys.some((k) => merged[k] !== orig[k]);
      const next = { ...prev };
      if (!differs) {
        delete next[row.rowKey];
      } else {
        next[row.rowKey] = merged;
      }
      return next;
    });
  };

  const handleStageMachineChange = (row: FlatRow, machineName: string) => {
    const cat = maquinas.find((m) => m.nombre === machineName);
    if (cat) {
      stageRowPatch(row, {
        maquina: cat.nombre,
        codigoMaquina: cat.codigo || row.fila.codigoMaquina || '',
        area: cat.area || row.fila.area || '',
        tipoProducto: cat.tipoProducto ? cat.tipoProducto.toUpperCase() : row.fila.tipoProducto || '',
      });
    } else {
      stageRowPatch(row, { maquina: machineName });
    }
  };

  const handleBulkReassignMachine = () => {
    if (!maquinaDestinoId) {
      showToast('error', 'Selecciona una máquina destino primero');
      return;
    }
    if (selectedStats.count === 0) {
      showToast('error', 'Selecciona al menos un registro');
      return;
    }
    const target = maquinas.find((m) => m.id === maquinaDestinoId || m.nombre === maquinaDestinoId);
    if (!target) return;

    setDraftEdits((prev) => {
      const next = { ...prev };
      for (const r of selectedStats.rows) {
        next[r.rowKey] = {
          ...(next[r.rowKey] || {}),
          maquina: target.nombre,
          codigoMaquina: target.codigo || '',
          area: target.area || '',
          tipoProducto: target.tipoProducto ? target.tipoProducto.toUpperCase() : '',
        };
      }
      return next;
    });
    showToast(
      'info',
      `Se preparó el cambio a "${target.nombre}" en ${selectedStats.count} registro(s). Pulsa "Guardar Cambios" para confirmar en la nube.`
    );
  };

  const handleBulkApplyOp = () => {
    const opVal = opLote.trim();
    if (!opVal) {
      showToast('error', 'Escribe una OP para aplicar en lote');
      return;
    }
    if (selectedStats.count === 0) {
      showToast('error', 'Selecciona al menos un registro');
      return;
    }
    setDraftEdits((prev) => {
      const next = { ...prev };
      for (const r of selectedStats.rows) {
        next[r.rowKey] = {
          ...(next[r.rowKey] || {}),
          op: opVal,
        };
      }
      return next;
    });
    showToast('info', `OP "${opVal}" aplicada en borrador a ${selectedStats.count} registro(s).`);
  };

  const handleBulkApplyObs = () => {
    if (selectedStats.count === 0) {
      showToast('error', 'Selecciona al menos un registro');
      return;
    }
    setDraftEdits((prev) => {
      const next = { ...prev };
      for (const r of selectedStats.rows) {
        next[r.rowKey] = {
          ...(next[r.rowKey] || {}),
          observaciones: obsLote.trim(),
        };
      }
      return next;
    });
    showToast('info', `Observación aplicada en borrador a ${selectedStats.count} registro(s).`);
  };

  const handleDescartarBorrador = () => {
    setDraftEdits({});
    showToast('info', 'Cambios sin guardar descartados');
  };

  const handleGuardarCambiosNube = async () => {
    if (dirtyKeys.length === 0) return;
    const cfg = getSupabaseConfig(false);
    if (!cfg) return;

    setGuardandoCambios(true);
    try {
      // Group dirty edits by shiftKey (fecha::tipo)
      const byShift = new Map<string, { turno: TurnoDBRow; patches: Map<number, Partial<FilaTurnoExcel>> }>();
      const turnoByKey = new Map<string, TurnoDBRow>();
      for (const t of turnos) {
        turnoByKey.set(`${t.fecha}::${t.tipo}`, t);
      }

      for (const rowKey of dirtyKeys) {
        const [fecha, tipo, idxStr] = rowKey.split('::');
        const shiftKey = `${fecha}::${tipo}`;
        const rowIndex = Number(idxStr);
        const turnoObj = turnoByKey.get(shiftKey);
        if (!turnoObj) continue;
        if (!byShift.has(shiftKey)) {
          byShift.set(shiftKey, { turno: turnoObj, patches: new Map() });
        }
        byShift.get(shiftKey)!.patches.set(rowIndex, draftEdits[rowKey]);
      }

      let turnosActualizados = 0;
      for (const { turno, patches } of byShift.values()) {
        const nuevasFilas = (turno.filas || []).map((f, idx) => {
          const patch = patches.get(idx);
          return patch ? { ...f, ...patch } : f;
        });

        await sbFetch(
          cfg,
          `turnos?fecha=eq.${encodeURIComponent(turno.fecha)}&tipo=eq.${encodeURIComponent(turno.tipo)}`,
          {
            method: 'PATCH',
            json: {
              filas: nuevasFilas,
              updated_at: new Date().toISOString(),
            },
            prefer: 'return=minimal',
          },
          session
        );
        turnosActualizados += 1;
      }

      showToast(
        'ok',
        `Se guardaron ${dirtyKeys.length} registro(s) en ${turnosActualizados} turno(s) de Supabase.`
      );
      await cargarTurnosAuditoria();
      onDataChanged();
    } catch (err: any) {
      showToast('error', 'Error al guardar cambios masivos: ' + (err.message || ''));
    } finally {
      setGuardandoCambios(false);
    }
  };

  const abrirModalEliminarSeleccionados = () => {
    if (selectedStats.count === 0) return;
    const opsSet = new Set<string>();
    const maqsSet = new Set<string>();
    for (const r of selectedStats.rows) {
      const f = getEffectiveFila(r);
      if (f.op) opsSet.add(f.op);
      if (f.maquina) maqsSet.add(f.maquina);
    }
    setEliminarTurnoSiQuedaVacio(false);
    setModalEliminar({
      rows: selectedStats.rows,
      turnosAfectados: selectedStats.shiftsCount,
      totalUnidades: selectedStats.totalUnits,
      opsResumen: Array.from(opsSet).slice(0, 12),
      maquinasResumen: Array.from(maqsSet).slice(0, 8),
    });
  };

  const ejecutarEliminarMasivo = async () => {
    if (!modalEliminar) return;
    const cfg = getSupabaseConfig(false);
    if (!cfg) {
      showToast('error', 'Falta configuración de Supabase');
      return;
    }

    setEliminandoMasivo(true);
    try {
      const turnoByKey = new Map<string, TurnoDBRow>();
      for (const t of turnos) {
        turnoByKey.set(`${t.fecha}::${t.tipo}`, t);
      }

      // Group row indices to delete per shift
      const indicesToDeleteByShift = new Map<string, Set<number>>();
      for (const r of modalEliminar.rows) {
        if (!indicesToDeleteByShift.has(r.shiftKey)) {
          indicesToDeleteByShift.set(r.shiftKey, new Set());
        }
        indicesToDeleteByShift.get(r.shiftKey)!.add(r.rowIndex);
      }

      let turnosModificados = 0;
      let turnosEliminadosPorVacio = 0;

      for (const [shiftKey, idxSet] of indicesToDeleteByShift.entries()) {
        const turno = turnoByKey.get(shiftKey);
        if (!turno) continue;
        // Also preserve any unsaved non-deleted edits on this same shift if any existed
        const nuevasFilas: FilaTurnoExcel[] = [];
        const origFilas = turno.filas || [];
        for (let i = 0; i < origFilas.length; i++) {
          if (idxSet.has(i)) continue;
          const rk = `${shiftKey}::${i}`;
          const patch = draftEdits[rk];
          nuevasFilas.push(patch ? { ...origFilas[i], ...patch } : origFilas[i]);
        }

        if (nuevasFilas.length === 0 && eliminarTurnoSiQuedaVacio) {
          await sbFetch(
            cfg,
            `turnos?fecha=eq.${encodeURIComponent(turno.fecha)}&tipo=eq.${encodeURIComponent(turno.tipo)}`,
            { method: 'DELETE', prefer: 'return=representation' },
            session
          );
          turnosEliminadosPorVacio += 1;
        } else {
          await sbFetch(
            cfg,
            `turnos?fecha=eq.${encodeURIComponent(turno.fecha)}&tipo=eq.${encodeURIComponent(turno.tipo)}`,
            {
              method: 'PATCH',
              json: {
                filas: nuevasFilas,
                updated_at: new Date().toISOString(),
              },
              prefer: 'return=minimal',
            },
            session
          );
          turnosModificados += 1;
        }
      }

      const msgParts = [`${modalEliminar.rows.length} registro(s) eliminados`];
      if (turnosModificados > 0) msgParts.push(`${turnosModificados} turno(s) actualizados`);
      if (turnosEliminadosPorVacio > 0) msgParts.push(`${turnosEliminadosPorVacio} turno(s) vacíos borrados`);

      showToast('ok', msgParts.join(' · '));
      setModalEliminar(null);
      await cargarTurnosAuditoria();
      onDataChanged();
    } catch (err: any) {
      showToast('error', 'Error en eliminación masiva: ' + (err.message || ''));
    } finally {
      setEliminandoMasivo(false);
    }
  };

  return (
    <div className="space-y-4">
      {/* Header Banner */}
      <div className="bg-[#1A202C] border border-[#F2A93B]/50 rounded-2xl p-4 sm:p-5 shadow-md space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-[#2D3748] pb-3">
          <div className="flex items-center gap-2.5">
            <button
              type="button"
              onClick={onClose}
              className="flex items-center gap-1.5 text-xs text-white bg-[#232A38] border border-[#4A5568] hover:border-[#F2A93B] py-2 px-3 rounded-xl cursor-pointer font-semibold min-h-[38px]"
            >
              <ArrowLeft size={14} />
              <span>Volver</span>
            </button>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="m-0 text-sm sm:text-base font-bold text-white font-['Oswald',sans-serif] tracking-wide uppercase">
                  Auditoría y Gestión Masiva de Registros
                </h2>
                <span className="text-[10px] font-bold bg-[#F2A93B]/20 text-[#F6AD55] border border-[#F2A93B]/40 px-2 py-0.5 rounded-md font-['IBM_Plex_Mono',monospace] flex items-center gap-1">
                  <ShieldCheck size={12} /> ADMIN
                </span>
              </div>
              <p className="text-[11px] text-[#A0AEC0] m-0 mt-0.5">
                Revisa registros entre múltiples turnos por máquina u OP, reasigna máquinas o elimina registros erróneos en lote.
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={cargarTurnosAuditoria}
            disabled={cargando}
            className="bg-[#232A38] text-[#F6AD55] border border-[#F2A93B]/40 hover:bg-[#3A2E1B] py-2 px-3.5 rounded-xl text-xs font-bold cursor-pointer flex items-center gap-1.5 min-h-[38px] disabled:opacity-50"
          >
            <RefreshCw size={14} className={cargando ? 'animate-spin' : ''} />
            <span>{cargando ? 'Cargando...' : 'Recargar Nube'}</span>
          </button>
        </div>

        {/* Cloud Query Bar (Date & Shift) */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
          <div>
            <label className="block text-[10px] font-bold text-[#A0AEC0] uppercase mb-1">Desde Fecha</label>
            <input
              type="date"
              value={desde}
              onChange={(e) => setDesde(e.target.value)}
              className="w-full py-2 px-2.5 rounded-xl border border-[#4A5568] text-xs font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none min-h-[38px]"
            />
          </div>
          <div>
            <label className="block text-[10px] font-bold text-[#A0AEC0] uppercase mb-1">Hasta Fecha</label>
            <input
              type="date"
              value={hasta}
              onChange={(e) => setHasta(e.target.value)}
              className="w-full py-2 px-2.5 rounded-xl border border-[#4A5568] text-xs font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none min-h-[38px]"
            />
          </div>
          <div>
            <label className="block text-[10px] font-bold text-[#A0AEC0] uppercase mb-1">Turno</label>
            <select
              value={filtroTurno}
              onChange={(e) => setFiltroTurno(e.target.value)}
              className="w-full py-2 px-2.5 rounded-xl border border-[#4A5568] text-xs bg-[#232A38] text-white outline-none min-h-[38px]"
            >
              <option value="">Todos los turnos</option>
              <option value="dia">Turno Día</option>
              <option value="noche">Turno Noche</option>
            </select>
          </div>
          <div>
            <label className="block text-[10px] font-bold text-[#A0AEC0] uppercase mb-1">Máx. Turnos</label>
            <select
              value={limite}
              onChange={(e) => setLimite(Number(e.target.value))}
              className="w-full py-2 px-2.5 rounded-xl border border-[#4A5568] text-xs font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none min-h-[38px]"
            >
              <option value={100}>Últimos 100 turnos</option>
              <option value={300}>Últimos 300 turnos</option>
              <option value={500}>Últimos 500 turnos</option>
            </select>
          </div>
        </div>

        {/* In-Memory Audit Filters (Machine / OPs / Product) */}
        <div className="bg-[#12161F] border border-[#2D3748] rounded-xl p-3 space-y-3">
          <div className="flex items-center justify-between flex-wrap gap-2">
            <div className="flex items-center gap-1.5 text-xs font-bold text-[#F6AD55] uppercase font-['Oswald',sans-serif] tracking-wide">
              <Filter size={14} />
              <span>Filtros Rápidos de Auditoría ({turnos.length} turnos · {allRows.length} registros cargados)</span>
            </div>
            {(filtroMaquina || filtroOps || filtroTexto || soloConProduccion) && (
              <button
                type="button"
                onClick={() => {
                  setFiltroMaquina('');
                  setFiltroOps('');
                  setFiltroTexto('');
                  setSoloConProduccion(false);
                }}
                className="text-[11px] text-[#CBD5E0] hover:text-white underline bg-transparent border-none cursor-pointer"
              >
                Limpiar filtros de vista
              </button>
            )}
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
            <div>
              <label className="block text-[10px] font-bold text-[#CBD5E0] uppercase mb-1">
                Filtrar por Máquina
              </label>
              <select
                value={filtroMaquina}
                onChange={(e) => setFiltroMaquina(e.target.value)}
                className="w-full py-2 px-2.5 rounded-xl border border-[#4A5568] text-xs bg-[#232A38] text-white outline-none focus:border-[#F2A93B] min-h-[40px]"
              >
                <option value="">Todas las máquinas ({allRows.length} reg)</option>
                {maquinasConConteo.map((m) => (
                  <option key={m.nombre} value={m.nombre}>
                    {m.nombre} ({m.count} reg · {m.ops.size} OPs)
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-[10px] font-bold text-[#CBD5E0] uppercase mb-1">
                Filtrar OP(s) (separa varias con coma o espacio)
              </label>
              <div className="relative">
                <Search size={13} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-[#A0AEC0]" />
                <input
                  type="text"
                  value={filtroOps}
                  onChange={(e) => setFiltroOps(e.target.value)}
                  placeholder="Ej: 4510, 4512, 4519"
                  className="w-full py-2 pl-8 pr-2.5 rounded-xl border border-[#4A5568] text-xs font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none focus:border-[#F2A93B] min-h-[40px]"
                />
              </div>
            </div>

            <div>
              <label className="block text-[10px] font-bold text-[#CBD5E0] uppercase mb-1">
                Producto / SAP / Obs. / Usuario
              </label>
              <input
                type="text"
                value={filtroTexto}
                onChange={(e) => setFiltroTexto(e.target.value)}
                placeholder="Buscar texto en producto, SAP..."
                className="w-full py-2 px-2.5 rounded-xl border border-[#4A5568] text-xs bg-[#232A38] text-white outline-none focus:border-[#F2A93B] min-h-[40px]"
              />
            </div>
          </div>

          <div className="flex items-center justify-between flex-wrap gap-2">
            <label className="inline-flex items-center gap-2 text-xs text-[#CBD5E0] cursor-pointer select-none">
              <input
                type="checkbox"
                checked={soloConProduccion}
                onChange={(e) => setSoloConProduccion(e.target.checked)}
                className="accent-[#F2A93B] w-4 h-4 cursor-pointer"
              />
              <span>Ocultar filas con total en 0 (paradas/sin producción)</span>
            </label>
            <span className="text-xs font-['IBM_Plex_Mono',monospace] text-[#A0AEC0]">
              Mostrando <strong className="text-white">{visibleRows.length}</strong> registro(s)
            </span>
          </div>

          {/* Distinct OPs found under current Machine/Date filter */}
          {opsSummary.length > 0 && (
            <div className="pt-2 border-t border-[#2D3748] space-y-1.5">
              <div className="text-[11px] text-[#A0AEC0]">
                OPs detectadas {filtroMaquina ? `en "${filtroMaquina}"` : 'en la vista actual'} —{' '}
                <span className="text-[#F6AD55]">toca una OP para seleccionar/deseleccionar todos sus registros</span>:
              </div>
              <div className="flex flex-wrap gap-1.5 max-h-28 overflow-y-auto pr-1">
                {opsSummary.map((item) => {
                  const allSelected =
                    item.rowKeys.length > 0 && item.rowKeys.every((k) => selectedKeys.has(k));
                  const partialSelected =
                    !allSelected && item.rowKeys.some((k) => selectedKeys.has(k));
                  return (
                    <button
                      key={item.op}
                      type="button"
                      onClick={() => toggleSeleccionPorOp(item.rowKeys)}
                      title={item.productoSample || item.op}
                      className={`text-[11px] px-2.5 py-1 rounded-lg border font-['IBM_Plex_Mono',monospace] cursor-pointer transition-colors flex items-center gap-1.5 ${
                        allSelected
                          ? 'bg-[#E53E3E]/25 text-white border-[#E53E3E]'
                          : partialSelected
                          ? 'bg-[#F2A93B]/20 text-[#F6AD55] border-[#F2A93B]'
                          : 'bg-[#1A202C] text-[#CBD5E0] border-[#4A5568] hover:border-[#F2A93B]'
                      }`}
                    >
                      <span className="font-bold">OP {item.op}</span>
                      <span className="text-[10px] opacity-80">
                        ({item.count} reg · {item.totalUnidades.toLocaleString('es-PE')} u)
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Bulk Action & Pending Changes Bar */}
      {(selectedStats.count > 0 || dirtyKeys.length > 0) && (
        <div className="bg-[#182335] border-2 border-[#F2A93B]/70 rounded-2xl p-3.5 sm:p-4 shadow-lg space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex flex-wrap items-center gap-2 text-xs">
              {selectedStats.count > 0 && (
                <span className="bg-[#F2A93B]/20 text-[#F6AD55] border border-[#F2A93B]/40 px-2.5 py-1 rounded-lg font-bold font-['IBM_Plex_Mono',monospace]">
                  {selectedStats.count} seleccionado(s) · {selectedStats.shiftsCount} turno(s) ·{' '}
                  {selectedStats.totalUnits.toLocaleString('es-PE')} u
                </span>
              )}
              {dirtyKeys.length > 0 && (
                <span className="bg-[#38A169]/20 text-[#48BB78] border border-[#38A169]/40 px-2.5 py-1 rounded-lg font-bold font-['IBM_Plex_Mono',monospace]">
                  {dirtyKeys.length} fila(s) modificadas en borrador
                </span>
              )}
            </div>

            <div className="flex flex-wrap items-center gap-2">
              {dirtyKeys.length > 0 && (
                <>
                  <button
                    type="button"
                    disabled={guardandoCambios}
                    onClick={handleDescartarBorrador}
                    className="bg-[#232A38] text-[#CBD5E0] hover:text-white border border-[#4A5568] py-2 px-3 rounded-xl text-xs font-semibold cursor-pointer flex items-center gap-1.5"
                  >
                    <RotateCcw size={13} />
                    <span>Descartar</span>
                  </button>
                  <button
                    type="button"
                    disabled={guardandoCambios}
                    onClick={handleGuardarCambiosNube}
                    className="bg-[#38A169] hover:bg-[#2f855a] text-white py-2 px-4 rounded-xl text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer flex items-center gap-1.5 shadow-md disabled:opacity-50"
                  >
                    <Save size={14} />
                    <span>{guardandoCambios ? 'Guardando...' : `Guardar Cambios (${dirtyKeys.length})`}</span>
                  </button>
                </>
              )}

              {selectedStats.count > 0 && (
                <button
                  type="button"
                  onClick={abrirModalEliminarSeleccionados}
                  className="bg-[#E53E3E] hover:bg-[#c53030] text-white py-2 px-4 rounded-xl text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer flex items-center gap-1.5 shadow-md"
                >
                  <Trash2 size={14} />
                  <span>Eliminar Seleccionados ({selectedStats.count})</span>
                </button>
              )}
            </div>
          </div>

          {/* Batch Edit Controls for selected rows */}
          {selectedStats.count > 0 && (
            <div className="grid grid-cols-1 md:grid-cols-3 gap-2.5 pt-2 border-t border-[#2D3748]">
              <div className="flex items-center gap-1.5">
                <select
                  value={maquinaDestinoId}
                  onChange={(e) => setMaquinaDestinoId(e.target.value)}
                  className="flex-1 py-2 px-2.5 rounded-xl border border-[#4A5568] text-xs bg-[#12161F] text-white outline-none min-h-[38px]"
                >
                  <option value="">Mover a máquina...</option>
                  {maquinas.map((m) => (
                    <option key={m.id} value={m.id}>
                      {m.nombre} ({m.tipoProducto ? m.tipoProducto.toUpperCase() : 'S/T'})
                    </option>
                  ))}
                </select>
                <button
                  type="button"
                  onClick={handleBulkReassignMachine}
                  className="bg-[#232A38] text-[#F6AD55] border border-[#F2A93B]/40 hover:bg-[#3A2E1B] py-2 px-3 rounded-xl text-xs font-bold cursor-pointer flex items-center gap-1 min-h-[38px] shrink-0"
                >
                  <ArrowRightLeft size={13} />
                  <span>Mover</span>
                </button>
              </div>

              <div className="flex items-center gap-1.5">
                <input
                  type="text"
                  value={opLote}
                  onChange={(e) => setOpLote(e.target.value)}
                  placeholder="Nueva OP en lote..."
                  className="flex-1 py-2 px-2.5 rounded-xl border border-[#4A5568] text-xs font-['IBM_Plex_Mono',monospace] bg-[#12161F] text-white outline-none min-h-[38px]"
                />
                <button
                  type="button"
                  onClick={handleBulkApplyOp}
                  className="bg-[#232A38] text-white border border-[#4A5568] hover:border-[#F2A93B] py-2 px-3 rounded-xl text-xs font-bold cursor-pointer min-h-[38px] shrink-0"
                >
                  Aplicar OP
                </button>
              </div>

              <div className="flex items-center gap-1.5">
                <input
                  type="text"
                  value={obsLote}
                  onChange={(e) => setObsLote(e.target.value)}
                  placeholder="Nueva observación en lote..."
                  className="flex-1 py-2 px-2.5 rounded-xl border border-[#4A5568] text-xs bg-[#12161F] text-white outline-none min-h-[38px]"
                />
                <button
                  type="button"
                  onClick={handleBulkApplyObs}
                  className="bg-[#232A38] text-white border border-[#4A5568] hover:border-[#F2A93B] py-2 px-3 rounded-xl text-xs font-bold cursor-pointer min-h-[38px] shrink-0"
                >
                  Aplicar Obs.
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Selection Toolbar + Table */}
      <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl overflow-hidden shadow-sm">
        <div className="p-3 bg-[#141B26] border-b border-[#2D3748] flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={
                selectedVisibleCount === visibleRows.length && visibleRows.length > 0
                  ? deseleccionarTodos
                  : seleccionarTodosVisibles
              }
              disabled={visibleRows.length === 0}
              className="bg-[#232A38] text-white border border-[#4A5568] hover:border-[#F2A93B] py-1.5 px-3 rounded-lg text-xs font-semibold cursor-pointer flex items-center gap-1.5 disabled:opacity-40"
            >
              {selectedVisibleCount === visibleRows.length && visibleRows.length > 0 ? (
                <>
                  <CheckSquare size={14} className="text-[#F2A93B]" />
                  <span>Deseleccionar todos ({visibleRows.length})</span>
                </>
              ) : (
                <>
                  <Square size={14} />
                  <span>Seleccionar todos visibles ({visibleRows.length})</span>
                </>
              )}
            </button>

            {selectedStats.count > 0 && (
              <button
                type="button"
                onClick={deseleccionarTodos}
                className="text-xs text-[#A0AEC0] hover:text-white bg-transparent border-none cursor-pointer underline"
              >
                Limpiar selección ({selectedStats.count})
              </button>
            )}
          </div>

          <span className="text-[11px] text-[#A0AEC0]">
            Puedes editar OP, máquina, total, merma u observación directamente en cada fila.
          </span>
        </div>

        {cargando ? (
          <div className="p-10 text-center text-sm text-[#CBD5E0]">
            Cargando registros de los turnos en Supabase…
          </div>
        ) : visibleRows.length === 0 ? (
          <div className="p-10 text-center text-sm text-[#A0AEC0]">
            No hay registros que coincidan con los filtros seleccionados.
          </div>
        ) : (
          <div className="overflow-x-auto max-h-[65vh]">
            <table className="w-full text-left border-collapse min-w-[920px]">
              <thead className="bg-[#12161F] sticky top-0 z-10 border-b border-[#2D3748]">
                <tr className="text-[10px] font-bold text-[#A0AEC0] uppercase tracking-wider font-['IBM_Plex_Mono',monospace]">
                  <th className="p-2.5 w-9 text-center">Sel</th>
                  <th className="p-2.5 w-32">Fecha / Turno</th>
                  <th className="p-2.5 w-44">Máquina</th>
                  <th className="p-2.5 w-28">OP</th>
                  <th className="p-2.5">Producto / SAP</th>
                  <th className="p-2.5 w-24 text-right">Total</th>
                  <th className="p-2.5 w-20 text-right">Merma</th>
                  <th className="p-2.5 w-44">Observaciones</th>
                  <th className="p-2.5 w-24">Usuario</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#2D3748] text-xs">
                {visibleRows.map((row) => {
                  const f = getEffectiveFila(row);
                  const isSelected = selectedKeys.has(row.rowKey);
                  const isDirty = Boolean(draftEdits[row.rowKey]);

                  return (
                    <tr
                      key={row.rowKey}
                      className={`transition-colors ${
                        isSelected
                          ? 'bg-[#E53E3E]/10'
                          : isDirty
                          ? 'bg-[#38A169]/10'
                          : 'hover:bg-[#232A38]/50'
                      }`}
                    >
                      <td className="p-2 text-center">
                        <input
                          type="checkbox"
                          checked={isSelected}
                          onChange={() => toggleRowSelection(row.rowKey)}
                          className="w-4 h-4 accent-[#F2A93B] cursor-pointer"
                        />
                      </td>

                      <td className="p-2 whitespace-nowrap">
                        <div className="font-['IBM_Plex_Mono',monospace] font-bold text-white text-[11px]">
                          {fechaCorta(row.fecha)}
                        </div>
                        <span
                          className={`inline-flex items-center gap-1 text-[10px] font-bold px-1.5 py-0.5 rounded mt-0.5 ${
                            row.tipo === 'noche'
                              ? 'bg-[#63B3ED]/20 text-[#63B3ED]'
                              : 'bg-[#F2A93B]/20 text-[#F6AD55]'
                          }`}
                        >
                          {row.tipo === 'noche' ? <Moon size={10} /> : <Sun size={10} />}
                          {row.tipo === 'noche' ? 'Noche' : 'Día'}
                        </span>
                      </td>

                      <td className="p-2">
                        <select
                          value={f.maquina || ''}
                          onChange={(e) => handleStageMachineChange(row, e.target.value)}
                          className="w-full py-1.5 px-2 rounded-lg border border-[#4A5568] bg-[#232A38] text-white text-[11px] outline-none focus:border-[#F2A93B]"
                        >
                          {!maquinas.some((m) => m.nombre === f.maquina) && f.maquina && (
                            <option value={f.maquina}>{f.maquina}</option>
                          )}
                          {maquinas.map((m) => (
                            <option key={m.id} value={m.nombre}>
                              {m.nombre}
                            </option>
                          ))}
                        </select>
                        <div className="text-[10px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace] mt-0.5 truncate">
                          {f.codigoMaquina || '-'} {f.tipoProducto ? `· ${f.tipoProducto}` : ''}
                        </div>
                      </td>

                      <td className="p-2">
                        <input
                          type="text"
                          value={f.op || ''}
                          onChange={(e) => stageRowPatch(row, { op: e.target.value })}
                          className="w-full py-1.5 px-2 rounded-lg border border-[#4A5568] bg-[#232A38] text-[#F6AD55] font-bold font-['IBM_Plex_Mono',monospace] text-[11px] outline-none focus:border-[#F2A93B]"
                        />
                      </td>

                      <td className="p-2 max-w-[240px]">
                        <div className="text-white text-[11px] font-medium truncate" title={f.producto}>
                          {f.producto || '-'}
                        </div>
                        <div className="text-[10px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace]">
                          SAP: {f.sap || '-'}
                        </div>
                      </td>

                      <td className="p-2">
                        <input
                          type="number"
                          min={0}
                          value={f.total ?? 0}
                          onChange={(e) => stageRowPatch(row, { total: Math.max(0, Number(e.target.value) || 0) })}
                          className="w-full py-1.5 px-2 rounded-lg border border-[#4A5568] bg-[#232A38] text-white font-bold font-['IBM_Plex_Mono',monospace] text-[11px] text-right outline-none focus:border-[#F2A93B] tabular-nums"
                        />
                      </td>

                      <td className="p-2">
                        <input
                          type="number"
                          min={0}
                          step="0.01"
                          value={f.merma ?? 0}
                          onChange={(e) => stageRowPatch(row, { merma: Math.max(0, Number(e.target.value) || 0) })}
                          className="w-full py-1.5 px-2 rounded-lg border border-[#4A5568] bg-[#232A38] text-[#FC8181] font-['IBM_Plex_Mono',monospace] text-[11px] text-right outline-none focus:border-[#F2A93B] tabular-nums"
                        />
                      </td>

                      <td className="p-2">
                        <input
                          type="text"
                          value={f.observaciones || ''}
                          onChange={(e) => stageRowPatch(row, { observaciones: e.target.value })}
                          placeholder="—"
                          className="w-full py-1.5 px-2 rounded-lg border border-[#4A5568] bg-[#232A38] text-[#CBD5E0] text-[11px] outline-none focus:border-[#F2A93B]"
                        />
                      </td>

                      <td className="p-2 text-[11px] text-[#A0AEC0] truncate max-w-[100px]" title={row.autor}>
                        {row.autor}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Modal Confirmar Eliminación Masiva de Registros */}
      {modalEliminar && (
        <div className="fixed inset-0 z-50 bg-black/75 flex items-center justify-center p-3 sm:p-4 backdrop-blur-xs">
          <div className="bg-[#1A202C] border border-[#E53E3E]/60 rounded-2xl p-5 sm:p-6 max-w-lg w-full shadow-2xl space-y-4">
            <div className="flex items-center gap-3 text-[#E53E3E]">
              <div className="w-10 h-10 rounded-xl bg-[#E53E3E]/20 flex items-center justify-center shrink-0 border border-[#E53E3E]/30">
                <AlertTriangle size={20} />
              </div>
              <div>
                <h3 className="text-white font-['Oswald',sans-serif] text-base sm:text-lg font-bold tracking-wide uppercase m-0">
                  ¿Eliminar {modalEliminar.rows.length} registro(s) en masa?
                </h3>
                <p className="text-xs text-[#CBD5E0] m-0">
                  Se actualizarán los turnos afectados en Supabase y se conservará respaldo en Control de Versiones.
                </p>
              </div>
            </div>

            <div className="bg-[#12161F] border border-[#2D3748] rounded-xl p-3.5 space-y-2 text-xs">
              <div className="flex justify-between">
                <span className="text-[#A0AEC0]">Registros a eliminar:</span>
                <span className="font-['IBM_Plex_Mono',monospace] font-bold text-[#FC8181]">
                  {modalEliminar.rows.length} fila(s)
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-[#A0AEC0]">Turnos afectados:</span>
                <span className="font-['IBM_Plex_Mono',monospace] font-bold text-white">
                  {modalEliminar.turnosAfectados} turno(s)
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-[#A0AEC0]">Unidades totales a remover:</span>
                <span className="font-['IBM_Plex_Mono',monospace] font-bold text-[#F2A93B]">
                  {modalEliminar.totalUnidades.toLocaleString('es-PE')} u
                </span>
              </div>
              {modalEliminar.maquinasResumen.length > 0 && (
                <div className="pt-1 border-t border-[#2D3748]">
                  <span className="text-[#A0AEC0] block mb-1">Máquinas involucradas:</span>
                  <span className="text-white font-medium">{modalEliminar.maquinasResumen.join(', ')}</span>
                </div>
              )}
              {modalEliminar.opsResumen.length > 0 && (
                <div className="pt-1 border-t border-[#2D3748]">
                  <span className="text-[#A0AEC0] block mb-1">OPs involucradas:</span>
                  <span className="text-[#F6AD55] font-['IBM_Plex_Mono',monospace]">
                    {modalEliminar.opsResumen.join(', ')}
                  </span>
                </div>
              )}
            </div>

            <label className="flex items-start gap-2.5 text-xs text-[#CBD5E0] bg-[#232A38]/60 border border-[#4A5568] rounded-xl p-3 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={eliminarTurnoSiQuedaVacio}
                onChange={(e) => setEliminarTurnoSiQuedaVacio(e.target.checked)}
                className="accent-[#E53E3E] w-4 h-4 mt-0.5 cursor-pointer shrink-0"
              />
              <span>
                Si un turno queda con <strong>0 registros</strong> tras la eliminación, borrar también ese turno completo de la lista.
              </span>
            </label>

            <div className="flex items-center gap-3 pt-1">
              <button
                type="button"
                disabled={eliminandoMasivo}
                onClick={() => setModalEliminar(null)}
                className="flex-1 bg-[#232A38] text-[#CBD5E0] hover:text-white py-2.5 px-4 rounded-xl text-xs sm:text-sm font-semibold border border-[#4A5568] cursor-pointer min-h-[42px]"
              >
                Cancelar
              </button>
              <button
                type="button"
                disabled={eliminandoMasivo}
                onClick={ejecutarEliminarMasivo}
                className="flex-1 bg-[#E53E3E] hover:bg-[#c53030] text-white py-2.5 px-4 rounded-xl text-xs sm:text-sm font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer shadow-md min-h-[42px] flex items-center justify-center gap-2 disabled:opacity-50"
              >
                {eliminandoMasivo ? (
                  <span>Eliminando...</span>
                ) : (
                  <>
                    <Trash2 size={16} />
                    <span>Confirmar Eliminación</span>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
