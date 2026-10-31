import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { useApp } from '../../context/AppContext';
import {
  History,
  Search,
  ArrowLeft,
  Download,
  Printer,
  Trash2,
  FileSpreadsheet,
  Clock,
  Plus,
  Layers,
  Copy,
  Sun,
  Moon,
  User,
  LogOut,
  ChevronDown,
  ChevronUp,
  ShieldCheck,
} from 'lucide-react';
import {
  getSupabaseConfig,
  sbLogin,
  sbSignUp,
  sbFetch,
} from '../../services/supabase';
import {
  exportarTurnoExcel,
  exportarMigracion,
  buscarEnCatalogo,
  resolverModelo,
  escapeHtml,
} from '../../services/excel';
import { TurnoDBRow, FilaTurnoExcel } from '../../types';

export const HistorialDBView: React.FC = () => {
  const { session, setSession, showToast, catalogo } = useApp();

  // Login form
  const [hdbUser, setHdbUser] = useState('');
  const [hdbPass, setHdbPass] = useState('');
  const [hdbLoginError, setHdbLoginError] = useState('');
  const [modoAuthHdb, setModoAuthHdb] = useState<'login' | 'signup'>('login');

  // Filters
  const [desde, setDesde] = useState('');
  const [hasta, setHasta] = useState('');
  const [filtroTurno, setFiltroTurno] = useState('');
  const [cargando, setCargando] = useState(false);
  const [turnos, setTurnos] = useState<TurnoDBRow[]>([]);
  const [turnoExpandidoKey, setTurnoExpandidoKey] = useState<string | null>(null);

  // Editor mode
  const [editorActivo, setEditorActivo] = useState(false);
  const [turnoEditando, setTurnoEditando] = useState<{
    fecha: string;
    tipo: string;
    filas: FilaTurnoExcel[];
    editable: boolean;
    creador: string;
    updated_at?: string;
  } | null>(null);
  const [filasEditadas, setFilasEditadas] = useState<FilaTurnoExcel[]>([]);
  const [editorExpandidos, setEditorExpandidos] = useState<Record<string, boolean>>({});
  const [guardandoEdicion, setGuardandoEdicion] = useState(false);

  // Versions mode
  const [modoVersiones, setModoVersiones] = useState(false);
  const [subModoVersiones, setSubModoVersiones] = useState<'resumen' | 'detalle'>('resumen');
  const [versionesTitulo, setVersionesTitulo] = useState('');
  const [versionesTurnoActual, setVersionesTurnoActual] = useState<any>(null);
  const [versionesItems, setVersionesItems] = useState<any[]>([]);
  const [cargandoVersiones, setCargandoVersiones] = useState(false);
  const [versionesResumenStats, setVersionesResumenStats] = useState<{
    total: number;
    turnosCount: number;
    kbTotal: number;
    grupos: any[];
  }>({ total: 0, turnosCount: 0, kbTotal: 0, grupos: [] });

  // Modal states for safe deletions without window.confirm or prompt
  const [turnoAEliminar, setTurnoAEliminar] = useState<TurnoDBRow | null>(null);
  const [eliminandoTurno, setEliminandoTurno] = useState(false);
  const [confirmVersionModal, setConfirmVersionModal] = useState<{
    titulo: string;
    mensaje: string;
    detalle?: string;
    onConfirm: () => Promise<void>;
  } | null>(null);
  const [procesandoConfirmModal, setProcesandoConfirmModal] = useState(false);
  const [modalLimpiarDias, setModalLimpiarDias] = useState(false);
  const [diasLimpiezaSeleccionados, setDiasLimpiezaSeleccionados] = useState<number>(30);
  const [modalCopiar, setModalCopiar] = useState<{
    filas: FilaTurnoExcel[];
    turnoData: any;
    fechaDestino: string;
    tipoDestino: string;
    autorEmail: string;
  } | null>(null);

  // Admin status: strictly the designated administrator (josetejera44@gmail.com or user/email 'admin')
  const esAdmin = Boolean(
    session && (
      session.email?.toLowerCase() === 'josetejera44@gmail.com' ||
      session.usuario?.toLowerCase() === 'admin' ||
      session.email?.toLowerCase() === 'admin@planta.local' ||
      session.email?.toLowerCase().startsWith('admin@')
    )
  );

  // Helper to verify if a shift is within the 24-hour editing window
  const esMenorDe24Horas = (fechaStr: string, updatedAtStr?: string): boolean => {
    try {
      if (updatedAtStr) {
        const time = new Date(updatedAtStr).getTime();
        if (!isNaN(time)) {
          const diffHoras = (Date.now() - time) / (1000 * 60 * 60);
          return diffHoras <= 24;
        }
      }
      const endOfDay = new Date(`${fechaStr}T23:59:59`).getTime();
      const diffHoras = (Date.now() - endOfDay) / (1000 * 60 * 60);
      return diffHoras <= 24;
    } catch {
      return false;
    }
  };

  const fechaCorta = (f: string) => {
    const p = String(f).split('-');
    return p.length === 3 ? `${p[2]}/${p[1]}/${p[0]}` : f;
  };

  const fechaHoraCorta = (iso: string) => {
    return iso
      ? new Date(iso).toLocaleString('es-PE', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit',
        })
      : '—';
  };

  const totalFilas = (fs: FilaTurnoExcel[]) => (fs || []).reduce((a, f) => a + (f.total || 0), 0);
  const kbDe = (o: any) => {
    try {
      return Math.round(JSON.stringify(o).length / 1024);
    } catch {
      return 0;
    }
  };

  const handleLogin = async () => {
    const cfg = getSupabaseConfig(false);
    if (!cfg) return;
    setHdbLoginError('');
    if (!hdbUser || !hdbPass) {
      setHdbLoginError('Ingresa usuario o correo y contraseña');
      return;
    }
    try {
      let s;
      if (modoAuthHdb === 'signup') {
        s = await sbSignUp(cfg, hdbUser, hdbPass);
        showToast('ok', 'Cuenta creada e inicio de sesión exitoso');
      } else {
        s = await sbLogin(cfg, hdbUser, hdbPass);
        showToast('ok', 'Sesión iniciada');
      }
      setSession(s);
      setHdbPass('');
    } catch (err: any) {
      setHdbLoginError(err.message || 'Error de autenticación');
    }
  };

  const buscarHistorial = useCallback(async () => {
    const cfg = getSupabaseConfig(false);
    if (!cfg) return;
    setCargando(true);
    let q = 'turnos?select=fecha,tipo,filas,turno,created_by,created_by_email,updated_at&order=fecha.desc,tipo.asc&limit=100';
    if (desde) q += '&fecha=gte.' + desde;
    if (hasta) q += '&fecha=lte.' + hasta;
    if (filtroTurno) q += '&tipo=eq.' + filtroTurno;

    try {
      const rows: TurnoDBRow[] = await sbFetch(cfg, q, undefined, session);
      setTurnos(rows || []);
    } catch (err: any) {
      showToast('error', 'Error al leer historial: ' + (err.message || ''));
    } finally {
      setCargando(false);
    }
  }, [desde, hasta, filtroTurno, session, showToast]);

  useEffect(() => {
    if (session) {
      buscarHistorial();
    }
  }, [session, buscarHistorial]);

  // Abrir editor
  const abrirEditor = (rec: TurnoDBRow, editable: boolean) => {
    setTurnoEditando({
      fecha: rec.fecha,
      tipo: rec.tipo,
      filas: rec.filas || [],
      editable,
      creador: rec.created_by_email || '',
      updated_at: rec.updated_at,
    });
    setFilasEditadas(JSON.parse(JSON.stringify(rec.filas || [])));
    setEditorExpandidos({});
    setEditorActivo(true);
    setModoVersiones(false);
  };

  const totalEditor = useMemo(() => {
    return filasEditadas.reduce((a, f) => a + (f.total || 0), 0);
  }, [filasEditadas]);

  const handleGuardarCambiosEditor = async () => {
    if (!turnoEditando || !turnoEditando.editable) return;
    if (!session) {
      showToast('error', 'Inicia sesión primero');
      return;
    }
    if (!esAdmin) {
      if (turnoEditando.creador && turnoEditando.creador.toLowerCase() !== session.email?.toLowerCase()) {
        showToast('error', 'Solo el autor o el administrador pueden modificar este turno');
        return;
      }
      if (!esMenorDe24Horas(turnoEditando.fecha, turnoEditando.updated_at)) {
        showToast('error', 'Han pasado más de 24 horas. Solo el administrador puede modificar este turno.');
        return;
      }
    }
    const cfg = getSupabaseConfig(false);
    if (!cfg) return;

    setGuardandoEdicion(true);
    try {
      await sbFetch(
        cfg,
        `turnos?fecha=eq.${turnoEditando.fecha}&tipo=eq.${turnoEditando.tipo}`,
        {
          method: 'PATCH',
          json: { filas: filasEditadas, updated_at: new Date().toISOString() },
          prefer: 'return=minimal',
        },
        session
      );
      showToast('ok', 'Cambios guardados en el historial');
      setEditorActivo(false);
      setTurnoEditando(null);
      buscarHistorial();
    } catch (err: any) {
      showToast('error', 'No se pudo guardar: ' + (err.message || ''));
    } finally {
      setGuardandoEdicion(false);
    }
  };

  const handleEliminarTurno = (rec: TurnoDBRow) => {
    if (!esAdmin) {
      showToast('error', 'Solo el administrador tiene permisos para eliminar turnos');
      return;
    }
    setTurnoAEliminar(rec);
  };

  const ejecutarEliminarTurno = async () => {
    if (!turnoAEliminar) return;
    const rec = turnoAEliminar;
    const cfg = getSupabaseConfig(false);
    if (!cfg) {
      showToast('error', 'Falta configuración de Supabase');
      return;
    }

    setEliminandoTurno(true);
    try {
      await sbFetch(
        cfg,
        `turnos?fecha=eq.${encodeURIComponent(rec.fecha)}&tipo=eq.${encodeURIComponent(rec.tipo)}`,
        { method: 'DELETE', prefer: 'return=representation' },
        session
      );
      showToast('ok', `Turno ${fechaCorta(rec.fecha)} (${rec.tipo === 'noche' ? 'Noche' : 'Día'}) eliminado`);
      setTurnoAEliminar(null);
      await buscarHistorial();
    } catch (err: any) {
      showToast('error', 'No se pudo borrar: ' + (err.message || ''));
    } finally {
      setEliminandoTurno(false);
    }
  };

  // Printable report for a past shift
  const handleImprimirHistorial = (rec: TurnoDBRow) => {
    const filas = rec.filas || [];
    if (!filas.length) {
      showToast('error', 'Este turno no tiene filas');
      return;
    }
    const totalGen = filas.reduce((s, f) => s + (f.total || 0), 0);
    const mermaGen = filas.reduce((s, f) => s + (f.merma || 0), 0);

    const html =
      `<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'><title>Turno ${fechaCorta(
        rec.fecha
      )}</title><style>` +
      `@page{size:A4;margin:12mm;}body{font-family:Arial,Helvetica,sans-serif;color:#111;font-size:11px;}` +
      `h1{font-size:16px;margin:0 0 2px;}h2{font-size:13px;background:#eee;padding:4px 6px;margin:14px 0 4px;border-left:4px solid #333;}` +
      `table{width:100%;border-collapse:collapse;margin-bottom:4px;}th,td{border:1px solid #999;padding:3px 5px;text-align:left;}` +
      `th{background:#f2f2f2;}.num{text-align:right;}.total td{font-weight:700;background:#f7f7f7;}` +
      `.resumen{display:flex;justify-content:space-between;align-items:baseline;border-bottom:2px solid #333;padding-bottom:6px;margin-bottom:6px;}` +
      `</style></head><body>` +
      `<div class='resumen'><div><h1>Producción por Máquina</h1><div>${fechaCorta(rec.fecha)} · ${
        rec.tipo === 'noche' ? 'Turno Noche' : 'Turno Día'
      }</div></div>` +
      `<div><strong>${totalGen} u</strong> total${mermaGen ? ' · ' + mermaGen + ' u merma' : ''}</div></div>` +
      `<table><thead><tr><th>Máquina</th><th>OP</th><th>Producto</th><th class='num'>Total</th><th class='num'>Merma</th><th>Observaciones</th></tr></thead><tbody>` +
      filas
        .map(
          (f) =>
            `<tr><td>${escapeHtml(f.maquina || '')}</td><td>${escapeHtml(f.op || '—')}</td><td>${escapeHtml(
              f.producto || '—'
            )}</td><td class='num'>${f.total || 0}</td><td class='num'>${f.merma || 0}</td><td>${escapeHtml(
              f.observaciones || ''
            )}</td></tr>`
        )
        .join('') +
      `<tr class='total'><td colspan='3'>Total</td><td class='num'>${totalGen}</td><td class='num'>${mermaGen}</td><td></td></tr></tbody></table></body></html>`;

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
  };

  // ================= VERSIONS MANAGEMENT (ADMIN ONLY) =================
  const abrirResumenVersiones = async () => {
    if (!esAdmin) return;
    const cfg = getSupabaseConfig(false);
    if (!cfg) return;

    setModoVersiones(true);
    setSubModoVersiones('resumen');
    setEditorActivo(false);
    setVersionesTitulo('Resumen de versiones');
    setCargandoVersiones(true);

    try {
      const vs: any[] = await sbFetch(
        cfg,
        'turnos_versiones?select=id,fecha,tipo,archivado_en,filas&order=archivado_en.desc&limit=500',
        undefined,
        session
      );
      const rows = vs || [];
      const grupos: Record<string, any> = {};
      const orden: string[] = [];
      let kbTotal = 0;

      rows.forEach((v) => {
        const k = `${v.fecha}|${v.tipo}`;
        if (!grupos[k]) {
          grupos[k] = { fecha: v.fecha, tipo: v.tipo, n: 0, kb: 0, ultima: v.archivado_en };
          orden.push(k);
        }
        const kb = kbDe(v.filas);
        grupos[k].n++;
        grupos[k].kb += kb;
        kbTotal += kb;
      });

      setVersionesResumenStats({
        total: rows.length,
        turnosCount: orden.length,
        kbTotal,
        grupos: orden.map((k) => grupos[k]),
      });
    } catch (err: any) {
      showToast('error', 'Error al leer versiones: ' + (err.message || ''));
    } finally {
      setCargandoVersiones(false);
    }
  };

  const abrirVersionesTurno = async (rec: any, origen: 'resumen' | 'detalle') => {
    if (!esAdmin) return;
    const cfg = getSupabaseConfig(false);
    if (!cfg) return;

    setModoVersiones(true);
    setSubModoVersiones('detalle');
    setEditorActivo(false);
    setVersionesTurnoActual(rec);
    setVersionesTitulo(`${fechaCorta(rec.fecha)} · ${rec.tipo === 'noche' ? 'Noche' : 'Día'}`);
    setCargandoVersiones(true);

    try {
      const vs: any[] = await sbFetch(
        cfg,
        `turnos_versiones?select=id,filas,turno,created_by_email,updated_at,archivado_en,accion&fecha=eq.${rec.fecha}&tipo=eq.${rec.tipo}&order=archivado_en.desc&limit=50`,
        undefined,
        session
      );
      const items: any[] = [];
      if (!rec.sinActual) {
        items.push({
          etiqueta: 'ACTUAL',
          filas: rec.filas || [],
          turno: rec.turno,
          autor: rec.created_by_email,
          cuando: rec.updated_at,
          cuandoTxt: 'guardado',
        });
      }
      (vs || []).forEach((v) => {
        items.push({
          id: v.id,
          etiqueta: v.accion === 'DELETE' ? 'BORRADO' : 'VERSIÓN ANTERIOR',
          filas: v.filas || [],
          turno: v.turno,
          autor: v.created_by_email,
          cuando: v.archivado_en,
          cuandoTxt: 'reemplazado',
        });
      });
      setVersionesItems(items);
    } catch (err: any) {
      showToast('error', 'Error al leer versiones del turno: ' + (err.message || ''));
    } finally {
      setCargandoVersiones(false);
    }
  };

  // Copiar turno version to another date/shift
  const iniciarCopiarTurno = (filas: FilaTurnoExcel[], turnoData: any, fechaSug: string, tipoSug: string, autorEmail: string) => {
    if (!esAdmin) return;
    setModalCopiar({
      filas,
      turnoData,
      fechaDestino: fechaSug || new Date().toISOString().split('T')[0],
      tipoDestino: tipoSug || 'dia',
      autorEmail: autorEmail || session?.email || '',
    });
  };

  const ejecutarCopiaConfirmada = async () => {
    if (!modalCopiar || !esAdmin) return;
    const { filas, turnoData, fechaDestino, tipoDestino, autorEmail } = modalCopiar;
    const fTrim = fechaDestino.trim();
    const tTrim = tipoDestino.trim().toLowerCase();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fTrim)) {
      showToast('error', 'Fecha inválida. Debe ser AAAA-MM-DD');
      return;
    }
    const cfg = getSupabaseConfig(false);
    if (!cfg) return;

    try {
      const ex: any[] = await sbFetch(cfg, `turnos?select=filas&fecha=eq.${fTrim}&tipo=eq.${tTrim}`, undefined, session);
      const e = ex && ex[0];
      const msg = e
        ? `El turno ${fechaCorta(fTrim)} (${tTrim === 'noche' ? 'Noche' : 'Día'}) ya existe con ${totalFilas(e.filas).toLocaleString('es-PE')} u. ¿Deseas reemplazarlo con ${totalFilas(
            filas
          ).toLocaleString('es-PE')} u?`
        : `¿Crear ${fechaCorta(fTrim)} (${tTrim === 'noche' ? 'Noche' : 'Día'}) con ${totalFilas(filas).toLocaleString('es-PE')} u?`;

      setModalCopiar(null);
      setConfirmVersionModal({
        titulo: 'Confirmar Copia de Turno',
        mensaje: msg,
        detalle: e ? 'El contenido anterior quedará respaldado en el historial de versiones.' : undefined,
        onConfirm: async () => {
          const ahora = new Date().toISOString();
          if (e) {
            await sbFetch(
              cfg,
              `turnos?fecha=eq.${fTrim}&tipo=eq.${tTrim}`,
              { method: 'PATCH', json: { filas, turno: turnoData || null, updated_at: ahora }, prefer: 'return=minimal' },
              session
            );
          } else {
            await sbFetch(
              cfg,
              'turnos',
              {
                method: 'POST',
                json: [
                  {
                    fecha: fTrim,
                    tipo: tTrim,
                    filas,
                    turno: turnoData || null,
                    created_by: session!.user_id,
                    created_by_email: autorEmail || session!.email,
                    updated_at: ahora,
                  },
                ],
                prefer: 'return=minimal',
              },
              session
            );
          }
          showToast('ok', `Guardado en ${fechaCorta(fTrim)} ${tTrim}`);
          buscarHistorial();
        },
      });
    } catch (err: any) {
      showToast('error', 'No se pudo guardar: ' + (err.message || ''));
    }
  };

  // Delete single version
  const handleEliminarVersion = (id: string, totalU: number) => {
    if (!esAdmin) return;
    setConfirmVersionModal({
      titulo: 'Eliminar Versión Archivada',
      mensaje: `¿Deseas eliminar permanentemente esta versión de archivo (${totalU.toLocaleString('es-PE')} u)?`,
      detalle: 'Esta acción no se puede deshacer.',
      onConfirm: async () => {
        const cfg = getSupabaseConfig(false);
        if (!cfg) return;
        await sbFetch(cfg, 'rpc/borrar_versiones', { method: 'POST', json: { p_id: id }, prefer: 'return=representation' }, session);
        showToast('ok', 'Versión eliminada');
        if (versionesTurnoActual) abrirVersionesTurno(versionesTurnoActual, subModoVersiones);
      },
    });
  };

  // Bulk delete versions of a shift
  const handleEliminarTodasVersionesDeTurno = (fecha: string, tipo: string) => {
    if (!esAdmin) return;
    setConfirmVersionModal({
      titulo: 'Eliminar Versiones Archivadas',
      mensaje: `¿Borrar todas las versiones archivadas de ${fechaCorta(fecha)} ${tipo === 'noche' ? 'Noche' : 'Día'}?`,
      detalle: 'El turno activo en producción NO será modificado. Solo se eliminarán las copias históricas de respaldo.',
      onConfirm: async () => {
        const cfg = getSupabaseConfig(false);
        if (!cfg) return;
        const rows: any[] = await sbFetch(
          cfg,
          `turnos_versiones?select=id&fecha=eq.${encodeURIComponent(fecha)}&tipo=eq.${encodeURIComponent(tipo)}&limit=1000`,
          undefined,
          session
        );
        const ids = (rows || []).map((r) => r.id);
        if (!ids.length) {
          showToast('error', 'No hay versiones que borrar');
          return;
        }
        for (const id of ids) {
          await sbFetch(cfg, 'rpc/borrar_versiones', { method: 'POST', json: { p_id: id }, prefer: 'return=representation' }, session).catch(() => {});
        }
        showToast('ok', `${ids.length} versión(es) eliminada(s)`);
        if (subModoVersiones === 'resumen') abrirResumenVersiones();
        else if (versionesTurnoActual) abrirVersionesTurno(versionesTurnoActual, 'detalle');
      },
    });
  };

  // Bulk delete by N days
  const handleBorrarVersionesAntiguas = () => {
    if (!esAdmin) return;
    setModalLimpiarDias(true);
  };

  const ejecutarBorrarVersionesAntiguas = async (dias: number) => {
    if (!esAdmin) return;
    const corte = new Date(Date.now() - dias * 86400000).toISOString();
    const cfg = getSupabaseConfig(false);
    if (!cfg) return;

    setProcesandoConfirmModal(true);
    try {
      const rows: any[] = await sbFetch(
        cfg,
        `turnos_versiones?select=id&archivado_en=lt.${encodeURIComponent(corte)}&limit=1000`,
        undefined,
        session
      );
      const ids = (rows || []).map((r) => r.id);
      if (!ids.length) {
        showToast('ok', 'No hay versiones archivadas antes de esa fecha');
        setModalLimpiarDias(false);
        return;
      }
      let borrados = 0;
      for (const id of ids) {
        await sbFetch(cfg, 'rpc/borrar_versiones', { method: 'POST', json: { p_id: id }, prefer: 'return=representation' }, session).catch(() => {});
        borrados++;
      }
      showToast('ok', `${borrados} versión(es) antigua(s) eliminada(s)`);
      setModalLimpiarDias(false);
      abrirResumenVersiones();
    } catch (err: any) {
      showToast('error', 'No se pudo borrar: ' + (err.message || ''));
    } finally {
      setProcesandoConfirmModal(false);
    }
  };

  // If not logged in
  if (!session) {
    return (
      <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-6 sm:p-8 max-w-md mx-auto shadow-xl">
        <div className="flex items-center gap-2.5 mb-4">
          <div className="w-8 h-8 rounded-lg bg-[#F2A93B] flex items-center justify-center text-[#12161F]">
            <History size={18} />
          </div>
          <h2 className="text-white text-base sm:text-lg font-bold font-['Oswald',sans-serif] tracking-wide uppercase m-0">
            {modoAuthHdb === 'login' ? 'Acceso a Historial en la Nube' : 'Crear Cuenta en Supabase'}
          </h2>
        </div>

        {/* Toggle Login vs Register */}
        <div className="flex p-1 bg-[#12161F] rounded-xl border border-[#2D3748] mb-4">
          <button
            type="button"
            onClick={() => {
              setModoAuthHdb('login');
              setHdbLoginError('');
            }}
            className={`flex-1 py-1.5 text-xs font-bold font-['Oswald',sans-serif] tracking-wider uppercase rounded-lg transition-colors cursor-pointer ${
              modoAuthHdb === 'login' ? 'bg-[#F2A93B] text-[#12161F]' : 'text-[#A0AEC0] hover:text-white'
            }`}
          >
            Iniciar Sesión
          </button>
          <button
            type="button"
            onClick={() => {
              setModoAuthHdb('signup');
              setHdbLoginError('');
            }}
            className={`flex-1 py-1.5 text-xs font-bold font-['Oswald',sans-serif] tracking-wider uppercase rounded-lg transition-colors cursor-pointer ${
              modoAuthHdb === 'signup' ? 'bg-[#F2A93B] text-[#12161F]' : 'text-[#A0AEC0] hover:text-white'
            }`}
          >
            Crear Cuenta
          </button>
        </div>

        <div className="space-y-3.5 mb-5">
          <div>
            <label className="block text-xs font-semibold text-[#CBD5E0] mb-1.5 uppercase tracking-wide">
              Usuario o Correo
            </label>
            <input
              type="text"
              value={hdbUser}
              onChange={(e) => setHdbUser(e.target.value)}
              placeholder="Ej. supervisor1 o usuario@correo.com"
              className="w-full py-3 px-3.5 rounded-xl border border-[#4A5568] text-base sm:text-sm bg-[#232A38] text-white outline-none focus:border-[#F2A93B] min-h-[46px]"
            />
          </div>
          <div>
            <label className="block text-xs font-semibold text-[#CBD5E0] mb-1.5 uppercase tracking-wide">
              Contraseña
            </label>
            <input
              type="password"
              value={hdbPass}
              onChange={(e) => setHdbPass(e.target.value)}
              placeholder="Tu contraseña"
              className="w-full py-3 px-3.5 rounded-xl border border-[#4A5568] text-base sm:text-sm bg-[#232A38] text-white outline-none focus:border-[#F2A93B] min-h-[46px]"
            />
          </div>
        </div>

        {hdbLoginError && (
          <p className="text-[#E53E3E] text-xs mb-4 bg-[#E53E3E]/10 p-2.5 rounded-lg border border-[#E53E3E]/30 leading-snug">
            {hdbLoginError}
          </p>
        )}

        <button
          type="button"
          onClick={handleLogin}
          className="w-full bg-[#F2A93B] text-[#12161F] py-3.5 px-4 rounded-xl font-bold uppercase font-['Oswald',sans-serif] tracking-wider text-sm cursor-pointer hover:brightness-105 active:scale-95 transition-all shadow-md min-h-[48px]"
        >
          {modoAuthHdb === 'login' ? 'Iniciar Sesión' : 'Registrar Cuenta'}
        </button>

        <p className="text-[11px] text-[#A0AEC0] mt-4 mb-0 text-center">
          Ingresa tu usuario (ej. supervisor1) o correo completo.
        </p>
      </div>
    );
  }

  // Versions Management View
  if (modoVersiones) {
    return (
      <div className="space-y-4 sm:space-y-6">
        <div className="flex items-center justify-between gap-3 bg-[#1A202C] border border-[#2D3748] rounded-2xl p-4 shadow-sm flex-wrap">
          <button
            type="button"
            onClick={() => {
              if (subModoVersiones === 'detalle') {
                abrirResumenVersiones();
              } else {
                setModoVersiones(false);
                buscarHistorial();
              }
            }}
            className="flex items-center gap-1.5 text-xs sm:text-sm font-semibold text-[#CBD5E0] hover:text-white bg-[#232A38] border border-[#4A5568] py-2 px-3.5 rounded-xl cursor-pointer min-h-[40px]"
          >
            <ArrowLeft size={16} /> <span>Volver</span>
          </button>
          <span className="font-['IBM_Plex_Mono',monospace] font-bold text-white text-sm sm:text-base">
            {versionesTitulo}
          </span>
          <div />
        </div>

        {cargandoVersiones ? (
          <div className="border border-dashed border-[#4A5568] rounded-2xl p-8 text-center text-sm text-[#CBD5E0] bg-[#1A202C]">
            Cargando historial de versiones…
          </div>
        ) : subModoVersiones === 'resumen' ? (
          <div className="space-y-4">
            <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-4 sm:p-5 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-sm">
              <div className="text-sm text-white">
                <strong className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace] text-base">{versionesResumenStats.total}</strong> versiones archivadas ·{' '}
                {versionesResumenStats.turnosCount} turnos · ~{versionesResumenStats.kbTotal} KB{' '}
                {versionesResumenStats.total >= 500 && (
                  <span className="text-[#A0AEC0] text-xs block sm:inline mt-1 sm:mt-0">
                    (mostrando 500 más recientes)
                  </span>
                )}
              </div>
              <div>
                <button
                  type="button"
                  onClick={handleBorrarVersionesAntiguas}
                  className="bg-[#E53E3E] text-white py-2 px-3.5 rounded-xl text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:opacity-90 min-h-[38px]"
                >
                  Limpiar antiguas (&gt; N días)…
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
              {versionesResumenStats.grupos.map((g, idx) => (
                <div key={idx} className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-4 flex flex-col justify-between gap-3 shadow-sm">
                  <div>
                    <div className="flex items-center gap-2 mb-1">
                      <span className="font-['IBM_Plex_Mono',monospace] font-bold text-white text-sm">
                        {fechaCorta(g.fecha)}
                      </span>
                      <span
                        className={`text-[10px] font-bold px-2 py-0.5 rounded-md ${
                          g.tipo === 'noche' ? 'bg-[#63B3ED]/20 text-[#63B3ED]' : 'bg-[#F2A93B]/20 text-[#F6AD55]'
                        }`}
                      >
                        {g.tipo === 'noche' ? 'NOCHE' : 'DÍA'}
                      </span>
                    </div>
                    <div className="text-xs text-[#CBD5E0]">
                      {g.n} versión(es) archivada(s) · ~{g.kb} KB
                    </div>
                    <div className="text-[11px] text-[#A0AEC0] mt-0.5">
                      Última modificación: {fechaHoraCorta(g.ultima)}
                    </div>
                  </div>

                  <div className="flex items-center gap-2 pt-2 border-t border-[#2D3748]">
                    <button
                      type="button"
                      onClick={() => {
                        const cfg = getSupabaseConfig(false);
                        if (!cfg) return;
                        sbFetch(
                          cfg,
                          `turnos?select=fecha,tipo,filas,turno,created_by_email,updated_at&fecha=eq.${g.fecha}&tipo=eq.${g.tipo}`,
                          undefined,
                          session
                        ).then((r) => {
                          const rec = (r && r[0]) || { fecha: g.fecha, tipo: g.tipo, sinActual: true };
                          abrirVersionesTurno(rec, 'resumen');
                        });
                      }}
                      className="flex-1 bg-[#F2A93B] text-[#12161F] py-2 px-3 rounded-xl text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer min-h-[38px]"
                    >
                      Ver Versiones
                    </button>
                    <button
                      type="button"
                      onClick={() => handleEliminarTodasVersionesDeTurno(g.fecha, g.tipo)}
                      className="bg-[#232A38] text-[#E53E3E] border border-[#E53E3E]/40 hover:bg-[#E53E3E]/20 py-2 px-3 rounded-xl text-xs font-semibold cursor-pointer min-h-[38px]"
                    >
                      Borrar
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        ) : (
          <div className="space-y-4">
            {versionesItems.length > 1 && (
              <div className="flex justify-end">
                <button
                  type="button"
                  onClick={() => handleEliminarTodasVersionesDeTurno(versionesTurnoActual.fecha, versionesTurnoActual.tipo)}
                  className="bg-[#E53E3E] text-white py-2 px-3.5 rounded-xl text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:opacity-90 min-h-[38px]"
                >
                  Eliminar todas las anteriores ({versionesItems.length - 1})
                </button>
              </div>
            )}

            <div className="space-y-3">
              {versionesItems.map((it, idx) => (
                <div key={idx} className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-sm">
                  <div>
                    <div className="flex items-center gap-2 flex-wrap mb-1">
                      <span className="text-xs font-bold bg-[#3A2E1B] text-[#F6AD55] py-0.5 px-2 rounded-md border border-[#F2A93B]/30 font-['IBM_Plex_Mono',monospace]">
                        {it.etiqueta}
                      </span>
                      <strong className="text-white text-sm font-['IBM_Plex_Mono',monospace]">
                        {totalFilas(it.filas).toLocaleString('es-PE')} u
                      </strong>
                      <span className="text-xs text-[#CBD5E0]">
                        ({it.filas.length} filas · ~{kbDe(it.filas)} KB)
                      </span>
                    </div>
                    <div className="text-xs text-[#A0AEC0]">
                      Por {(it.autor || '?').split('@')[0]} · {it.cuandoTxt} {fechaHoraCorta(it.cuando)}
                    </div>
                  </div>

                  <div className="flex items-center gap-2 flex-wrap">
                    <button
                      type="button"
                      onClick={() => {
                        exportarTurnoExcel(
                          it.filas,
                          `PRODUCCION ${it.etiqueta} ${fechaCorta(versionesTurnoActual.fecha)}`,
                          `version_${versionesTurnoActual.fecha}_${versionesTurnoActual.tipo}`
                        );
                      }}
                      className="bg-[#38A169] text-white py-2 px-3 rounded-xl text-xs font-bold cursor-pointer min-h-[38px]"
                    >
                      Excel
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        iniciarCopiarTurno(it.filas, it.turno, versionesTurnoActual.fecha, versionesTurnoActual.tipo, it.autor)
                      }
                      className="bg-[#F2A93B] text-[#12161F] py-2 px-3 rounded-xl text-xs font-bold cursor-pointer min-h-[38px]"
                    >
                      Copiar a…
                    </button>
                    {it.id && (
                      <button
                        type="button"
                        onClick={() => handleEliminarVersion(it.id, totalFilas(it.filas))}
                        className="p-2 text-[#E53E3E] hover:bg-[#E53E3E]/20 rounded-xl cursor-pointer min-w-[38px] min-h-[38px] flex items-center justify-center"
                        title="Eliminar esta versión"
                      >
                        <Trash2 size={16} />
                      </button>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    );
  }

  // Shift Editor View
  if (editorActivo && turnoEditando) {
    const grupos: Record<string, { nombre: string; codigoMaquina: string; indices: number[] }> = {};
    const ordenGrupos: string[] = [];

    filasEditadas.forEach((f, i) => {
      const key = f.maquina || '(sin máquina)';
      if (!grupos[key]) {
        grupos[key] = { nombre: key, codigoMaquina: f.codigoMaquina || '', indices: [] };
        ordenGrupos.push(key);
      }
      grupos[key].indices.push(i);
    });

    const handleUpdateFila = (index: number, updates: Partial<FilaTurnoExcel>) => {
      setFilasEditadas((prev) => {
        const copia = [...prev];
        copia[index] = { ...copia[index], ...updates };
        return copia;
      });
    };

    const handleEliminarFila = (index: number) => {
      setFilasEditadas((prev) => prev.filter((_, i) => i !== index));
    };

    const handleAgregarOpAMaquina = (nombreMaquina: string, codigo: string) => {
      setFilasEditadas((prev) => [
        ...prev,
        {
          maquina: nombreMaquina,
          codigoMaquina: codigo,
          tipoProducto: '',
          area: '',
          op: '',
          sap: '',
          producto: '',
          total: 0,
          merma: 0,
          observaciones: '',
        },
      ]);
      setEditorExpandidos((prev) => ({ ...prev, [nombreMaquina]: true }));
    };

    const handleAgregarNuevaMaquina = () => {
      const nombre = prompt('Nombre de la máquina para esta nueva fila:', '');
      if (!nombre) return;
      handleAgregarOpAMaquina(nombre.trim(), '');
    };

    return (
      <div className="space-y-4 sm:space-y-6">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-[#1A202C] border border-[#2D3748] rounded-2xl p-4 shadow-sm">
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={() => {
                setEditorActivo(false);
                setTurnoEditando(null);
              }}
              className="flex items-center gap-1.5 text-xs sm:text-sm font-semibold text-[#CBD5E0] hover:text-white bg-[#232A38] border border-[#4A5568] py-2 px-3.5 rounded-xl cursor-pointer min-h-[40px]"
            >
              <ArrowLeft size={16} /> <span>Volver</span>
            </button>
            <div>
              <span className="font-['IBM_Plex_Mono',monospace] font-bold text-white text-sm sm:text-base">
                {fechaCorta(turnoEditando.fecha)} · {turnoEditando.tipo === 'noche' ? 'Noche' : 'Día'}
              </span>
              {!turnoEditando.editable && (
                <span className="text-xs text-[#F6AD55] bg-[#3A2E1B] border border-[#F2A93B]/30 px-2 py-0.5 rounded-md ml-2 inline-block">
                  Solo lectura{turnoEditando.creador && !esAdmin ? ' (Bloqueado tras 24h · solo Admin)' : ''}
                </span>
              )}
            </div>
          </div>

          {turnoEditando.editable && (
            <button
              type="button"
              disabled={guardandoEdicion}
              onClick={handleGuardarCambiosEditor}
              className="bg-[#F2A93B] text-[#12161F] py-2.5 px-5 rounded-xl text-xs sm:text-sm font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:brightness-105 disabled:opacity-50 min-h-[42px] shadow-sm"
            >
              {guardandoEdicion ? 'Guardando...' : 'Guardar Cambios'}
            </button>
          )}
        </div>

        {/* Machine groups accordion */}
        <div className="space-y-3.5">
          {ordenGrupos.map((mKey) => {
            const g = grupos[mKey];
            const abierto = !!editorExpandidos[mKey];
            const sumTot = g.indices.reduce((a, idx) => a + (filasEditadas[idx].total || 0), 0);

            return (
              <div key={mKey} className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-4 shadow-sm">
                <div
                  onClick={() => setEditorExpandidos((prev) => ({ ...prev, [mKey]: !prev[mKey] }))}
                  className="flex items-center justify-between cursor-pointer py-1"
                >
                  <div>
                    <span className="font-['IBM_Plex_Mono',monospace] font-bold text-white text-base">
                      {g.nombre}
                    </span>
                    {g.codigoMaquina && (
                      <span className="text-xs text-[#A0AEC0] font-['IBM_Plex_Mono',monospace] ml-2">
                        ({g.codigoMaquina})
                      </span>
                    )}
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="font-['IBM_Plex_Mono',monospace] text-sm font-bold text-[#F2A93B] tabular-nums">
                      {sumTot.toLocaleString('es-PE')} u · {g.indices.length} OP
                    </span>
                  </div>
                </div>

                {abierto && (
                  <div className="mt-4 pt-4 border-t border-[#2D3748] space-y-3">
                    {g.indices.map((fIdx) => {
                      const f = filasEditadas[fIdx];
                      const matchC = buscarEnCatalogo(catalogo, f.op);

                      return (
                        <div key={fIdx} className="bg-[#232A38] border border-[#4A5568] rounded-2xl p-3.5 sm:p-4 space-y-3">
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                            <div>
                              <label className="block text-xs font-semibold text-[#CBD5E0] mb-1">OP</label>
                              <input
                                type="text"
                                disabled={!turnoEditando.editable}
                                value={f.op || ''}
                                onChange={(e) => {
                                  const op = e.target.value;
                                  const m = buscarEnCatalogo(catalogo, op);
                                  handleUpdateFila(fIdx, {
                                    op,
                                    producto: m ? m.descripcion : f.producto,
                                    sap: m ? m.sap || '' : f.sap,
                                  });
                                }}
                                className="w-full py-2.5 px-3 rounded-xl border border-[#4A5568] text-base sm:text-sm font-['IBM_Plex_Mono',monospace] bg-[#1A202C] text-white min-h-[44px]"
                              />
                            </div>
                            <div>
                              <label className="block text-xs font-semibold text-[#CBD5E0] mb-1">SAP</label>
                              <input
                                type="text"
                                disabled={!turnoEditando.editable}
                                value={f.sap || ''}
                                onChange={(e) => handleUpdateFila(fIdx, { sap: e.target.value })}
                                className="w-full py-2.5 px-3 rounded-xl border border-[#4A5568] text-base sm:text-sm font-['IBM_Plex_Mono',monospace] bg-[#1A202C] text-white min-h-[44px]"
                              />
                            </div>
                          </div>

                          {matchC && matchC.modelos && matchC.modelos.length > 1 && (
                            <div className="flex gap-1.5 flex-wrap">
                              {matchC.modelos.map((mo, i) => (
                                <button
                                  key={i}
                                  type="button"
                                  disabled={!turnoEditando.editable}
                                  onClick={() => handleUpdateFila(fIdx, { producto: mo.descripcion, sap: mo.sap || '' })}
                                  className={`py-1 px-2.5 rounded-lg text-xs font-semibold border min-h-[36px] ${
                                    f.producto === mo.descripcion
                                      ? 'bg-[#F2A93B] text-[#12161F] border-[#F2A93B]'
                                      : 'bg-[#1A202C] text-[#CBD5E0] border-[#4A5568]'
                                  }`}
                                >
                                  {mo.descripcion || `Modelo ${i + 1}`}
                                </button>
                              ))}
                            </div>
                          )}

                          <div>
                            <label className="block text-xs font-semibold text-[#CBD5E0] mb-1">Producto</label>
                            <input
                              type="text"
                              disabled={!turnoEditando.editable}
                              value={f.producto || ''}
                              onChange={(e) => handleUpdateFila(fIdx, { producto: e.target.value })}
                              className="w-full py-2.5 px-3 rounded-xl border border-[#4A5568] text-base sm:text-sm bg-[#1A202C] text-white min-h-[44px]"
                            />
                          </div>

                          <div className="grid grid-cols-2 gap-3">
                            <div>
                              <label className="block text-xs font-semibold text-[#CBD5E0] mb-1">Total Unidades</label>
                              <input
                                type="number"
                                inputMode="numeric"
                                min="0"
                                disabled={!turnoEditando.editable}
                                value={f.total || 0}
                                onChange={(e) => handleUpdateFila(fIdx, { total: parseInt(e.target.value, 10) || 0 })}
                                className="w-full py-2.5 px-3 rounded-xl border border-[#4A5568] text-base sm:text-sm font-['IBM_Plex_Mono',monospace] bg-[#1A202C] text-white min-h-[44px]"
                              />
                            </div>
                            <div>
                              <label className="block text-xs font-semibold text-[#CBD5E0] mb-1">Merma</label>
                              <input
                                type="number"
                                inputMode="numeric"
                                min="0"
                                disabled={!turnoEditando.editable}
                                value={f.merma || 0}
                                onChange={(e) => handleUpdateFila(fIdx, { merma: parseInt(e.target.value, 10) || 0 })}
                                className="w-full py-2.5 px-3 rounded-xl border border-[#4A5568] text-base sm:text-sm font-['IBM_Plex_Mono',monospace] bg-[#1A202C] text-white min-h-[44px]"
                              />
                            </div>
                          </div>

                          <div className="grid grid-cols-2 gap-3">
                            <div>
                              <label className="block text-xs font-semibold text-[#CBD5E0] mb-1">Tipo</label>
                              <select
                                disabled={!turnoEditando.editable}
                                value={f.tipoProducto || ''}
                                onChange={(e) => handleUpdateFila(fIdx, { tipoProducto: e.target.value })}
                                className="w-full py-2.5 px-3 rounded-xl border border-[#4A5568] text-base sm:text-sm bg-[#1A202C] text-white min-h-[44px]"
                              >
                                <option value="">Sin tipo</option>
                                <option value="pp">PP (Proceso)</option>
                                <option value="pt">PT (Terminado)</option>
                              </select>
                            </div>
                            <div>
                              <label className="block text-xs font-semibold text-[#CBD5E0] mb-1">Área</label>
                              <select
                                disabled={!turnoEditando.editable}
                                value={f.area || ''}
                                onChange={(e) => handleUpdateFila(fIdx, { area: e.target.value })}
                                className="w-full py-2.5 px-3 rounded-xl border border-[#4A5568] text-base sm:text-sm bg-[#1A202C] text-white min-h-[44px]"
                              >
                                <option value="">Sin área</option>
                                <option value="vasos_envases">Vasos y Envases</option>
                                <option value="papel_carton">Papel y Cartón</option>
                              </select>
                            </div>
                          </div>

                          <div>
                            <label className="block text-xs font-semibold text-[#CBD5E0] mb-1">Observaciones</label>
                            <input
                              type="text"
                              disabled={!turnoEditando.editable}
                              value={f.observaciones || ''}
                              onChange={(e) => handleUpdateFila(fIdx, { observaciones: e.target.value })}
                              className="w-full py-2.5 px-3 rounded-xl border border-[#4A5568] text-base sm:text-sm bg-[#1A202C] text-white min-h-[44px]"
                            />
                          </div>

                          {turnoEditando.editable && (
                            <div className="flex justify-end pt-1">
                              <button
                                type="button"
                                onClick={() => handleEliminarFila(fIdx)}
                                className="text-[#E53E3E] text-xs font-semibold border border-[#E53E3E]/40 hover:bg-[#E53E3E]/15 py-2 px-3 rounded-lg cursor-pointer min-h-[36px]"
                              >
                                Eliminar fila
                              </button>
                            </div>
                          )}
                        </div>
                      );
                    })}

                    {turnoEditando.editable && (
                      <button
                        type="button"
                        onClick={() => handleAgregarOpAMaquina(g.nombre, g.codigoMaquina)}
                        className="w-full py-3 bg-[#232A38] text-white border border-[#4A5568] rounded-xl text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:bg-[#2D3748] min-h-[44px]"
                      >
                        + Agregar OP a {g.nombre}
                      </button>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>

        {turnoEditando.editable && (
          <button
            type="button"
            onClick={handleAgregarNuevaMaquina}
            className="w-full py-3.5 bg-[#232A38] text-white border border-[#4A5568] rounded-xl text-xs sm:text-sm font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:bg-[#2D3748] min-h-[46px]"
          >
            + Agregar Otra Máquina
          </button>
        )}

        <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-4 flex items-center justify-between">
          <span className="text-xs font-bold text-[#CBD5E0] font-['Oswald',sans-serif] tracking-wider uppercase">
            Total del Turno Editado
          </span>
          <span className="font-['IBM_Plex_Mono',monospace] text-xl font-bold text-[#F2A93B] tabular-nums">
            {totalEditor.toLocaleString('es-PE')} u
          </span>
        </div>
      </div>
    );
  }

  // Normal List View
  return (
    <div className="space-y-4 sm:space-y-6">
      {/* Filter panel */}
      <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3.5">
        <div className="flex items-center justify-between pb-2 border-b border-[#2D3748] flex-wrap gap-2">
          <div className="flex items-center gap-2">
            <History size={18} className="text-[#F2A93B]" />
            <h2 className="m-0 text-sm sm:text-base font-bold text-white font-['Oswald',sans-serif] tracking-wide uppercase">
              Historial de Turnos (Nube)
            </h2>
            {esAdmin && (
              <span className="text-[10px] font-bold bg-[#F2A93B]/20 text-[#F6AD55] border border-[#F2A93B]/40 px-2 py-0.5 rounded-md font-['IBM_Plex_Mono',monospace] flex items-center gap-1">
                <ShieldCheck size={12} /> ADMIN
              </span>
            )}
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => setSession(null)}
              className="flex items-center gap-1 text-xs text-[#CBD5E0] hover:text-[#E53E3E] cursor-pointer bg-transparent border-none p-1"
            >
              <LogOut size={14} />
              <span>Salir ({session.usuario})</span>
            </button>
          </div>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div>
            <label className="block text-xs font-semibold text-[#CBD5E0] mb-1 uppercase tracking-wide">Desde</label>
            <input
              type="date"
              value={desde}
              onChange={(e) => setDesde(e.target.value)}
              className="w-full py-2.5 px-3 rounded-xl border border-[#4A5568] text-sm font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none min-h-[44px]"
            />
          </div>
          <div>
            <label className="block text-xs font-semibold text-[#CBD5E0] mb-1 uppercase tracking-wide">Hasta</label>
            <input
              type="date"
              value={hasta}
              onChange={(e) => setHasta(e.target.value)}
              className="w-full py-2.5 px-3 rounded-xl border border-[#4A5568] text-sm font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none min-h-[44px]"
            />
          </div>
          <div>
            <label className="block text-xs font-semibold text-[#CBD5E0] mb-1 uppercase tracking-wide">Turno</label>
            <select
              value={filtroTurno}
              onChange={(e) => setFiltroTurno(e.target.value)}
              className="w-full py-2.5 px-3 rounded-xl border border-[#4A5568] text-sm bg-[#232A38] text-white outline-none min-h-[44px]"
            >
              <option value="">Todos los turnos</option>
              <option value="dia">Turno Día</option>
              <option value="noche">Turno Noche</option>
            </select>
          </div>
        </div>

        <div className="flex flex-col sm:flex-row gap-2 pt-1">
          <button
            type="button"
            onClick={buscarHistorial}
            onDoubleClick={() => getSupabaseConfig(true)}
            title="Doble clic para cambiar URL/clave de Supabase"
            className="flex-1 bg-[#F2A93B] text-[#12161F] py-2.5 px-4 rounded-xl text-xs sm:text-sm font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:brightness-105 active:scale-95 transition-all shadow-sm min-h-[42px] flex items-center justify-center gap-2"
          >
            <Search size={16} />
            <span>{cargando ? 'Buscando...' : 'Buscar Registros'}</span>
          </button>

          {(desde || hasta || filtroTurno) && (
            <button
              type="button"
              onClick={() => {
                setDesde('');
                setHasta('');
                setFiltroTurno('');
              }}
              className="bg-[#232A38] text-[#CBD5E0] hover:text-white border border-[#4A5568] py-2.5 px-3 rounded-xl text-xs font-semibold cursor-pointer active:scale-95 min-h-[42px] flex items-center justify-center gap-1.5"
            >
              <span>Ver Todos (Quitar Filtros)</span>
            </button>
          )}

          {esAdmin && (
            <button
              type="button"
              onClick={abrirResumenVersiones}
              className="bg-[#232A38] text-white border border-[#4A5568] py-2.5 px-3 rounded-xl text-xs sm:text-sm font-semibold cursor-pointer hover:bg-[#2D3748] active:scale-95 min-h-[42px]"
            >
              Control de Versiones (Admin)
            </button>
          )}
        </div>

        <p className="text-[11px] text-[#A0AEC0] m-0">
          Doble clic o toque en &quot;Buscar Registros&quot; para reconfigurar la URL / clave de Supabase.
        </p>
      </div>

      {/* Shifts List: Responsive on mobile & desktop */}
      <div className="space-y-3">
        {cargando ? (
          <div className="border border-dashed border-[#4A5568] rounded-2xl p-8 text-center text-sm text-[#CBD5E0] bg-[#1A202C]">
            Consultando registros en la base de datos…
          </div>
        ) : turnos.length === 0 ? (
          <div className="border border-dashed border-[#4A5568] rounded-2xl p-8 text-center text-sm text-[#CBD5E0] bg-[#1A202C]">
            No se encontraron turnos guardados con los filtros seleccionados.
          </div>
        ) : (
          turnos.map((rec) => {
            const filas = rec.filas || [];
            const tot = filas.reduce((a, f) => a + (f.total || 0), 0);
            const duenio = Boolean(
              rec.created_by_email &&
              session?.email &&
              rec.created_by_email.toLowerCase() === session.email.toLowerCase()
            );
            const dentroDe24h = esMenorDe24Horas(rec.fecha, rec.updated_at);
            // Admin can always edit. Creators can only edit within the first 24 hours.
            const puedeEditar = esAdmin || (duenio && dentroDe24h);
            // ONLY Admin can delete shifts.
            const puedeBorrar = esAdmin;
            const autor = rec.created_by_email ? rec.created_by_email.split('@')[0] : 'desconocido';
            const itemKey = `${rec.fecha}-${rec.tipo}`;
            const estaExpandido = turnoExpandidoKey === itemKey;

            return (
              <div
                key={itemKey}
                className={`border rounded-xl transition-all shadow-sm overflow-hidden ${
                  estaExpandido
                    ? 'bg-[#182335] border-[#F2A93B]/70 shadow-[#F2A93B]/5 ring-1 ring-[#F2A93B]/20'
                    : 'bg-[#1A202C] border-[#2D3748] hover:border-[#4A5568]'
                }`}
              >
                {/* Header row: Clickable to expand/collapse */}
                <div
                  onClick={() => setTurnoExpandidoKey((prev) => (prev === itemKey ? null : itemKey))}
                  className="p-2.5 sm:p-3 flex items-center justify-between gap-2 cursor-pointer select-none hover:bg-[#232A38]/50 transition-colors"
                >
                  <div className="flex items-center gap-2 flex-wrap min-w-0">
                    <span className="font-['IBM_Plex_Mono',monospace] font-bold text-white text-xs sm:text-sm">
                      {fechaCorta(rec.fecha)}
                    </span>
                    <span
                      className={`text-[11px] font-bold px-2 py-0.5 rounded-md flex items-center gap-1 ${
                        rec.tipo === 'noche'
                          ? 'bg-[#63B3ED]/20 text-[#63B3ED]'
                          : 'bg-[#F2A93B]/20 text-[#F6AD55]'
                      }`}
                    >
                      {rec.tipo === 'noche' ? <Moon size={11} /> : <Sun size={11} />}
                      <span>{rec.tipo === 'noche' ? 'Noche' : 'Día'}</span>
                    </span>
                    <span className="text-[11px] text-[#A0AEC0] truncate">
                      {filas.length} filas · {duenio ? <strong className="text-[#38A169]">por ti</strong> : `por ${autor}`}
                    </span>
                    {duenio && !dentroDe24h && !esAdmin && (
                      <span
                        className="text-[10px] text-[#CBD5E0] bg-[#232A38] border border-[#4A5568] px-1.5 py-0.5 rounded font-['IBM_Plex_Mono',monospace]"
                        title="Edición bloqueada tras 24 horas (solo Admin)"
                      >
                        Bloqueado (&gt;24h)
                      </span>
                    )}
                  </div>

                  <div className="flex items-center gap-1.5 sm:gap-2 shrink-0">
                    <span className="font-['IBM_Plex_Mono',monospace] font-bold text-xs sm:text-sm text-[#F2A93B] tabular-nums">
                      {tot.toLocaleString('es-PE')} u
                    </span>
                    {puedeBorrar && (
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          handleEliminarTurno(rec);
                        }}
                        className="p-1.5 text-[#E53E3E]/75 hover:text-[#E53E3E] hover:bg-[#E53E3E]/15 rounded-lg transition-colors cursor-pointer"
                        title="Eliminar del historial (Solo Admin)"
                      >
                        <Trash2 size={15} />
                      </button>
                    )}
                    <div className={`p-1 rounded text-[#A0AEC0] transition-colors ${estaExpandido ? 'text-[#F2A93B]' : ''}`}>
                      {estaExpandido ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
                    </div>
                  </div>
                </div>

                {/* Expanded action buttons drawer */}
                {estaExpandido && (
                  <div className="px-2.5 sm:px-3 pb-2.5 pt-1.5 border-t border-[#2D3748] bg-[#141B26] flex items-center gap-2 flex-wrap">
                    <button
                      type="button"
                      onClick={() => abrirEditor(rec, puedeEditar)}
                      className="flex-1 sm:flex-initial bg-[#F2A93B] text-[#12161F] py-1.5 px-3 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:brightness-105 active:scale-95"
                    >
                      {puedeEditar ? 'Ver / Editar' : 'Ver Turno'}
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        exportarTurnoExcel(
                          filas,
                          `PRODUCCION TURNO ${rec.tipo === 'noche' ? 'NOCHE' : 'DIA'} ${fechaCorta(rec.fecha)}`,
                          `turno_${rec.fecha}_${rec.tipo}`
                        );
                      }}
                      className="flex-1 sm:flex-initial bg-[#232A38] text-[#38A169] border border-[#38A169]/40 hover:bg-[#1F7A45]/30 py-1.5 px-2.5 rounded-lg text-xs font-semibold cursor-pointer active:scale-95"
                    >
                      Excel
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        try {
                          exportarMigracion(filas, rec.fecha, rec.tipo);
                          showToast('ok', 'Excel de migración exportado');
                        } catch (err: any) {
                          showToast('error', err.message || 'Error');
                        }
                      }}
                      className="flex-1 sm:flex-initial bg-[#232A38] text-[#F6AD55] border border-[#F2A93B]/40 hover:bg-[#3A2E1B] py-1.5 px-2.5 rounded-lg text-xs font-semibold cursor-pointer active:scale-95"
                    >
                      Migración
                    </button>
                    <button
                      type="button"
                      onClick={() => handleImprimirHistorial(rec)}
                      className="flex-1 sm:flex-initial bg-[#232A38] text-[#93C5FD] border border-[#93C5FD]/40 hover:bg-[#232A38]/60 py-1.5 px-2.5 rounded-lg text-xs font-semibold cursor-pointer active:scale-95"
                    >
                      Imprimir
                    </button>
                    {esAdmin && (
                      <button
                        type="button"
                        onClick={() => abrirVersionesTurno(rec, 'detalle')}
                        className="bg-[#232A38] text-[#C4B5FD] border border-[#C4B5FD]/40 hover:bg-[#232A38]/60 py-1.5 px-2.5 rounded-lg text-xs font-semibold cursor-pointer"
                      >
                        Versiones
                      </button>
                    )}
                    {puedeBorrar && (
                      <button
                        type="button"
                        onClick={() => handleEliminarTurno(rec)}
                        className="bg-[#232A38] text-[#E53E3E] border border-[#E53E3E]/40 hover:bg-[#E53E3E]/20 py-1.5 px-2.5 rounded-lg text-xs font-semibold cursor-pointer active:scale-95 flex items-center gap-1.5 ml-auto"
                        title="Eliminar del historial (Solo Admin)"
                      >
                        <Trash2 size={14} />
                        <span>Eliminar</span>
                      </button>
                    )}
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>

      {/* Modal Confirmar Eliminación de Turno */}
      {turnoAEliminar && (
        <div className="fixed inset-0 z-50 bg-black/75 flex items-center justify-center p-3 sm:p-4 backdrop-blur-xs">
          <div className="bg-[#1A202C] border border-[#E53E3E]/50 rounded-2xl p-5 sm:p-6 max-w-md w-full shadow-2xl">
            <div className="flex items-center gap-3 text-[#E53E3E] mb-3">
              <div className="w-10 h-10 rounded-xl bg-[#E53E3E]/20 flex items-center justify-center shrink-0 border border-[#E53E3E]/30">
                <Trash2 size={20} />
              </div>
              <div>
                <h3 className="text-white font-['Oswald',sans-serif] text-base sm:text-lg font-bold tracking-wide uppercase m-0">
                  ¿Eliminar turno del historial?
                </h3>
                <p className="text-xs text-[#CBD5E0] m-0">
                  Esta acción eliminará el registro de Supabase permanentemente.
                </p>
              </div>
            </div>

            <div className="bg-[#12161F] border border-[#2D3748] rounded-xl p-3.5 my-4 space-y-2 text-xs">
              <div className="flex justify-between">
                <span className="text-[#A0AEC0]">Fecha:</span>
                <span className="font-['IBM_Plex_Mono',monospace] font-bold text-white">
                  {fechaCorta(turnoAEliminar.fecha)}
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-[#A0AEC0]">Turno:</span>
                <span className="font-bold text-[#F6AD55]">
                  {turnoAEliminar.tipo === 'noche' ? 'Noche' : 'Día'}
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-[#A0AEC0]">Total piezas:</span>
                <span className="font-['IBM_Plex_Mono',monospace] font-bold text-[#F2A93B]">
                  {totalFilas(turnoAEliminar.filas).toLocaleString('es-PE')} u
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-[#A0AEC0]">Máquinas / filas:</span>
                <span className="text-white">
                  {turnoAEliminar.filas?.length || 0} filas
                </span>
              </div>
              {turnoAEliminar.created_by_email && (
                <div className="flex justify-between">
                  <span className="text-[#A0AEC0]">Creado por:</span>
                  <span className="text-white truncate max-w-[200px]">
                    {turnoAEliminar.created_by_email}
                  </span>
                </div>
              )}
            </div>

            <div className="flex items-center gap-3 pt-1">
              <button
                type="button"
                disabled={eliminandoTurno}
                onClick={() => setTurnoAEliminar(null)}
                className="flex-1 bg-[#232A38] text-[#CBD5E0] hover:text-white py-2.5 px-4 rounded-xl text-xs sm:text-sm font-semibold border border-[#4A5568] cursor-pointer min-h-[42px]"
              >
                Cancelar
              </button>
              <button
                type="button"
                disabled={eliminandoTurno}
                onClick={ejecutarEliminarTurno}
                className="flex-1 bg-[#E53E3E] hover:bg-[#c53030] text-white py-2.5 px-4 rounded-xl text-xs sm:text-sm font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer shadow-md min-h-[42px] flex items-center justify-center gap-2 disabled:opacity-50"
              >
                {eliminandoTurno ? (
                  <span>Eliminando...</span>
                ) : (
                  <>
                    <Trash2 size={16} />
                    <span>Sí, Eliminar</span>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Confirmar Eliminación de Versiones */}
      {confirmVersionModal && (
        <div className="fixed inset-0 z-50 bg-black/75 flex items-center justify-center p-3 sm:p-4 backdrop-blur-xs">
          <div className="bg-[#1A202C] border border-[#E53E3E]/50 rounded-2xl p-5 sm:p-6 max-w-md w-full shadow-2xl">
            <div className="flex items-center gap-3 text-[#E53E3E] mb-3">
              <div className="w-10 h-10 rounded-xl bg-[#E53E3E]/20 flex items-center justify-center shrink-0 border border-[#E53E3E]/30">
                <Trash2 size={20} />
              </div>
              <div>
                <h3 className="text-white font-['Oswald',sans-serif] text-base sm:text-lg font-bold tracking-wide uppercase m-0">
                  {confirmVersionModal.titulo}
                </h3>
              </div>
            </div>

            <p className="text-xs sm:text-sm text-[#CBD5E0] leading-relaxed my-3">
              {confirmVersionModal.mensaje}
            </p>
            {confirmVersionModal.detalle && (
              <p className="text-xs text-[#A0AEC0] bg-[#12161F] p-2.5 rounded-lg border border-[#2D3748]">
                {confirmVersionModal.detalle}
              </p>
            )}

            <div className="flex items-center gap-3 pt-3">
              <button
                type="button"
                disabled={procesandoConfirmModal}
                onClick={() => setConfirmVersionModal(null)}
                className="flex-1 bg-[#232A38] text-[#CBD5E0] hover:text-white py-2.5 px-4 rounded-xl text-xs sm:text-sm font-semibold border border-[#4A5568] cursor-pointer min-h-[42px]"
              >
                Cancelar
              </button>
              <button
                type="button"
                disabled={procesandoConfirmModal}
                onClick={async () => {
                  setProcesandoConfirmModal(true);
                  try {
                    await confirmVersionModal.onConfirm();
                    setConfirmVersionModal(null);
                  } catch (err: any) {
                    showToast('error', err.message || 'Error');
                  } finally {
                    setProcesandoConfirmModal(false);
                  }
                }}
                className="flex-1 bg-[#E53E3E] hover:bg-[#c53030] text-white py-2.5 px-4 rounded-xl text-xs sm:text-sm font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer shadow-md min-h-[42px] flex items-center justify-center gap-2 disabled:opacity-50"
              >
                {procesandoConfirmModal ? <span>Procesando...</span> : <span>Confirmar</span>}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Limpiar Versiones Antiguas */}
      {modalLimpiarDias && (
        <div className="fixed inset-0 z-50 bg-black/75 flex items-center justify-center p-3 sm:p-4 backdrop-blur-xs">
          <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-5 sm:p-6 max-w-md w-full shadow-2xl">
            <h3 className="text-white font-['Oswald',sans-serif] text-base sm:text-lg font-bold tracking-wide uppercase mb-2">
              Limpiar Versiones Archivadas Antiguas
            </h3>
            <p className="text-xs text-[#CBD5E0] mb-4">
              Selecciona el umbral para borrar versiones de respaldo anteriores a esa antigüedad:
            </p>

            <div className="grid grid-cols-4 gap-2 mb-4">
              {[15, 30, 60, 90].map((d) => (
                <button
                  key={d}
                  type="button"
                  onClick={() => setDiasLimpiezaSeleccionados(d)}
                  className={`py-2 px-2 rounded-xl text-xs font-bold font-['IBM_Plex_Mono',monospace] border cursor-pointer transition-colors ${
                    diasLimpiezaSeleccionados === d
                      ? 'bg-[#F2A93B] text-[#12161F] border-[#F2A93B]'
                      : 'bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:text-white'
                  }`}
                >
                  &gt; {d} días
                </button>
              ))}
            </div>

            <div className="flex items-center gap-3 pt-2">
              <button
                type="button"
                disabled={procesandoConfirmModal}
                onClick={() => setModalLimpiarDias(false)}
                className="flex-1 bg-[#232A38] text-[#CBD5E0] hover:text-white py-2.5 px-4 rounded-xl text-xs sm:text-sm font-semibold border border-[#4A5568] cursor-pointer min-h-[42px]"
              >
                Cancelar
              </button>
              <button
                type="button"
                disabled={procesandoConfirmModal}
                onClick={() => ejecutarBorrarVersionesAntiguas(diasLimpiezaSeleccionados)}
                className="flex-1 bg-[#E53E3E] hover:bg-[#c53030] text-white py-2.5 px-4 rounded-xl text-xs sm:text-sm font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer shadow-md min-h-[42px] flex items-center justify-center gap-2 disabled:opacity-50"
              >
                {procesandoConfirmModal ? <span>Limpiando...</span> : <span>Limpiar Antiguas</span>}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Copiar Versión a Fecha/Turno */}
      {modalCopiar && (
        <div className="fixed inset-0 z-50 bg-black/75 flex items-center justify-center p-3 sm:p-4 backdrop-blur-xs">
          <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-5 sm:p-6 max-w-md w-full shadow-2xl space-y-4">
            <h3 className="text-white font-['Oswald',sans-serif] text-base sm:text-lg font-bold tracking-wide uppercase m-0">
              Copiar Versión a Turno
            </h3>
            <p className="text-xs text-[#CBD5E0] m-0">
              Indica la fecha y el turno de destino donde deseas copiar estas {modalCopiar.filas.length} filas ({totalFilas(modalCopiar.filas).toLocaleString('es-PE')} u):
            </p>

            <div className="space-y-3">
              <div>
                <label className="block text-xs font-semibold text-[#CBD5E0] mb-1 uppercase tracking-wide">
                  Fecha Destino
                </label>
                <input
                  type="date"
                  value={modalCopiar.fechaDestino}
                  onChange={(e) => setModalCopiar({ ...modalCopiar, fechaDestino: e.target.value })}
                  className="w-full py-2.5 px-3 rounded-xl border border-[#4A5568] text-sm font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none min-h-[44px]"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-[#CBD5E0] mb-1 uppercase tracking-wide">
                  Turno Destino
                </label>
                <select
                  value={modalCopiar.tipoDestino}
                  onChange={(e) => setModalCopiar({ ...modalCopiar, tipoDestino: e.target.value })}
                  className="w-full py-2.5 px-3 rounded-xl border border-[#4A5568] text-sm bg-[#232A38] text-white outline-none min-h-[44px]"
                >
                  <option value="dia">Turno Día</option>
                  <option value="noche">Turno Noche</option>
                </select>
              </div>
            </div>

            <div className="flex items-center gap-3 pt-2">
              <button
                type="button"
                onClick={() => setModalCopiar(null)}
                className="flex-1 bg-[#232A38] text-[#CBD5E0] hover:text-white py-2.5 px-4 rounded-xl text-xs sm:text-sm font-semibold border border-[#4A5568] cursor-pointer min-h-[42px]"
              >
                Cancelar
              </button>
              <button
                type="button"
                onClick={ejecutarCopiaConfirmada}
                className="flex-1 bg-[#F2A93B] text-[#12161F] py-2.5 px-4 rounded-xl text-xs sm:text-sm font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer shadow-md min-h-[42px]"
              >
                Continuar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
