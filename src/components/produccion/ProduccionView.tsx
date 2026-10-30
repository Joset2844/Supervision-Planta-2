import React, { useState, useMemo } from 'react';
import { useApp } from '../../context/AppContext';
import { Odometer } from '../common/Odometer';
import {
  TrendingUp,
  History,
  RotateCcw,
  Save,
  Download,
  Calendar,
  ChevronDown,
  ChevronUp,
  Plus,
  Trash2,
  Copy,
  Printer,
  FileSpreadsheet,
  Sun,
  Moon,
  Share2,
  SlidersHorizontal,
  X,
} from 'lucide-react';
import {
  buscarEnCatalogo,
  resolverModelo,
  descripcionCorta,
  etiquetaPliego,
  generarFilasTurno,
  exportarTurnoExcel,
  exportarMigracion,
  exportarActivacionExcel,
  imprimirActivacion,
  turnosDeMaquina,
  claveTarjeta,
  nombreTarjeta,
  filasActivacion,
} from '../../services/excel';
import {
  formatearFechaTurno,
  etiquetaTurno,
  generarReporteProduccion,
  generarReporteActivacion,
  copiarAlPortapapeles,
} from '../../services/reports';
import { getSupabaseConfig, sbLogin } from '../../services/supabase';
import { Maquina, TurnoMaquina, RegistroTurno } from '../../types';

const MOTIVOS_PARADA = ['Mantenimiento', 'Falta de personal', 'Desglose', 'Calibración'];

export const ProduccionView: React.FC = () => {
  const {
    maquinas,
    turno,
    setTurno,
    catalogo,
    metaTurno,
    setMetaTurno,
    plantilla,
    setPlantilla,
    historial,
    setHistorial,
    session,
    setSession,
    showToast,
    reiniciarTurno,
    guardarTurnoDB,
  } = useApp();

  const [expandidos, setExpandidos] = useState<Record<string, boolean>>({});
  const [panelLoginOpen, setPanelLoginOpen] = useState(false);
  const [panelHistorialOpen, setPanelHistorialOpen] = useState(false);
  const [panelRespaldoOpen, setPanelRespaldoOpen] = useState(false);
  const [panelRestaurarOpen, setPanelRestaurarOpen] = useState(false);
  const [textoRespaldo, setTextoRespaldo] = useState('');
  const [textoRestaurar, setTextoRestaurar] = useState('');
  const [reportesAbiertosMobile, setReportesAbiertosMobile] = useState(false);

  // Login form
  const [loginEmail, setLoginEmail] = useState('');
  const [loginPass, setLoginPass] = useState('');
  const [loginError, setLoginError] = useState('');

  // Temporary inputs per machine card for adding records
  const [cantInputs, setCantInputs] = useState<Record<string, string>>({});
  const [undInputs, setUndInputs] = useState<Record<string, string>>({});
  const [cajasInputs, setCajasInputs] = useState<Record<string, string>>({});
  const [mermaInputs, setMermaInputs] = useState<Record<string, string>>({});
  const [modoCantidadMap, setModoCantidadMap] = useState<Record<string, 'total' | 'caja'>>({});
  const [opSuggestMap, setOpSuggestMap] = useState<Record<string, string>>({});
  const [showSuggestMap, setShowSuggestMap] = useState<Record<string, boolean>>({});

  const activas = useMemo(() => {
    const getR = (area: string) => (area === 'vasos_envases' ? 0 : area === 'papel_carton' ? 1 : 2);
    return maquinas
      .filter((m) => m.activa)
      .sort((a, b) => {
        const ra = getR(a.area);
        const rb = getR(b.area);
        if (ra !== rb) return ra - rb;
        return a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' });
      });
  }, [maquinas]);

  const totalProducidoGeneral = useMemo(() => {
    return turno.reduce((acc, t) => {
      const sumaReg = (t.registros || []).reduce((rAcc, r) => rAcc + (r.cantidad || 0), 0);
      return acc + sumaReg;
    }, 0);
  }, [turno]);

  // Turno helper for each machine & instance
  const getTurnoMaquina = (maquinaId: string, inst = 1): TurnoMaquina => {
    let t = turno.find((x) => x.maquinaId === maquinaId && (x.instancia || 1) === inst);
    if (!t) {
      t = {
        maquinaId,
        instancia: inst,
        op: '',
        objetivo: null,
        motivoParada: '',
        observaciones: '',
        personal: '',
        modeloIdx: -1,
        pliego: '',
        registros: [],
      };
    }
    return t;
  };

  const updateTurnoMaquina = (maquinaId: string, inst = 1, updates: Partial<TurnoMaquina>) => {
    setTurno((prev) => {
      const idx = prev.findIndex((x) => x.maquinaId === maquinaId && (x.instancia || 1) === inst);
      if (idx === -1) {
        const nuevo: TurnoMaquina = {
          maquinaId,
          instancia: inst,
          op: '',
          objetivo: null,
          motivoParada: '',
          observaciones: '',
          personal: '',
          modeloIdx: -1,
          pliego: '',
          registros: [],
          ...updates,
        };
        return [...prev, nuevo];
      }
      const copia = [...prev];
      copia[idx] = { ...copia[idx], ...updates };
      return copia;
    });
  };

  const toggleExpand = (clave: string) => {
    setExpandidos((prev) => ({
      ...prev,
      [clave]: !prev[clave],
    }));
  };

  // Add production record to a machine instance
  const handleRegistrarProduccion = (m: Maquina, inst = 1) => {
    const cardKey = claveTarjeta(m, inst);
    const t = getTurnoMaquina(m.id, inst);
    const esVasos = m.area === 'vasos_envases';
    const modo = modoCantidadMap[cardKey] || 'total';

    let cantidad = 0;
    if (esVasos && modo === 'caja') {
      const u = parseInt(undInputs[cardKey] || '', 10);
      const c = parseInt(cajasInputs[cardKey] || '', 10);
      if (!undInputs[cardKey] || !cajasInputs[cardKey] || isNaN(u) || isNaN(c) || u <= 0 || c <= 0) {
        showToast('error', 'Completa unidades x caja y cantidad de cajas');
        return;
      }
      cantidad = u * c;
    } else {
      cantidad = parseInt(cantInputs[cardKey] || '', 10);
      if (!cantInputs[cardKey] || isNaN(cantidad) || cantidad <= 0) {
        showToast('error', 'Ingresa una cantidad válida');
        return;
      }
    }

    let merma = parseInt(mermaInputs[cardKey] || '0', 10);
    if (isNaN(merma) || merma < 0) merma = 0;

    const currentOp = t.op || '';
    const match = buscarEnCatalogo(catalogo, currentOp);
    const modeloElegido = resolverModelo(match, t.modeloIdx);

    const nuevoRegistro: RegistroTurno = {
      id: Date.now() + '-' + Math.random().toString(36).slice(2, 7),
      op: currentOp,
      descripcion: modeloElegido.descripcion,
      sap: modeloElegido.sap || '',
      pliego: t.pliego || '',
      cantidad,
      merma,
      observacion: t.observaciones || '',
      fecha: new Date().toISOString(),
    };

    updateTurnoMaquina(m.id, inst, {
      registros: [nuevoRegistro, ...(t.registros || [])],
      observaciones: '',
    });

    // Reset inputs
    setCantInputs((prev) => ({ ...prev, [cardKey]: '' }));
    setUndInputs((prev) => ({ ...prev, [cardKey]: '' }));
    setCajasInputs((prev) => ({ ...prev, [cardKey]: '' }));
    setMermaInputs((prev) => ({ ...prev, [cardKey]: '' }));
    showToast('ok', 'Producción registrada');
  };

  const handleEliminarRegistro = (maquinaId: string, inst = 1, regId: string) => {
    const t = getTurnoMaquina(maquinaId, inst);
    updateTurnoMaquina(maquinaId, inst, {
      registros: (t.registros || []).filter((r) => r.id !== regId),
    });
  };

  const handleAgregarInstancia = (m: Maquina) => {
    const instancias = turnosDeMaquina(m, turno);
    const sig = instancias.reduce((mx, x) => Math.max(mx, x.instancia || 1), 0) + 1;
    const nuevaClave = claveTarjeta(m, sig);
    setTurno((prev) => [
      ...prev,
      {
        maquinaId: m.id,
        instancia: sig,
        op: '',
        objetivo: null,
        motivoParada: '',
        observaciones: '',
        personal: '',
        modeloIdx: -1,
        pliego: '',
        registros: [],
      },
    ]);
    setExpandidos((prev) => ({ ...prev, [nuevaClave]: true }));
    showToast('ok', `Nueva tarjeta ${m.nombre}(${sig}) agregada`);
  };

  const handleEliminarInstancia = (m: Maquina, inst: number) => {
    const t = getTurnoMaquina(m.id, inst);
    const n = (t.registros || []).length;
    if (n > 0) {
      try {
        if (typeof window !== 'undefined' && typeof window.confirm === 'function') {
          const ok = window.confirm(
            `¿Quitar la tarjeta ${nombreTarjeta(m, inst)}?\n\nTiene ${n} registro(s) de producción que se ELIMINARÁN.`
          );
          if (!ok) return;
        }
      } catch {
        // Ignored in sandboxed environments
      }
    }
    setTurno((prev) => {
      let filtered = prev.filter((x) => !(x.maquinaId === m.id && (x.instancia || 1) === inst));
      filtered = filtered.map((x) => {
        if (x.maquinaId === m.id && (x.instancia || 1) > inst) {
          return { ...x, instancia: (x.instancia || 1) - 1 };
        }
        return x;
      });
      return filtered;
    });
    showToast('ok', `Tarjeta ${nombreTarjeta(m, inst)} eliminada`);
  };

  // Export Turno Excel
  const handleExportarExcelTurno = () => {
    const filas = generarFilasTurno(maquinas, turno, catalogo);
    if (filas.length === 0) {
      showToast('error', 'No hay máquinas activas para exportar');
      return;
    }
    const fechaTxt = formatearFechaTurno(metaTurno.fecha);
    const tipoTxt = metaTurno.tipo ? `TURNO ${etiquetaTurno(metaTurno.tipo).toUpperCase()} ` : '';
    const titulo = `PRODUCCION ${tipoTxt}${fechaTxt}`.trim();
    const nombreBase = `turno_${metaTurno.fecha || 'sin-fecha'}_${metaTurno.tipo || 'sin-turno'}`;
    exportarTurnoExcel(filas, titulo, nombreBase);
    showToast('ok', 'Excel exportado');
  };

  // Export Migracion
  const handleExportarMigracion = () => {
    try {
      const filas = generarFilasTurno(maquinas, turno, catalogo);
      exportarMigracion(filas, metaTurno.fecha, metaTurno.tipo);
      showToast('ok', 'Excel de migración exportado');
    } catch (err: any) {
      showToast('error', err.message || 'Error al exportar migración');
    }
  };

  // Activación Excel & PDF
  const handleActivacionExcel = () => {
    const filas = filasActivacion(maquinas, turno, catalogo);
    if (filas.length === 0) {
      showToast('error', 'No hay máquinas activas');
      return;
    }
    const titulo = `ACTIVACION DE MAQUINAS ${formatearFechaTurno(metaTurno.fecha)} - ${etiquetaTurno(
      metaTurno.tipo
    ).toUpperCase()}`;
    const nombre = `activacion_${metaTurno.fecha || 'sin-fecha'}_${metaTurno.tipo || 'sin-turno'}`;
    exportarActivacionExcel(filas, titulo, nombre);
    showToast('ok', 'Activación exportada a Excel');
  };

  const handleActivacionPdf = () => {
    const filas = filasActivacion(maquinas, turno, catalogo);
    if (filas.length === 0) {
      showToast('error', 'No hay máquinas activas');
      return;
    }
    const fecha = formatearFechaTurno(metaTurno.fecha);
    const tipo = metaTurno.tipo ? etiquetaTurno(metaTurno.tipo) : '';
    const nombre = `activacion_${metaTurno.fecha || 'sin-fecha'}_${metaTurno.tipo || 'sin-turno'}`;
    imprimirActivacion(filas, fecha, tipo, nombre);
  };

  const handleCopiarReporteWsp = async () => {
    if (maquinas.length === 0) {
      showToast('error', 'No hay máquinas registradas');
      return;
    }
    const txt = generarReporteProduccion(maquinas, turno, catalogo);
    const ok = await copiarAlPortapapeles(txt);
    showToast(ok ? 'ok' : 'error', ok ? 'Reporte copiado' : 'No se pudo copiar');
  };

  const handleEnviarActivacionWsp = () => {
    if (maquinas.length === 0) {
      showToast('error', 'No hay máquinas registradas');
      return;
    }
    const texto = generarReporteActivacion(maquinas, turno, catalogo, metaTurno.fecha, metaTurno.tipo);
    const url = 'https://wa.me/?text=' + encodeURIComponent(texto);
    window.open(url, '_blank', 'noopener');
  };

  // Login handler
  const handleConfirmLogin = async () => {
    const cfg = getSupabaseConfig(false);
    if (!cfg) return;
    setLoginError('');
    if (!loginEmail || !loginPass) {
      setLoginError('Ingresa correo y contraseña');
      return;
    }
    try {
      const s = await sbLogin(cfg, loginEmail, loginPass);
      setSession(s);
      setPanelLoginOpen(false);
      setLoginPass('');
      showToast('ok', 'Sesión iniciada');
    } catch (err: any) {
      setLoginError(err.message || 'Error al iniciar sesión');
    }
  };

  // Turno backup / restore
  const handleAbrirRespaldoProd = () => {
    if (turno.length === 0) {
      showToast('error', 'No hay datos de turno que respaldar');
      return;
    }
    setPanelRestaurarOpen(false);
    setTextoRespaldo(JSON.stringify(turno));
    setPanelRespaldoOpen(true);
  };

  const handleConfirmarRestaurarProd = () => {
    const texto = textoRestaurar.trim();
    if (!texto) {
      showToast('error', 'Pega el texto de respaldo primero');
      return;
    }
    try {
      const datos = JSON.parse(texto);
      if (!Array.isArray(datos)) {
        showToast('error', 'El formato no corresponde a datos de turno');
        return;
      }
      if (turno.length > 0 && !window.confirm('Esto reemplazará el turno actual registrado. ¿Continuar?')) return;
      setTurno(datos);
      setPanelRestaurarOpen(false);
      showToast('ok', 'Turno restaurado');
    } catch {
      showToast('error', 'Ese texto no es válido');
    }
  };

  return (
    <div className="space-y-3 sm:space-y-4">
      {/* Sleek, Space-Efficient Top Controls Strip */}
      <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-2.5 sm:p-3.5 shadow-sm space-y-2.5">
        {/* Row 1: Turno switch, Date picker, and session actions */}
        <div className="flex flex-wrap items-center justify-between gap-2">
          {/* Turno Día / Noche switch */}
          <div className="flex items-center gap-2">
            <div className="inline-flex p-0.5 bg-[#232A38] rounded-lg border border-[#4A5568]">
              <button
                type="button"
                onClick={() => setMetaTurno((prev) => ({ ...prev, tipo: 'dia' }))}
                className={`flex items-center gap-1.5 py-1 px-2.5 rounded-md text-xs font-bold font-['Oswald',sans-serif] uppercase tracking-wide cursor-pointer transition-all ${
                  metaTurno.tipo === 'dia'
                    ? 'bg-[#F2A93B] text-[#12161F] shadow-sm'
                    : 'text-[#CBD5E0] hover:text-white'
                }`}
              >
                <Sun size={13} />
                <span>Día</span>
              </button>
              <button
                type="button"
                onClick={() => setMetaTurno((prev) => ({ ...prev, tipo: 'noche' }))}
                className={`flex items-center gap-1.5 py-1 px-2.5 rounded-md text-xs font-bold font-['Oswald',sans-serif] uppercase tracking-wide cursor-pointer transition-all ${
                  metaTurno.tipo === 'noche'
                    ? 'bg-[#63B3ED] text-[#12161F] shadow-sm'
                    : 'text-[#CBD5E0] hover:text-white'
                }`}
              >
                <Moon size={13} />
                <span>Noche</span>
              </button>
            </div>

            <input
              type="date"
              value={metaTurno.fecha}
              onChange={(e) => setMetaTurno((prev) => ({ ...prev, fecha: e.target.value }))}
              className="py-1 px-2 rounded-lg border border-[#4A5568] text-xs font-semibold font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none"
            />
          </div>

          {/* Quick utility links */}
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => setPanelHistorialOpen((prev) => !prev)}
              className="text-xs text-[#CBD5E0] hover:text-white py-1 px-2 rounded-lg bg-[#232A38] border border-[#4A5568] cursor-pointer flex items-center gap-1"
            >
              <History size={12} />
              <span className="hidden sm:inline">Guardados</span>
            </button>
            <button
              type="button"
              onClick={() => {
                if (session) {
                  if (window.confirm(`¿Cerrar sesión de ${session.usuario}?`)) setSession(null);
                } else {
                  setPanelLoginOpen(true);
                }
              }}
              className={`text-xs font-semibold py-1 px-2 rounded-lg cursor-pointer border ${
                session
                  ? 'bg-[#1F7A45]/20 text-[#38A169] border-[#38A169]/40'
                  : 'bg-[#232A38] text-[#F6AD55] border-[#F2A93B]/40'
              }`}
            >
              {session ? session.usuario : 'Login'}
            </button>
            <button
              type="button"
              onClick={reiniciarTurno}
              className="text-xs text-[#E53E3E] hover:bg-[#E53E3E]/20 py-1 px-2 rounded-lg border border-[#E53E3E]/30 cursor-pointer flex items-center gap-1"
              title="Reiniciar turno"
            >
              <RotateCcw size={12} />
              <span className="hidden sm:inline">Reiniciar</span>
            </button>
          </div>
        </div>

        {/* Row 2: Compact operational action bar */}
        <div className="pt-2 border-t border-[#2D3748] flex items-center gap-1.5 overflow-x-auto pb-0.5 no-scrollbar">
          <button
            type="button"
            onClick={() => {
              if (!session) {
                setPanelLoginOpen(true);
                showToast('error', 'Inicia sesión para subir los datos a la nube');
                return;
              }
              guardarTurnoDB();
            }}
            className="flex items-center gap-1.5 bg-[#F2A93B] text-[#12161F] py-1.5 px-3 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider whitespace-nowrap hover:brightness-105 active:scale-95 transition-all shadow-sm shrink-0"
          >
            <Save size={13} />
            <span>Guardar DB</span>
          </button>
          <button
            type="button"
            onClick={handleExportarExcelTurno}
            className="flex items-center gap-1 bg-[#232A38] border border-[#4A5568] text-[#CBD5E0] hover:text-white py-1.5 px-2.5 rounded-lg text-xs font-semibold whitespace-nowrap shrink-0 hover:border-[#F2A93B]/60 transition-colors"
          >
            <Download size={12} />
            <span>Excel</span>
          </button>
          <button
            type="button"
            onClick={handleExportarMigracion}
            className="flex items-center gap-1 bg-[#232A38] border border-[#F2A93B]/40 text-[#F6AD55] hover:bg-[#3A2E1B] py-1.5 px-2.5 rounded-lg text-xs font-semibold whitespace-nowrap shrink-0 transition-colors"
          >
            <Download size={12} />
            <span>Migración</span>
          </button>
          <button
            type="button"
            onClick={handleAbrirRespaldoProd}
            className="flex items-center gap-1 text-[#A0AEC0] hover:text-white py-1.5 px-2 text-xs whitespace-nowrap shrink-0 ml-auto"
            title="Respaldar datos del turno"
          >
            <Copy size={12} />
            <span>Respaldar</span>
          </button>
        </div>
      </div>

      {/* Supabase Login Modal Drawer */}
      {panelLoginOpen && (
        <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-3.5 shadow-xl">
          <p className="text-xs font-bold m-0 mb-2 text-white uppercase font-['Oswald',sans-serif] tracking-wider">
            Iniciar Sesión
          </p>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 mb-2">
            <input
              type="text"
              value={loginEmail}
              onChange={(e) => setLoginEmail(e.target.value)}
              placeholder="Usuario (sin @)"
              className="py-1.5 px-2.5 rounded-lg border border-[#4A5568] text-sm bg-[#232A38] text-white"
            />
            <input
              type="password"
              value={loginPass}
              onChange={(e) => setLoginPass(e.target.value)}
              placeholder="Contraseña"
              className="py-1.5 px-2.5 rounded-lg border border-[#4A5568] text-sm bg-[#232A38] text-white"
            />
          </div>
          {loginError && <p className="text-[#E53E3E] text-xs m-0 mb-2">{loginError}</p>}
          <div className="flex gap-2">
            <button
              type="button"
              onClick={handleConfirmLogin}
              className="flex-1 bg-[#F2A93B] text-[#12161F] py-1.5 px-3 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif]"
            >
              Ingresar
            </button>
            <button
              type="button"
              onClick={() => setPanelLoginOpen(false)}
              className="border border-[#4A5568] text-white py-1.5 px-3 rounded-lg text-xs"
            >
              Cerrar
            </button>
          </div>
        </div>
      )}

      {/* Local Shifts History Drawer */}
      {panelHistorialOpen && (
        <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-3 shadow-xl">
          <div className="flex items-center justify-between mb-2 pb-1 border-b border-[#2D3748]">
            <span className="text-xs font-bold text-white uppercase font-['Oswald',sans-serif]">
              Turnos Guardados en este dispositivo
            </span>
            <button
              type="button"
              onClick={() => setPanelHistorialOpen(false)}
              className="border border-[#4A5568] text-white py-0.5 px-2 rounded text-xs"
            >
              Cerrar
            </button>
          </div>
          <div className="max-h-60 overflow-y-auto space-y-2">
            {historial.length === 0 ? (
              <p className="text-xs text-[#CBD5E0] text-center py-3">Aún no hay turnos guardados localmente.</p>
            ) : (
              historial.map((h) => {
                const f = new Date(h.fecha);
                const fTxt =
                  f.toLocaleDateString('es-PE', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
                  ' ' +
                  f.toLocaleTimeString('es-PE', { hour: '2-digit', minute: '2-digit' });

                const etiquetaHist = h.fechaTurno
                  ? `${formatearFechaTurno(h.fechaTurno)}${h.tipoTurno ? ` · ${h.tipoTurno.toUpperCase()}` : ''}`
                  : fTxt;
                const nombreExport = `turno_${h.fechaTurno || h.fecha.slice(0, 10)}${h.tipoTurno ? `_${h.tipoTurno}` : ''}`;
                const tituloExport = h.fechaTurno
                  ? `PRODUCCION ${h.tipoTurno ? `TURNO ${h.tipoTurno.toUpperCase()} ` : ''}${formatearFechaTurno(h.fechaTurno)}`
                  : `PRODUCCION TURNO ${fTxt}`;

                return (
                  <div key={h.id} className="border border-[#2D3748] rounded-lg p-2 bg-[#232A38] text-xs">
                    <div className="flex items-center justify-between mb-1">
                      <span className="font-['IBM_Plex_Mono',monospace] font-bold text-[#F6AD55]">
                        {etiquetaHist}
                      </span>
                      <div className="flex gap-1">
                        <button
                          type="button"
                          onClick={() => {
                            copiarAlPortapapeles(h.resumen);
                            showToast('ok', 'Copiado');
                          }}
                          className="bg-[#F2A93B] text-[#12161F] py-0.5 px-2 rounded text-[11px] font-bold cursor-pointer"
                        >
                          Copiar
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            if (h.filas && h.filas.length) {
                              exportarTurnoExcel(h.filas, tituloExport, nombreExport);
                            }
                          }}
                          className="bg-[#1F7A45] text-white py-0.5 px-2 rounded text-[11px] font-bold cursor-pointer"
                        >
                          Excel
                        </button>
                        <button
                          type="button"
                          onClick={() => setHistorial((prev) => prev.filter((x) => x.id !== h.id))}
                          className="text-[#E53E3E] p-1"
                        >
                          <Trash2 size={12} />
                        </button>
                      </div>
                    </div>
                  </div>
                );
              })
            )}
          </div>
        </div>
      )}

      {/* Turno Odometer: Space-efficient */}
      <Odometer label="Total Producido (Turno Activo)" value={totalProducidoGeneral} />

      {/* Active Machines Cards: 1 column on mobile, 2 columns on tablet/desktop, 3 on wide screens */}
      {activas.length === 0 ? (
        <div className="border border-dashed border-[#4A5568] rounded-xl p-6 text-center text-xs sm:text-sm text-[#CBD5E0] bg-[#1A202C]">
          No hay máquinas activas en este turno. Ve a la pestaña <strong>Máquinas</strong> y activa las necesarias.
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-3">
          {activas.map((m) => {
            const instancias = turnosDeMaquina(m, turno);
            return (
              <React.Fragment key={m.id}>
                {instancias.map((t) => {
                  const inst = t.instancia || 1;
                  const cardKey = claveTarjeta(m, inst);
                  const cardNombre = nombreTarjeta(m, inst);
                  const expandido = !!expandidos[cardKey];
                  const sumaProducido = (t.registros || []).reduce((acc, r) => acc + (r.cantidad || 0), 0);
                  const sumaMerma = (t.registros || []).reduce((acc, r) => acc + (r.merma || 0), 0);

                  let subTexto = 'Sin OP asignada';
                  if (t.motivoParada) {
                    subTexto = '⏸ ' + t.motivoParada;
                  } else if (t.op) {
                    const match = buscarEnCatalogo(catalogo, t.op);
                    const descTxt = match ? resolverModelo(match, t.modeloIdx).descripcion : '';
                    subTexto = t.op + (descTxt ? ' — ' + descTxt : '') + (t.pliego ? ' (' + etiquetaPliego(t.pliego) + ')' : '');
                  }

                  const esVasos = m.area === 'vasos_envases';
                  const modo = modoCantidadMap[cardKey] || 'total';

                  // Suggestions for this machine instance (búsqueda optimizada por dígitos)
                  const queryOp = (opSuggestMap[cardKey] || '').trim();
                  const queryDigits = queryOp.replace(/\D/g, '');
                  const matchingSuggs = queryDigits
                    ? catalogo
                        .filter((c) => {
                          const cDigits = c.op.replace(/\D/g, '');
                          return (queryDigits && cDigits.includes(queryDigits)) || c.op.toLowerCase().includes(queryOp.toLowerCase());
                        })
                        .sort((a, b) => {
                          const aDigits = a.op.replace(/\D/g, '');
                          const bDigits = b.op.replace(/\D/g, '');
                          const aStarts = queryDigits && aDigits.startsWith(queryDigits);
                          const bStarts = queryDigits && bDigits.startsWith(queryDigits);
                          if (aStarts && !bStarts) return -1;
                          if (!aStarts && bStarts) return 1;
                          return aDigits.length - bDigits.length;
                        })
                        .slice(0, 8)
                    : [];

                  const currentMatch = buscarEnCatalogo(catalogo, t.op);

                  return (
                    <div
                      key={cardKey}
                      className={`bg-[#1A202C] border rounded-xl p-3 flex flex-col justify-between transition-all shadow-sm ${
                        t.motivoParada
                          ? 'border-[#E53E3E]/60 bg-[#1A202C]'
                          : expandido
                          ? 'border-[#F2A93B]/70 ring-1 ring-[#F2A93B]/30'
                          : 'border-[#2D3748] hover:border-[#4A5568]'
                      }`}
                    >
                      {/* Card Header (Tap to expand/collapse) */}
                      <div
                        onClick={() => toggleExpand(cardKey)}
                        className="flex items-center justify-between gap-2.5 cursor-pointer select-none py-0.5"
                      >
                        <div className="min-w-0 flex-1">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className="font-['IBM_Plex_Mono',monospace] font-bold text-sm sm:text-base text-white truncate">
                              {cardNombre}
                            </span>
                            {m.codigo && (
                              <span className="text-[10px] font-bold px-1.5 py-0.2 rounded bg-[#232A38] border border-[#4A5568] text-[#CBD5E0] font-['IBM_Plex_Mono',monospace]">
                                {m.codigo}
                              </span>
                            )}
                            {m.tipoProducto && (
                              <span
                                className={`text-[9.5px] font-bold px-1.5 py-0.2 rounded ${
                                  m.tipoProducto === 'pp'
                                    ? 'bg-[#F2A93B]/20 text-[#F6AD55]'
                                    : 'bg-[#63B3ED]/20 text-[#63B3ED]'
                                }`}
                              >
                                {m.tipoProducto.toUpperCase()}
                              </span>
                            )}
                          </div>
                          <div className="text-[11px] text-[#A0AEC0] truncate mt-0.5">
                            {subTexto}
                          </div>
                        </div>

                        <div className="flex items-center gap-2 shrink-0">
                          <div className="text-right">
                            <div className="font-['IBM_Plex_Mono',monospace] text-sm font-bold text-[#F2A93B] tabular-nums">
                              {sumaProducido.toLocaleString('es-PE')} u
                            </div>
                            {sumaMerma > 0 && (
                              <div className="text-[9.5px] text-[#E53E3E]">
                                -{sumaMerma} merma
                              </div>
                            )}
                          </div>
                          <div className="w-6 h-6 rounded bg-[#232A38] flex items-center justify-center text-[#A0AEC0]">
                            <ChevronDown
                              size={15}
                              className={`transition-transform duration-200 ${expandido ? 'rotate-180 text-white' : ''}`}
                            />
                          </div>
                        </div>
                      </div>

                      {/* Expanded Detail Panel */}
                      {expandido && (
                        <div className="mt-3 pt-3 border-t border-[#2D3748] space-y-2.5">
                          {/* Personal & Parada in a 2-column grid on desktop */}
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                            <div>
                              <label className="block text-[10px] font-semibold text-[#CBD5E0] mb-0.5 uppercase tracking-wide">
                                Personal a cargo
                              </label>
                              <input
                                type="text"
                                value={t.personal || ''}
                                onChange={(e) => updateTurnoMaquina(m.id, inst, { personal: e.target.value })}
                                placeholder="Operador / equipo"
                                className="w-full py-1.5 px-2.5 rounded-lg border border-[#4A5568] text-sm outline-none bg-[#232A38] text-white focus:border-[#F2A93B]"
                              />
                            </div>

                            <div>
                              <label className="block text-[10px] font-semibold text-[#CBD5E0] mb-0.5 uppercase tracking-wide">
                                Estado operativo
                              </label>
                              <select
                                value={
                                  t.motivoParada
                                    ? MOTIVOS_PARADA.includes(t.motivoParada)
                                      ? t.motivoParada
                                      : '__otros'
                                    : ''
                                }
                                onChange={(e) => {
                                  const v = e.target.value;
                                  if (!v) updateTurnoMaquina(m.id, inst, { motivoParada: '' });
                                  else if (v === '__otros') updateTurnoMaquina(m.id, inst, { motivoParada: 'Otros' });
                                  else updateTurnoMaquina(m.id, inst, { motivoParada: v });
                                }}
                                className="w-full py-1.5 px-2 rounded-lg border border-[#4A5568] text-sm outline-none bg-[#232A38] text-white"
                              >
                                <option value="">🟢 Operativa (Sin parada)</option>
                                {MOTIVOS_PARADA.map((mot) => (
                                  <option key={mot} value={mot}>
                                    🔴 Parada: {mot}
                                  </option>
                                ))}
                                <option value="__otros">🔴 Parada: Otros motivos…</option>
                              </select>
                            </div>
                          </div>

                          {t.motivoParada && !MOTIVOS_PARADA.includes(t.motivoParada) && (
                            <input
                              type="text"
                              value={t.motivoParada.replace(/^Otros:?\s*/i, '')}
                              onChange={(e) => {
                                const x = e.target.value.trim();
                                updateTurnoMaquina(m.id, inst, { motivoParada: x ? 'Otros: ' + x : 'Otros' });
                              }}
                              placeholder="Especifica el motivo de la parada"
                              className="w-full py-1.5 px-2.5 rounded-lg border border-[#4A5568] text-sm outline-none bg-[#232A38] text-white"
                            />
                          )}

                          {/* Observaciones */}
                          <div>
                            <input
                              type="text"
                              value={t.observaciones || ''}
                              onChange={(e) => updateTurnoMaquina(m.id, inst, { observaciones: e.target.value })}
                              placeholder="Observaciones de la máquina (opcional)"
                              className="w-full py-1.5 px-2.5 rounded-lg border border-[#4A5568] text-sm outline-none bg-[#232A38] text-white focus:border-[#F2A93B]"
                            />
                          </div>

                          {/* OP Assignment */}
                          <div>
                            {t.op ? (
                              <div className="space-y-1.5">
                                <div className="flex items-center justify-between gap-2 py-1.5 px-2.5 rounded-lg bg-[#232A38] border border-[#4A5568]">
                                  <span className="font-['IBM_Plex_Mono',monospace] font-bold text-xs sm:text-sm text-[#F6AD55]">
                                    {t.op}
                                  </span>
                                  <button
                                    type="button"
                                    onClick={() => updateTurnoMaquina(m.id, inst, { op: '', modeloIdx: -1 })}
                                    className="text-[#E53E3E] hover:underline text-xs font-semibold cursor-pointer"
                                  >
                                    ✕ Cambiar
                                  </button>
                                </div>

                                {/* Multi-model chips */}
                                {currentMatch && currentMatch.modelos && currentMatch.modelos.length > 1 && (
                                  <div className="flex gap-1 flex-wrap pt-0.5">
                                    <button
                                      type="button"
                                      onClick={() => updateTurnoMaquina(m.id, inst, { modeloIdx: -1 })}
                                      className={`py-1 px-2 rounded text-[11px] font-semibold cursor-pointer border ${
                                        t.modeloIdx === -1
                                          ? 'bg-[#F2A93B] text-[#12161F] border-[#F2A93B]'
                                          : 'bg-[#1A202C] text-[#CBD5E0] border-[#4A5568]'
                                      }`}
                                    >
                                      Todos
                                    </button>
                                    {currentMatch.modelos.map((mo, i) => (
                                      <button
                                        key={i}
                                        type="button"
                                        onClick={() => updateTurnoMaquina(m.id, inst, { modeloIdx: i })}
                                        className={`py-1 px-2 rounded text-[11px] font-semibold cursor-pointer border ${
                                          t.modeloIdx === i
                                            ? 'bg-[#F2A93B] text-[#12161F] border-[#F2A93B]'
                                            : 'bg-[#1A202C] text-[#CBD5E0] border-[#4A5568]'
                                        }`}
                                      >
                                        {mo.descripcion || `Mod ${i + 1}`}
                                      </button>
                                    ))}
                                  </div>
                                )}

                                {currentMatch && (
                                  <div className="text-[11px] text-[#CBD5E0] truncate bg-[#232A38] px-2 py-1 rounded">
                                    {resolverModelo(currentMatch, t.modeloIdx).descripcion} · <span className="text-[#A0AEC0]">{currentMatch.cliente}</span>
                                  </div>
                                )}
                              </div>
                            ) : (
                              <div className="relative">
                                <div className="relative">
                                  <input
                                    type="text"
                                    inputMode="numeric"
                                    pattern="[0-9]*"
                                    value={opSuggestMap[cardKey] || ''}
                                    onChange={(e) => {
                                      const val = e.target.value.replace(/\D/g, '');
                                      setOpSuggestMap((prev) => ({ ...prev, [cardKey]: val }));
                                      setShowSuggestMap((prev) => ({ ...prev, [cardKey]: true }));
                                    }}
                                    onFocus={() => setShowSuggestMap((prev) => ({ ...prev, [cardKey]: true }))}
                                    onBlur={() => {
                                      setTimeout(() => {
                                        setShowSuggestMap((prev) => ({ ...prev, [cardKey]: false }));
                                      }, 240);
                                    }}
                                    onKeyDown={(e) => {
                                      const allowed = ['Backspace', 'Delete', 'Tab', 'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Enter', 'Escape'];
                                      if (allowed.includes(e.key) || e.ctrlKey || e.metaKey) {
                                        if (e.key === 'Enter') {
                                          e.preventDefault();
                                          const typed = (opSuggestMap[cardKey] || '').replace(/\D/g, '').trim();
                                          if (typed) {
                                            const selected = matchingSuggs.length > 0 ? matchingSuggs[0].op : typed;
                                            updateTurnoMaquina(m.id, inst, { op: selected, modeloIdx: -1 });
                                            setOpSuggestMap((prev) => ({ ...prev, [cardKey]: '' }));
                                            setShowSuggestMap((prev) => ({ ...prev, [cardKey]: false }));
                                          }
                                        }
                                        return;
                                      }
                                      if (!/[0-9]/.test(e.key)) {
                                        e.preventDefault();
                                      }
                                    }}
                                    placeholder="Escribe números de OP (ej. 10234)..."
                                    autoComplete="off"
                                    className="w-full py-1.5 px-2.5 pr-8 rounded-lg border border-[#4A5568] text-sm font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none focus:border-[#F2A93B] focus:ring-1 focus:ring-[#F2A93B]/30"
                                  />
                                  {opSuggestMap[cardKey] && (
                                    <button
                                      type="button"
                                      onClick={() => {
                                        setOpSuggestMap((prev) => ({ ...prev, [cardKey]: '' }));
                                        setShowSuggestMap((prev) => ({ ...prev, [cardKey]: false }));
                                      }}
                                      className="absolute right-2 top-1/2 -translate-y-1/2 text-[#A0AEC0] hover:text-white p-1 rounded"
                                      title="Borrar"
                                    >
                                      <X size={13} />
                                    </button>
                                  )}
                                </div>
                                {showSuggestMap[cardKey] && queryDigits !== '' && (
                                  <div className="absolute z-40 left-0 right-0 mt-1 bg-[#1A202C] border border-[#4A5568] rounded-xl max-h-56 overflow-y-auto shadow-2xl overflow-hidden divide-y divide-[#2D3748]">
                                    {matchingSuggs.length > 0 ? (
                                      <>
                                        <div className="px-2.5 py-1 bg-[#232A38] flex items-center justify-between text-[10px] text-[#A0AEC0]">
                                          <span className="font-semibold uppercase tracking-wider text-[#CBD5E0]">
                                            Catálogo ({matchingSuggs.length})
                                          </span>
                                          <span>Toca para asignar</span>
                                        </div>
                                        {matchingSuggs.map((item) => (
                                          <button
                                            key={item.op}
                                            type="button"
                                            onMouseDown={(e) => {
                                              e.preventDefault();
                                              updateTurnoMaquina(m.id, inst, { op: item.op, modeloIdx: -1 });
                                              setOpSuggestMap((prev) => ({ ...prev, [cardKey]: '' }));
                                              setShowSuggestMap((prev) => ({ ...prev, [cardKey]: false }));
                                            }}
                                            className="w-full text-left py-2 px-2.5 hover:bg-[#232A38] active:bg-[#2D3748] text-xs block text-white transition-colors cursor-pointer group"
                                          >
                                            <div className="flex items-center justify-between gap-1.5 mb-0.5">
                                              <span className="font-bold text-[#F2A93B] group-hover:text-[#F6AD55] font-['IBM_Plex_Mono',monospace] text-xs px-1.5 py-0.2 rounded bg-[#232A38] border border-[#4A5568]">
                                                {item.op}
                                              </span>
                                              <div className="flex items-center gap-1">
                                                {item.modelos && item.modelos.length > 1 && (
                                                  <span className="text-[9px] px-1 py-0.2 rounded bg-[#4A5568]/40 text-[#CBD5E0]">
                                                    {item.modelos.length} mod.
                                                  </span>
                                                )}
                                                <span className="text-[#A0AEC0] text-[10px] truncate max-w-[120px]">
                                                  {item.cliente || 'Sin cliente'}
                                                </span>
                                              </div>
                                            </div>
                                            <div className="text-[#CBD5E0] text-[11px] truncate">
                                              {descripcionCorta(item)}
                                            </div>
                                          </button>
                                        ))}
                                      </>
                                    ) : (
                                      <button
                                        type="button"
                                        onMouseDown={(e) => {
                                          e.preventDefault();
                                          updateTurnoMaquina(m.id, inst, { op: queryDigits, modeloIdx: -1 });
                                          setOpSuggestMap((prev) => ({ ...prev, [cardKey]: '' }));
                                          setShowSuggestMap((prev) => ({ ...prev, [cardKey]: false }));
                                        }}
                                        className="w-full text-left p-2.5 hover:bg-[#232A38] text-xs block text-white cursor-pointer"
                                      >
                                        <div className="text-[#F2A93B] font-bold font-mono">
                                          + Asignar OP {queryDigits} (manual)
                                        </div>
                                        <div className="text-[10px] text-[#A0AEC0] mt-0.5">
                                          No está en el catálogo, pero puedes asignarla directamente a esta máquina.
                                        </div>
                                      </button>
                                    )}
                                  </div>
                                )}
                              </div>
                            )}
                          </div>

                          {/* Pliego selector for PP machines: compact touch row */}
                          {m.tipoProducto === 'pp' && (
                            <div className="p-1.5 bg-[#232A38] rounded-lg border border-[#4A5568]">
                              <span className="text-[10px] font-semibold text-[#CBD5E0] block mb-1 uppercase tracking-wide">
                                Pliego (PP):
                              </span>
                              <div className="flex gap-1 flex-wrap">
                                <button
                                  type="button"
                                  onClick={() => updateTurnoMaquina(m.id, inst, { pliego: '' })}
                                  className={`py-1 px-2 rounded text-[11px] font-bold cursor-pointer ${
                                    !t.pliego ? 'bg-[#F2A93B] text-[#12161F]' : 'text-[#CBD5E0] bg-[#1A202C]'
                                  }`}
                                >
                                  Único
                                </button>
                                <button
                                  type="button"
                                  onClick={() => updateTurnoMaquina(m.id, inst, { pliego: 'Tira' })}
                                  className={`py-1 px-2 rounded text-[11px] font-bold cursor-pointer ${
                                    t.pliego === 'Tira' ? 'bg-[#F2A93B] text-[#12161F]' : 'text-[#CBD5E0] bg-[#1A202C]'
                                  }`}
                                >
                                  Tira
                                </button>
                                <button
                                  type="button"
                                  onClick={() => updateTurnoMaquina(m.id, inst, { pliego: 'Retira' })}
                                  className={`py-1 px-2 rounded text-[11px] font-bold cursor-pointer ${
                                    t.pliego === 'Retira' ? 'bg-[#F2A93B] text-[#12161F]' : 'text-[#CBD5E0] bg-[#1A202C]'
                                  }`}
                                >
                                  Retira
                                </button>
                                {[1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map((pn) => (
                                  <button
                                    key={pn}
                                    type="button"
                                    onClick={() => updateTurnoMaquina(m.id, inst, { pliego: pn })}
                                    className={`py-1 px-2 rounded text-[11px] font-bold font-['IBM_Plex_Mono',monospace] cursor-pointer ${
                                      t.pliego === pn ? 'bg-[#F2A93B] text-[#12161F]' : 'text-[#CBD5E0] bg-[#1A202C]'
                                    }`}
                                  >
                                    P{pn}
                                  </button>
                                ))}
                              </div>
                            </div>
                          )}

                          {/* Mode selector if Vasos y Envases */}
                          {esVasos && (
                            <div className="flex gap-1 p-0.5 bg-[#232A38] rounded-lg border border-[#4A5568]">
                              <button
                                type="button"
                                onClick={() => setModoCantidadMap((prev) => ({ ...prev, [cardKey]: 'total' }))}
                                className={`flex-1 py-1 px-2 rounded text-[11px] font-bold cursor-pointer ${
                                  modo === 'total'
                                    ? 'bg-[#F2A93B] text-[#12161F] shadow-sm'
                                    : 'text-[#CBD5E0]'
                                }`}
                              >
                                UND Totales
                              </button>
                              <button
                                type="button"
                                onClick={() => setModoCantidadMap((prev) => ({ ...prev, [cardKey]: 'caja' }))}
                                className={`flex-1 py-1 px-2 rounded text-[11px] font-bold cursor-pointer ${
                                  modo === 'caja'
                                    ? 'bg-[#F2A93B] text-[#12161F] shadow-sm'
                                    : 'text-[#CBD5E0]'
                                }`}
                              >
                                UND × Caja
                              </button>
                            </div>
                          )}

                          {/* Production registration input box: compact and guaranteed no overflow */}
                          {(() => {
                            const uVal = parseInt(undInputs[cardKey] || '', 10) || 0;
                            const cVal = parseInt(cajasInputs[cardKey] || '', 10) || 0;
                            const totalCajasCalculado = uVal * cVal;

                            return (
                              <div className="p-2 bg-[#232A38] rounded-lg border border-[#2D3748] space-y-2">
                                <div className="flex items-center gap-1.5 w-full">
                                  {esVasos && modo === 'caja' ? (
                                    <>
                                      <input
                                        type="number"
                                        inputMode="numeric"
                                        min="0"
                                        value={undInputs[cardKey] || ''}
                                        onChange={(e) => setUndInputs((prev) => ({ ...prev, [cardKey]: e.target.value }))}
                                        placeholder="Und/caja"
                                        className="min-w-0 flex-1 py-1.5 px-2 rounded-lg border border-[#4A5568] text-xs sm:text-sm font-['IBM_Plex_Mono',monospace] bg-[#1A202C] text-white"
                                      />
                                      <input
                                        type="number"
                                        inputMode="numeric"
                                        min="0"
                                        value={cajasInputs[cardKey] || ''}
                                        onChange={(e) => setCajasInputs((prev) => ({ ...prev, [cardKey]: e.target.value }))}
                                        placeholder="Cajas"
                                        className="min-w-0 flex-1 py-1.5 px-2 rounded-lg border border-[#4A5568] text-xs sm:text-sm font-['IBM_Plex_Mono',monospace] bg-[#1A202C] text-white"
                                      />
                                    </>
                                  ) : (
                                    <input
                                      type="number"
                                      inputMode="numeric"
                                      min="0"
                                      value={cantInputs[cardKey] || ''}
                                      onChange={(e) => setCantInputs((prev) => ({ ...prev, [cardKey]: e.target.value }))}
                                      placeholder="Cantidad"
                                      className="min-w-0 flex-1 py-1.5 px-2.5 rounded-lg border border-[#4A5568] text-xs sm:text-sm font-['IBM_Plex_Mono',monospace] bg-[#1A202C] text-white"
                                    />
                                  )}

                                  <input
                                    type="number"
                                    inputMode="numeric"
                                    min="0"
                                    value={mermaInputs[cardKey] || ''}
                                    onChange={(e) => setMermaInputs((prev) => ({ ...prev, [cardKey]: e.target.value }))}
                                    placeholder="Merma"
                                    className="min-w-0 w-16 sm:w-20 py-1.5 px-2 rounded-lg border border-[#4A5568] text-xs sm:text-sm font-['IBM_Plex_Mono',monospace] bg-[#1A202C] text-white"
                                  />

                                  <button
                                    type="button"
                                    onClick={() => handleRegistrarProduccion(m, inst)}
                                    className="w-9 h-8 sm:w-10 sm:h-9 bg-[#F2A93B] text-[#12161F] rounded-lg cursor-pointer hover:brightness-105 active:scale-95 transition-all shadow-sm shrink-0 flex items-center justify-center font-bold"
                                    title="Registrar avance"
                                    aria-label="Registrar avance"
                                  >
                                    <Plus size={18} strokeWidth={2.5} />
                                  </button>
                                </div>

                                {/* Sumatoria / Multiplicación calculada para UND x Caja */}
                                {esVasos && modo === 'caja' && (
                                  <div className="flex items-center justify-between px-2 py-1 rounded bg-[#1A202C] border border-[#2D3748] text-xs font-['IBM_Plex_Mono',monospace]">
                                    <span className="text-[#A0AEC0] text-[11px]">Total cálculo:</span>
                                    <span className="font-bold text-[#F2A93B] tabular-nums">
                                      {uVal > 0 && cVal > 0 ? (
                                        <>
                                          <span className="text-[#A0AEC0] font-normal text-[11px]">
                                            {uVal.toLocaleString('es-PE')} × {cVal.toLocaleString('es-PE')} ={' '}
                                          </span>
                                          <span className="text-[#F6AD55] font-black">
                                            {totalCajasCalculado.toLocaleString('es-PE')}
                                          </span>{' '}
                                          u
                                        </>
                                      ) : (
                                        <span className="text-[#718096]">= 0 u</span>
                                      )}
                                    </span>
                                  </div>
                                )}
                              </div>
                            );
                          })()}

                          {/* Records History List for this machine */}
                          {t.registros && t.registros.length > 0 && (
                            <div className="space-y-1 pt-1">
                              <span className="text-[10px] font-bold text-[#CBD5E0] uppercase tracking-wide block">
                                Avances registrados ({t.registros.length})
                              </span>
                              <div className="space-y-1 max-h-40 overflow-y-auto">
                                {t.registros.map((r, regIdx) => {
                                  const descLimpia = (r.descripcion || '').replace(/^[^A-Za-zÁÉÍÓÚÑÜáéíóúñü]+/, '').trim();

                                  return (
                                    <div
                                      key={r.id}
                                      className="py-1 px-2 rounded-lg bg-[#232A38] border border-[#2D3748] flex items-center justify-between gap-1.5 text-xs"
                                    >
                                      <div className="min-w-0 flex-1 truncate text-[#CBD5E0]">
                                        <span className="text-[#A0AEC0] font-['IBM_Plex_Mono',monospace] text-[10px] font-bold mr-1.5">
                                          #{regIdx + 1}
                                        </span>
                                        <strong className="text-white font-['IBM_Plex_Mono',monospace]">{r.op || 'Sin OP'}</strong>
                                        {descLimpia && ` — ${descLimpia}`}
                                      </div>
                                      <div className="flex items-center gap-1.5 shrink-0">
                                        <span className="font-['IBM_Plex_Mono',monospace] font-bold text-[#F2A93B] tabular-nums text-xs">
                                          +{r.cantidad.toLocaleString('es-PE')}
                                        </span>
                                        {r.merma > 0 && (
                                          <span className="text-[#E53E3E] text-[10px]">
                                            (-{r.merma})
                                          </span>
                                        )}
                                        <button
                                          type="button"
                                          onClick={() => handleEliminarRegistro(m.id, inst, r.id)}
                                          className="p-1 text-[#E53E3E] hover:bg-[#E53E3E]/20 rounded cursor-pointer"
                                          title="Eliminar registro"
                                        >
                                          <Trash2 size={13} />
                                        </button>
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                            </div>
                          )}

                          {/* Quitar esta tarjeta si es Multi-OP y hay más de 1 tarjeta */}
                          {m.multiOp && instancias.length > 1 && (
                            <button
                              type="button"
                              onClick={() => handleEliminarInstancia(m, inst)}
                              className="w-full mt-2 py-1.5 px-3 rounded-lg border border-[#E53E3E] text-[#E53E3E] hover:bg-[#E53E3E]/10 text-xs font-semibold cursor-pointer transition-colors"
                            >
                              ✕ Quitar esta tarjeta ({cardNombre})
                            </button>
                          )}
                        </div>
                      )}
                    </div>
                  );
                })}

                {/* Si la máquina tiene Multi-OP activado, botón para agregar otra tarjeta */}
                {m.multiOp && (
                  <div className="col-span-full md:col-span-2 xl:col-span-3">
                    <button
                      type="button"
                      onClick={() => handleAgregarInstancia(m)}
                      className="w-full max-w-full bg-[#1A202C] hover:bg-[#232A38] text-[#F6AD55] border border-dashed border-[#F2A93B]/60 p-2.5 rounded-xl text-xs font-bold cursor-pointer transition-all flex items-center justify-center gap-1.5 shadow-sm truncate"
                    >
                      <Plus size={14} className="shrink-0" />
                      <span className="truncate">+ Agregar OP a {m.nombre} (nueva tarjeta)</span>
                    </button>
                  </div>
                )}
              </React.Fragment>
            );
          })}
        </div>
      )}

      {/* ============================================================== */}
      {/* DESKTOP FLOATING BAR: Sleek, compact 1-line docked bar at bottom-0 */}
      {/* ============================================================== */}
      <div className="hidden md:block fixed bottom-0 left-0 right-0 bg-[#1A202C]/95 backdrop-blur-md border-t border-[#F2A93B] z-30 shadow-2xl">
        <div className="max-w-[1440px] mx-auto px-4 lg:px-6 py-2 flex items-center justify-between gap-3">
          <div className="text-xs text-[#CBD5E0] flex items-center gap-2">
            <span>Reportes de Turno:</span>
            <strong className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace]">
              {totalProducidoGeneral.toLocaleString('es-PE')} u producidas
            </strong>
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={handleCopiarReporteWsp}
              className="flex items-center gap-1.5 bg-[#232A38] border border-[#4A5568] text-white py-1 px-3 rounded-lg text-xs font-semibold cursor-pointer hover:bg-[#2D3748]"
            >
              <Share2 size={13} className="text-[#38A169]" />
              <span>Copiar WhatsApp</span>
            </button>
            <button
              type="button"
              onClick={handleEnviarActivacionWsp}
              className="flex items-center gap-1.5 bg-[#232A38] border border-[#93C5FD] text-[#93C5FD] py-1 px-3 rounded-lg text-xs font-semibold cursor-pointer hover:bg-[#2D3748]"
            >
              <Share2 size={13} />
              <span>Enviar Activación</span>
            </button>
            <button
              type="button"
              onClick={handleActivacionExcel}
              className="flex items-center gap-1.5 bg-[#F2A93B] text-[#12161F] py-1 px-3 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:brightness-105"
            >
              <FileSpreadsheet size={13} />
              <span>Activación Excel</span>
            </button>
            <button
              type="button"
              onClick={handleActivacionPdf}
              className="flex items-center gap-1.5 bg-[#232A38] text-white border border-[#4A5568] py-1 px-3 rounded-lg text-xs font-semibold cursor-pointer hover:bg-[#2D3748]"
            >
              <Printer size={13} />
              <span>Activación PDF</span>
            </button>
          </div>
        </div>
      </div>

      {/* ============================================================== */}
      {/* MOBILE REPORT TOGGLE DOCK: Never permanently covers cards! */}
      {/* A single sleek compact floating pill button that opens reports when tapped */}
      {/* ============================================================== */}
      <div className="md:hidden fixed bottom-[58px] left-3 right-3 z-30 pointer-events-none">
        {reportesAbiertosMobile ? (
          <div className="bg-[#1A202C] border-2 border-[#F2A93B] rounded-2xl p-3 shadow-2xl space-y-2 pointer-events-auto animate-in slide-in-from-bottom-2 duration-150">
            <div className="flex items-center justify-between pb-1.5 border-b border-[#2D3748]">
              <span className="text-xs font-bold text-white uppercase font-['Oswald',sans-serif] tracking-wide flex items-center gap-1.5">
                <Share2 size={14} className="text-[#F2A93B]" />
                Reportes y Activación
              </span>
              <button
                type="button"
                onClick={() => setReportesAbiertosMobile(false)}
                className="text-[#CBD5E0] hover:text-white p-1 rounded-lg bg-[#232A38]"
              >
                <X size={15} />
              </button>
            </div>
            <div className="grid grid-cols-2 gap-2">
              <button
                type="button"
                onClick={() => {
                  handleCopiarReporteWsp();
                  setReportesAbiertosMobile(false);
                }}
                className="flex items-center justify-center gap-1.5 bg-[#232A38] border border-[#4A5568] text-white py-2 px-2 rounded-xl text-xs font-semibold active:scale-95"
              >
                <Share2 size={13} className="text-[#38A169]" />
                <span>Copiar WhatsApp</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  handleEnviarActivacionWsp();
                  setReportesAbiertosMobile(false);
                }}
                className="flex items-center justify-center gap-1.5 bg-[#232A38] border border-[#93C5FD] text-[#93C5FD] py-2 px-2 rounded-xl text-xs font-semibold active:scale-95"
              >
                <Share2 size={13} />
                <span>Enviar WSP</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  handleActivacionExcel();
                  setReportesAbiertosMobile(false);
                }}
                className="flex items-center justify-center gap-1.5 bg-[#F2A93B] text-[#12161F] py-2 px-2 rounded-xl text-xs font-bold uppercase font-['Oswald',sans-serif] active:scale-95 shadow-sm"
              >
                <FileSpreadsheet size={14} />
                <span>Activación Excel</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  handleActivacionPdf();
                  setReportesAbiertosMobile(false);
                }}
                className="flex items-center justify-center gap-1.5 bg-[#232A38] text-white border border-[#4A5568] py-2 px-2 rounded-xl text-xs font-semibold active:scale-95"
              >
                <Printer size={14} />
                <span>Activación PDF</span>
              </button>
            </div>
          </div>
        ) : (
          <div className="flex justify-end">
            <button
              type="button"
              onClick={() => setReportesAbiertosMobile(true)}
              className="pointer-events-auto flex items-center gap-2 bg-[#F2A93B] text-[#12161F] py-2 px-3.5 rounded-full text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider shadow-xl shadow-[#F2A93B]/25 active:scale-95 transition-all border border-[#F2A93B]"
            >
              <Share2 size={14} />
              <span>Reportes / WhatsApp</span>
            </button>
          </div>
        )}
      </div>
    </div>
  );
};
