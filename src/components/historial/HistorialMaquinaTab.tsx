import React, { useState, useMemo, useCallback, useEffect } from 'react';
import {
  Search,
  X,
  RefreshCw,
  Sun,
  Moon,
  Calendar,
  Cpu,
  AlertCircle,
  Layers,
  TrendingUp,
  FileDown,
  Target,
  ChevronDown,
  Package,
} from 'lucide-react';
import { useApp } from '../../context/AppContext';
import { formatearFechaTurno } from '../../services/reports';
import {
  consultarHistorialMaquina,
  ConsultaMaquinaResultado,
  MaquinaOpConsolidado,
  TurnoMaquinaResumen,
} from '../../services/opHistory';
import { exportarPdfReporteMaquina } from '../../services/pdfMaquina';
import { Maquina } from '../../types';

function toLocalIsoDate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

export const HistorialMaquinaTab: React.FC = () => {
  const { catalogo, historial, maquinas, turno, metaTurno, session, showToast } = useApp();

  const [maqInput, setMaqInput] = useState('');
  const [showDropdown, setShowDropdown] = useState(false);
  const [maquinaSeleccionada, setMaquinaSeleccionada] = useState<Maquina | null>(null);

  const [fechaDesde, setFechaDesde] = useState('');
  const [fechaHasta, setFechaHasta] = useState('');
  const [tipoTurnoFiltro, setTipoTurnoFiltro] = useState<'' | 'dia' | 'noche'>('');

  const [cargando, setCargando] = useState(false);
  const [resultado, setResultado] = useState<ConsultaMaquinaResultado | null>(null);

  const areaTexto = (area: Maquina['area']) =>
    area === 'vasos_envases'
      ? 'Vasos y Envases'
      : area === 'papel_carton'
      ? 'Papel y Cartón'
      : '';

  // Autocomplete dropdown items from configured machines (active first, then alphabetical)
  const sugerenciasMaquinas = useMemo(() => {
    const q = maqInput.trim().toLowerCase();
    return [...maquinas]
      .filter((m) => {
        if (!q) return true;
        return (
          m.nombre.toLowerCase().includes(q) ||
          (m.codigo || '').toLowerCase().includes(q) ||
          areaTexto(m.area).toLowerCase().includes(q)
        );
      })
      .sort((a, b) => {
        if (a.activa !== b.activa) return a.activa ? -1 : 1;
        return a.nombre.localeCompare(b.nombre, 'es');
      });
  }, [maquinas, maqInput]);

  // Quick-pick active machines in current shift (or configured active machines)
  const maquinasRapidas = useMemo(() => {
    const enTurnoIds = new Set((turno || []).map((t) => t.maquinaId));
    const activasTurno = maquinas.filter((m) => enTurnoIds.has(m.id));
    if (activasTurno.length > 0) return activasTurno.slice(0, 10);
    return maquinas.filter((m) => m.activa).slice(0, 10);
  }, [maquinas, turno]);

  const ejecutarConsulta = useCallback(
    async (maq: Maquina, desdeStr: string, hastaStr: string) => {
      setCargando(true);
      const res = await consultarHistorialMaquina({
        maquina: maq,
        historial,
        metaTurno,
        session,
        turno,
        catalogo,
        fechaDesde: desdeStr || undefined,
        fechaHasta: hastaStr || undefined,
      });
      setResultado(res);
      setCargando(false);
    },
    [historial, metaTurno, session, turno, catalogo]
  );

  useEffect(() => {
    if (maquinaSeleccionada) {
      ejecutarConsulta(maquinaSeleccionada, fechaDesde, fechaHasta);
    }
  }, [maquinaSeleccionada, fechaDesde, fechaHasta, ejecutarConsulta]);

  const handleSelectMaquina = (m: Maquina) => {
    setMaquinaSeleccionada(m);
    setMaqInput(m.codigo ? `${m.nombre} (${m.codigo})` : m.nombre);
    setShowDropdown(false);
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const q = maqInput.trim().toLowerCase();
    if (!q) {
      setShowDropdown(true);
      return;
    }
    const exact =
      maquinas.find(
        (m) =>
          m.nombre.toLowerCase() === q ||
          (m.codigo || '').toLowerCase() === q ||
          `${m.nombre} (${m.codigo})`.toLowerCase() === q
      ) || sugerenciasMaquinas[0];

    if (exact) {
      handleSelectMaquina(exact);
    } else {
      showToast('error', 'No se encontró una máquina coincidente');
    }
  };

  const aplicarPresetFecha = (preset: 'hoy' | '7d' | '30d' | 'mes' | 'todo') => {
    const hoy = new Date();
    const hoyStr = toLocalIsoDate(hoy);
    if (preset === 'hoy') {
      setFechaDesde(hoyStr);
      setFechaHasta(hoyStr);
    } else if (preset === '7d') {
      const ini = new Date(hoy);
      ini.setDate(ini.getDate() - 6);
      setFechaDesde(toLocalIsoDate(ini));
      setFechaHasta(hoyStr);
    } else if (preset === '30d') {
      const ini = new Date(hoy);
      ini.setDate(ini.getDate() - 29);
      setFechaDesde(toLocalIsoDate(ini));
      setFechaHasta(hoyStr);
    } else if (preset === 'mes') {
      const ini = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
      setFechaDesde(toLocalIsoDate(ini));
      setFechaHasta(hoyStr);
    } else {
      setFechaDesde('');
      setFechaHasta('');
    }
  };

  // Client-side shift type filter + recalculation of consolidated OP totals & KPIs
  const { turnosFiltrados, porOpFiltrado, resumenFiltrado } = useMemo(() => {
    if (!resultado) {
      return {
        turnosFiltrados: [] as TurnoMaquinaResumen[],
        porOpFiltrado: [] as MaquinaOpConsolidado[],
        resumenFiltrado: {
          totalAnterior: 0,
          mermaAnterior: 0,
          totalActual: 0,
          mermaActual: 0,
          totalProducido: 0,
          totalMerma: 0,
          promedioPorTurno: 0,
          turnosDia: 0,
          turnosNoche: 0,
        },
      };
    }

    const todos: TurnoMaquinaResumen[] = [];
    if (resultado.turnoActualResumen) todos.push(resultado.turnoActualResumen);
    todos.push(...resultado.turnosAnteriores);

    const filtrados = tipoTurnoFiltro
      ? todos.filter((t) => (t.tipo || '').toLowerCase() === tipoTurnoFiltro)
      : todos;

    if (!tipoTurnoFiltro) {
      const turnosDia = filtrados.filter((t) => t.tipo === 'dia').length;
      const turnosNoche = filtrados.filter((t) => t.tipo === 'noche').length;
      return {
        turnosFiltrados: filtrados,
        porOpFiltrado: resultado.porOp,
        resumenFiltrado: {
          totalAnterior: resultado.totalAnterior,
          mermaAnterior: resultado.mermaAnterior,
          totalActual: resultado.totalActual,
          mermaActual: resultado.mermaActual,
          totalProducido: resultado.totalProducido,
          totalMerma: resultado.totalMerma,
          promedioPorTurno: resultado.promedioPorTurno,
          turnosDia,
          turnosNoche,
        },
      };
    }

    // Recompute OP breakdown & totals when Día/Noche filter is active
    const opMap = new Map<string, MaquinaOpConsolidado & { _turnosSet: Set<string> }>();
    const metaPorKey = new Map(resultado.porOp.map((o) => [o.key, o]));

    let totalAnterior = 0;
    let mermaAnterior = 0;
    let totalActual = 0;
    let mermaActual = 0;

    filtrados.forEach((t) => {
      const esAct = t.fuente === 'actual';
      if (esAct) {
        totalActual += t.totalProducido;
        mermaActual += t.totalMerma;
      } else {
        totalAnterior += t.totalProducido;
        mermaAnterior += t.totalMerma;
      }

      t.filas.forEach((f) => {
        const opKey = (f.op || '').trim();
        const sapKey = (f.sap || '').trim();
        const prodKey = (f.producto || '').trim();
        const mapKey = `${opKey.toLowerCase()}||${sapKey.toLowerCase()}||${prodKey.toLowerCase()}`;
        const ref = metaPorKey.get(mapKey);

        if (!opMap.has(mapKey)) {
          opMap.set(mapKey, {
            key: mapKey,
            op: opKey,
            sap: sapKey,
            producto: prodKey || ref?.producto || '',
            cliente: ref?.cliente || '',
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
        if (esAct) {
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
    });

    const totalProducido = totalAnterior + totalActual;
    const totalMerma = mermaAnterior + mermaActual;
    const promedioPorTurno =
      filtrados.length > 0 ? Math.round(totalProducido / filtrados.length) : 0;
    const turnosDia = filtrados.filter((t) => t.tipo === 'dia').length;
    const turnosNoche = filtrados.filter((t) => t.tipo === 'noche').length;

    const porOpList: MaquinaOpConsolidado[] = Array.from(opMap.values())
      .map(({ _turnosSet, ...rest }) => rest)
      .sort((a, b) => b.totalAcumulado - a.totalAcumulado);

    return {
      turnosFiltrados: filtrados,
      porOpFiltrado: porOpList,
      resumenFiltrado: {
        totalAnterior,
        mermaAnterior,
        totalActual,
        mermaActual,
        totalProducido,
        totalMerma,
        promedioPorTurno,
        turnosDia,
        turnosNoche,
      },
    };
  }, [resultado, tipoTurnoFiltro]);

  const handleDescargarPdfMaquina = () => {
    if (!maquinaSeleccionada || !resultado) return;
    try {
      const { nombreArchivo } = exportarPdfReporteMaquina({
        maquina: maquinaSeleccionada,
        turnosFiltrados,
        porOpFiltrado,
        resumenFiltrado,
        fechaDesde,
        fechaHasta,
        tipoTurnoFiltro,
        resultadoOriginal: resultado,
      });
      showToast('ok', `PDF generado: ${nombreArchivo}`);
    } catch (err) {
      console.error('Error al generar PDF de Máquina:', err);
      showToast('error', 'No se pudo generar el reporte PDF de la máquina');
    }
  };

  const meta = Number(maquinaSeleccionada?.meta) || 0;
  const pctMetaProm =
    meta > 0 && turnosFiltrados.length > 0
      ? (resumenFiltrado.promedioPorTurno / meta) * 100
      : 0;
  const turnosEnMeta =
    meta > 0 ? turnosFiltrados.filter((t) => t.totalProducido >= meta).length : 0;

  const distinctOpsCount = useMemo(
    () =>
      new Set(
        porOpFiltrado.map((o) => (o.op || '').trim().toLowerCase()).filter(Boolean)
      ).size,
    [porOpFiltrado]
  );

  const hoyIso = toLocalIsoDate(new Date());
  const tieneFiltroFechas = Boolean(fechaDesde || fechaHasta);

  return (
    <div className="space-y-4 sm:space-y-5">
      {/* Search & Filter Card */}
      <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3.5">
        <div className="flex items-center justify-between gap-2 pb-2 border-b border-[#2D3748] flex-wrap">
          <div className="flex items-center gap-2">
            <Cpu size={18} className="text-[#F2A93B]" />
            <h2 className="m-0 text-sm sm:text-base font-bold text-white font-['Oswald',sans-serif] tracking-wide uppercase">
              Consultar Producción por Máquina
            </h2>
          </div>

          {maquinaSeleccionada && (
            <div className="flex items-center gap-2 flex-wrap">
              {resultado && !cargando && (
                <button
                  type="button"
                  onClick={handleDescargarPdfMaquina}
                  className="flex items-center gap-1.5 py-1 px-3 rounded-lg bg-[#2B6CB0] hover:bg-[#3182CE] text-white text-xs font-bold cursor-pointer shadow-sm transition-colors"
                  title="Descargar reporte completo de esta máquina en PDF"
                >
                  <FileDown size={13} />
                  <span>Reporte PDF</span>
                </button>
              )}
              <button
                type="button"
                disabled={cargando}
                onClick={() => ejecutarConsulta(maquinaSeleccionada, fechaDesde, fechaHasta)}
                className="flex items-center gap-1.5 py-1 px-2.5 rounded-lg border border-[#4A5568] bg-[#232A38] text-[#CBD5E0] hover:text-white text-xs font-semibold cursor-pointer disabled:opacity-50"
              >
                <RefreshCw size={12} className={cargando ? 'animate-spin' : ''} />
                <span>Actualizar</span>
              </button>
            </div>
          )}
        </div>

        {/* Autocomplete Machine Search */}
        <form onSubmit={handleSubmit} className="flex flex-col sm:flex-row gap-2.5">
          <div className="relative flex-1">
            <input
              type="text"
              value={maqInput}
              onChange={(e) => {
                setMaqInput(e.target.value);
                setShowDropdown(true);
              }}
              onFocus={() => setShowDropdown(true)}
              onBlur={() => {
                setTimeout(() => setShowDropdown(false), 220);
              }}
              placeholder="Buscar o seleccionar máquina (nombre, código o área)..."
              autoComplete="off"
              className="w-full py-2.5 px-3.5 pr-16 rounded-xl border border-[#4A5568] text-sm font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none focus:border-[#F2A93B] min-h-[44px]"
            />

            <div className="absolute right-2 top-1/2 -translate-y-1/2 flex items-center gap-0.5">
              {maqInput && (
                <button
                  type="button"
                  onClick={() => {
                    setMaqInput('');
                    setShowDropdown(true);
                  }}
                  className="text-[#A0AEC0] hover:text-white p-1 rounded cursor-pointer"
                  title="Limpiar búsqueda"
                >
                  <X size={15} />
                </button>
              )}
              <button
                type="button"
                onMouseDown={(e) => {
                  e.preventDefault();
                  setShowDropdown((prev) => !prev);
                }}
                className="text-[#A0AEC0] hover:text-white p-1 rounded cursor-pointer"
                title="Mostrar lista de máquinas"
              >
                <ChevronDown
                  size={16}
                  className={`transition-transform ${showDropdown ? 'rotate-180 text-[#F2A93B]' : ''}`}
                />
              </button>
            </div>

            {showDropdown && sugerenciasMaquinas.length > 0 && (
              <div className="absolute z-40 left-0 right-0 mt-1 bg-[#1A202C] border border-[#4A5568] rounded-xl max-h-72 overflow-y-auto shadow-2xl divide-y divide-[#2D3748]">
                <div className="px-3 py-1.5 bg-[#232A38] flex items-center justify-between text-[10px] text-[#A0AEC0]">
                  <span className="font-semibold uppercase tracking-wider text-[#CBD5E0]">
                    Máquinas ({sugerenciasMaquinas.length})
                  </span>
                  <span>Toca para consultar</span>
                </div>
                {sugerenciasMaquinas.map((m) => {
                  const isSel = maquinaSeleccionada?.id === m.id;
                  const aTxt = areaTexto(m.area);
                  return (
                    <button
                      key={m.id}
                      type="button"
                      onMouseDown={(e) => {
                        e.preventDefault();
                        handleSelectMaquina(m);
                      }}
                      className={`w-full text-left py-2.5 px-3 hover:bg-[#232A38] text-xs block text-white transition-colors cursor-pointer ${
                        isSel ? 'bg-[#232A38]/80' : ''
                      }`}
                    >
                      <div className="flex items-center justify-between gap-2 flex-wrap">
                        <div className="flex items-center gap-1.5 flex-wrap">
                          <Cpu
                            size={13}
                            className={m.activa ? 'text-[#48BB78]' : 'text-[#A0AEC0]'}
                          />
                          <span className="font-bold font-['IBM_Plex_Mono',monospace] text-white">
                            {m.nombre}
                          </span>
                          {m.codigo && (
                            <span className="text-[10.5px] px-1.5 py-0.2 rounded bg-[#12161F] border border-[#4A5568] text-[#CBD5E0] font-['IBM_Plex_Mono',monospace]">
                              {m.codigo}
                            </span>
                          )}
                          {m.tipoProducto && (
                            <span
                              className={`text-[10px] font-bold px-1.5 py-0.2 rounded font-['IBM_Plex_Mono',monospace] ${
                                m.tipoProducto === 'pp'
                                  ? 'bg-[#F2A93B]/20 text-[#F6AD55]'
                                  : 'bg-[#63B3ED]/20 text-[#63B3ED]'
                              }`}
                            >
                              {m.tipoProducto.toUpperCase()}
                            </span>
                          )}
                          {m.activa && (
                            <span className="text-[9.5px] font-bold px-1.5 py-0.2 rounded bg-[#38A169]/20 text-[#48BB78]">
                              Activa
                            </span>
                          )}
                        </div>

                        <div className="flex items-center gap-2 text-[10.5px] text-[#A0AEC0]">
                          {aTxt && <span>{aTxt}</span>}
                          {m.meta && m.meta > 0 ? (
                            <span className="text-[#F6AD55] font-['IBM_Plex_Mono',monospace] font-semibold">
                              Meta: {m.meta.toLocaleString('es-PE')}
                            </span>
                          ) : null}
                        </div>
                      </div>
                    </button>
                  );
                })}
              </div>
            )}
          </div>

          <button
            type="submit"
            disabled={cargando}
            className="bg-[#F2A93B] text-[#12161F] py-2.5 px-5 rounded-xl text-xs sm:text-sm font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:brightness-105 active:scale-95 transition-all shadow-sm min-h-[44px] flex items-center justify-center gap-2 disabled:opacity-50"
          >
            <Search size={16} />
            <span>{cargando ? 'Consultando...' : 'Consultar'}</span>
          </button>
        </form>

        {/* Quick-pick machine chips */}
        {maquinasRapidas.length > 0 && (
          <div className="flex items-center gap-1.5 flex-wrap pt-0.5">
            <span className="text-[11px] text-[#A0AEC0] mr-1">Máquinas activas:</span>
            {maquinasRapidas.map((m) => (
              <button
                key={m.id}
                type="button"
                onClick={() => handleSelectMaquina(m)}
                className={`py-0.5 px-2 rounded-lg text-[11px] font-['IBM_Plex_Mono',monospace] font-bold border cursor-pointer transition-colors ${
                  maquinaSeleccionada?.id === m.id
                    ? 'bg-[#F2A93B] text-[#12161F] border-[#F2A93B]'
                    : 'bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:border-[#F2A93B]/60 hover:text-white'
                }`}
              >
                {m.nombre}
              </button>
            ))}
          </div>
        )}

        {/* Date / Range & Shift Filters */}
        <div className="pt-2.5 border-t border-[#2D3748] space-y-2.5">
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
            <div>
              <label className="block text-[10.5px] font-semibold text-[#A0AEC0] uppercase tracking-wider mb-1">
                Desde
              </label>
              <input
                type="date"
                value={fechaDesde}
                max={fechaHasta || undefined}
                onChange={(e) => setFechaDesde(e.target.value)}
                className="w-full py-2 px-3 rounded-xl border border-[#4A5568] text-xs font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none focus:border-[#F2A93B] min-h-[40px]"
              />
            </div>

            <div>
              <label className="block text-[10.5px] font-semibold text-[#A0AEC0] uppercase tracking-wider mb-1">
                Hasta
              </label>
              <input
                type="date"
                value={fechaHasta}
                min={fechaDesde || undefined}
                onChange={(e) => setFechaHasta(e.target.value)}
                className="w-full py-2 px-3 rounded-xl border border-[#4A5568] text-xs font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none focus:border-[#F2A93B] min-h-[40px]"
              />
            </div>

            <div>
              <label className="block text-[10.5px] font-semibold text-[#A0AEC0] uppercase tracking-wider mb-1">
                Turno
              </label>
              <select
                value={tipoTurnoFiltro}
                onChange={(e) => setTipoTurnoFiltro(e.target.value as '' | 'dia' | 'noche')}
                className="w-full py-2 px-3 rounded-xl border border-[#4A5568] text-xs bg-[#232A38] text-white outline-none focus:border-[#F2A93B] min-h-[40px] cursor-pointer"
              >
                <option value="">Todos (Día y Noche)</option>
                <option value="dia">Solo Turno Día</option>
                <option value="noche">Solo Turno Noche</option>
              </select>
            </div>
          </div>

          {/* Quick date range chips */}
          <div className="flex items-center gap-1.5 flex-wrap">
            <span className="text-[11px] text-[#A0AEC0] mr-1">Rango rápido:</span>
            <button
              type="button"
              onClick={() => aplicarPresetFecha('hoy')}
              className={`py-0.5 px-2.5 rounded-lg text-[11px] font-semibold border cursor-pointer transition-colors ${
                fechaDesde === hoyIso && fechaHasta === hoyIso
                  ? 'bg-[#F2A93B]/20 text-[#F6AD55] border-[#F2A93B]'
                  : 'bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:text-white'
              }`}
            >
              Hoy
            </button>
            <button
              type="button"
              onClick={() => aplicarPresetFecha('7d')}
              className="py-0.5 px-2.5 rounded-lg text-[11px] font-semibold border bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:text-white cursor-pointer transition-colors"
            >
              Últimos 7 días
            </button>
            <button
              type="button"
              onClick={() => aplicarPresetFecha('30d')}
              className="py-0.5 px-2.5 rounded-lg text-[11px] font-semibold border bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:text-white cursor-pointer transition-colors"
            >
              Últimos 30 días
            </button>
            <button
              type="button"
              onClick={() => aplicarPresetFecha('mes')}
              className="py-0.5 px-2.5 rounded-lg text-[11px] font-semibold border bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:text-white cursor-pointer transition-colors"
            >
              Este mes
            </button>
            {tieneFiltroFechas && (
              <button
                type="button"
                onClick={() => aplicarPresetFecha('todo')}
                className="py-0.5 px-2.5 rounded-lg text-[11px] font-semibold border bg-[#E53E3E]/15 text-[#FC8181] border-[#E53E3E]/40 hover:bg-[#E53E3E]/25 cursor-pointer transition-colors flex items-center gap-1"
              >
                <X size={11} />
                <span>Todo el historial</span>
              </button>
            )}
          </div>
        </div>
      </div>

      {/* Results Area */}
      {!maquinaSeleccionada ? (
        <div className="border border-dashed border-[#4A5568] rounded-2xl p-8 text-center text-xs sm:text-sm text-[#CBD5E0] bg-[#1A202C]">
          Selecciona o busca una máquina arriba para visualizar sus turnos, fechas, consolidado por OP y descargar su reporte PDF.
        </div>
      ) : cargando ? (
        <div className="border border-dashed border-[#4A5568] rounded-2xl p-10 text-center space-y-2 text-xs sm:text-sm text-[#CBD5E0] bg-[#1A202C]">
          <RefreshCw size={24} className="mx-auto text-[#F2A93B] animate-spin" />
          <p className="m-0">
            Consultando historial de{' '}
            <strong className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace]">
              {maquinaSeleccionada.nombre}
            </strong>
            …
          </p>
        </div>
      ) : resultado ? (
        <>
          {resultado.nubeMensaje && (
            <div className="flex items-center gap-2 text-xs text-[#F6AD55] bg-[#3A2E1B]/50 border border-[#F2A93B]/30 rounded-xl px-3.5 py-2.5">
              <AlertCircle size={15} className="shrink-0" />
              <span>{resultado.nubeMensaje}</span>
            </div>
          )}

          {/* Machine Info Banner */}
          <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-4 shadow-sm flex items-start justify-between gap-3 flex-wrap">
            <div className="min-w-0 flex-1 space-y-1.5">
              <div className="flex items-center gap-2 flex-wrap">
                <span className="font-['IBM_Plex_Mono',monospace] font-bold text-base sm:text-lg text-[#F2A93B] bg-[#232A38] border border-[#4A5568] px-2.5 py-0.5 rounded-lg">
                  {maquinaSeleccionada.nombre}
                </span>
                {maquinaSeleccionada.codigo && (
                  <span className="text-xs font-bold text-[#CBD5E0] bg-[#232A38] px-2 py-0.5 rounded-lg border border-[#4A5568] font-['IBM_Plex_Mono',monospace]">
                    {maquinaSeleccionada.codigo}
                  </span>
                )}
                {maquinaSeleccionada.tipoProducto && (
                  <span
                    className={`text-xs font-bold px-2 py-0.5 rounded-lg font-['IBM_Plex_Mono',monospace] ${
                      maquinaSeleccionada.tipoProducto === 'pp'
                        ? 'bg-[#F2A93B]/20 text-[#F6AD55]'
                        : 'bg-[#63B3ED]/20 text-[#63B3ED]'
                    }`}
                  >
                    {maquinaSeleccionada.tipoProducto.toUpperCase()}
                  </span>
                )}
                {areaTexto(maquinaSeleccionada.area) && (
                  <span className="text-xs text-[#CBD5E0] bg-[#232A38] px-2.5 py-0.5 rounded-lg border border-[#2D3748]">
                    {areaTexto(maquinaSeleccionada.area)}
                  </span>
                )}
              </div>

              <div className="text-xs text-[#A0AEC0] flex items-center gap-2.5 flex-wrap">
                <span>
                  Período:{' '}
                  <strong className="text-white">
                    {fechaDesde && fechaHasta
                      ? `${formatearFechaTurno(fechaDesde)} al ${formatearFechaTurno(fechaHasta)}`
                      : fechaDesde
                      ? `Desde ${formatearFechaTurno(fechaDesde)}`
                      : fechaHasta
                      ? `Hasta ${formatearFechaTurno(fechaHasta)}`
                      : 'Todo el historial'}
                  </strong>
                </span>
                {tipoTurnoFiltro && (
                  <span className="text-[#63B3ED] font-semibold">
                    · Solo Turno {tipoTurnoFiltro === 'dia' ? 'Día' : 'Noche'}
                  </span>
                )}
              </div>
            </div>

            {meta > 0 && (
              <div className="bg-[#12161F] border border-[#F2A93B]/40 rounded-xl px-3 py-2 text-right shrink-0">
                <div className="text-[10px] uppercase tracking-wider text-[#A0AEC0] font-semibold flex items-center justify-end gap-1">
                  <Target size={11} className="text-[#F2A93B]" />
                  <span>Meta por Turno</span>
                </div>
                <div className="font-['IBM_Plex_Mono',monospace] text-sm sm:text-base font-bold text-[#F6AD55] tabular-nums">
                  {meta.toLocaleString('es-PE')} u
                </div>
              </div>
            )}
          </div>

          {/* 4 KPI Summary Cards */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
            {/* Total Producido */}
            <div className="bg-[#1A202C] border border-[#F2A93B]/40 rounded-2xl p-4 shadow-sm space-y-1.5">
              <div className="text-[11px] uppercase tracking-wider text-[#A0AEC0] font-semibold">
                Total Producido
              </div>
              <div className="font-['IBM_Plex_Mono',monospace] text-2xl font-bold text-[#F2A93B] tabular-nums">
                {resumenFiltrado.totalProducido.toLocaleString('es-PE')} u
              </div>
              <div className="text-[11px] text-[#CBD5E0] font-['IBM_Plex_Mono',monospace] flex items-center gap-2 flex-wrap">
                <span>Ant: {resumenFiltrado.totalAnterior.toLocaleString('es-PE')} u</span>
                {resumenFiltrado.totalActual > 0 && (
                  <span className="text-[#48BB78] font-semibold">
                    · Act: +{resumenFiltrado.totalActual.toLocaleString('es-PE')} u
                  </span>
                )}
                {resumenFiltrado.totalMerma > 0 && (
                  <span className="text-[#FC8181]">
                    · -{resumenFiltrado.totalMerma.toLocaleString('es-PE')} merma
                  </span>
                )}
              </div>
            </div>

            {/* Turnos Registrados */}
            <div className="bg-[#1A202C] border border-[#63B3ED]/30 rounded-2xl p-4 shadow-sm space-y-1.5">
              <div className="text-[11px] uppercase tracking-wider text-[#A0AEC0] font-semibold">
                Turnos Registrados
              </div>
              <div className="font-['IBM_Plex_Mono',monospace] text-2xl font-bold text-[#63B3ED] tabular-nums">
                {turnosFiltrados.length}
              </div>
              <div className="text-[11px] text-[#CBD5E0] font-['IBM_Plex_Mono',monospace] flex items-center gap-2.5">
                <span className="inline-flex items-center gap-1 text-[#F6AD55]">
                  <Sun size={11} /> {resumenFiltrado.turnosDia} Día
                </span>
                <span className="inline-flex items-center gap-1 text-[#93C5FD]">
                  <Moon size={11} /> {resumenFiltrado.turnosNoche} Noche
                </span>
              </div>
            </div>

            {/* Promedio por Turno */}
            <div className="bg-[#1A202C] border border-[#38A169]/35 rounded-2xl p-4 shadow-sm space-y-1.5">
              <div className="flex items-center justify-between gap-2">
                <span className="text-[11px] uppercase tracking-wider text-[#A0AEC0] font-semibold">
                  Promedio / Turno
                </span>
                {meta > 0 && turnosFiltrados.length > 0 && (
                  <span
                    className={`text-xs font-bold font-['IBM_Plex_Mono',monospace] ${
                      pctMetaProm >= 100
                        ? 'text-[#48BB78]'
                        : pctMetaProm >= 80
                        ? 'text-[#F6AD55]'
                        : 'text-[#FC8181]'
                    }`}
                  >
                    {pctMetaProm.toFixed(1)}%
                  </span>
                )}
              </div>
              <div className="font-['IBM_Plex_Mono',monospace] text-2xl font-bold text-[#48BB78] tabular-nums">
                {resumenFiltrado.promedioPorTurno.toLocaleString('es-PE')} u
              </div>
              {meta > 0 && turnosFiltrados.length > 0 ? (
                <div className="space-y-1 pt-0.5">
                  <div className="w-full bg-[#12161F] rounded-full h-1.5 overflow-hidden border border-[#2D3748]">
                    <div
                      className={`h-full rounded-full ${
                        pctMetaProm >= 100
                          ? 'bg-[#38A169]'
                          : pctMetaProm >= 80
                          ? 'bg-[#F2A93B]'
                          : 'bg-[#E53E3E]'
                      }`}
                      style={{ width: `${Math.min(100, Math.max(0, pctMetaProm))}%` }}
                    />
                  </div>
                  <div className="text-[10.5px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace]">
                    {turnosEnMeta} de {turnosFiltrados.length} turno(s) cumplieron meta
                  </div>
                </div>
              ) : (
                <div className="text-[11px] text-[#A0AEC0]">
                  {meta > 0 ? 'Sin turnos para promediar' : 'Máquina sin meta configurada'}
                </div>
              )}
            </div>

            {/* OPs Trabajadas */}
            <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-4 shadow-sm space-y-1.5">
              <div className="text-[11px] uppercase tracking-wider text-[#A0AEC0] font-semibold">
                OPs Trabajadas
              </div>
              <div className="font-['IBM_Plex_Mono',monospace] text-2xl font-bold text-white tabular-nums">
                {distinctOpsCount || porOpFiltrado.length}
              </div>
              <div className="text-[11px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace]">
                {porOpFiltrado.length} producto(s) / registro(s)
              </div>
            </div>
          </div>

          {turnosFiltrados.length === 0 ? (
            <div className="border border-dashed border-[#4A5568] rounded-2xl p-8 text-center text-xs sm:text-sm text-[#CBD5E0] bg-[#1A202C]">
              No se encontraron registros para{' '}
              <strong className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace]">
                {maquinaSeleccionada.nombre}
              </strong>{' '}
              con los filtros aplicados.
            </div>
          ) : (
            <>
              {/* Section 1: Consolidated by OP / Product */}
              <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl overflow-hidden shadow-sm">
                <div className="px-4 py-3 bg-[#12161F] border-b border-[#2D3748] flex items-center justify-between gap-2">
                  <div className="flex items-center gap-2">
                    <Layers size={16} className="text-[#F2A93B]" />
                    <h3 className="m-0 text-xs sm:text-sm font-bold text-white uppercase font-['Oswald',sans-serif] tracking-wide">
                      Consolidado por OP y Producto ({porOpFiltrado.length})
                    </h3>
                  </div>
                </div>

                <div className="divide-y divide-[#2D3748]">
                  {porOpFiltrado.map((item) => (
                    <div
                      key={item.key}
                      className="px-4 py-2.5 flex items-center justify-between gap-3 text-xs sm:text-sm hover:bg-[#232A38]/50"
                    >
                      <div className="min-w-0 flex-1">
                        <div className="flex items-center gap-2 flex-wrap">
                          <Package size={14} className="text-[#F2A93B] shrink-0" />
                          <span className="font-['IBM_Plex_Mono',monospace] font-bold text-[#F6AD55] bg-[#232A38] border border-[#4A5568] px-1.5 py-0.2 rounded">
                            {item.op ? `OP ${item.op}` : 'Sin OP'}
                          </span>
                          {item.sap && (
                            <span className="text-[10.5px] px-1.5 py-0.2 rounded bg-[#12161F] border border-[#2D3748] text-[#CBD5E0] font-['IBM_Plex_Mono',monospace]">
                              SAP: {item.sap}
                            </span>
                          )}
                          {item.cliente && (
                            <span className="text-[11px] text-[#A0AEC0] truncate max-w-[200px]">
                              · {item.cliente}
                            </span>
                          )}
                          <span className="text-[11px] text-[#A0AEC0]">
                            · {item.turnosCount} turno(s)
                          </span>
                        </div>

                        {item.producto && (
                          <div className="text-xs text-[#CBD5E0] truncate mt-0.5 pl-5">
                            {item.producto}
                          </div>
                        )}

                        {(item.totalActual > 0 || item.mermaAcumulada > 0) && (
                          <div className="text-[11px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace] mt-0.5 pl-5 flex items-center gap-2 flex-wrap">
                            <span>Guardado: {item.totalAnterior.toLocaleString('es-PE')} u</span>
                            {item.totalActual > 0 && (
                              <span className="text-[#48BB78]">
                                · Actual: +{item.totalActual.toLocaleString('es-PE')} u
                              </span>
                            )}
                            {item.mermaAcumulada > 0 && (
                              <span className="text-[#FC8181]">
                                · -{item.mermaAcumulada.toLocaleString('es-PE')} merma
                              </span>
                            )}
                          </div>
                        )}
                      </div>

                      <div className="text-right shrink-0 font-['IBM_Plex_Mono',monospace]">
                        <span className="font-bold text-sm sm:text-base text-[#63B3ED] tabular-nums">
                          {item.totalAcumulado.toLocaleString('es-PE')} u
                        </span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Section 2: Chronological Breakdown by Dates & Shifts */}
              <div className="space-y-2.5">
                <div className="flex items-center gap-2 px-1">
                  <TrendingUp size={16} className="text-[#F2A93B]" />
                  <h3 className="m-0 text-xs sm:text-sm font-bold text-white uppercase font-['Oswald',sans-serif] tracking-wide">
                    Detalle por Turnos y Fechas ({turnosFiltrados.length})
                  </h3>
                </div>

                {turnosFiltrados.map((t) => {
                  const esNoche = t.tipo === 'noche';
                  const esActual = t.fuente === 'actual';
                  const etiquetaTurnoTxt = esNoche ? 'Turno Noche' : t.tipo === 'dia' ? 'Turno Día' : 'Turno';
                  const pctMeta = meta > 0 ? (t.totalProducido / meta) * 100 : 0;

                  return (
                    <div
                      key={t.key}
                      className={`bg-[#1A202C] border rounded-2xl overflow-hidden shadow-sm ${
                        esActual ? 'border-[#38A169]/60' : 'border-[#2D3748]'
                      }`}
                    >
                      <div className="px-3.5 py-2.5 bg-[#232A38] border-b border-[#2D3748] flex items-center justify-between gap-2 flex-wrap">
                        <div className="flex items-center gap-2 flex-wrap">
                          <Calendar size={14} className="text-[#A0AEC0]" />
                          <span className="font-['IBM_Plex_Mono',monospace] font-bold text-xs sm:text-sm text-white">
                            {formatearFechaTurno(t.fecha)}
                          </span>
                          <span
                            className={`inline-flex items-center gap-1 text-[10.5px] font-bold px-2 py-0.5 rounded-md ${
                              esNoche
                                ? 'bg-[#63B3ED]/20 text-[#63B3ED]'
                                : 'bg-[#F2A93B]/20 text-[#F6AD55]'
                            }`}
                          >
                            {esNoche ? <Moon size={11} /> : <Sun size={11} />}
                            {etiquetaTurnoTxt}
                          </span>
                          {esActual && (
                            <span className="text-[10px] font-bold px-2 py-0.5 rounded-md bg-[#38A169]/20 text-[#48BB78] border border-[#38A169]/40">
                              TURNO ACTUAL (EN CURSO)
                            </span>
                          )}
                          {t.fuente === 'local' && (
                            <span className="text-[10px] px-1.5 py-0.5 rounded bg-[#2D3748] text-[#A0AEC0]">
                              Local
                            </span>
                          )}
                        </div>

                        <div className="flex items-center gap-2 font-['IBM_Plex_Mono',monospace] text-xs flex-wrap">
                          {meta > 0 && (
                            <span
                              className={`text-[10px] font-bold px-2 py-0.5 rounded border ${
                                pctMeta >= 100
                                  ? 'bg-[#38A169]/20 text-[#48BB78] border-[#38A169]/40'
                                  : pctMeta >= 80
                                  ? 'bg-[#F2A93B]/20 text-[#F6AD55] border-[#F2A93B]/40'
                                  : 'bg-[#E53E3E]/20 text-[#FC8181] border-[#E53E3E]/40'
                              }`}
                            >
                              {pctMeta.toFixed(0)}% meta
                            </span>
                          )}
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-[#F2A93B]/15 border border-[#F2A93B]/30 text-[#F6AD55] font-bold tabular-nums">
                            +{t.totalProducido.toLocaleString('es-PE')} u
                          </span>
                          {t.totalMerma > 0 && (
                            <span className="text-[#FC8181] text-[11px] tabular-nums">
                              (-{t.totalMerma.toLocaleString('es-PE')} merma)
                            </span>
                          )}
                        </div>
                      </div>

                      <div className="divide-y divide-[#2D3748]/70">
                        {t.filas.map((f, idx) => (
                          <div
                            key={idx}
                            className="px-3.5 py-2.5 flex items-center justify-between gap-3 text-xs sm:text-sm hover:bg-[#232A38]/40"
                          >
                            <div className="min-w-0 flex-1">
                              <div className="flex items-center gap-1.5 flex-wrap">
                                <span
                                  className={`font-['IBM_Plex_Mono',monospace] font-bold px-1.5 py-0.2 rounded text-[11.5px] ${
                                    f.op
                                      ? 'bg-[#232A38] border border-[#4A5568] text-[#F6AD55]'
                                      : 'bg-[#232A38] text-[#A0AEC0]'
                                  }`}
                                >
                                  {f.op ? `OP ${f.op}` : 'Sin OP'}
                                </span>
                                {f.sap && (
                                  <span className="text-[10.5px] font-['IBM_Plex_Mono',monospace] text-[#A0AEC0]">
                                    SAP: {f.sap}
                                  </span>
                                )}
                              </div>
                              {f.producto && (
                                <div className="text-xs text-[#CBD5E0] truncate mt-0.5">
                                  {f.producto}
                                </div>
                              )}
                              {f.observaciones && (
                                <div className="text-[11px] text-[#A0AEC0] truncate mt-0.5 italic">
                                  {f.observaciones}
                                </div>
                              )}
                            </div>

                            <div className="text-right shrink-0 font-['IBM_Plex_Mono',monospace]">
                              <div
                                className={`font-bold tabular-nums ${
                                  f.total > 0 ? 'text-[#48BB78]' : 'text-[#718096]'
                                }`}
                              >
                                {f.total.toLocaleString('es-PE')} u
                              </div>
                              {f.merma > 0 && (
                                <div className="text-[10.5px] text-[#FC8181] tabular-nums">
                                  -{f.merma.toLocaleString('es-PE')} merma
                                </div>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>
            </>
          )}
        </>
      ) : null}
    </div>
  );
};
