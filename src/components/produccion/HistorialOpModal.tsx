import React, { useState, useEffect, useCallback, useMemo } from 'react';
import {
  History,
  X,
  Sun,
  Moon,
  RefreshCw,
  Calendar,
  Cpu,
  AlertCircle,
  Package,
} from 'lucide-react';
import { useApp } from '../../context/AppContext';
import { buscarEnCatalogo } from '../../services/excel';
import { formatearFechaTurno } from '../../services/reports';
import {
  consultarHistorialOp,
  ConsultaOpResultado,
  calcularResumenModelosOp,
  filtrarResultadoOpPorModelo,
} from '../../services/opHistory';

interface HistorialOpModalProps {
  op: string;
  initialModeloIdx?: number;
  onClose: () => void;
}

export const HistorialOpModal: React.FC<HistorialOpModalProps> = ({
  op,
  initialModeloIdx,
  onClose,
}) => {
  const { catalogo, historial, maquinas, turno, metaTurno, session } = useApp();

  const [cargando, setCargando] = useState(true);
  const [resultado, setResultado] = useState<ConsultaOpResultado | null>(null);
  const [modeloSeleccionadoIdx, setModeloSeleccionadoIdx] = useState<number>(
    typeof initialModeloIdx === 'number' ? initialModeloIdx : -1
  );

  useEffect(() => {
    setModeloSeleccionadoIdx(typeof initialModeloIdx === 'number' ? initialModeloIdx : -1);
  }, [op, initialModeloIdx]);

  const catalogMatch = useMemo(() => buscarEnCatalogo(catalogo, op), [catalogo, op]);

  const cargarHistorialOp = useCallback(async () => {
    setCargando(true);
    const res = await consultarHistorialOp({
      op,
      catalogo,
      historial,
      maquinas,
      turno,
      metaTurno,
      session,
    });
    setResultado(res);
    setCargando(false);
  }, [op, catalogo, historial, maquinas, turno, metaTurno, session]);

  useEffect(() => {
    cargarHistorialOp();
  }, [cargarHistorialOp]);

  const modelosDisponibles = useMemo(() => catalogMatch?.modelos || [], [catalogMatch]);
  const tieneMultiModelos = modelosDisponibles.length > 1;
  const modeloActivo = useMemo(() => {
    if (!tieneMultiModelos || modeloSeleccionadoIdx < 0) return null;
    return modelosDisponibles[modeloSeleccionadoIdx] || null;
  }, [tieneMultiModelos, modeloSeleccionadoIdx, modelosDisponibles]);

  const resultadoVista = useMemo(() => {
    if (!resultado) return null;
    return modeloActivo ? filtrarResultadoOpPorModelo(resultado, modeloActivo) : resultado;
  }, [resultado, modeloActivo]);

  const resumenModelos = useMemo(() => {
    if (!resultado || !tieneMultiModelos) return [];
    return calcularResumenModelosOp(resultado, catalogMatch);
  }, [resultado, tieneMultiModelos, catalogMatch]);

  const turnosAnteriores = resultadoVista?.turnosAnteriores || [];
  const totales = resultadoVista?.totales || {
    antPP: 0,
    mermaAntPP: 0,
    actualPP: 0,
    mermaActualPP: 0,
    acumPP: 0,
    mermaAcumPP: 0,
    antPT: 0,
    mermaAntPT: 0,
    actualPT: 0,
    mermaActualPT: 0,
    acumPT: 0,
    mermaAcumPT: 0,
    antSinTipo: 0,
    mermaAntSinTipo: 0,
    actualSinTipo: 0,
    mermaActualSinTipo: 0,
    acumSinTipo: 0,
    mermaAcumSinTipo: 0,
  };

  const ppPorMaquina = useMemo(
    () => (resultadoVista?.porMaquina || []).filter((m) => m.tipoProducto === 'pp'),
    [resultadoVista]
  );

  const cantidadOrden = modeloActivo
    ? Number(modeloActivo.cantidadTotal || 0)
    : Number(catalogMatch?.cantidadTotal || 0);
  const pctPT = cantidadOrden > 0 ? (totales.acumPT / cantidadOrden) * 100 : 0;
  const pendientePT = cantidadOrden > 0 ? cantidadOrden - totales.acumPT : 0;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-black/75 backdrop-blur-xs"
      onClick={onClose}
    >
      <div
        className="bg-[#1A202C] border border-[#2D3748] rounded-2xl max-w-2xl w-full max-h-[88vh] flex flex-col shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-150"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="px-4 py-3.5 bg-[#12161F] border-b border-[#2D3748] flex items-start justify-between gap-3 shrink-0">
          <div className="min-w-0 flex-1">
            <div className="flex items-center gap-2 flex-wrap">
              <div className="w-7 h-7 rounded-lg bg-[#F2A93B]/20 text-[#F2A93B] flex items-center justify-center shrink-0">
                <History size={16} />
              </div>
              <h3 className="m-0 text-sm sm:text-base font-bold text-white font-['Oswald',sans-serif] uppercase tracking-wide">
                Historial de Producción — OP{' '}
                <span className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace]">{op}</span>
              </h3>
            </div>
            {catalogMatch && (
              <div className="text-xs text-[#CBD5E0] mt-1.5 flex items-center gap-2 flex-wrap">
                <span className="truncate max-w-full">
                  {modeloActivo
                    ? modeloActivo.descripcion || modeloActivo.modelo
                    : catalogMatch.descripcionBase || catalogMatch.descripcion}
                </span>
                {catalogMatch.cliente && (
                  <span className="text-[#A0AEC0]">· {catalogMatch.cliente}</span>
                )}
                {cantidadOrden > 0 && (
                  <span className="inline-flex items-center gap-1 text-[11px] font-['IBM_Plex_Mono',monospace] font-bold text-[#F6AD55] bg-[#3A2E1B]/60 px-2 py-0.5 rounded border border-[#F2A93B]/30">
                    <Package size={11} />
                    {modeloActivo ? 'Modelo' : 'Orden'}: {cantidadOrden.toLocaleString('es-PE')} u
                  </span>
                )}
              </div>
            )}

            {tieneMultiModelos && (
              <div className="mt-2.5 flex items-center gap-1.5 flex-wrap">
                <span className="text-[10px] uppercase tracking-wider text-[#A0AEC0] font-semibold mr-0.5">
                  Modelo:
                </span>
                <button
                  type="button"
                  onClick={() => setModeloSeleccionadoIdx(-1)}
                  className={`text-[10.5px] px-2 py-0.5 rounded-md border cursor-pointer transition-colors ${
                    modeloSeleccionadoIdx === -1
                      ? 'bg-[#F2A93B] text-[#12161F] border-[#F2A93B] font-bold'
                      : 'bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:text-white'
                  }`}
                >
                  Todos ({modelosDisponibles.length})
                </button>
                {modelosDisponibles.map((mo, idx) => {
                  const active = modeloSeleccionadoIdx === idx;
                  return (
                    <button
                      key={idx}
                      type="button"
                      onClick={() => setModeloSeleccionadoIdx(idx)}
                      className={`text-[10.5px] px-2 py-0.5 rounded-md border cursor-pointer transition-colors ${
                        active
                          ? 'bg-[#F2A93B] text-[#12161F] border-[#F2A93B] font-bold'
                          : 'bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:text-white'
                      }`}
                    >
                      {mo.modelo || mo.descripcion || `Modelo ${idx + 1}`}
                      {mo.cantidadTotal && Number(mo.cantidadTotal) > 0 ? (
                        <span
                          className={`ml-1 font-['IBM_Plex_Mono',monospace] ${
                            active ? 'text-[#12161F]' : 'text-[#F6AD55]'
                          }`}
                        >
                          ({Number(mo.cantidadTotal).toLocaleString('es-PE')})
                        </span>
                      ) : null}
                    </button>
                  );
                })}
              </div>
            )}
          </div>

          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg text-[#A0AEC0] hover:text-white hover:bg-[#232A38] transition-colors cursor-pointer shrink-0"
            aria-label="Cerrar"
          >
            <X size={18} />
          </button>
        </div>

        {/* Summary KPI Strip: PT summed + PP separated by machine in small cards */}
        <div className="px-4 py-3 bg-[#141923] border-b border-[#2D3748] space-y-2.5 shrink-0 max-h-[38vh] overflow-y-auto">
          {tieneMultiModelos && !modeloActivo && resumenModelos.length > 0 && !cargando && (
            <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl px-3 py-2 space-y-1.5">
              <div className="text-[10px] uppercase tracking-wider text-[#A0AEC0] font-semibold">
                Avance PT por Modelo (selecciona para filtrar)
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-1.5">
                {resumenModelos.map((rm) => (
                  <button
                    key={rm.idx}
                    type="button"
                    onClick={() => setModeloSeleccionadoIdx(rm.idx)}
                    className="text-left bg-[#12161F] border border-[#2D3748] hover:border-[#F2A93B]/60 rounded-lg px-2.5 py-1.5 flex items-center justify-between gap-2 cursor-pointer transition-colors"
                  >
                    <span className="text-[11px] font-semibold text-white truncate">
                      {rm.modeloCorto || rm.nombre}
                    </span>
                    <span className="text-[10.5px] font-['IBM_Plex_Mono',monospace] text-[#63B3ED] font-bold shrink-0">
                      {rm.ptAcumulado.toLocaleString('es-PE')}
                      {rm.cantidadOrden > 0 ? ` / ${rm.cantidadOrden.toLocaleString('es-PE')} u` : ' u'}
                    </span>
                  </button>
                ))}
              </div>
            </div>
          )}

          <div
            className={`grid grid-cols-1 gap-2.5 ${
              totales.acumSinTipo > 0 ? 'sm:grid-cols-2' : ''
            }`}
          >
            {/* PT - Producto Terminado */}
            <div className="bg-[#1A202C] border border-[#63B3ED]/30 rounded-xl px-3 py-2.5 space-y-1">
              <div className="flex items-center justify-between gap-2">
                <div className="flex items-center gap-1.5">
                  <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-[#63B3ED]/20 text-[#63B3ED] font-['IBM_Plex_Mono',monospace]">
                    PT
                  </span>
                  <span className="text-[10.5px] uppercase tracking-wider text-[#CBD5E0] font-semibold">
                    Producto Terminado
                  </span>
                </div>
                {cantidadOrden > 0 && (
                  <span
                    className={`text-[11px] font-bold font-['IBM_Plex_Mono',monospace] ${
                      pctPT >= 100 ? 'text-[#48BB78]' : 'text-[#93C5FD]'
                    }`}
                  >
                    {pctPT.toFixed(1)}%
                  </span>
                )}
              </div>

              <div className="font-['IBM_Plex_Mono',monospace] text-base sm:text-lg font-bold text-[#63B3ED] tabular-nums">
                {totales.acumPT.toLocaleString('es-PE')} u
              </div>

              <div className="text-[10.5px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace] flex items-center gap-2 flex-wrap">
                <span>Ant: {totales.antPT.toLocaleString('es-PE')} u</span>
                {totales.actualPT > 0 && (
                  <span className="text-[#48BB78]">
                    · Actual: +{totales.actualPT.toLocaleString('es-PE')} u
                  </span>
                )}
                {totales.mermaAcumPT > 0 && (
                  <span className="text-[#FC8181]">
                    · -{totales.mermaAcumPT.toLocaleString('es-PE')} merma
                  </span>
                )}
              </div>

              {cantidadOrden > 0 && (
                <div className="pt-1 space-y-1">
                  <div className="w-full bg-[#12161F] rounded-full h-1.5 overflow-hidden border border-[#2D3748]">
                    <div
                      className={`h-full rounded-full transition-all duration-300 ${
                        pctPT >= 100 ? 'bg-[#38A169]' : 'bg-[#63B3ED]'
                      }`}
                      style={{ width: `${Math.min(100, Math.max(0, pctPT))}%` }}
                    />
                  </div>
                  <div className="text-[10px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace]">
                    {pendientePT > 0
                      ? `Faltan ${pendientePT.toLocaleString('es-PE')} u de la orden`
                      : pendientePT < 0
                      ? `+${Math.abs(pendientePT).toLocaleString('es-PE')} u sobre la orden`
                      : 'Orden completa en PT'}
                  </div>
                </div>
              )}
            </div>

            {/* Sin clasificar */}
            {totales.acumSinTipo > 0 && (
              <div className="bg-[#1A202C] border border-[#4A5568] rounded-xl px-3 py-2.5 space-y-1">
                <div className="flex items-center gap-1.5">
                  <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-[#4A5568]/40 text-[#CBD5E0] font-['IBM_Plex_Mono',monospace]">
                    S/T
                  </span>
                  <span className="text-[10.5px] uppercase tracking-wider text-[#CBD5E0] font-semibold">
                    Sin tipo
                  </span>
                </div>
                <div className="font-['IBM_Plex_Mono',monospace] text-base sm:text-lg font-bold text-white tabular-nums">
                  {totales.acumSinTipo.toLocaleString('es-PE')} u
                </div>
                <div className="text-[10px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace]">
                  Ant: {totales.antSinTipo.toLocaleString('es-PE')} u
                  {totales.actualSinTipo > 0 && (
                    <span className="text-[#48BB78] ml-1.5">
                      · Actual: +{totales.actualSinTipo.toLocaleString('es-PE')} u
                    </span>
                  )}
                </div>
              </div>
            )}
          </div>

          {/* PP - Producto en Proceso separado por máquinas en cuadros pequeños */}
          <div className="bg-[#1A202C] border border-[#F2A93B]/30 rounded-xl px-3 py-2.5 space-y-2">
            <div className="flex items-center justify-between gap-2 flex-wrap">
              <div className="flex items-center gap-1.5">
                <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-[#F2A93B]/20 text-[#F6AD55] font-['IBM_Plex_Mono',monospace]">
                  PP
                </span>
                <span className="text-[10.5px] uppercase tracking-wider text-[#CBD5E0] font-semibold">
                  Producto en Proceso por Máquina
                </span>
              </div>
              <span className="text-[10px] text-[#A0AEC0]">
                {ppPorMaquina.length > 0
                  ? `${ppPorMaquina.length} máquina(s) independientes`
                  : 'Sin registros PP'}
              </span>
            </div>

            {ppPorMaquina.length === 0 ? (
              <div className="text-[11px] text-[#718096]">
                No hay producción registrada en PP para esta OP.
              </div>
            ) : (
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                {ppPorMaquina.map((mq) => (
                  <div
                    key={mq.key}
                    className="bg-[#12161F] border border-[#F2A93B]/30 rounded-lg px-2.5 py-1.5 space-y-0.5"
                  >
                    <div className="flex items-center justify-between gap-1">
                      <span className="text-[10.5px] font-bold text-white truncate font-['IBM_Plex_Mono',monospace]">
                        {mq.maquina}
                      </span>
                      {mq.codigoMaquina && (
                        <span className="text-[9px] px-1 rounded bg-[#232A38] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace] shrink-0">
                          {mq.codigoMaquina}
                        </span>
                      )}
                    </div>
                    <div className="font-['IBM_Plex_Mono',monospace] text-sm font-bold text-[#F2A93B] tabular-nums">
                      {mq.totalAcumulado.toLocaleString('es-PE')} u
                    </div>
                    <div className="text-[9.5px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace] flex items-center gap-1 flex-wrap">
                      <span>Ant: {mq.totalAnterior.toLocaleString('es-PE')}</span>
                      {mq.totalActual > 0 && (
                        <span className="text-[#48BB78]">
                          · Act: +{mq.totalActual.toLocaleString('es-PE')}
                        </span>
                      )}
                      {mq.mermaAcumulada > 0 && (
                        <span className="text-[#FC8181]">
                          · -{mq.mermaAcumulada.toLocaleString('es-PE')} m
                        </span>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          {resultado?.nubeMensaje && !cargando && (
            <div className="flex items-center gap-2 text-[11px] text-[#F6AD55] bg-[#3A2E1B]/50 border border-[#F2A93B]/30 rounded-lg px-2.5 py-1.5">
              <AlertCircle size={13} className="shrink-0" />
              <span>{resultado.nubeMensaje}</span>
            </div>
          )}
        </div>

        {/* Body: List of shifts */}
        <div className="flex-1 overflow-y-auto p-4 space-y-3">
          {cargando ? (
            <div className="py-10 text-center space-y-2 text-xs text-[#CBD5E0]">
              <RefreshCw size={22} className="mx-auto text-[#F2A93B] animate-spin" />
              <p className="m-0">Buscando producción registrada de la OP {op}…</p>
            </div>
          ) : turnosAnteriores.length === 0 ? (
            <div className="border border-dashed border-[#4A5568] rounded-xl p-6 text-center text-xs sm:text-sm text-[#CBD5E0] bg-[#141923]">
              No se encontraron registros en turnos anteriores para la OP{' '}
              <strong className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace]">{op}</strong>.
              {!resultado?.nubeDisponible && !session && (
                <p className="text-[11px] text-[#A0AEC0] mt-1.5 mb-0">
                  Tip: Inicia sesión en la pestaña Historial para consultar también todos los turnos guardados en la nube.
                </p>
              )}
            </div>
          ) : (
            turnosAnteriores.map((t) => {
              const esNoche = t.tipo === 'noche';
              const etiquetaTurnoTxt = esNoche ? 'Turno Noche' : t.tipo === 'dia' ? 'Turno Día' : 'Turno';

              return (
                <div
                  key={t.key}
                  className="bg-[#141923] border border-[#2D3748] rounded-xl overflow-hidden shadow-xs"
                >
                  <div className="px-3 py-2 bg-[#232A38] border-b border-[#2D3748] flex items-center justify-between gap-2 flex-wrap">
                    <div className="flex items-center gap-2">
                      <Calendar size={13} className="text-[#A0AEC0]" />
                      <span className="font-['IBM_Plex_Mono',monospace] font-bold text-xs sm:text-sm text-white">
                        {formatearFechaTurno(t.fecha)}
                      </span>
                      <span
                        className={`inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-md ${
                          esNoche
                            ? 'bg-[#63B3ED]/20 text-[#63B3ED]'
                            : 'bg-[#F2A93B]/20 text-[#F6AD55]'
                        }`}
                      >
                        {esNoche ? <Moon size={10} /> : <Sun size={10} />}
                        {etiquetaTurnoTxt}
                      </span>
                      {t.fuente === 'local' && (
                        <span className="text-[9.5px] px-1.5 py-0.5 rounded bg-[#2D3748] text-[#A0AEC0]">
                          Local
                        </span>
                      )}
                    </div>

                    <div className="flex items-center gap-2 font-['IBM_Plex_Mono',monospace] text-xs flex-wrap">
                      {(t.totalPT > 0 || t.filas.some((f) => f.tipoProducto === 'pt')) && (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-[#63B3ED]/15 border border-[#63B3ED]/30 text-[#63B3ED] font-bold tabular-nums">
                          PT: +{t.totalPT.toLocaleString('es-PE')} u
                        </span>
                      )}
                      {t.ppPorMaquina.map((pm, pIdx) => (
                        <span
                          key={pIdx}
                          className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-[#F2A93B]/15 border border-[#F2A93B]/30 text-[#F6AD55] font-bold tabular-nums"
                        >
                          PP ({pm.maquina}): +{pm.total.toLocaleString('es-PE')} u
                        </span>
                      ))}
                      {t.totalSinTipo > 0 && (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-[#2D3748] text-[#CBD5E0] font-bold tabular-nums">
                          +{t.totalSinTipo.toLocaleString('es-PE')} u
                        </span>
                      )}
                    </div>
                  </div>

                  <div className="divide-y divide-[#2D3748]/70">
                    {t.filas.map((f, idx) => (
                      <div
                        key={idx}
                        className="px-3 py-2 flex items-center justify-between gap-3 text-xs hover:bg-[#1A202C]/60"
                      >
                        <div className="min-w-0 flex-1">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <Cpu size={12} className="text-[#A0AEC0] shrink-0" />
                            <span className="font-['IBM_Plex_Mono',monospace] font-bold text-white">
                              {f.maquina}
                            </span>
                            {f.codigoMaquina && (
                              <span className="text-[10px] px-1.5 py-0.2 rounded bg-[#232A38] border border-[#4A5568] text-[#CBD5E0] font-['IBM_Plex_Mono',monospace]">
                                {f.codigoMaquina}
                              </span>
                            )}
                            {f.tipoProducto && (
                              <span
                                className={`text-[9.5px] font-bold px-1.5 py-0.2 rounded font-['IBM_Plex_Mono',monospace] ${
                                  f.tipoProducto === 'pp'
                                    ? 'bg-[#F2A93B]/20 text-[#F6AD55]'
                                    : 'bg-[#63B3ED]/20 text-[#63B3ED]'
                                }`}
                              >
                                {f.tipoProducto.toUpperCase()}
                              </span>
                            )}
                          </div>
                          {f.producto && (
                            <div className="text-[11px] text-[#CBD5E0] truncate mt-0.5 pl-4">
                              {f.producto}
                            </div>
                          )}
                          {f.observaciones && (
                            <div className="text-[10.5px] text-[#A0AEC0] truncate mt-0.5 pl-4 italic">
                              {f.observaciones}
                            </div>
                          )}
                        </div>

                        <div className="text-right shrink-0 font-['IBM_Plex_Mono',monospace]">
                          <div
                            className={`font-bold tabular-nums ${
                              f.total > 0
                                ? f.tipoProducto === 'pp'
                                  ? 'text-[#F6AD55]'
                                  : 'text-[#63B3ED]'
                                : 'text-[#718096]'
                            }`}
                          >
                            {f.total.toLocaleString('es-PE')} u
                          </div>
                          {f.merma > 0 && (
                            <div className="text-[10px] text-[#FC8181] tabular-nums">
                              -{f.merma.toLocaleString('es-PE')} merma
                            </div>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })
          )}
        </div>

        {/* Footer */}
        <div className="px-4 py-2.5 bg-[#12161F] border-t border-[#2D3748] flex items-center justify-between gap-2 shrink-0">
          <button
            type="button"
            disabled={cargando}
            onClick={cargarHistorialOp}
            className="flex items-center gap-1.5 py-1.5 px-3 rounded-lg border border-[#4A5568] bg-[#232A38] text-[#CBD5E0] hover:text-white text-xs font-semibold cursor-pointer disabled:opacity-50"
          >
            <RefreshCw size={12} className={cargando ? 'animate-spin' : ''} />
            <span>Actualizar</span>
          </button>
          <button
            type="button"
            onClick={onClose}
            className="bg-[#F2A93B] text-[#12161F] py-1.5 px-4 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:brightness-105"
          >
            Cerrar
          </button>
        </div>
      </div>
    </div>
  );
};
