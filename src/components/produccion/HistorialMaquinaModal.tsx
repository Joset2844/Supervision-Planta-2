import React, { useState, useEffect, useCallback } from 'react';
import {
  History,
  X,
  Sun,
  Moon,
  RefreshCw,
  Calendar,
  Cpu,
  AlertCircle,
  Target,
} from 'lucide-react';
import { useApp } from '../../context/AppContext';
import { formatearFechaTurno } from '../../services/reports';
import {
  consultarHistorialMaquina,
  ConsultaMaquinaResultado,
} from '../../services/opHistory';
import { Maquina } from '../../types';

interface HistorialMaquinaModalProps {
  maquina: Maquina;
  onClose: () => void;
}

export const HistorialMaquinaModal: React.FC<HistorialMaquinaModalProps> = ({
  maquina,
  onClose,
}) => {
  const { historial, metaTurno, session } = useApp();

  const [cargando, setCargando] = useState(true);
  const [resultado, setResultado] = useState<ConsultaMaquinaResultado | null>(null);

  const cargarHistorial = useCallback(async () => {
    setCargando(true);
    const res = await consultarHistorialMaquina({
      maquina,
      historial,
      metaTurno,
      session,
    });
    setResultado(res);
    setCargando(false);
  }, [maquina, historial, metaTurno, session]);

  useEffect(() => {
    cargarHistorial();
  }, [cargarHistorial]);

  const turnosAnteriores = resultado?.turnosAnteriores || [];
  const totalProducido = resultado?.totalProducido || 0;
  const totalMerma = resultado?.totalMerma || 0;
  const promedioPorTurno = resultado?.promedioPorTurno || 0;
  const meta = maquina.meta || 0;

  const areaTxt =
    maquina.area === 'vasos_envases'
      ? 'Vasos y Envases'
      : maquina.area === 'papel_carton'
      ? 'Papel y Cartón'
      : '';

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
              <div className="w-7 h-7 rounded-lg bg-[#63B3ED]/20 text-[#63B3ED] flex items-center justify-center shrink-0">
                <Cpu size={16} />
              </div>
              <h3 className="m-0 text-sm sm:text-base font-bold text-white font-['Oswald',sans-serif] uppercase tracking-wide">
                Historial de Máquina —{' '}
                <span className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace]">
                  {maquina.nombre}
                </span>
              </h3>
              {maquina.codigo && (
                <span className="text-[10.5px] font-bold px-1.5 py-0.5 rounded bg-[#232A38] border border-[#4A5568] text-[#CBD5E0] font-['IBM_Plex_Mono',monospace]">
                  {maquina.codigo}
                </span>
              )}
              {maquina.tipoProducto && (
                <span
                  className={`text-[10px] font-bold px-1.5 py-0.5 rounded font-['IBM_Plex_Mono',monospace] ${
                    maquina.tipoProducto === 'pp'
                      ? 'bg-[#F2A93B]/20 text-[#F6AD55]'
                      : 'bg-[#63B3ED]/20 text-[#63B3ED]'
                  }`}
                >
                  {maquina.tipoProducto.toUpperCase()}
                </span>
              )}
            </div>
            {(areaTxt || meta > 0) && (
              <div className="text-xs text-[#A0AEC0] mt-1 flex items-center gap-2 flex-wrap">
                {areaTxt && <span>{areaTxt}</span>}
                {meta > 0 && (
                  <span className="inline-flex items-center gap-1 text-[11px] font-['IBM_Plex_Mono',monospace] text-[#CBD5E0] bg-[#232A38] px-2 py-0.5 rounded border border-[#2D3748]">
                    <Target size={11} className="text-[#F2A93B]" />
                    Meta/turno: {meta.toLocaleString('es-PE')} u
                  </span>
                )}
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

        {/* Summary Strip */}
        <div className="px-4 py-3 bg-[#141923] border-b border-[#2D3748] space-y-2.5 shrink-0">
          <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
            <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl px-3 py-2">
              <div className="text-[10px] uppercase tracking-wider text-[#A0AEC0] font-semibold">
                Total Producido
              </div>
              <div className="font-['IBM_Plex_Mono',monospace] text-sm sm:text-base font-bold text-[#F2A93B] tabular-nums mt-0.5">
                {totalProducido.toLocaleString('es-PE')} u
              </div>
              {totalMerma > 0 && (
                <div className="text-[10px] text-[#FC8181] font-['IBM_Plex_Mono',monospace]">
                  -{totalMerma.toLocaleString('es-PE')} u merma
                </div>
              )}
            </div>

            <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl px-3 py-2">
              <div className="text-[10px] uppercase tracking-wider text-[#A0AEC0] font-semibold">
                Turnos Anteriores
              </div>
              <div className="font-['IBM_Plex_Mono',monospace] text-sm sm:text-base font-bold text-white tabular-nums mt-0.5">
                {turnosAnteriores.length}
              </div>
              <div className="text-[10px] text-[#A0AEC0]">turnos registrados</div>
            </div>

            <div className="col-span-2 sm:col-span-1 bg-[#1A202C] border border-[#2D3748] rounded-xl px-3 py-2">
              <div className="text-[10px] uppercase tracking-wider text-[#A0AEC0] font-semibold">
                Promedio / Turno
              </div>
              <div className="font-['IBM_Plex_Mono',monospace] text-sm sm:text-base font-bold text-[#63B3ED] tabular-nums mt-0.5">
                {promedioPorTurno.toLocaleString('es-PE')} u
              </div>
              {meta > 0 && turnosAnteriores.length > 0 && (
                <div className="text-[10px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace]">
                  {((promedioPorTurno / meta) * 100).toFixed(1)}% de la meta
                </div>
              )}
            </div>
          </div>

          {resultado?.nubeMensaje && !cargando && (
            <div className="flex items-center gap-2 text-[11px] text-[#F6AD55] bg-[#3A2E1B]/50 border border-[#F2A93B]/30 rounded-lg px-2.5 py-1.5">
              <AlertCircle size={13} className="shrink-0" />
              <span>{resultado.nubeMensaje}</span>
            </div>
          )}
        </div>

        {/* Body: Shift history list */}
        <div className="flex-1 overflow-y-auto p-4 space-y-3">
          {cargando ? (
            <div className="py-10 text-center space-y-2 text-xs text-[#CBD5E0]">
              <RefreshCw size={22} className="mx-auto text-[#F2A93B] animate-spin" />
              <p className="m-0">Consultando historial de {maquina.nombre}…</p>
            </div>
          ) : turnosAnteriores.length === 0 ? (
            <div className="border border-dashed border-[#4A5568] rounded-xl p-6 text-center text-xs sm:text-sm text-[#CBD5E0] bg-[#141923]">
              No se encontraron registros anteriores para{' '}
              <strong className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace]">
                {maquina.nombre}
              </strong>
              .
              {!resultado?.nubeDisponible && !session && (
                <p className="text-[11px] text-[#A0AEC0] mt-1.5 mb-0">
                  Tip: Inicia sesión en la pestaña Historial para consultar también los turnos guardados en la nube.
                </p>
              )}
            </div>
          ) : (
            turnosAnteriores.map((t) => {
              const esNoche = t.tipo === 'noche';
              const etiquetaTurnoTxt = esNoche ? 'Turno Noche' : t.tipo === 'dia' ? 'Turno Día' : 'Turno';
              const pctMeta = meta > 0 ? (t.totalProducido / meta) * 100 : 0;

              return (
                <div
                  key={t.key}
                  className="bg-[#141923] border border-[#2D3748] rounded-xl overflow-hidden shadow-xs"
                >
                  {/* Shift header */}
                  <div className="px-3 py-2 bg-[#232A38] border-b border-[#2D3748] flex items-center justify-between gap-2 flex-wrap">
                    <div className="flex items-center gap-2 flex-wrap">
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
                      {meta > 0 && (
                        <span
                          className={`text-[10px] font-bold px-1.5 py-0.5 rounded border ${
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
                      <span className="font-bold text-[#F2A93B] tabular-nums">
                        +{t.totalProducido.toLocaleString('es-PE')} u
                      </span>
                      {t.totalMerma > 0 && (
                        <span className="text-[#FC8181] text-[11px] tabular-nums">
                          (-{t.totalMerma.toLocaleString('es-PE')} merma)
                        </span>
                      )}
                    </div>
                  </div>

                  {/* OP / Product rows in that shift */}
                  <div className="divide-y divide-[#2D3748]/70">
                    {t.filas.map((f, idx) => (
                      <div
                        key={idx}
                        className="px-3 py-2 flex items-center justify-between gap-3 text-xs hover:bg-[#1A202C]/60"
                      >
                        <div className="min-w-0 flex-1">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span
                              className={`font-['IBM_Plex_Mono',monospace] font-bold px-1.5 py-0.2 rounded text-[11px] ${
                                f.op
                                  ? 'bg-[#232A38] border border-[#4A5568] text-[#F6AD55]'
                                  : 'bg-[#232A38] text-[#A0AEC0]'
                              }`}
                            >
                              {f.op ? `OP ${f.op}` : 'Sin OP'}
                            </span>
                            {f.sap && (
                              <span className="text-[10px] font-['IBM_Plex_Mono',monospace] text-[#A0AEC0]">
                                SAP: {f.sap}
                              </span>
                            )}
                          </div>
                          {f.producto && (
                            <div className="text-[11px] text-[#CBD5E0] truncate mt-0.5">
                              {f.producto}
                            </div>
                          )}
                          {f.observaciones && (
                            <div className="text-[10.5px] text-[#F6AD55]/90 truncate mt-0.5 italic">
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
            onClick={cargarHistorial}
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
