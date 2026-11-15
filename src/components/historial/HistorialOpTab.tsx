import React, { useState, useMemo, useCallback } from 'react';
import {
  Search,
  X,
  RefreshCw,
  Sun,
  Moon,
  Calendar,
  Cpu,
  AlertCircle,
  Package,
  Layers,
  TrendingUp,
  FileDown,
} from 'lucide-react';
import { useApp } from '../../context/AppContext';
import { buscarEnCatalogo, descripcionCorta } from '../../services/excel';
import { formatearFechaTurno } from '../../services/reports';
import {
  consultarHistorialOp,
  ConsultaOpResultado,
  TurnoOpResumen,
} from '../../services/opHistory';
import { exportarPdfReporteOp } from '../../services/pdfOp';
import { CatalogoItem } from '../../types';

export const HistorialOpTab: React.FC = () => {
  const { catalogo, historial, maquinas, turno, metaTurno, session, showToast } = useApp();

  const [opInput, setOpInput] = useState('');
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [opConsultada, setOpConsultada] = useState('');
  const [cargando, setCargando] = useState(false);
  const [resultado, setResultado] = useState<ConsultaOpResultado | null>(null);

  // Autocomplete suggestions from catalog
  const suggestions = useMemo(() => {
    const raw = opInput.trim();
    const digits = raw.replace(/\D/g, '');
    if (!digits && !raw) return [];
    return catalogo
      .filter((c) => {
        const cDigits = c.op.replace(/\D/g, '');
        return (
          (digits && cDigits.includes(digits)) ||
          c.op.toLowerCase().includes(raw.toLowerCase()) ||
          (c.descripcion || '').toLowerCase().includes(raw.toLowerCase())
        );
      })
      .sort((a, b) => {
        const aDigits = a.op.replace(/\D/g, '');
        const bDigits = b.op.replace(/\D/g, '');
        const aStarts = digits && aDigits.startsWith(digits);
        const bStarts = digits && bDigits.startsWith(digits);
        if (aStarts && !bStarts) return -1;
        if (!aStarts && bStarts) return 1;
        return aDigits.length - bDigits.length;
      })
      .slice(0, 10);
  }, [catalogo, opInput]);

  // Quick-pick OPs currently active in turno
  const opsTurnoActual = useMemo(() => {
    const set = new Set<string>();
    (turno || []).forEach((t) => {
      if (t.op && t.op.trim()) set.add(t.op.trim());
      (t.registros || []).forEach((r) => {
        if (r.op && r.op.trim()) set.add(r.op.trim());
      });
    });
    return Array.from(set).slice(0, 12);
  }, [turno]);

  const catalogMatch = useMemo(
    () => (opConsultada ? buscarEnCatalogo(catalogo, opConsultada) : null),
    [catalogo, opConsultada]
  );

  const ejecutarConsulta = useCallback(
    async (opObjetivo: string) => {
      const limpia = opObjetivo.trim();
      if (!limpia) return;
      setOpConsultada(limpia);
      setShowSuggestions(false);
      setCargando(true);
      const res = await consultarHistorialOp({
        op: limpia,
        catalogo,
        historial,
        maquinas,
        turno,
        metaTurno,
        session,
      });
      setResultado(res);
      setCargando(false);
    },
    [catalogo, historial, maquinas, turno, metaTurno, session]
  );

  const handleSelectSuggestion = (item: CatalogoItem) => {
    setOpInput(item.op);
    setShowSuggestions(false);
    ejecutarConsulta(item.op);
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const typed = opInput.trim();
    if (!typed) return;
    if (showSuggestions && suggestions.length > 0 && suggestions[0].op.toLowerCase() === typed.toLowerCase()) {
      handleSelectSuggestion(suggestions[0]);
      return;
    }
    ejecutarConsulta(typed);
  };

  const totales = resultado?.totales;
  const cantidadOrden = Number(catalogMatch?.cantidadTotal || 0);
  const pctPT = cantidadOrden > 0 && totales ? (totales.acumPT / cantidadOrden) * 100 : 0;
  const pendientePT = cantidadOrden > 0 && totales ? cantidadOrden - totales.acumPT : 0;

  const todosLosTurnos: TurnoOpResumen[] = useMemo(() => {
    if (!resultado) return [];
    const lista: TurnoOpResumen[] = [];
    if (resultado.turnoActualResumen) lista.push(resultado.turnoActualResumen);
    lista.push(...resultado.turnosAnteriores);
    return lista;
  }, [resultado]);

  const ppPorMaquina = useMemo(
    () => (resultado?.porMaquina || []).filter((m) => m.tipoProducto === 'pp'),
    [resultado]
  );

  const handleDescargarPdfOp = () => {
    if (!opConsultada || !resultado) return;
    try {
      const { nombreArchivo } = exportarPdfReporteOp({
        op: opConsultada,
        catalogMatch,
        resultado,
      });
      showToast('ok', `PDF generado: ${nombreArchivo}`);
    } catch (err) {
      console.error('Error al generar PDF de OP:', err);
      showToast('error', 'No se pudo generar el reporte PDF de la OP');
    }
  };

  return (
    <div className="space-y-4 sm:space-y-5">
      {/* Search Card */}
      <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-4 sm:p-5 shadow-sm space-y-3.5">
        <div className="flex items-center justify-between gap-2 pb-2 border-b border-[#2D3748] flex-wrap">
          <div className="flex items-center gap-2">
            <Search size={18} className="text-[#F2A93B]" />
            <h2 className="m-0 text-sm sm:text-base font-bold text-white font-['Oswald',sans-serif] tracking-wide uppercase">
              Consultar Producción Acumulada por OP
            </h2>
          </div>
          {opConsultada && (
            <div className="flex items-center gap-2 flex-wrap">
              {resultado && !cargando && (
                <button
                  type="button"
                  onClick={handleDescargarPdfOp}
                  className="flex items-center gap-1.5 py-1 px-3 rounded-lg bg-[#2B6CB0] hover:bg-[#3182CE] text-white text-xs font-bold cursor-pointer shadow-sm transition-colors"
                  title="Descargar reporte completo de esta OP en PDF"
                >
                  <FileDown size={13} />
                  <span>Reporte PDF</span>
                </button>
              )}
              <button
                type="button"
                disabled={cargando}
                onClick={() => ejecutarConsulta(opConsultada)}
                className="flex items-center gap-1.5 py-1 px-2.5 rounded-lg border border-[#4A5568] bg-[#232A38] text-[#CBD5E0] hover:text-white text-xs font-semibold cursor-pointer disabled:opacity-50"
              >
                <RefreshCw size={12} className={cargando ? 'animate-spin' : ''} />
                <span>Actualizar</span>
              </button>
            </div>
          )}
        </div>

        <form onSubmit={handleSubmit} className="flex flex-col sm:flex-row gap-2.5">
          <div className="relative flex-1">
            <input
              type="text"
              value={opInput}
              onChange={(e) => {
                setOpInput(e.target.value);
                setShowSuggestions(true);
              }}
              onFocus={() => setShowSuggestions(true)}
              onBlur={() => {
                setTimeout(() => setShowSuggestions(false), 220);
              }}
              placeholder="Ingresa número de OP (ej. 10234)..."
              autoComplete="off"
              className="w-full py-2.5 px-3.5 pr-9 rounded-xl border border-[#4A5568] text-sm font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none focus:border-[#F2A93B] min-h-[44px]"
            />
            {opInput && (
              <button
                type="button"
                onClick={() => {
                  setOpInput('');
                  setShowSuggestions(false);
                }}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-[#A0AEC0] hover:text-white p-1 rounded cursor-pointer"
                title="Limpiar"
              >
                <X size={15} />
              </button>
            )}

            {showSuggestions && opInput.trim() !== '' && suggestions.length > 0 && (
              <div className="absolute z-40 left-0 right-0 mt-1 bg-[#1A202C] border border-[#4A5568] rounded-xl max-h-64 overflow-y-auto shadow-2xl divide-y divide-[#2D3748]">
                <div className="px-3 py-1.5 bg-[#232A38] flex items-center justify-between text-[10px] text-[#A0AEC0]">
                  <span className="font-semibold uppercase tracking-wider text-[#CBD5E0]">
                    Catálogo ({suggestions.length})
                  </span>
                  <span>Toca para consultar</span>
                </div>
                {suggestions.map((item) => (
                  <button
                    key={item.op}
                    type="button"
                    onMouseDown={(e) => {
                      e.preventDefault();
                      handleSelectSuggestion(item);
                    }}
                    className="w-full text-left py-2 px-3 hover:bg-[#232A38] text-xs block text-white transition-colors cursor-pointer"
                  >
                    <div className="flex items-center justify-between gap-2 mb-0.5 flex-wrap">
                      <div className="flex items-center gap-1.5">
                        <span className="font-bold text-[#F2A93B] font-['IBM_Plex_Mono',monospace] px-1.5 py-0.5 rounded bg-[#232A38] border border-[#4A5568]">
                          {item.op}
                        </span>
                        {item.cantidadTotal && Number(item.cantidadTotal) > 0 ? (
                          <span className="text-[#F6AD55] font-['IBM_Plex_Mono',monospace] text-[10.5px] font-bold">
                            {Number(item.cantidadTotal).toLocaleString('es-PE')} u
                          </span>
                        ) : null}
                      </div>
                      <span className="text-[#A0AEC0] text-[10.5px] truncate max-w-[160px]">
                        {item.cliente || 'Sin cliente'}
                      </span>
                    </div>
                    <div className="text-[#CBD5E0] text-[11px] truncate">{descripcionCorta(item)}</div>
                  </button>
                ))}
              </div>
            )}
          </div>

          <button
            type="submit"
            disabled={!opInput.trim() || cargando}
            className="bg-[#F2A93B] text-[#12161F] py-2.5 px-5 rounded-xl text-xs sm:text-sm font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:brightness-105 active:scale-95 transition-all shadow-sm min-h-[44px] flex items-center justify-center gap-2 disabled:opacity-50"
          >
            <Search size={16} />
            <span>{cargando ? 'Consultando...' : 'Consultar OP'}</span>
          </button>
        </form>

        {/* Quick-pick chips for OPs in active shift */}
        {opsTurnoActual.length > 0 && (
          <div className="flex items-center gap-1.5 flex-wrap pt-0.5">
            <span className="text-[11px] text-[#A0AEC0] mr-1">En turno actual:</span>
            {opsTurnoActual.map((opItem) => (
              <button
                key={opItem}
                type="button"
                onClick={() => {
                  setOpInput(opItem);
                  ejecutarConsulta(opItem);
                }}
                className={`py-0.5 px-2 rounded-lg text-[11px] font-['IBM_Plex_Mono',monospace] font-bold border cursor-pointer transition-colors ${
                  opConsultada === opItem
                    ? 'bg-[#F2A93B] text-[#12161F] border-[#F2A93B]'
                    : 'bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:border-[#F2A93B]/60 hover:text-white'
                }`}
              >
                {opItem}
              </button>
            ))}
          </div>
        )}
      </div>

      {/* Results area */}
      {!opConsultada ? (
        <div className="border border-dashed border-[#4A5568] rounded-2xl p-8 text-center text-xs sm:text-sm text-[#CBD5E0] bg-[#1A202C]">
          Ingresa un número de OP arriba para ver todo lo registrado en{' '}
          <strong className="text-[#63B3ED]">PT (Producto Terminado)</strong> y{' '}
          <strong className="text-[#F6AD55]">PP (Producto en Proceso)</strong>, los totales por máquina y el desglose por turnos/fechas.
        </div>
      ) : cargando ? (
        <div className="border border-dashed border-[#4A5568] rounded-2xl p-10 text-center space-y-2 text-xs sm:text-sm text-[#CBD5E0] bg-[#1A202C]">
          <RefreshCw size={24} className="mx-auto text-[#F2A93B] animate-spin" />
          <p className="m-0">
            Buscando todos los registros de la OP{' '}
            <strong className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace]">{opConsultada}</strong>…
          </p>
        </div>
      ) : resultado && totales ? (
        <>
          {/* Cloud status notice if offline or unauthenticated */}
          {resultado.nubeMensaje && (
            <div className="flex items-center gap-2 text-xs text-[#F6AD55] bg-[#3A2E1B]/50 border border-[#F2A93B]/30 rounded-xl px-3.5 py-2.5">
              <AlertCircle size={15} className="shrink-0" />
              <span>{resultado.nubeMensaje}</span>
            </div>
          )}

          {/* Catalog Info Banner for searched OP */}
          <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl p-4 shadow-sm space-y-2">
            <div className="flex items-start justify-between gap-3 flex-wrap">
              <div className="min-w-0 flex-1">
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="font-['IBM_Plex_Mono',monospace] font-bold text-base sm:text-lg text-[#F2A93B] bg-[#232A38] border border-[#4A5568] px-2.5 py-0.5 rounded-lg">
                    OP {opConsultada}
                  </span>
                  {catalogMatch?.cliente && (
                    <span className="text-xs font-semibold text-[#CBD5E0] bg-[#232A38] px-2.5 py-1 rounded-lg border border-[#2D3748]">
                      Cliente: {catalogMatch.cliente}
                    </span>
                  )}
                </div>
                <p className="text-xs sm:text-sm text-white font-medium mt-1.5 mb-0">
                  {catalogMatch
                    ? catalogMatch.descripcionBase || catalogMatch.descripcion
                    : 'OP no encontrada en el catálogo local (mostrando registros de producción)'}
                </p>
              </div>

              {cantidadOrden > 0 && (
                <div className="bg-[#12161F] border border-[#F2A93B]/40 rounded-xl px-3 py-2 text-right shrink-0">
                  <div className="text-[10px] uppercase tracking-wider text-[#A0AEC0] font-semibold flex items-center justify-end gap-1">
                    <Package size={11} className="text-[#F2A93B]" />
                    <span>Cantidad Orden</span>
                  </div>
                  <div className="font-['IBM_Plex_Mono',monospace] text-sm sm:text-base font-bold text-[#F6AD55] tabular-nums">
                    {cantidadOrden.toLocaleString('es-PE')} u
                  </div>
                </div>
              )}
            </div>

            {catalogMatch && catalogMatch.modelos && catalogMatch.modelos.length > 1 && (
              <div className="pt-2 border-t border-[#2D3748] flex items-center gap-1.5 flex-wrap">
                <span className="text-[10.5px] text-[#A0AEC0] font-semibold uppercase tracking-wide mr-1">
                  Modelos ({catalogMatch.modelos.length}):
                </span>
                {catalogMatch.modelos.map((mo, idx) => (
                  <span
                    key={idx}
                    className="text-[11px] bg-[#232A38] border border-[#4A5568] text-[#CBD5E0] px-2 py-0.5 rounded-md"
                  >
                    {mo.descripcion || `Modelo ${idx + 1}`}
                    {mo.cantidadTotal && Number(mo.cantidadTotal) > 0 ? (
                      <strong className="text-[#F6AD55] font-['IBM_Plex_Mono',monospace] ml-1">
                        ({Number(mo.cantidadTotal).toLocaleString('es-PE')} u)
                      </strong>
                    ) : null}
                  </span>
                ))}
              </div>
            )}
          </div>

          {/* Separated Totals: PT (summed) & PP (by machine small cards) */}
          <div className="space-y-3">
            <div
              className={`grid grid-cols-1 gap-3 ${
                totales.acumSinTipo > 0 ? 'md:grid-cols-2' : ''
              }`}
            >
              {/* Card PT (mantiene lógica acumulada) */}
              <div className="bg-[#1A202C] border border-[#63B3ED]/40 rounded-2xl p-4 shadow-sm space-y-2">
                <div className="flex items-center justify-between gap-2">
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold px-2 py-0.5 rounded-md bg-[#63B3ED]/20 text-[#63B3ED] font-['IBM_Plex_Mono',monospace]">
                      PT
                    </span>
                    <span className="text-xs uppercase tracking-wider text-white font-bold font-['Oswald',sans-serif]">
                      Producto Terminado
                    </span>
                  </div>
                  {cantidadOrden > 0 && (
                    <span
                      className={`text-xs font-bold font-['IBM_Plex_Mono',monospace] ${
                        pctPT >= 100 ? 'text-[#48BB78]' : 'text-[#93C5FD]'
                      }`}
                    >
                      {pctPT.toFixed(1)}%
                    </span>
                  )}
                </div>

                <div className="font-['IBM_Plex_Mono',monospace] text-2xl font-bold text-[#63B3ED] tabular-nums">
                  {totales.acumPT.toLocaleString('es-PE')} u
                </div>

                <div className="text-xs text-[#CBD5E0] font-['IBM_Plex_Mono',monospace] flex items-center gap-2.5 flex-wrap">
                  <span>Turnos ant.: {totales.antPT.toLocaleString('es-PE')} u</span>
                  {totales.actualPT > 0 && (
                    <span className="text-[#48BB78] font-semibold">
                      · Turno actual: +{totales.actualPT.toLocaleString('es-PE')} u
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
                    <div className="w-full bg-[#12161F] rounded-full h-2 overflow-hidden border border-[#2D3748]">
                      <div
                        className={`h-full rounded-full transition-all duration-300 ${
                          pctPT >= 100 ? 'bg-[#38A169]' : 'bg-[#63B3ED]'
                        }`}
                        style={{ width: `${Math.min(100, Math.max(0, pctPT))}%` }}
                      />
                    </div>
                    <div className="text-[11px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace]">
                      {pendientePT > 0
                        ? `Faltan ${pendientePT.toLocaleString('es-PE')} u para completar la orden`
                        : pendientePT < 0
                        ? `+${Math.abs(pendientePT).toLocaleString('es-PE')} u sobre la cantidad de orden`
                        : 'Orden completada al 100% en PT'}
                    </div>
                  </div>
                )}
              </div>

              {/* Card Sin Tipo (legacy) */}
              {totales.acumSinTipo > 0 && (
                <div className="bg-[#1A202C] border border-[#4A5568] rounded-2xl p-4 shadow-sm space-y-2">
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold px-2 py-0.5 rounded-md bg-[#4A5568]/40 text-[#CBD5E0] font-['IBM_Plex_Mono',monospace]">
                      S/T
                    </span>
                    <span className="text-xs uppercase tracking-wider text-white font-bold font-['Oswald',sans-serif]">
                      Sin Clasificar
                    </span>
                  </div>
                  <div className="font-['IBM_Plex_Mono',monospace] text-2xl font-bold text-white tabular-nums">
                    {totales.acumSinTipo.toLocaleString('es-PE')} u
                  </div>
                  <div className="text-xs text-[#CBD5E0] font-['IBM_Plex_Mono',monospace]">
                    Ant.: {totales.antSinTipo.toLocaleString('es-PE')} u
                    {totales.actualSinTipo > 0 && (
                      <span className="text-[#48BB78] ml-2">
                        · Actual: +{totales.actualSinTipo.toLocaleString('es-PE')} u
                      </span>
                    )}
                  </div>
                </div>
              )}
            </div>

            {/* PP separado por máquinas en cuadros pequeños (sin sumar entre máquinas) */}
            <div className="bg-[#1A202C] border border-[#F2A93B]/35 rounded-2xl p-3.5 sm:p-4 shadow-sm space-y-2.5">
              <div className="flex items-center justify-between gap-2 flex-wrap">
                <div className="flex items-center gap-2">
                  <span className="text-xs font-bold px-2 py-0.5 rounded-md bg-[#F2A93B]/20 text-[#F6AD55] font-['IBM_Plex_Mono',monospace]">
                    PP
                  </span>
                  <span className="text-xs uppercase tracking-wider text-white font-bold font-['Oswald',sans-serif]">
                    Producto en Proceso por Máquina
                  </span>
                </div>
                <span className="text-[11px] text-[#A0AEC0]">
                  {ppPorMaquina.length > 0
                    ? `${ppPorMaquina.length} máquina(s) · procesos independientes`
                    : 'Sin registros en proceso'}
                </span>
              </div>

              {ppPorMaquina.length === 0 ? (
                <div className="bg-[#12161F] border border-dashed border-[#2D3748] rounded-xl px-3 py-2.5 text-xs text-[#A0AEC0]">
                  No hay producción registrada en PP para esta OP.
                </div>
              ) : (
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-2.5">
                  {ppPorMaquina.map((mq) => (
                    <div
                      key={mq.key}
                      className="bg-[#12161F] border border-[#F2A93B]/35 rounded-xl px-3 py-2.5 space-y-1 hover:border-[#F2A93B]/60 transition-colors"
                    >
                      <div className="flex items-center justify-between gap-1.5">
                        <div className="flex items-center gap-1.5 min-w-0">
                          <Cpu size={12} className="text-[#F2A93B] shrink-0" />
                          <span className="text-xs font-bold text-white truncate font-['IBM_Plex_Mono',monospace]">
                            {mq.maquina}
                          </span>
                        </div>
                        {mq.codigoMaquina && (
                          <span className="text-[9.5px] px-1.5 py-0.5 rounded bg-[#232A38] border border-[#4A5568] text-[#CBD5E0] font-['IBM_Plex_Mono',monospace] shrink-0">
                            {mq.codigoMaquina}
                          </span>
                        )}
                      </div>

                      <div className="font-['IBM_Plex_Mono',monospace] text-lg font-bold text-[#F2A93B] tabular-nums">
                        {mq.totalAcumulado.toLocaleString('es-PE')} u
                      </div>

                      <div className="text-[10.5px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace] flex items-center gap-1.5 flex-wrap">
                        <span>Ant: {mq.totalAnterior.toLocaleString('es-PE')}</span>
                        {mq.totalActual > 0 && (
                          <span className="text-[#48BB78] font-semibold">
                            · Act: +{mq.totalActual.toLocaleString('es-PE')}
                          </span>
                        )}
                        {mq.mermaAcumulada > 0 && (
                          <span className="text-[#FC8181]">
                            · -{mq.mermaAcumulada.toLocaleString('es-PE')} m
                          </span>
                        )}
                        <span className="text-[#718096]">· {mq.turnosCount} t.</span>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>

          {resultado.porMaquina.length === 0 && todosLosTurnos.length === 0 ? (
            <div className="border border-dashed border-[#4A5568] rounded-2xl p-8 text-center text-xs sm:text-sm text-[#CBD5E0] bg-[#1A202C]">
              No se encontraron registros de producción para la OP{' '}
              <strong className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace]">{opConsultada}</strong>{' '}
              en turnos guardados ni en el turno actual.
            </div>
          ) : (
            <>
              {/* Section 1: Totals by Machine */}
              <div className="bg-[#1A202C] border border-[#2D3748] rounded-2xl overflow-hidden shadow-sm">
                <div className="px-4 py-3 bg-[#12161F] border-b border-[#2D3748] flex items-center justify-between gap-2">
                  <div className="flex items-center gap-2">
                    <Layers size={16} className="text-[#F2A93B]" />
                    <h3 className="m-0 text-xs sm:text-sm font-bold text-white uppercase font-['Oswald',sans-serif] tracking-wide">
                      Totales por Máquina ({resultado.porMaquina.length})
                    </h3>
                  </div>
                </div>

                <div className="divide-y divide-[#2D3748]">
                  {resultado.porMaquina.map((mq) => (
                    <div
                      key={mq.key}
                      className="px-4 py-2.5 flex items-center justify-between gap-3 text-xs sm:text-sm hover:bg-[#232A38]/50"
                    >
                      <div className="min-w-0 flex-1">
                        <div className="flex items-center gap-2 flex-wrap">
                          <Cpu size={14} className="text-[#A0AEC0] shrink-0" />
                          <span className="font-['IBM_Plex_Mono',monospace] font-bold text-white">
                            {mq.maquina}
                          </span>
                          {mq.codigoMaquina && (
                            <span className="text-[10.5px] px-1.5 py-0.5 rounded bg-[#232A38] border border-[#4A5568] text-[#CBD5E0] font-['IBM_Plex_Mono',monospace]">
                              {mq.codigoMaquina}
                            </span>
                          )}
                          {mq.tipoProducto && (
                            <span
                              className={`text-[10px] font-bold px-1.5 py-0.5 rounded font-['IBM_Plex_Mono',monospace] ${
                                mq.tipoProducto === 'pp'
                                  ? 'bg-[#F2A93B]/20 text-[#F6AD55]'
                                  : 'bg-[#63B3ED]/20 text-[#63B3ED]'
                              }`}
                            >
                              {mq.tipoProducto.toUpperCase()}
                            </span>
                          )}
                          <span className="text-[11px] text-[#A0AEC0]">
                            · {mq.turnosCount} turno(s)
                          </span>
                        </div>
                        {(mq.totalActual > 0 || mq.mermaAcumulada > 0) && (
                          <div className="text-[11px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace] mt-0.5 pl-5 flex items-center gap-2 flex-wrap">
                            <span>Guardado: {mq.totalAnterior.toLocaleString('es-PE')} u</span>
                            {mq.totalActual > 0 && (
                              <span className="text-[#48BB78]">
                                · Actual: +{mq.totalActual.toLocaleString('es-PE')} u
                              </span>
                            )}
                            {mq.mermaAcumulada > 0 && (
                              <span className="text-[#FC8181]">
                                · -{mq.mermaAcumulada.toLocaleString('es-PE')} merma
                              </span>
                            )}
                          </div>
                        )}
                      </div>

                      <div className="text-right shrink-0 font-['IBM_Plex_Mono',monospace]">
                        <span
                          className={`font-bold text-sm sm:text-base tabular-nums ${
                            mq.tipoProducto === 'pp'
                              ? 'text-[#F2A93B]'
                              : mq.tipoProducto === 'pt'
                              ? 'text-[#63B3ED]'
                              : 'text-white'
                          }`}
                        >
                          {mq.totalAcumulado.toLocaleString('es-PE')} u
                        </span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Section 2: Breakdown by Shifts / Dates */}
              <div className="space-y-2.5">
                <div className="flex items-center gap-2 px-1">
                  <TrendingUp size={16} className="text-[#F2A93B]" />
                  <h3 className="m-0 text-xs sm:text-sm font-bold text-white uppercase font-['Oswald',sans-serif] tracking-wide">
                    Detalle por Turnos y Fechas ({todosLosTurnos.length})
                  </h3>
                </div>

                {todosLosTurnos.map((t) => {
                  const esNoche = t.tipo === 'noche';
                  const esActual = t.fuente === 'actual';
                  const etiquetaTurnoTxt = esNoche ? 'Turno Noche' : t.tipo === 'dia' ? 'Turno Día' : 'Turno';

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
                            className="px-3.5 py-2.5 flex items-center justify-between gap-3 text-xs sm:text-sm hover:bg-[#232A38]/40"
                          >
                            <div className="min-w-0 flex-1">
                              <div className="flex items-center gap-1.5 flex-wrap">
                                <Cpu size={13} className="text-[#A0AEC0] shrink-0" />
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
                                    className={`text-[10px] font-bold px-1.5 py-0.2 rounded font-['IBM_Plex_Mono',monospace] ${
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
                                <div className="text-xs text-[#CBD5E0] truncate mt-0.5 pl-5">
                                  {f.producto}
                                </div>
                              )}
                              {f.observaciones && (
                                <div className="text-[11px] text-[#A0AEC0] truncate mt-0.5 pl-5 italic">
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
                                      : f.tipoProducto === 'pt'
                                      ? 'text-[#63B3ED]'
                                      : 'text-white'
                                    : 'text-[#718096]'
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
