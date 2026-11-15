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
import { getSupabaseConfig, sbFetch } from '../../services/supabase';
import { buscarEnCatalogo } from '../../services/excel';
import { formatearFechaTurno } from '../../services/reports';
import { TurnoDBRow } from '../../types';

interface HistorialOpModalProps {
  op: string;
  onClose: () => void;
}

interface MaquinaRegistroOp {
  maquina: string;
  codigoMaquina: string;
  tipoProducto: 'pp' | 'pt' | '';
  producto: string;
  total: number;
  merma: number;
  observaciones: string;
}

interface TurnoOpResumen {
  key: string;
  fecha: string;
  tipo: string;
  fuente: 'nube' | 'local';
  totalPP: number;
  mermaPP: number;
  totalPT: number;
  mermaPT: number;
  totalSinTipo: number;
  mermaSinTipo: number;
  filas: MaquinaRegistroOp[];
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

export const HistorialOpModal: React.FC<HistorialOpModalProps> = ({ op, onClose }) => {
  const { catalogo, historial, maquinas, turno, metaTurno, session } = useApp();

  const [cargando, setCargando] = useState(true);
  const [nubeDisponible, setNubeDisponible] = useState(true);
  const [nubeMensaje, setNubeMensaje] = useState('');
  const [turnosAnteriores, setTurnosAnteriores] = useState<TurnoOpResumen[]>([]);

  const catalogMatch = useMemo(() => buscarEnCatalogo(catalogo, op), [catalogo, op]);

  const maqTipoMap = useMemo(() => {
    const map = new Map<string, 'pp' | 'pt' | ''>();
    maquinas.forEach((m) => {
      if (m.nombre) map.set(m.nombre.trim().toLowerCase(), m.tipoProducto || '');
      if (m.codigo) map.set(`cod:${m.codigo.trim().toLowerCase()}`, m.tipoProducto || '');
    });
    return map;
  }, [maquinas]);

  const resolverTipo = useCallback(
    (rawTipo?: string, nombreMaq?: string, codMaq?: string): 'pp' | 'pt' | '' => {
      const t = (rawTipo || '').trim().toLowerCase();
      if (t === 'pp' || t === 'pt') return t;
      const byName = nombreMaq ? maqTipoMap.get(nombreMaq.trim().toLowerCase()) : '';
      if (byName === 'pp' || byName === 'pt') return byName;
      const byCode = codMaq ? maqTipoMap.get(`cod:${codMaq.trim().toLowerCase()}`) : '';
      if (byCode === 'pp' || byCode === 'pt') return byCode;
      return '';
    },
    [maqTipoMap]
  );

  const cargarHistorialOp = useCallback(async () => {
    setCargando(true);
    setNubeMensaje('');

    let dbRows: TurnoDBRow[] = [];
    let cloudOk = false;

    const cfg = getSupabaseConfig(false);
    if (cfg) {
      try {
        const res = await sbFetch<TurnoDBRow[]>(
          cfg,
          'turnos?select=fecha,tipo,filas&order=fecha.desc,tipo.asc&limit=400',
          undefined,
          session
        );
        if (Array.isArray(res)) {
          dbRows = res;
          cloudOk = true;
        }
      } catch {
        cloudOk = false;
        setNubeMensaje(
          session
            ? 'No se pudo consultar la base de datos en la nube. Mostrando registros locales.'
            : 'Inicia sesión en la pestaña Historial para consultar también los turnos guardados en la nube.'
        );
      }
    } else {
      cloudOk = false;
      setNubeMensaje('Mostrando solo registros guardados en este dispositivo.');
    }

    setNubeDisponible(cloudOk);

    const mapaTurnos = new Map<string, TurnoOpResumen>();
    const currentFecha = (metaTurno.fecha || '').trim();
    const currentTipo = (metaTurno.tipo || '').trim().toLowerCase();

    const construirResumen = (
      key: string,
      fecha: string,
      tipo: string,
      fuente: 'nube' | 'local',
      coincidentes: any[]
    ): TurnoOpResumen => {
      const filasMap: MaquinaRegistroOp[] = coincidentes.map((f) => ({
        maquina: f.maquina || 'Sin máquina',
        codigoMaquina: f.codigoMaquina || '',
        tipoProducto: resolverTipo(f.tipoProducto, f.maquina, f.codigoMaquina),
        producto: f.producto || '',
        total: Number(f.total) || 0,
        merma: Number(f.merma) || 0,
        observaciones: f.observaciones || '',
      }));

      let totalPP = 0;
      let mermaPP = 0;
      let totalPT = 0;
      let mermaPT = 0;
      let totalSinTipo = 0;
      let mermaSinTipo = 0;

      filasMap.forEach((r) => {
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

      // Ordenar filas dentro del turno: PT primero, luego PP, luego sin tipo
      const peso = (tp: string) => (tp === 'pt' ? 0 : tp === 'pp' ? 1 : 2);
      filasMap.sort((a, b) => peso(a.tipoProducto) - peso(b.tipoProducto));

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
        filas: filasMap,
      };
    };

    // 1. Procesar turnos de Supabase (fuente autoritativa)
    dbRows.forEach((rec) => {
      const fFecha = (rec.fecha || '').trim();
      const fTipo = (rec.tipo || '').trim().toLowerCase();
      if (currentFecha && currentTipo && fFecha === currentFecha && fTipo === currentTipo) {
        return; // Excluir turno actual en curso
      }

      const coincidentes = (rec.filas || []).filter((f) => opCoincide(f.op, op));
      if (coincidentes.length === 0) return;

      const key = `${fFecha}|${fTipo}`;
      mapaTurnos.set(key, construirResumen(key, fFecha, fTipo, 'nube', coincidentes));
    });

    // 2. Complementar con historial local guardado en el dispositivo
    (historial || []).forEach((h) => {
      const fFecha = (h.fechaTurno || (h.fecha ? h.fecha.slice(0, 10) : '')).trim();
      const fTipo = (h.tipoTurno || '').trim().toLowerCase();
      if (!fFecha) return;
      if (currentFecha && currentTipo && fFecha === currentFecha && fTipo === currentTipo) {
        return;
      }

      const key = `${fFecha}|${fTipo}`;
      if (mapaTurnos.has(key)) return; // Ya cubierto por la nube

      const coincidentes = (h.filas || []).filter((f) => opCoincide(f.op, op));
      if (coincidentes.length === 0) return;

      mapaTurnos.set(key, construirResumen(key, fFecha, fTipo, 'local', coincidentes));
    });

    // Ordenar más reciente primero (y Noche antes de Día en la misma fecha)
    const ordenados = Array.from(mapaTurnos.values()).sort((a, b) => {
      const cmpFecha = b.fecha.localeCompare(a.fecha);
      if (cmpFecha !== 0) return cmpFecha;
      const pesoTipo = (t: string) => (t === 'noche' ? 2 : t === 'dia' ? 1 : 0);
      return pesoTipo(b.tipo) - pesoTipo(a.tipo);
    });

    setTurnosAnteriores(ordenados);
    setCargando(false);
  }, [op, session, historial, metaTurno.fecha, metaTurno.tipo, resolverTipo]);

  useEffect(() => {
    cargarHistorialOp();
  }, [cargarHistorialOp]);

  // Calcular lo producido en el turno activo actual para esta OP separado por PP y PT
  const resumenTurnoActual = useMemo(() => {
    let totalPP = 0;
    let mermaPP = 0;
    let totalPT = 0;
    let mermaPT = 0;
    let totalSinTipo = 0;
    let mermaSinTipo = 0;

    (turno || []).forEach((t) => {
      const m = maquinas.find((x) => x.id === t.maquinaId);
      const tipoMaq: 'pp' | 'pt' | '' = m?.tipoProducto || '';
      const regs = (t.registros || []).filter((r) => opCoincide(r.op || t.op, op));
      const sumaMaq = regs.reduce((s, r) => s + (Number(r.cantidad) || 0), 0);
      const mermaMaq = regs.reduce((s, r) => s + (Number(r.merma) || 0), 0);
      if (sumaMaq > 0 || mermaMaq > 0) {
        if (tipoMaq === 'pp') {
          totalPP += sumaMaq;
          mermaPP += mermaMaq;
        } else if (tipoMaq === 'pt') {
          totalPT += sumaMaq;
          mermaPT += mermaMaq;
        } else {
          totalSinTipo += sumaMaq;
          mermaSinTipo += mermaMaq;
        }
      }
    });

    return { totalPP, mermaPP, totalPT, mermaPT, totalSinTipo, mermaSinTipo };
  }, [turno, maquinas, op]);

  const totales = useMemo(() => {
    let antPP = 0;
    let mermaAntPP = 0;
    let antPT = 0;
    let mermaAntPT = 0;
    let antSinTipo = 0;
    let mermaAntSinTipo = 0;

    turnosAnteriores.forEach((t) => {
      antPP += t.totalPP;
      mermaAntPP += t.mermaPP;
      antPT += t.totalPT;
      mermaAntPT += t.mermaPT;
      antSinTipo += t.totalSinTipo;
      mermaAntSinTipo += t.mermaSinTipo;
    });

    const acumPP = antPP + resumenTurnoActual.totalPP;
    const acumPT = antPT + resumenTurnoActual.totalPT;
    const acumSinTipo = antSinTipo + resumenTurnoActual.totalSinTipo;

    return {
      antPP,
      mermaAntPP,
      acumPP,
      antPT,
      mermaAntPT,
      acumPT,
      antSinTipo,
      mermaAntSinTipo,
      acumSinTipo,
    };
  }, [turnosAnteriores, resumenTurnoActual]);

  const cantidadOrden = Number(catalogMatch?.cantidadTotal || 0);
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
                  {catalogMatch.descripcionBase || catalogMatch.descripcion}
                </span>
                {catalogMatch.cliente && (
                  <span className="text-[#A0AEC0]">· {catalogMatch.cliente}</span>
                )}
                {cantidadOrden > 0 && (
                  <span className="inline-flex items-center gap-1 text-[11px] font-['IBM_Plex_Mono',monospace] font-bold text-[#F6AD55] bg-[#3A2E1B]/60 px-2 py-0.5 rounded border border-[#F2A93B]/30">
                    <Package size={11} />
                    Orden: {cantidadOrden.toLocaleString('es-PE')} u
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

        {/* Summary KPI Strip separated by PT and PP */}
        <div className="px-4 py-3 bg-[#141923] border-b border-[#2D3748] space-y-2.5 shrink-0">
          <div
            className={`grid grid-cols-1 gap-2.5 ${
              totales.acumSinTipo > 0 ? 'sm:grid-cols-3' : 'sm:grid-cols-2'
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
                {resumenTurnoActual.totalPT > 0 && (
                  <span className="text-[#48BB78]">
                    · Actual: +{resumenTurnoActual.totalPT.toLocaleString('es-PE')} u
                  </span>
                )}
                {totales.mermaAntPT + resumenTurnoActual.mermaPT > 0 && (
                  <span className="text-[#FC8181]">
                    · -{(totales.mermaAntPT + resumenTurnoActual.mermaPT).toLocaleString('es-PE')} merma
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

            {/* PP - Producto en Proceso */}
            <div className="bg-[#1A202C] border border-[#F2A93B]/30 rounded-xl px-3 py-2.5 space-y-1">
              <div className="flex items-center justify-between gap-2">
                <div className="flex items-center gap-1.5">
                  <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-[#F2A93B]/20 text-[#F6AD55] font-['IBM_Plex_Mono',monospace]">
                    PP
                  </span>
                  <span className="text-[10.5px] uppercase tracking-wider text-[#CBD5E0] font-semibold">
                    Producto en Proceso
                  </span>
                </div>
              </div>

              <div className="font-['IBM_Plex_Mono',monospace] text-base sm:text-lg font-bold text-[#F2A93B] tabular-nums">
                {totales.acumPP.toLocaleString('es-PE')} u
              </div>

              <div className="text-[10.5px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace] flex items-center gap-2 flex-wrap">
                <span>Ant: {totales.antPP.toLocaleString('es-PE')} u</span>
                {resumenTurnoActual.totalPP > 0 && (
                  <span className="text-[#48BB78]">
                    · Actual: +{resumenTurnoActual.totalPP.toLocaleString('es-PE')} u
                  </span>
                )}
                {totales.mermaAntPP + resumenTurnoActual.mermaPP > 0 && (
                  <span className="text-[#FC8181]">
                    · -{(totales.mermaAntPP + resumenTurnoActual.mermaPP).toLocaleString('es-PE')} merma
                  </span>
                )}
              </div>

              <div className="text-[10px] text-[#718096] pt-0.5">
                Contabilizado por separado (pliegos / semielaborados)
              </div>
            </div>

            {/* Sin clasificar (solo si existen registros legacy sin PP/PT) */}
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
                <div className="text-[10.5px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace]">
                  Ant: {totales.antSinTipo.toLocaleString('es-PE')} u
                  {resumenTurnoActual.totalSinTipo > 0 && (
                    <span className="text-[#48BB78] ml-1.5">
                      · Actual: +{resumenTurnoActual.totalSinTipo.toLocaleString('es-PE')} u
                    </span>
                  )}
                </div>
              </div>
            )}
          </div>

          {nubeMensaje && !cargando && (
            <div className="flex items-center gap-2 text-[11px] text-[#F6AD55] bg-[#3A2E1B]/50 border border-[#F2A93B]/30 rounded-lg px-2.5 py-1.5">
              <AlertCircle size={13} className="shrink-0" />
              <span>{nubeMensaje}</span>
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
              {!nubeDisponible && !session && (
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
                  {/* Shift header with separated PT / PP totals */}
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
                      {(t.totalPP > 0 || t.filas.some((f) => f.tipoProducto === 'pp')) && (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-[#F2A93B]/15 border border-[#F2A93B]/30 text-[#F6AD55] font-bold tabular-nums">
                          PP: +{t.totalPP.toLocaleString('es-PE')} u
                        </span>
                      )}
                      {t.totalSinTipo > 0 && (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-[#2D3748] text-[#CBD5E0] font-bold tabular-nums">
                          +{t.totalSinTipo.toLocaleString('es-PE')} u
                        </span>
                      )}
                    </div>
                  </div>

                  {/* Machine rows */}
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
