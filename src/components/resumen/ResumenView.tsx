import React, { useMemo } from 'react';
import { useApp } from '../../context/AppContext';
import {
  FileSpreadsheet,
  Download,
  Printer,
  Calendar,
  Layers,
  CheckCircle2,
} from 'lucide-react';
import {
  AREAS_ORDEN,
  generarFilasTurno,
  exportarTurnoExcel,
  exportarMigracion,
  filasActivacion,
  exportarActivacionExcel,
  imprimirActivacion,
} from '../../services/excel';
import { formatearFechaTurno, etiquetaTurno } from '../../services/reports';
import { FilaTurnoExcel } from '../../types';

export const ResumenView: React.FC = () => {
  const { maquinas, turno, catalogo, metaTurno, showToast } = useApp();

  const filas = useMemo(() => {
    return generarFilasTurno(maquinas, turno, catalogo);
  }, [maquinas, turno, catalogo]);

  const activas = useMemo(() => maquinas.filter((m) => m.activa), [maquinas]);

  const totalGeneral = useMemo(() => {
    return filas.reduce((acc, f) => acc + (f.total || 0), 0);
  }, [filas]);

  const mermaGeneral = useMemo(() => {
    return filas.reduce((acc, f) => acc + (f.merma || 0), 0);
  }, [filas]);

  const tituloTurno = useMemo(() => {
    const fTxt = formatearFechaTurno(metaTurno.fecha);
    const tTxt = etiquetaTurno(metaTurno.tipo || 'dia').toUpperCase();
    return `PRODUCCIÓN ${tTxt} — ${fTxt}`;
  }, [metaTurno]);

  const handleExportarExcel = () => {
    if (filas.length === 0) {
      showToast('error', 'No hay máquinas activas para exportar');
      return;
    }
    const fTxt = formatearFechaTurno(metaTurno.fecha);
    const tipoTxt = metaTurno.tipo ? `TURNO ${etiquetaTurno(metaTurno.tipo).toUpperCase()} ` : '';
    const titulo = `PRODUCCION ${tipoTxt}${fTxt}`.trim();
    const nombreBase = `turno_${metaTurno.fecha || 'sin-fecha'}_${metaTurno.tipo || 'sin-turno'}`;
    exportarTurnoExcel(filas, titulo, nombreBase);
    showToast('ok', 'Excel formateado exportado');
  };

  const handleExportarMigracion = () => {
    try {
      exportarMigracion(filas, metaTurno.fecha, metaTurno.tipo);
      showToast('ok', 'Excel de migración SAP exportado');
    } catch (err: any) {
      showToast('error', err.message || 'Error al exportar migración');
    }
  };

  const handleActivacionExcel = () => {
    const filasAct = filasActivacion(maquinas, turno, catalogo);
    if (filasAct.length === 0) {
      showToast('error', 'No hay máquinas activas');
      return;
    }
    const titulo = `ACTIVACION DE MAQUINAS ${formatearFechaTurno(metaTurno.fecha)} - ${etiquetaTurno(
      metaTurno.tipo
    ).toUpperCase()}`;
    const nombre = `activacion_${metaTurno.fecha || 'sin-fecha'}_${metaTurno.tipo || 'sin-turno'}`;
    exportarActivacionExcel(filasAct, titulo, nombre);
    showToast('ok', 'Activación exportada a Excel');
  };

  const handleActivacionPdf = () => {
    const filasAct = filasActivacion(maquinas, turno, catalogo);
    if (filasAct.length === 0) {
      showToast('error', 'No hay máquinas activas');
      return;
    }
    const fecha = formatearFechaTurno(metaTurno.fecha);
    const tipo = metaTurno.tipo ? etiquetaTurno(metaTurno.tipo) : '';
    const nombre = `activacion_${metaTurno.fecha || 'sin-fecha'}_${metaTurno.tipo || 'sin-turno'}`;
    imprimirActivacion(filasAct, fecha, tipo, nombre);
  };

  return (
    <div className="space-y-3.5 sm:space-y-5">
      {/* Top Header Card matching Excel Title Banner */}
      <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-3 sm:p-4 shadow-sm space-y-3">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 pb-2.5 border-b border-[#2D3748]">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-[#F2A93B] flex items-center justify-center text-[#12161F] font-bold">
              <FileSpreadsheet size={18} />
            </div>
            <div>
              <h2 className="m-0 text-sm sm:text-base font-bold text-white font-['Oswald',sans-serif] tracking-wide uppercase">
                {tituloTurno}
              </h2>
              <div className="flex items-center gap-3 text-xs text-[#CBD5E0] mt-0.5">
                <span>
                  Total:{' '}
                  <strong className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace] text-sm tabular-nums">
                    {totalGeneral.toLocaleString('es-PE')} u
                  </strong>
                </span>
                {mermaGeneral > 0 && (
                  <span>
                    Merma:{' '}
                    <strong className="text-[#E53E3E] font-['IBM_Plex_Mono',monospace] text-sm tabular-nums">
                      {mermaGeneral.toLocaleString('es-PE')} u
                    </strong>
                  </span>
                )}
                <span>{filas.length} registros</span>
              </div>
            </div>
          </div>

          {/* Quick Action buttons */}
          <div className="flex items-center gap-1.5 flex-wrap">
            <button
              type="button"
              onClick={handleExportarExcel}
              className="flex items-center gap-1 bg-[#F2A93B] text-[#12161F] py-1.5 px-2.5 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:brightness-105 active:scale-95 transition-all shadow-sm"
              title="Descargar archivo Excel formateado con estilos"
            >
              <FileSpreadsheet size={13} />
              <span>Excel</span>
            </button>
            <button
              type="button"
              onClick={handleExportarMigracion}
              className="flex items-center gap-1 bg-[#3A2E1B] border border-[#F2A93B]/40 text-[#F6AD55] py-1.5 px-2.5 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:brightness-110 active:scale-95 transition-all shadow-sm"
              title="Exportar archivo de migración para SAP"
            >
              <Download size={13} />
              <span>SAP</span>
            </button>
            <button
              type="button"
              onClick={handleActivacionExcel}
              className="flex items-center gap-1 bg-[#232A38] text-[#CBD5E0] border border-[#4A5568] hover:text-white py-1.5 px-2.5 rounded-lg text-xs font-semibold cursor-pointer active:scale-95 transition-all"
              title="Exportar activación de máquinas"
            >
              <FileSpreadsheet size={13} />
              <span>Activación</span>
            </button>
            <button
              type="button"
              onClick={handleActivacionPdf}
              className="flex items-center gap-1 bg-[#232A38] text-[#CBD5E0] border border-[#4A5568] hover:text-white py-1.5 px-2.5 rounded-lg text-xs font-semibold cursor-pointer active:scale-95 transition-all"
              title="Imprimir activación"
            >
              <Printer size={13} />
              <span>Imprimir</span>
            </button>
          </div>
        </div>
      </div>

      {filas.length === 0 ? (
        <div className="border border-dashed border-[#4A5568] rounded-xl p-8 text-center text-xs sm:text-sm text-[#CBD5E0] bg-[#1A202C]">
          No hay máquinas activas con producción en este turno. Ve a la pestaña <strong>Máquinas</strong> o <strong>Producción</strong>.
        </div>
      ) : (
        <div className="space-y-4">
          {AREAS_ORDEN.map((areaDef) => {
            const filasDeArea = filas.filter((f) => (f.area || '') === areaDef.key);
            if (filasDeArea.length === 0) return null;

            const totalArea = filasDeArea.reduce((acc, f) => acc + (f.total || 0), 0);
            const mermaArea = filasDeArea.reduce((acc, f) => acc + (f.merma || 0), 0);

            return (
              <div
                key={areaDef.key || 'sin-area'}
                className="bg-[#1A202C] border border-[#2D3748] rounded-xl overflow-hidden shadow-sm"
              >
                {/* Area Header Bar (Exactly formatted like Excel's Area Bar) */}
                <div className="bg-[#D97706] text-white px-3.5 py-2 flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <Layers size={15} className="text-white/80" />
                    <span className="font-['Oswald',sans-serif] font-bold text-xs sm:text-sm tracking-wider uppercase">
                      {areaDef.label}
                    </span>
                    <span className="text-[11px] bg-black/20 text-white font-['IBM_Plex_Mono',monospace] px-1.5 py-0.2 rounded font-semibold">
                      {filasDeArea.length}
                    </span>
                  </div>
                  <div className="flex items-center gap-3 text-xs font-['IBM_Plex_Mono',monospace] font-bold">
                    <span>{totalArea.toLocaleString('es-PE')} u</span>
                    {mermaArea > 0 && <span className="text-red-200">-{mermaArea} merma</span>}
                  </div>
                </div>

                {/* MOBILE VIEW: Formatted Spreadsheet List */}
                <div className="block md:hidden divide-y divide-[#2D3748]">
                  {filasDeArea.map((f, idx) => (
                    <div
                      key={idx}
                      className="p-2.5 sm:p-3 hover:bg-[#232A38]/50 transition-colors space-y-1.5"
                    >
                      {/* Row Top: Machine name, Machine code, OP chip, SAP */}
                      <div className="flex items-center justify-between gap-2">
                        <div className="flex items-center gap-1.5 flex-wrap min-w-0">
                          <span className="font-['IBM_Plex_Mono',monospace] font-bold text-white text-xs sm:text-sm">
                            {f.maquina}
                          </span>
                          {f.codigoMaquina && (
                            <span className="text-[9.5px] font-bold px-1.5 py-0.2 rounded bg-[#232A38] border border-[#4A5568] text-[#CBD5E0] font-['IBM_Plex_Mono',monospace]">
                              {f.codigoMaquina}
                            </span>
                          )}
                          {f.op && (
                            <span className="text-[10px] font-bold px-1.5 py-0.2 rounded bg-[#3A2E1B] text-[#F6AD55] border border-[#F2A93B]/30 font-['IBM_Plex_Mono',monospace]">
                              {f.op}
                            </span>
                          )}
                          {f.sap && (
                            <span className="text-[9.5px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace]">
                              SAP: {f.sap}
                            </span>
                          )}
                        </div>

                        {/* Total units */}
                        <div className="text-right shrink-0">
                          <span className="font-['IBM_Plex_Mono',monospace] font-bold text-[#F2A93B] text-sm tabular-nums">
                            {f.total.toLocaleString('es-PE')} u
                          </span>
                          {f.merma > 0 && (
                            <div className="text-[10px] text-[#E53E3E] font-['IBM_Plex_Mono',monospace]">
                              -{f.merma} merma
                            </div>
                          )}
                        </div>
                      </div>

                      {/* Row Middle: Product Description */}
                      <div className="text-xs text-[#CBD5E0] leading-snug">
                        {f.producto || <span className="text-[#A0AEC0] italic">Sin OP / parada</span>}
                      </div>

                      {/* Row Bottom: Observations if any */}
                      {f.observaciones && (
                        <div className="text-[11px] text-[#A0AEC0] bg-[#232A38] px-2 py-1 rounded border border-[#2D3748] truncate">
                          {f.observaciones}
                        </div>
                      )}
                    </div>
                  ))}
                </div>

                {/* DESKTOP VIEW: Formatted Excel Spreadsheet Table */}
                <div className="hidden md:block overflow-x-auto">
                  <table className="w-full min-w-[760px] text-xs border-collapse">
                    <thead>
                      <tr className="bg-[#232A38] border-b border-[#2D3748]">
                        <th className="py-2 px-3 text-left font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-36">
                          Máquina
                        </th>
                        <th className="py-2 px-2 text-left font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-20">
                          Código
                        </th>
                        <th className="py-2 px-2.5 text-left font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-24">
                          OP's
                        </th>
                        <th className="py-2 px-2.5 text-left font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-24">
                          SAP
                        </th>
                        <th className="py-2 px-3 text-left font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider">
                          Producto
                        </th>
                        <th className="py-2 px-3 text-right font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-28">
                          Total Producción
                        </th>
                        <th className="py-2 px-2.5 text-right font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-20">
                          Merma
                        </th>
                        <th className="py-2 px-3 text-left font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-44">
                          Observaciones
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#2D3748]">
                      {filasDeArea.map((f, idx) => (
                        <tr
                          key={idx}
                          className="hover:bg-[#232A38]/60 transition-colors"
                        >
                          <td className="py-2 px-3 font-['IBM_Plex_Mono',monospace] font-bold text-white">
                            {f.maquina}
                          </td>
                          <td className="py-2 px-2 font-['IBM_Plex_Mono',monospace] text-[#CBD5E0]">
                            {f.codigoMaquina || '—'}
                          </td>
                          <td className="py-2 px-2.5 font-['IBM_Plex_Mono',monospace] font-bold text-[#F6AD55]">
                            {f.op || '—'}
                          </td>
                          <td className="py-2 px-2.5 font-['IBM_Plex_Mono',monospace] text-[#CBD5E0]">
                            {f.sap || '—'}
                          </td>
                          <td className="py-2 px-3 text-[#CBD5E0] max-w-[260px] truncate" title={f.producto}>
                            {f.producto || '—'}
                          </td>
                          <td className="py-2 px-3 font-['IBM_Plex_Mono',monospace] font-bold text-[#F2A93B] text-right tabular-nums">
                            {f.total.toLocaleString('es-PE')}
                          </td>
                          <td className="py-2 px-2.5 font-['IBM_Plex_Mono',monospace] text-[#E53E3E] text-right tabular-nums">
                            {f.merma || 0}
                          </td>
                          <td className="py-2 px-3 text-[#A0AEC0] text-[11px] max-w-[180px] truncate" title={f.observaciones}>
                            {f.observaciones || '—'}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Subtotal Row per Area (Formatted like Excel's Total Row) */}
                <div className="bg-[#232A38] border-t-2 border-[#2D3748] px-3.5 py-2 flex items-center justify-between text-xs font-bold">
                  <span className="font-['Oswald',sans-serif] uppercase text-white tracking-wider">
                    Total {areaDef.label}
                  </span>
                  <div className="flex items-center gap-4 font-['IBM_Plex_Mono',monospace]">
                    <span className="text-[#F2A93B] tabular-nums text-sm">
                      {totalArea.toLocaleString('es-PE')} u
                    </span>
                    {mermaArea > 0 && (
                      <span className="text-[#E53E3E] tabular-nums">
                        {mermaArea.toLocaleString('es-PE')} merma
                      </span>
                    )}
                  </div>
                </div>
              </div>
            );
          })}

          {/* Grand Total Summary Box at Bottom (Excel Final Total) */}
          <div className="bg-[#1A202C] border-2 border-[#F2A93B]/40 rounded-xl p-3 sm:p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-2 shadow-sm">
            <div className="flex items-center gap-2">
              <CheckCircle2 size={18} className="text-[#F2A93B]" />
              <span className="font-['Oswald',sans-serif] font-bold text-sm sm:text-base text-white uppercase tracking-wider">
                Total General de la Planta
              </span>
            </div>
            <div className="flex items-center gap-4 font-['IBM_Plex_Mono',monospace]">
              <span className="text-xs text-[#CBD5E0]">
                Producción:{' '}
                <strong className="text-base sm:text-lg text-[#F2A93B] font-bold tabular-nums">
                  {totalGeneral.toLocaleString('es-PE')} u
                </strong>
              </span>
              {mermaGeneral > 0 && (
                <span className="text-xs text-[#CBD5E0]">
                  Merma:{' '}
                  <strong className="text-sm sm:text-base text-[#E53E3E] font-bold tabular-nums">
                    {mermaGeneral.toLocaleString('es-PE')} u
                  </strong>
                </span>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
