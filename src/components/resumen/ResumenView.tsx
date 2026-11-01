import React, { useMemo, useState } from 'react';
import { useApp } from '../../context/AppContext';
import {
  FileSpreadsheet,
  Download,
  Printer,
  Calendar,
  Layers,
  CheckCircle2,
  Target,
  TrendingUp,
  Gauge,
  Award,
  ChevronDown,
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
import { FilaTurnoExcel, Maquina } from '../../types';

export const ResumenView: React.FC = () => {
  const { maquinas, turno, catalogo, metaTurno, showToast } = useApp();
  const [mostrarDetalleMaquinas, setMostrarDetalleMaquinas] = useState(true);

  const filas = useMemo(() => {
    return generarFilasTurno(maquinas, turno, catalogo);
  }, [maquinas, turno, catalogo]);

  const activas = useMemo(() => maquinas.filter((m) => m.activa), [maquinas]);

  const mapaMaquinas = useMemo(() => {
    const map = new Map<string, Maquina>();
    maquinas.forEach((m) => {
      map.set(m.nombre.trim().toLowerCase(), m);
      if (m.codigo) map.set(m.codigo.trim().toLowerCase(), m);
    });
    return map;
  }, [maquinas]);

  const prodPorMaquina = useMemo(() => {
    const map = new Map<string, number>();
    filas.forEach((f) => {
      const key = f.maquina.trim().toLowerCase();
      map.set(key, (map.get(key) || 0) + (f.total || 0));
    });
    return map;
  }, [filas]);

  const metricasMaquinas = useMemo(() => {
    return activas.map((m) => {
      const meta = m.meta || 0;
      const tieneMeta = meta > 0;
      const totalProd = prodPorMaquina.get(m.nombre.trim().toLowerCase()) || 0;
      const pct = tieneMeta ? (totalProd / meta) * 100 : 0;
      const estado = !tieneMeta
        ? ('sin_meta' as const)
        : pct >= 100
        ? ('eficiente' as const)
        : pct >= 80
        ? ('progreso' as const)
        : ('bajo' as const);
      const diferencia = tieneMeta ? totalProd - meta : 0;
      return {
        maquina: m,
        meta,
        tieneMeta,
        totalProd,
        pct,
        estado,
        diferencia,
      };
    });
  }, [activas, prodPorMaquina]);

  const metricasAreas = useMemo(() => {
    const areas = [
      { key: 'vasos_envases', label: 'Vasos y Envases' },
      { key: 'papel_carton', label: 'Papel y Cartón' },
    ];

    return areas.map((area) => {
      const maqsArea = metricasMaquinas.filter((mm) => mm.maquina.area === area.key);
      const conMeta = maqsArea.filter((mm) => mm.tieneMeta);
      const metaTotal = conMeta.reduce((acc, mm) => acc + mm.meta, 0);
      const prodTotal = maqsArea.reduce((acc, mm) => acc + mm.totalProd, 0);
      const prodConMeta = conMeta.reduce((acc, mm) => acc + mm.totalProd, 0);
      const pct = metaTotal > 0 ? (prodConMeta / metaTotal) * 100 : 0;
      const eficientes = conMeta.filter((mm) => mm.estado === 'eficiente').length;
      return {
        ...area,
        totalMaquinas: maqsArea.length,
        conMetaCount: conMeta.length,
        metaTotal,
        prodTotal,
        pct,
        eficientes,
      };
    });
  }, [metricasMaquinas]);

  const resumenEficiencia = useMemo(() => {
    const conMeta = metricasMaquinas.filter((mm) => mm.tieneMeta);
    const eficientesCount = conMeta.filter((mm) => mm.estado === 'eficiente').length;
    const enProgresoCount = conMeta.filter((mm) => mm.estado === 'progreso').length;
    const bajoCount = conMeta.filter((mm) => mm.estado === 'bajo').length;
    const maquinasSinMeta = metricasMaquinas.length - conMeta.length;
    const metaAcumulada = conMeta.reduce((acc, mm) => acc + mm.meta, 0);
    const prodAcumulada = conMeta.reduce((acc, mm) => acc + mm.totalProd, 0);
    const pctGlobal = metaAcumulada > 0 ? (prodAcumulada / metaAcumulada) * 100 : 0;
    const diferenciaGlobal = prodAcumulada - metaAcumulada;

    return {
      maquinasConMeta: conMeta.length,
      maquinasSinMeta,
      eficientesCount,
      enProgresoCount,
      bajoCount,
      metaAcumulada,
      prodAcumulada,
      pctGlobal,
      diferenciaGlobal,
    };
  }, [metricasMaquinas]);

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

      {/* Dashboard de Eficiencia y Métricas (Turno Activo) */}
      {resumenEficiencia.maquinasConMeta > 0 && (
        <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-3 sm:p-5 shadow-sm space-y-4">
          {/* Dashboard Header with Toggle */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 pb-3 border-b border-[#2D3748]">
            <div className="flex items-center gap-2.5">
              <div className="w-8 h-8 rounded-lg bg-[#3A2E1B] border border-[#F2A93B]/40 flex items-center justify-center text-[#F6AD55] shrink-0">
                <Gauge size={18} />
              </div>
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <h3 className="m-0 text-sm sm:text-base font-bold text-white font-['Oswald',sans-serif] tracking-wide uppercase">
                    Dashboard de Eficiencia y Rendimiento
                  </h3>
                  <span className="text-[10px] font-sans font-bold bg-[#F2A93B]/20 text-[#F6AD55] border border-[#F2A93B]/40 px-2 py-0.2 rounded-full uppercase">
                    Turno {etiquetaTurno(metaTurno.tipo || 'dia')}
                  </span>
                </div>
                <p className="m-0 text-xs text-[#CBD5E0]">
                  Medición basada en la columna CANTIDAD ({resumenEficiencia.maquinasConMeta} máquinas medidas • {resumenEficiencia.maquinasSinMeta} sin meta)
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={() => setMostrarDetalleMaquinas((prev) => !prev)}
              className="flex items-center gap-1.5 text-xs font-semibold py-1.5 px-3 rounded-lg bg-[#232A38] hover:bg-[#2D3748] text-[#CBD5E0] hover:text-white border border-[#4A5568] transition-all cursor-pointer self-start sm:self-auto"
            >
              <span>{mostrarDetalleMaquinas ? 'Ocultar Detalle Máquinas' : 'Ver Detalle Máquinas'}</span>
              <ChevronDown
                size={14}
                className={`transition-transform duration-200 ${mostrarDetalleMaquinas ? 'rotate-180' : ''}`}
              />
            </button>
          </div>

          {/* KPI Summary Cards */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-3 sm:gap-4">
            {/* Card 1: Eficiencia Global de Planta */}
            <div className="bg-[#232A38] border border-[#2D3748] rounded-xl p-3 sm:p-4 flex flex-col justify-between space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-xs uppercase font-['Oswald',sans-serif] font-bold text-[#CBD5E0] tracking-wider flex items-center gap-1.5">
                  <Target size={14} className="text-[#F2A93B]" />
                  Eficiencia Global Planta
                </span>
                <span
                  className={`text-[10px] font-bold px-2 py-0.5 rounded-full font-['IBM_Plex_Mono',monospace] border ${
                    resumenEficiencia.pctGlobal >= 100
                      ? 'bg-[#38A169]/20 text-[#48BB78] border-[#38A169]/50'
                      : resumenEficiencia.pctGlobal >= 80
                      ? 'bg-[#F2A93B]/20 text-[#F6AD55] border-[#F2A93B]/50'
                      : 'bg-[#E53E3E]/20 text-[#FC8181] border-[#E53E3E]/50'
                  }`}
                >
                  {resumenEficiencia.pctGlobal >= 100
                    ? '✓ EFICIENTE'
                    : resumenEficiencia.pctGlobal >= 80
                    ? '⏳ EN PROGRESO'
                    : '⚠ BAJO META'}
                </span>
              </div>

              <div>
                <div className="flex items-baseline gap-2">
                  <span
                    className={`text-2xl sm:text-3xl font-black font-['IBM_Plex_Mono',monospace] tracking-tight tabular-nums ${
                      resumenEficiencia.pctGlobal >= 100
                        ? 'text-[#48BB78]'
                        : resumenEficiencia.pctGlobal >= 80
                        ? 'text-[#F6AD55]'
                        : 'text-[#FC8181]'
                    }`}
                  >
                    {resumenEficiencia.pctGlobal.toFixed(1)}%
                  </span>
                  <span className="text-xs text-[#CBD5E0] font-['IBM_Plex_Mono',monospace]">
                    de la meta global
                  </span>
                </div>

                {/* Global Progress Bar */}
                <div className="w-full bg-[#12161F] rounded-full h-2 mt-2 overflow-hidden border border-[#2D3748]">
                  <div
                    className={`h-full rounded-full transition-all duration-300 ${
                      resumenEficiencia.pctGlobal >= 100
                        ? 'bg-[#38A169]'
                        : resumenEficiencia.pctGlobal >= 80
                        ? 'bg-[#F2A93B]'
                        : 'bg-[#E53E3E]'
                    }`}
                    style={{ width: `${Math.min(100, Math.max(0, resumenEficiencia.pctGlobal))}%` }}
                  />
                </div>
              </div>

              <div className="flex items-center justify-between text-xs text-[#CBD5E0] font-['IBM_Plex_Mono',monospace] pt-1.5 border-t border-[#2D3748]/70">
                <span>
                  Prod: <strong className="text-white">{resumenEficiencia.prodAcumulada.toLocaleString('es-PE')}</strong> /{' '}
                  <span className="text-[#A0AEC0]">{resumenEficiencia.metaAcumulada.toLocaleString('es-PE')} u</span>
                </span>
                <span
                  className={
                    resumenEficiencia.diferenciaGlobal >= 0 ? 'text-[#48BB78] font-bold' : 'text-[#FC8181] font-bold'
                  }
                >
                  {resumenEficiencia.diferenciaGlobal >= 0
                    ? `+${resumenEficiencia.diferenciaGlobal.toLocaleString('es-PE')} u`
                    : `${resumenEficiencia.diferenciaGlobal.toLocaleString('es-PE')} u`}
                </span>
              </div>
            </div>

            {/* Card 2: Semáforo de Cumplimiento de Máquinas */}
            <div className="bg-[#232A38] border border-[#2D3748] rounded-xl p-3 sm:p-4 flex flex-col justify-between space-y-2">
              <span className="text-xs uppercase font-['Oswald',sans-serif] font-bold text-[#CBD5E0] tracking-wider flex items-center gap-1.5">
                <Award size={14} className="text-[#F2A93B]" />
                Cumplimiento de Máquinas
              </span>

              <div className="grid grid-cols-2 gap-2 text-center font-['IBM_Plex_Mono',monospace]">
                <div className="bg-[#1A202C] p-2 rounded-lg border border-[#38A169]/30">
                  <div className="text-[10px] text-[#A0AEC0] font-sans uppercase">Eficientes (≥100%)</div>
                  <div className="text-lg font-bold text-[#48BB78]">{resumenEficiencia.eficientesCount}</div>
                </div>
                <div className="bg-[#1A202C] p-2 rounded-lg border border-[#F2A93B]/30">
                  <div className="text-[10px] text-[#A0AEC0] font-sans uppercase">En Meta (80-99%)</div>
                  <div className="text-lg font-bold text-[#F6AD55]">{resumenEficiencia.enProgresoCount}</div>
                </div>
                <div className="bg-[#1A202C] p-2 rounded-lg border border-[#E53E3E]/30">
                  <div className="text-[10px] text-[#A0AEC0] font-sans uppercase">Bajo Meta (&lt;80%)</div>
                  <div className="text-lg font-bold text-[#FC8181]">{resumenEficiencia.bajoCount}</div>
                </div>
                <div className="bg-[#1A202C] p-2 rounded-lg border border-[#2D3748]">
                  <div className="text-[10px] text-[#718096] font-sans uppercase">Sin Meta (0)</div>
                  <div className="text-lg font-bold text-[#A0AEC0]">{resumenEficiencia.maquinasSinMeta}</div>
                </div>
              </div>

              <div className="text-[11px] text-[#A0AEC0] text-center pt-1.5 border-t border-[#2D3748]/70">
                {resumenEficiencia.eficientesCount} de {resumenEficiencia.maquinasConMeta} máquinas alcanzaron su cuota diaria
              </div>
            </div>

            {/* Card 3: Eficiencia por Área Operativa */}
            <div className="bg-[#232A38] border border-[#2D3748] rounded-xl p-3 sm:p-4 flex flex-col justify-between space-y-2">
              <span className="text-xs uppercase font-['Oswald',sans-serif] font-bold text-[#CBD5E0] tracking-wider flex items-center gap-1.5">
                <Layers size={14} className="text-[#F2A93B]" />
                Eficiencia por Área
              </span>

              <div className="space-y-2.5">
                {metricasAreas.map((area) => (
                  <div key={area.key} className="bg-[#1A202C] p-2 rounded-lg border border-[#2D3748] space-y-1">
                    <div className="flex items-center justify-between text-xs">
                      <span className="font-['Oswald',sans-serif] font-bold text-white tracking-wide">
                        {area.label}
                      </span>
                      <span
                        className={`font-['IBM_Plex_Mono',monospace] font-bold ${
                          area.pct >= 100
                            ? 'text-[#48BB78]'
                            : area.pct >= 80
                            ? 'text-[#F6AD55]'
                            : area.metaTotal > 0
                            ? 'text-[#FC8181]'
                            : 'text-[#A0AEC0]'
                        }`}
                      >
                        {area.metaTotal > 0 ? `${area.pct.toFixed(1)}%` : 'Sin metas'}
                      </span>
                    </div>

                    <div className="w-full bg-[#12161F] rounded-full h-1.5 overflow-hidden border border-[#2D3748]">
                      <div
                        className={`h-full rounded-full transition-all duration-300 ${
                          area.pct >= 100 ? 'bg-[#38A169]' : area.pct >= 80 ? 'bg-[#F2A93B]' : 'bg-[#E53E3E]'
                        }`}
                        style={{ width: `${Math.min(100, Math.max(0, area.pct))}%` }}
                      />
                    </div>

                    <div className="flex items-center justify-between text-[10.5px] text-[#CBD5E0] font-['IBM_Plex_Mono',monospace]">
                      <span>
                        {area.prodTotal.toLocaleString('es-PE')} / {area.metaTotal.toLocaleString('es-PE')} u
                      </span>
                      <span>
                        {area.eficientes}/{area.conMetaCount} eficientes
                      </span>
                    </div>
                  </div>
                ))}
              </div>

              <div className="text-[11px] text-[#A0AEC0] text-center pt-0.5">
                Métricas calculadas sobre máquinas con meta activa
              </div>
            </div>
          </div>

          {/* Machine-by-Machine Detailed Breakdown Matrix (Toggleable) */}
          {mostrarDetalleMaquinas && (
            <div className="pt-2 border-t border-[#2D3748] space-y-2.5">
              <div className="flex items-center justify-between">
                <span className="text-xs uppercase font-['Oswald',sans-serif] font-bold text-white tracking-wider flex items-center gap-1.5">
                  <TrendingUp size={14} className="text-[#F2A93B]" />
                  Rendimiento y Cuotas por Máquina ({metricasMaquinas.length})
                </span>
                <span className="text-[11px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace]">
                  Meta diaria (CANTIDAD) vs Total Producido
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-2.5">
                {metricasMaquinas.map((mm) => {
                  const pctFormatted = mm.pct >= 100 ? Math.round(mm.pct) : mm.pct.toFixed(1);
                  return (
                    <div
                      key={mm.maquina.id}
                      className={`p-2.5 rounded-xl border transition-all ${
                        mm.estado === 'eficiente'
                          ? 'bg-[#232A38] border-[#38A169]/40'
                          : mm.estado === 'progreso'
                          ? 'bg-[#232A38] border-[#F2A93B]/40'
                          : mm.estado === 'bajo'
                          ? 'bg-[#232A38] border-[#E53E3E]/40'
                          : 'bg-[#1E2430]/60 border-[#2D3748]'
                      }`}
                    >
                      <div className="flex items-center justify-between gap-1 mb-1">
                        <span className="font-['IBM_Plex_Mono',monospace] font-bold text-white text-xs truncate">
                          {mm.maquina.nombre}
                        </span>
                        {mm.maquina.codigo && (
                          <span className="text-[9.5px] font-bold px-1.5 py-0.2 rounded bg-[#1A202C] text-[#CBD5E0] font-['IBM_Plex_Mono',monospace] border border-[#2D3748]">
                            {mm.maquina.codigo}
                          </span>
                        )}
                      </div>

                      {mm.tieneMeta ? (
                        <div className="space-y-1.5">
                          <div className="flex items-center justify-between text-xs font-['IBM_Plex_Mono',monospace]">
                            <span className="text-[#F2A93B] font-bold">
                              {mm.totalProd.toLocaleString('es-PE')} u
                            </span>
                            <span className="text-[#A0AEC0] text-[11px]">
                              / {mm.meta.toLocaleString('es-PE')} meta
                            </span>
                          </div>

                          <div className="w-full bg-[#12161F] rounded-full h-1.5 overflow-hidden border border-[#2D3748]">
                            <div
                              className={`h-full rounded-full transition-all duration-300 ${
                                mm.estado === 'eficiente'
                                  ? 'bg-[#38A169]'
                                  : mm.estado === 'progreso'
                                  ? 'bg-[#F2A93B]'
                                  : 'bg-[#E53E3E]'
                              }`}
                              style={{ width: `${Math.min(100, Math.max(0, mm.pct))}%` }}
                            />
                          </div>

                          <div className="flex items-center justify-between text-[10.5px] pt-0.5">
                            <span
                              className={`font-bold px-1.5 py-0.2 rounded font-['IBM_Plex_Mono',monospace] ${
                                mm.estado === 'eficiente'
                                  ? 'bg-[#38A169]/20 text-[#48BB78]'
                                  : mm.estado === 'progreso'
                                  ? 'bg-[#F2A93B]/20 text-[#F6AD55]'
                                  : 'bg-[#E53E3E]/20 text-[#FC8181]'
                              }`}
                            >
                              {pctFormatted}% {mm.estado === 'eficiente' ? '✓' : mm.estado === 'progreso' ? '⏳' : '⚠'}
                            </span>
                            <span
                              className={`font-['IBM_Plex_Mono',monospace] ${
                                mm.diferencia >= 0 ? 'text-[#48BB78]' : 'text-[#FC8181]'
                              }`}
                            >
                              {mm.diferencia >= 0
                                ? `+${mm.diferencia.toLocaleString('es-PE')}`
                                : `${mm.diferencia.toLocaleString('es-PE')}`}
                            </span>
                          </div>
                        </div>
                      ) : (
                        <div className="py-2 text-center text-[10.5px] text-[#718096] italic bg-[#1A202C]/60 rounded-lg">
                          Meta: 0 u — Sin medición
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      )}

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
                  {filasDeArea.map((f, idx) => {
                    const maq = mapaMaquinas.get(f.maquina.trim().toLowerCase()) || (f.codigoMaquina ? mapaMaquinas.get(f.codigoMaquina.trim().toLowerCase()) : undefined);
                    const meta = maq?.meta || 0;
                    const totalMaquina = prodPorMaquina.get(f.maquina.trim().toLowerCase()) || f.total || 0;
                    const tieneMeta = meta > 0;
                    const pct = tieneMeta ? (totalMaquina / meta) * 100 : 0;
                    const pctFormatted = pct >= 100 ? Math.round(pct) : pct.toFixed(1);
                    const esEficiente = tieneMeta && pct >= 100;
                    const esCercano = tieneMeta && pct >= 80 && pct < 100;

                    return (
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

                        {/* Row Efficiency indicator */}
                        {tieneMeta ? (
                          <div className="flex items-center justify-between text-[11px] pt-1.5 border-t border-[#2D3748]/70">
                            <span className="text-[#A0AEC0] font-['IBM_Plex_Mono',monospace]">
                              Meta: <strong className="text-white">{meta.toLocaleString('es-PE')} u</strong>
                            </span>
                            <span
                              className={`font-bold px-2 py-0.5 rounded-full font-['IBM_Plex_Mono',monospace] text-[10px] border flex items-center gap-1 ${
                                esEficiente
                                  ? 'bg-[#38A169]/20 text-[#48BB78] border-[#38A169]/50'
                                  : esCercano
                                  ? 'bg-[#F2A93B]/20 text-[#F6AD55] border-[#F2A93B]/50'
                                  : 'bg-[#E53E3E]/20 text-[#FC8181] border-[#E53E3E]/50'
                              }`}
                            >
                              <span
                                className={`w-1.5 h-1.5 rounded-full ${
                                  esEficiente ? 'bg-[#48BB78]' : esCercano ? 'bg-[#F6AD55]' : 'bg-[#FC8181]'
                                }`}
                              />
                              <span>{pctFormatted}%</span>
                              <span className="font-sans text-[9px] uppercase font-semibold">
                                {esEficiente ? 'Eficiente' : esCercano ? 'En meta' : 'Bajo meta'}
                              </span>
                            </span>
                          </div>
                        ) : (
                          <div className="text-[10px] text-[#718096] pt-1 border-t border-[#2D3748]/40 italic flex items-center justify-between font-['IBM_Plex_Mono',monospace]">
                            <span>Meta: 0 u</span>
                            <span>Sin meta (no se mide)</span>
                          </div>
                        )}

                        {/* Row Bottom: Observations if any */}
                        {f.observaciones && (
                          <div className="text-[11px] text-[#A0AEC0] bg-[#232A38] px-2 py-1 rounded border border-[#2D3748] truncate">
                            {f.observaciones}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>

                {/* DESKTOP VIEW: Formatted Excel Spreadsheet Table */}
                <div className="hidden md:block overflow-x-auto">
                  <table className="w-full min-w-[880px] text-xs border-collapse">
                    <thead>
                      <tr className="bg-[#232A38] border-b border-[#2D3748]">
                        <th className="py-2 px-3 text-left font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-32">
                          Máquina
                        </th>
                        <th className="py-2 px-2 text-left font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-20">
                          Código
                        </th>
                        <th className="py-2 px-2.5 text-left font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-24">
                          OP's
                        </th>
                        <th className="py-2 px-2 text-left font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-20">
                          SAP
                        </th>
                        <th className="py-2 px-3 text-left font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider">
                          Producto
                        </th>
                        <th className="py-2 px-2.5 text-right font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-24">
                          Meta Diaria
                        </th>
                        <th className="py-2 px-3 text-right font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-28">
                          Total Prod.
                        </th>
                        <th className="py-2 px-2.5 text-center font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-28">
                          Eficiencia
                        </th>
                        <th className="py-2 px-2.5 text-right font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-20">
                          Merma
                        </th>
                        <th className="py-2 px-3 text-left font-['Oswald',sans-serif] uppercase font-bold text-[#CBD5E0] tracking-wider w-40">
                          Observaciones
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#2D3748]">
                      {filasDeArea.map((f, idx) => {
                        const maq = mapaMaquinas.get(f.maquina.trim().toLowerCase()) || (f.codigoMaquina ? mapaMaquinas.get(f.codigoMaquina.trim().toLowerCase()) : undefined);
                        const meta = maq?.meta || 0;
                        const totalMaquina = prodPorMaquina.get(f.maquina.trim().toLowerCase()) || f.total || 0;
                        const tieneMeta = meta > 0;
                        const pct = tieneMeta ? (totalMaquina / meta) * 100 : 0;
                        const pctFormatted = pct >= 100 ? Math.round(pct) : pct.toFixed(1);
                        const esEficiente = tieneMeta && pct >= 100;
                        const esCercano = tieneMeta && pct >= 80 && pct < 100;

                        return (
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
                            <td className="py-2 px-2 font-['IBM_Plex_Mono',monospace] text-[#CBD5E0]">
                              {f.sap || '—'}
                            </td>
                            <td className="py-2 px-3 text-[#CBD5E0] max-w-[240px] truncate" title={f.producto}>
                              {f.producto || '—'}
                            </td>
                            <td className="py-2 px-2.5 font-['IBM_Plex_Mono',monospace] text-right text-xs">
                              {tieneMeta ? (
                                <span className="text-white font-semibold tabular-nums">
                                  {meta.toLocaleString('es-PE')}
                                </span>
                              ) : (
                                <span className="text-[#718096] text-[11px]">0</span>
                              )}
                            </td>
                            <td className="py-2 px-3 font-['IBM_Plex_Mono',monospace] font-bold text-[#F2A93B] text-right tabular-nums">
                              {f.total.toLocaleString('es-PE')}
                            </td>
                            <td className="py-2 px-2.5 text-center">
                              {tieneMeta ? (
                                <span
                                  className={`inline-flex items-center gap-1 font-bold font-['IBM_Plex_Mono',monospace] px-2 py-0.5 rounded-full text-[10px] border ${
                                    esEficiente
                                      ? 'bg-[#38A169]/20 text-[#48BB78] border-[#38A169]/50'
                                      : esCercano
                                      ? 'bg-[#F2A93B]/20 text-[#F6AD55] border-[#F2A93B]/50'
                                      : 'bg-[#E53E3E]/20 text-[#FC8181] border-[#E53E3E]/50'
                                  }`}
                                  title={`Producido: ${totalMaquina.toLocaleString('es-PE')} u / Meta: ${meta.toLocaleString('es-PE')} u`}
                                >
                                  <span
                                    className={`w-1.5 h-1.5 rounded-full ${
                                      esEficiente ? 'bg-[#48BB78]' : esCercano ? 'bg-[#F6AD55]' : 'bg-[#FC8181]'
                                    }`}
                                  />
                                  <span>{pctFormatted}%</span>
                                </span>
                              ) : (
                                <span className="text-[#718096] text-[10px] italic">Sin meta</span>
                              )}
                            </td>
                            <td className="py-2 px-2.5 font-['IBM_Plex_Mono',monospace] text-[#E53E3E] text-right tabular-nums">
                              {f.merma || 0}
                            </td>
                            <td className="py-2 px-3 text-[#A0AEC0] text-[11px] max-w-[180px] truncate" title={f.observaciones}>
                              {f.observaciones || '—'}
                            </td>
                          </tr>
                        );
                      })}
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
