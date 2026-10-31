import React, { useState, useRef, useMemo } from 'react';
import { useApp } from '../../context/AppContext';
import { Plus, Upload, RefreshCw, Lock, Unlock, Copy, RotateCcw, XCircle, Trash2, Search, X, ChevronDown, ChevronUp } from 'lucide-react';
import * as XLSX from 'xlsx';
import { parsearCodigosMaquinas } from '../../services/excel';
import { copiarAlPortapapeles } from '../../services/reports';
import { STORAGE_KEYS } from '../../services/storage';
import { Maquina } from '../../types';

const DEFAULT_URL_SHEETS_MAQUINAS =
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vTVxCxG3XBMgSSedTqw7xWb7s7qxbYB3ZavS8MP9bx2lrwkzlK8oNQRSoFLw_sjjY68b1V0vNUaCb7B/pub?gid=770560989&single=true&output=csv';

export const MaquinasView: React.FC = () => {
  const { maquinas, setMaquinas, turno, showToast } = useApp();

  const [nombreInput, setNombreInput] = useState('');
  const [busqueda, setBusqueda] = useState('');
  const [panelAdminOpen, setPanelAdminOpen] = useState(false);
  const [filtroArea, setFiltroArea] = useState<'todas' | 'activas' | 'vasos_envases' | 'papel_carton' | 'no_asignadas'>('todas');
  const [actualizando, setActualizando] = useState(false);

  // Backup / Restore panels
  const [panelRespaldoOpen, setPanelRespaldoOpen] = useState(false);
  const [panelRestaurarOpen, setPanelRestaurarOpen] = useState(false);
  const [textoRespaldo, setTextoRespaldo] = useState('');
  const [textoRestaurar, setTextoRestaurar] = useState('');

  const fileInputRef = useRef<HTMLInputElement>(null);

  const hayBloqueadas = useMemo(() => {
    return maquinas.some((m) => m.bloqueada);
  }, [maquinas]);

  const activasCount = useMemo(() => {
    return maquinas.filter((m) => m.activa).length;
  }, [maquinas]);

  const vasosCount = useMemo(() => {
    return maquinas.filter((m) => m.area === 'vasos_envases').length;
  }, [maquinas]);

  const papelCount = useMemo(() => {
    return maquinas.filter((m) => m.area === 'papel_carton').length;
  }, [maquinas]);

  const noAsignadasCount = useMemo(() => {
    return maquinas.filter((m) => !m.area || (m.area !== 'vasos_envases' && m.area !== 'papel_carton')).length;
  }, [maquinas]);

  const handleAgregarMaquinas = () => {
    const nombres = nombreInput
      .split(/[\n,]+/)
      .map((n) => n.trim())
      .filter(Boolean);
    if (nombres.length === 0) {
      showToast('error', 'Ponle un nombre a la máquina');
      return;
    }

    const nuevas: Maquina[] = nombres.map((nombre) => ({
      id: Date.now() + '-' + Math.random().toString(36).slice(2, 7) + '-' + Math.random().toString(36).slice(2, 5),
      nombre,
      activa: false,
      area: '',
      codigo: '',
      bloqueada: false,
      tipoProducto: '',
    }));

    setMaquinas((prev) => [...nuevas, ...prev]);
    setNombreInput('');
    showToast('ok', nombres.length > 1 ? `${nombres.length} máquinas agregadas` : 'Máquina agregada');
  };

  const handleImportarExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const res = parsearCodigosMaquinas(wb, maquinas);
        setMaquinas(res.maquinas);
        showToast('ok', `${res.actualizadas} actualizadas, ${res.creadas} creadas`);
      } catch (err: any) {
        showToast('error', err.message || 'No se pudo leer el archivo');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const pedirUrlSheetsMaquinas = () => {
    let actual = '';
    try {
      actual = localStorage.getItem(STORAGE_KEYS.URL_SHEETS_MAQUINAS) || DEFAULT_URL_SHEETS_MAQUINAS;
    } catch {}
    const nueva = prompt('Link de la hoja de MÁQUINAS (columnas MAQUINA, CODIGO, TIPO, AREA, publicada como CSV):', actual);
    if (nueva === null) return;
    const trimmed = nueva.trim();
    if (!trimmed) {
      showToast('error', 'El link es obligatorio');
      return;
    }
    try {
      localStorage.setItem(STORAGE_KEYS.URL_SHEETS_MAQUINAS, trimmed);
      showToast('ok', 'Link guardado');
    } catch {}
  };

  const handleActualizarDesdeSheets = async () => {
    let url = '';
    try {
      url = localStorage.getItem(STORAGE_KEYS.URL_SHEETS_MAQUINAS) || DEFAULT_URL_SHEETS_MAQUINAS;
    } catch {
      url = DEFAULT_URL_SHEETS_MAQUINAS;
    }
    setActualizando(true);
    try {
      const resp = await fetch(url, { cache: 'no-store' });
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      const csvTexto = await resp.text();
      const wb = XLSX.read(csvTexto, { type: 'string' });
      const res = parsearCodigosMaquinas(wb, maquinas);
      setMaquinas(res.maquinas);
      showToast('ok', `${res.actualizadas} actualizadas, ${res.creadas} creadas`);
    } catch {
      showToast('error', 'No se pudo descargar la lista de máquinas. Revisa el link.');
    } finally {
      setActualizando(false);
    }
  };

  const handleToggleBloqueoTodo = () => {
    if (maquinas.length === 0) return;
    const nuevoEstado = !hayBloqueadas;
    setMaquinas((prev) => prev.map((m) => ({ ...m, bloqueada: nuevoEstado })));
    showToast('ok', nuevoEstado ? 'Todas las máquinas bloqueadas' : 'Todas las máquinas desbloqueadas');
  };

  const handleDesactivarTodas = () => {
    const activas = maquinas.filter((m) => m.activa);
    if (activas.length === 0) {
      showToast('error', 'No hay máquinas activas');
      return;
    }
    try {
      if (typeof window !== 'undefined' && typeof window.confirm === 'function') {
        const ok = window.confirm(`¿Desactivar las ${activas.length} máquina(s) activa(s)?`);
        if (!ok) return;
      }
    } catch {
      // Ignored in sandboxed environments
    }
    setMaquinas((prev) => prev.map((m) => ({ ...m, activa: false })));
    showToast('ok', 'Todas las máquinas desactivadas');
  };

  // Card click lock toggle with single toast execution
  const handleToggleLockCard = (id: string, ev?: React.MouseEvent) => {
    if (ev && (ev.target as HTMLElement).closest('button, input, select')) return;
    const m = maquinas.find((item) => item.id === id);
    if (!m) return;
    const nuevoBloqueo = !m.bloqueada;
    setMaquinas((prev) => prev.map((item) => (item.id === id ? { ...item, bloqueada: nuevoBloqueo } : item)));
    showToast('ok', nuevoBloqueo ? 'Máquina bloqueada' : 'Máquina desbloqueada para editar');
  };

  const handleUpdateMaquina = (id: string, updates: Partial<Maquina>) => {
    setMaquinas((prev) => prev.map((m) => (m.id === id ? { ...m, ...updates } : m)));
  };

  const handleEliminarMaquina = (id: string) => {
    try {
      if (typeof window !== 'undefined' && typeof window.confirm === 'function') {
        const ok = window.confirm('¿Eliminar esta máquina?');
        if (!ok) return;
      }
    } catch {
      // Ignored in sandboxed environments
    }
    setMaquinas((prev) => prev.filter((m) => m.id !== id));
    showToast('ok', 'Máquina eliminada');
  };

  // Backup / Restore
  const handleAbrirRespaldo = () => {
    if (maquinas.length === 0) {
      showToast('error', 'No hay máquinas que respaldar');
      return;
    }
    setPanelRestaurarOpen(false);
    setTextoRespaldo(JSON.stringify(maquinas));
    setPanelRespaldoOpen(true);
  };

  const handleCopiarRespaldo = async () => {
    const ok = await copiarAlPortapapeles(textoRespaldo);
    showToast(ok ? 'ok' : 'error', ok ? 'Copiado al portapapeles' : 'No se pudo copiar');
  };

  const handleConfirmarRestaurar = () => {
    const texto = textoRestaurar.trim();
    if (!texto) {
      showToast('error', 'Pega el texto de respaldo primero');
      return;
    }
    try {
      const datos = JSON.parse(texto);
      if (!Array.isArray(datos)) {
        showToast('error', 'El formato no corresponde a una lista de máquinas');
        return;
      }
      const idsExistentes = new Set(maquinas.map((m) => m.id));
      let agregadas = 0;
      const nuevasMaquinas: Maquina[] = [];
      datos.forEach((m: any) => {
        if (m && m.id && !idsExistentes.has(m.id)) {
          nuevasMaquinas.push(m);
          idsExistentes.add(m.id);
          agregadas++;
        }
      });
      setMaquinas((prev) => [...nuevasMaquinas, ...prev]);
      setPanelRestaurarOpen(false);
      showToast('ok', `${agregadas} máquina(s) restaurada(s)`);
    } catch {
      showToast('error', 'Ese texto no es válido');
    }
  };

  // Grouping, filtering and sorting by category
  const seccionGrupos = useMemo(() => {
    const q = busqueda.trim().toLowerCase();
    const coincideBusqueda = (m: Maquina) => {
      if (!q) return true;
      return (
        m.nombre.toLowerCase().includes(q) ||
        (m.codigo && m.codigo.toLowerCase().includes(q))
      );
    };

    const coincideEstado = (m: Maquina) => {
      if (filtroArea === 'activas') return m.activa;
      return true;
    };

    const candidatas = maquinas.filter((m) => coincideBusqueda(m) && coincideEstado(m));

    const vasos = candidatas
      .filter((m) => m.area === 'vasos_envases')
      .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' }));

    const papel = candidatas
      .filter((m) => m.area === 'papel_carton')
      .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' }));

    const noAsignadas = candidatas
      .filter((m) => !m.area || (m.area !== 'vasos_envases' && m.area !== 'papel_carton'))
      .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' }));

    const secciones = [
      {
        id: 'vasos_envases' as const,
        titulo: 'Vasos y Envases',
        dotColor: 'bg-[#F2A93B]',
        badgeColor: 'bg-[#F2A93B]/20 text-[#F6AD55] border border-[#F2A93B]/30',
        maquinas: vasos,
      },
      {
        id: 'papel_carton' as const,
        titulo: 'Papel y Cartón',
        dotColor: 'bg-[#63B3ED]',
        badgeColor: 'bg-[#63B3ED]/20 text-[#63B3ED] border border-[#63B3ED]/30',
        maquinas: papel,
      },
      {
        id: 'no_asignadas' as const,
        titulo: 'No Asignadas',
        dotColor: 'bg-[#A0AEC0]',
        badgeColor: 'bg-[#2D3748] text-[#CBD5E0] border border-[#4A5568]',
        maquinas: noAsignadas,
      },
    ];

    if (filtroArea === 'vasos_envases') return secciones.filter((s) => s.id === 'vasos_envases');
    if (filtroArea === 'papel_carton') return secciones.filter((s) => s.id === 'papel_carton');
    if (filtroArea === 'no_asignadas') return secciones.filter((s) => s.id === 'no_asignadas');

    return secciones;
  }, [maquinas, busqueda, filtroArea]);

  const totalFiltradasCount = useMemo(() => {
    return seccionGrupos.reduce((acc, s) => acc + s.maquinas.length, 0);
  }, [seccionGrupos]);

  return (
    <div className="space-y-2.5 sm:space-y-3.5">
      {/* Search Bar & Fast Filters Toolbar */}
      <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-2.5 sm:p-3 shadow-sm space-y-2">
        {/* Top line: Search Input & Toggle for Admin Tools */}
        <div className="flex items-center gap-2">
          <div className="relative flex-1">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-[#A0AEC0] pointer-events-none" />
            <input
              type="text"
              value={busqueda}
              onChange={(e) => setBusqueda(e.target.value)}
              placeholder="Buscar máquina por nombre o código…"
              className="w-full pl-9 pr-8 py-2 rounded-lg border border-[#4A5568] bg-[#232A38] text-xs sm:text-sm text-white placeholder-[#718096] outline-none focus:border-[#F2A93B] transition-colors"
            />
            {busqueda && (
              <button
                type="button"
                onClick={() => setBusqueda('')}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-[#A0AEC0] hover:text-white p-1 cursor-pointer"
                title="Limpiar búsqueda"
              >
                <X size={14} />
              </button>
            )}
          </div>

          <button
            type="button"
            onClick={() => setPanelAdminOpen((prev) => !prev)}
            className={`py-2 px-2.5 rounded-lg border text-xs font-bold flex items-center gap-1.5 transition-colors cursor-pointer shrink-0 ${
              panelAdminOpen
                ? 'bg-[#F2A93B] text-[#12161F] border-[#F2A93B]'
                : 'bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:text-white'
            }`}
            title="Agregar máquina o herramientas de importación"
          >
            <Plus size={14} />
            <span className="hidden sm:inline">Agregar / Opciones</span>
            <span className="sm:hidden">Más</span>
            {panelAdminOpen ? <ChevronUp size={13} /> : <ChevronDown size={13} />}
          </button>
        </div>

        {/* Filter Pills row + Quick Global Actions */}
        <div className="flex items-center justify-between gap-2 flex-wrap">
          <div className="flex items-center gap-1 overflow-x-auto pb-0.5 no-scrollbar max-w-full">
            <button
              type="button"
              onClick={() => setFiltroArea('todas')}
              className={`py-1 px-2.5 rounded-lg text-xs font-bold cursor-pointer whitespace-nowrap transition-colors ${
                filtroArea === 'todas'
                  ? 'bg-[#F2A93B] text-[#12161F] shadow-sm'
                  : 'bg-[#232A38] text-[#CBD5E0] hover:text-white border border-[#4A5568]/50'
              }`}
            >
              Todas ({maquinas.length})
            </button>
            <button
              type="button"
              onClick={() => setFiltroArea('activas')}
              className={`py-1 px-2.5 rounded-lg text-xs font-bold cursor-pointer whitespace-nowrap transition-colors flex items-center gap-1 ${
                filtroArea === 'activas'
                  ? 'bg-[#38A169] text-white shadow-sm'
                  : 'bg-[#232A38] text-[#CBD5E0] hover:text-white border border-[#4A5568]/50'
              }`}
            >
              <span className={`w-1.5 h-1.5 rounded-full ${activasCount > 0 ? 'bg-[#4ADE80]' : 'bg-[#718096]'}`} />
              <span>Activas ({activasCount})</span>
            </button>
            <button
              type="button"
              onClick={() => setFiltroArea('vasos_envases')}
              className={`py-1 px-2.5 rounded-lg text-xs font-bold cursor-pointer whitespace-nowrap transition-colors ${
                filtroArea === 'vasos_envases'
                  ? 'bg-[#F2A93B] text-[#12161F] shadow-sm'
                  : 'bg-[#232A38] text-[#CBD5E0] hover:text-white border border-[#4A5568]/50'
              }`}
            >
              Vasos y Envases ({vasosCount})
            </button>
            <button
              type="button"
              onClick={() => setFiltroArea('papel_carton')}
              className={`py-1 px-2.5 rounded-lg text-xs font-bold cursor-pointer whitespace-nowrap transition-colors ${
                filtroArea === 'papel_carton'
                  ? 'bg-[#63B3ED] text-[#12161F] shadow-sm'
                  : 'bg-[#232A38] text-[#CBD5E0] hover:text-white border border-[#4A5568]/50'
              }`}
            >
              Papel y Cartón ({papelCount})
            </button>
            <button
              type="button"
              onClick={() => setFiltroArea('no_asignadas')}
              className={`py-1 px-2.5 rounded-lg text-xs font-bold cursor-pointer whitespace-nowrap transition-colors ${
                filtroArea === 'no_asignadas'
                  ? 'bg-[#CBD5E0] text-[#12161F] shadow-sm'
                  : 'bg-[#232A38] text-[#CBD5E0] hover:text-white border border-[#4A5568]/50'
              }`}
            >
              No Asignadas ({noAsignadasCount})
            </button>
          </div>

          <div className="flex items-center gap-1.5 shrink-0 text-xs">
            <button
              type="button"
              onClick={handleToggleBloqueoTodo}
              className="py-1 px-2 rounded-lg bg-[#232A38] border border-[#4A5568]/60 text-[#F6AD55] hover:text-white text-[11px] font-semibold flex items-center gap-1 cursor-pointer transition-colors"
              title={hayBloqueadas ? 'Desbloquear todo' : 'Bloquear todo'}
            >
              {hayBloqueadas ? <Unlock size={12} /> : <Lock size={12} />}
              <span className="hidden sm:inline">{hayBloqueadas ? 'Desbloquear' : 'Bloquear'}</span>
            </button>
            {activasCount > 0 && (
              <button
                type="button"
                onClick={handleDesactivarTodas}
                className="py-1 px-2 rounded-lg bg-[#232A38] border border-[#E53E3E]/40 text-[#E53E3E] hover:bg-[#E53E3E]/20 text-[11px] font-semibold flex items-center gap-1 cursor-pointer transition-colors"
                title="Desactivar todas las máquinas activas"
              >
                <XCircle size={12} />
                <span className="hidden sm:inline">Desactivar todas</span>
              </button>
            )}
          </div>
        </div>

        {/* Collapsible Admin Panel: Add machines, Google Sheets, Excel import, Backup & Restore */}
        {panelAdminOpen && (
          <div className="pt-2.5 border-t border-[#2D3748] space-y-2.5">
            <div className="flex flex-col sm:flex-row gap-2">
              <input
                type="text"
                value={nombreInput}
                onChange={(e) => setNombreInput(e.target.value)}
                placeholder="Nombre de máquina (ej. Troqueladora 3 o varias separadas por coma)"
                className="flex-1 py-1.5 px-2.5 rounded-lg border border-[#4A5568] text-xs sm:text-sm bg-[#232A38] text-white outline-none focus:border-[#F2A93B]"
              />
              <button
                type="button"
                onClick={handleAgregarMaquinas}
                className="bg-[#F2A93B] text-[#12161F] py-1.5 px-3 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer whitespace-nowrap hover:brightness-105"
              >
                + Agregar Máquina
              </button>
            </div>

            <div className="flex items-center gap-2 flex-wrap text-xs">
              <button
                type="button"
                onClick={() => fileInputRef.current?.click()}
                className="flex items-center gap-1 text-[11px] bg-[#232A38] border border-[#4A5568] hover:text-white py-1 px-2.5 rounded-lg cursor-pointer text-[#CBD5E0]"
              >
                <Upload size={12} />
                <span>Importar Excel</span>
              </button>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleImportarExcel}
                className="hidden"
              />

              <button
                type="button"
                disabled={actualizando}
                onClick={handleActualizarDesdeSheets}
                className="flex items-center gap-1 text-[11px] bg-[#232A38] border border-[#4A5568] hover:text-white py-1 px-2.5 rounded-lg cursor-pointer text-[#CBD5E0] disabled:opacity-50"
              >
                <RefreshCw size={12} className={actualizando ? 'animate-spin' : ''} />
                <span>{actualizando ? 'Cargando...' : 'Desde Sheets'}</span>
              </button>

              <button
                type="button"
                onClick={pedirUrlSheetsMaquinas}
                className="text-[11px] text-[#A0AEC0] underline hover:text-[#CBD5E0] cursor-pointer"
              >
                (Link Sheets)
              </button>

              <button
                type="button"
                onClick={handleAbrirRespaldo}
                className="flex items-center gap-1 text-[11px] bg-[#232A38] border border-[#4A5568] hover:text-white py-1 px-2.5 rounded-lg cursor-pointer text-[#CBD5E0]"
              >
                <Copy size={12} />
                <span>Respaldar</span>
              </button>

              <button
                type="button"
                onClick={() => {
                  setPanelRespaldoOpen(false);
                  setTextoRestaurar('');
                  setPanelRestaurarOpen(true);
                }}
                className="flex items-center gap-1 text-[11px] bg-[#232A38] border border-[#4A5568] hover:text-white py-1 px-2.5 rounded-lg cursor-pointer text-[#CBD5E0]"
              >
                <RotateCcw size={12} />
                <span>Restaurar</span>
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Backup panel modal/box */}
      {panelRespaldoOpen && (
        <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-3 shadow-xl">
          <p className="text-xs text-[#A0AEC0] m-0 mb-1.5">
            Copia el texto JSON para respaldar tus máquinas:
          </p>
          <textarea
            readOnly
            value={textoRespaldo}
            className="w-full h-20 text-xs font-['IBM_Plex_Mono',monospace] p-2 rounded-lg border border-[#4A5568] bg-[#232A38] text-white"
          />
          <div className="flex gap-2 mt-2">
            <button
              type="button"
              onClick={handleCopiarRespaldo}
              className="flex-1 bg-[#F2A93B] text-[#12161F] py-2 px-3 rounded-lg text-xs font-bold cursor-pointer"
            >
              Copiar al portapapeles
            </button>
            <button
              type="button"
              onClick={() => setPanelRespaldoOpen(false)}
              className="border border-[#4A5568] text-white py-2 px-3 rounded-lg text-xs cursor-pointer"
            >
              Cerrar
            </button>
          </div>
        </div>
      )}

      {/* Restore panel modal/box */}
      {panelRestaurarOpen && (
        <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-3 shadow-xl">
          <p className="text-xs text-[#A0AEC0] m-0 mb-1.5">
            Pega aquí el texto JSON previamente respaldado:
          </p>
          <textarea
            value={textoRestaurar}
            onChange={(e) => setTextoRestaurar(e.target.value)}
            placeholder="Pega aquí el texto de respaldo…"
            className="w-full h-20 text-xs font-['IBM_Plex_Mono',monospace] p-2 rounded-lg border border-[#4A5568] bg-[#232A38] text-white"
          />
          <div className="flex gap-2 mt-2">
            <button
              type="button"
              onClick={handleConfirmarRestaurar}
              className="flex-1 bg-[#F2A93B] text-[#12161F] py-2 px-3 rounded-lg text-xs font-bold cursor-pointer"
            >
              Restaurar máquinas
            </button>
            <button
              type="button"
              onClick={() => setPanelRestaurarOpen(false)}
              className="border border-[#4A5568] text-white py-2 px-3 rounded-lg text-xs cursor-pointer"
            >
              Cerrar
            </button>
          </div>
        </div>
      )}

      {/* Counter bar */}
      <div className="flex items-center justify-between text-xs text-[#A0AEC0] px-1">
        <span>
          {busqueda ? `Resultados para «${busqueda}»` : 'Toca ACTIVAR para enviar la máquina a Producción'}
        </span>
        <span className="font-['IBM_Plex_Mono',monospace] font-bold">
          {totalFiltradasCount} de {maquinas.length}
        </span>
      </div>

      {/* Machines Sections: Separated by Vasos y Envases, Papel y Cartón, No Asignadas, in a vertical single-column list (una sobre otra) */}
      {totalFiltradasCount === 0 ? (
        <div className="border border-dashed border-[#4A5568] rounded-xl p-6 text-center text-xs sm:text-sm text-[#CBD5E0] bg-[#1A202C]">
          {busqueda ? `No se encontraron máquinas con el término «${busqueda}».` : 'No hay máquinas registradas en este filtro.'}
        </div>
      ) : (
        <div className="space-y-4">
          {seccionGrupos.map((seccion) => {
            if (seccion.maquinas.length === 0) return null;
            const activasEnSeccion = seccion.maquinas.filter((m) => m.activa).length;

            return (
              <div key={seccion.id} className="space-y-2">
                {/* Section Separation Banner */}
                <div className="flex items-center justify-between px-3 py-2 rounded-xl bg-[#1A202C] border border-[#2D3748] shadow-sm">
                  <div className="flex items-center gap-2">
                    <span className={`w-2.5 h-2.5 rounded-full ${seccion.dotColor} shrink-0`} />
                    <h2 className="font-['Oswald',sans-serif] text-sm sm:text-base font-bold text-white tracking-wide uppercase">
                      {seccion.titulo}
                    </h2>
                    <span className={`text-[10.5px] font-semibold px-2 py-0.5 rounded-md ${seccion.badgeColor}`}>
                      {seccion.maquinas.length} {seccion.maquinas.length === 1 ? 'máquina' : 'máquinas'}
                    </span>
                  </div>
                  <div className="text-xs text-[#CBD5E0] font-['IBM_Plex_Mono',monospace]">
                    <span className="text-[#38A169] font-bold">{activasEnSeccion}</span> activas
                  </div>
                </div>

                {/* Vertical Stack: Una sobre otra */}
                <div className="flex flex-col gap-2">
                  {seccion.maquinas.map((m) => {
                    const esPP = m.tipoProducto === 'pp';
                    const esPT = m.tipoProducto === 'pt';

                    return (
                      <div
                        key={m.id}
                        onDoubleClick={(e) => handleToggleLockCard(m.id, e)}
                        className={`rounded-xl border p-2.5 sm:px-4 sm:py-3 transition-all flex flex-col md:flex-row md:items-center justify-between gap-2.5 shadow-sm ${
                          m.activa
                            ? 'bg-[#182335] border-[#38A169]/70 ring-1 ring-[#38A169]/30 shadow-[#38A169]/5'
                            : 'bg-[#1A202C] border-[#2D3748] hover:border-[#4A5568]'
                        }`}
                      >
                        {/* Left / Main info */}
                        <div className="flex items-center gap-2.5 min-w-0 flex-1 flex-wrap sm:flex-nowrap">
                          {/* 1-tap Activation */}
                          <button
                            type="button"
                            onClick={() => handleUpdateMaquina(m.id, { activa: !m.activa })}
                            className={`py-1.5 px-3 rounded-lg text-[11px] sm:text-xs font-bold font-['Oswald',sans-serif] tracking-wider cursor-pointer transition-all flex items-center gap-1.5 shrink-0 ${
                              m.activa
                                ? 'bg-[#38A169] text-white shadow-sm shadow-[#38A169]/40 hover:brightness-105 active:scale-95'
                                : 'bg-[#232A38] text-[#CBD5E0] border border-[#4A5568] hover:text-white hover:bg-[#2D3748] active:scale-95'
                            }`}
                            title={m.activa ? 'Desactivar máquina' : 'Activar máquina para Producción'}
                          >
                            <span className={`w-2 h-2 rounded-full ${m.activa ? 'bg-white animate-pulse' : 'bg-[#718096]'}`} />
                            <span>{m.activa ? 'ACTIVA' : 'ACTIVAR'}</span>
                          </button>

                          {/* Machine Name and Code */}
                          <div className="flex items-baseline gap-2 min-w-0">
                            <span className="font-['IBM_Plex_Mono',monospace] font-bold text-xs sm:text-base text-white truncate">
                              {m.nombre}
                            </span>
                            {m.codigo && m.bloqueada && (
                              <span className="text-[11px] font-['IBM_Plex_Mono',monospace] font-bold text-[#A0AEC0] bg-[#232A38] px-2 py-0.5 rounded border border-[#4A5568]/60 shrink-0">
                                {m.codigo}
                              </span>
                            )}
                          </div>

                          {/* Area & Code edit (or badge when locked) */}
                          <div className="flex items-center gap-1.5 shrink-0">
                            {m.bloqueada ? (
                              <span className="text-[10px] sm:text-[11px] font-medium text-[#A0AEC0] bg-[#232A38]/90 px-2 py-0.5 rounded border border-[#2D3748]">
                                {m.area === 'vasos_envases'
                                  ? 'Vasos y Env.'
                                  : m.area === 'papel_carton'
                                  ? 'Papel y Cart.'
                                  : 'Sin área'}
                              </span>
                            ) : (
                              <div className="flex items-center gap-1.5">
                                <select
                                  value={m.area || ''}
                                  onChange={(e) => handleUpdateMaquina(m.id, { area: e.target.value as any })}
                                  className="py-1 px-2 rounded-lg border border-[#4A5568] text-xs bg-[#232A38] text-[#CBD5E0] outline-none focus:border-[#F2A93B]"
                                  title="Área"
                                >
                                  <option value="">Sin área</option>
                                  <option value="vasos_envases">Vasos</option>
                                  <option value="papel_carton">Papel</option>
                                </select>
                                <input
                                  type="text"
                                  value={m.codigo || ''}
                                  onChange={(e) => handleUpdateMaquina(m.id, { codigo: e.target.value.trim() })}
                                  placeholder="Cód."
                                  title="Código de máquina"
                                  className="w-16 sm:w-20 py-1 px-2 rounded-lg border border-[#4A5568] text-xs font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white outline-none focus:border-[#F2A93B]"
                                />
                              </div>
                            )}
                          </div>
                        </div>

                        {/* Right: Chips, Multi-OP, Lock & Trash */}
                        <div className="flex items-center justify-between sm:justify-end gap-2 shrink-0 pt-1.5 sm:pt-0 border-t sm:border-t-0 border-[#2D3748]/60">
                          {/* PP / PT toggle chips */}
                          <div className="flex p-0.5 bg-[#232A38] rounded-lg border border-[#4A5568]/60">
                            <button
                              type="button"
                              disabled={m.bloqueada}
                              onClick={() => handleUpdateMaquina(m.id, { tipoProducto: esPP ? '' : 'pp' })}
                              className={`py-1 px-2 rounded-md text-[10.5px] sm:text-[11px] font-bold cursor-pointer transition-all ${
                                esPP ? 'bg-[#F2A93B] text-[#12161F] shadow-sm' : 'text-[#CBD5E0] hover:text-white'
                              } ${m.bloqueada ? 'opacity-60 cursor-default' : ''}`}
                              title="Producto en Proceso"
                            >
                              PP
                            </button>
                            <button
                              type="button"
                              disabled={m.bloqueada}
                              onClick={() => handleUpdateMaquina(m.id, { tipoProducto: esPT ? '' : 'pt' })}
                              className={`py-1 px-2 rounded-md text-[10.5px] sm:text-[11px] font-bold cursor-pointer transition-all ${
                                esPT ? 'bg-[#63B3ED] text-[#12161F] shadow-sm' : 'text-[#CBD5E0] hover:text-white'
                              } ${m.bloqueada ? 'opacity-60 cursor-default' : ''}`}
                              title="Producto Terminado"
                            >
                              PT
                            </button>
                          </div>

                          {/* Multi-OP toggle */}
                          <button
                            type="button"
                            onClick={() => {
                              if (!m.multiOp) {
                                handleUpdateMaquina(m.id, { multiOp: true });
                                showToast('ok', `Multi-OP activado en ${m.nombre}`);
                              } else {
                                const n = turno.filter((x) => x.maquinaId === m.id).length;
                                if (n > 1) {
                                  showToast('error', `Primero quita las tarjetas adicionales en Producción (hay ${n})`);
                                  return;
                                }
                                handleUpdateMaquina(m.id, { multiOp: false });
                                showToast('ok', `Multi-OP desactivado en ${m.nombre}`);
                              }
                            }}
                            title="Multi-OP: Permite asignar varias OP a la vez a esta máquina"
                            className={`py-1 px-2 sm:px-2.5 rounded-lg text-[10.5px] sm:text-[11px] font-bold cursor-pointer border transition-all ${
                              m.multiOp
                                ? 'bg-[#F2A93B] text-[#12161F] border-[#F2A93B] shadow-sm font-black'
                                : 'bg-[#232A38] text-[#CBD5E0] border-[#4A5568]/60 hover:text-white'
                            }`}
                          >
                            Multi-OP
                          </button>

                          {/* Lock icon & Trash */}
                          <div className="flex items-center gap-1 pl-1 border-l border-[#2D3748]">
                            <button
                              type="button"
                              onClick={() => handleToggleLockCard(m.id)}
                              className={`p-1.5 rounded-lg cursor-pointer transition-colors ${
                                m.bloqueada
                                  ? 'text-[#F6AD55] hover:text-[#FBD38D] bg-[#232A38]'
                                  : 'text-[#718096] hover:text-[#CBD5E0]'
                              }`}
                              title={
                                m.bloqueada
                                  ? 'Máquina bloqueada (doble clic para editar datos)'
                                  : 'Máquina desbloqueada (clic para bloquear)'
                              }
                            >
                              {m.bloqueada ? <Lock size={14} /> : <Unlock size={14} />}
                            </button>
                            {!m.bloqueada && (
                              <button
                                type="button"
                                onClick={() => handleEliminarMaquina(m.id)}
                                className="p-1.5 text-[#E53E3E] hover:bg-[#E53E3E]/20 rounded-lg cursor-pointer"
                                title="Eliminar máquina"
                              >
                                <Trash2 size={14} />
                              </button>
                            )}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};
