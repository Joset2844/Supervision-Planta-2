import React, { useState, useRef, useMemo } from 'react';
import { useApp } from '../../context/AppContext';
import { Odometer } from '../common/Odometer';
import { Plus, RefreshCw, Upload, Check, Copy, Trash2, Search, X, RotateCcw, Download } from 'lucide-react';
import * as XLSX from 'xlsx';
import { procesarWorkbookCatalogo, buscarEnCatalogo, resolverModelo, exportarInventarioExcel } from '../../services/excel';
import { copiarAlPortapapeles } from '../../services/reports';
import { CatalogoItem, ListaItem } from '../../types';

export const InventarioView: React.FC = () => {
  const {
    catalogo,
    lista,
    setLista,
    subTabInventario,
    showToast,
    sincronizarCatalogoSupabase,
    subirCatalogoASupabase,
  } = useApp();

  // Form states
  const [opInput, setOpInput] = useState('');
  const [descInput, setDescInput] = useState('');
  const [clienteInput, setClienteInput] = useState('');
  const [observacionInput, setObservacionInput] = useState('');
  const [tipoEmpaque, setTipoEmpaque] = useState<'caja' | 'paquete'>('caja');
  const [unidadesInput, setUnidadesInput] = useState('');
  const [cajasInput, setCajasInput] = useState('');
  const [modeloSeleccionado, setModeloSeleccionado] = useState<number>(-1);
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [actualizando, setActualizando] = useState(false);

  // Backup / Restore states
  const [panelRespaldoOpen, setPanelRespaldoOpen] = useState(false);
  const [panelRestaurarOpen, setPanelRestaurarOpen] = useState(false);
  const [textoRespaldo, setTextoRespaldo] = useState('');
  const [textoRestaurar, setTextoRestaurar] = useState('');

  // Search state for Consultar
  const [buscarQuery, setBuscarQuery] = useState('');

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Matched catalog item
  const matchActual = useMemo(() => {
    return buscarEnCatalogo(catalogo, opInput);
  }, [catalogo, opInput]);

  const opEncontrada = !!matchActual;

  const suggestions = useMemo(() => {
    const raw = opInput.trim();
    const digits = raw.replace(/\D/g, '');
    if (!digits && !raw) return [];
    return catalogo
      .filter((c) => {
        const cDigits = c.op.replace(/\D/g, '');
        return (digits && cDigits.includes(digits)) || c.op.toLowerCase().includes(raw.toLowerCase());
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

  const totalCalculado = useMemo(() => {
    const u = parseInt(unidadesInput, 10) || 0;
    const c = parseInt(cajasInput, 10) || 0;
    return u * c;
  }, [unidadesInput, cajasInput]);

  const totalUnidadesLista = useMemo(() => {
    return lista.reduce((acc, i) => acc + (i.totalUnidades || 0), 0);
  }, [lista]);

  // Handle OP selection
  const handleSelectOp = (item: CatalogoItem) => {
    setOpInput(item.op);
    setModeloSeleccionado(-1);
    const mod = resolverModelo(item, -1);
    setDescInput(mod.descripcion);
    setClienteInput(item.cliente);
    setShowSuggestions(false);
    setTimeout(() => {
      document.getElementById('inventario-unidades-input')?.focus();
    }, 60);
  };

  const handleOpChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    // Solo permitir números
    const digitsOnly = e.target.value.replace(/\D/g, '');
    setOpInput(digitsOnly);
    setShowSuggestions(true);
    const m = buscarEnCatalogo(catalogo, digitsOnly);
    if (m) {
      setModeloSeleccionado(-1);
      const mod = resolverModelo(m, -1);
      setDescInput(mod.descripcion);
      setClienteInput(m.cliente);
    } else {
      setModeloSeleccionado(-1);
    }
  };

  const handleOpKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    const allowed = ['Backspace', 'Delete', 'Tab', 'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Enter', 'Escape'];
    if (allowed.includes(e.key) || e.ctrlKey || e.metaKey) {
      if (e.key === 'Enter') {
        e.preventDefault();
        if (showSuggestions && suggestions.length > 0) {
          handleSelectOp(suggestions[0]);
        } else {
          setShowSuggestions(false);
          document.getElementById('inventario-unidades-input')?.focus();
        }
      }
      return;
    }
    // Bloquear caracteres no numéricos
    if (!/[0-9]/.test(e.key)) {
      e.preventDefault();
    }
  };

  const handleModeloChipClick = (idx: number) => {
    if (!matchActual) return;
    setModeloSeleccionado(idx);
    const mod = resolverModelo(matchActual, idx);
    setDescInput(mod.descripcion);
  };

  const limpiarFormulario = () => {
    setOpInput('');
    setDescInput('');
    setClienteInput('');
    setObservacionInput('');
    setUnidadesInput('');
    setCajasInput('');
    setModeloSeleccionado(-1);
    setShowSuggestions(false);
  };

  const handleAgregarPieza = () => {
    if (!opInput.trim()) {
      showToast('error', 'Ingresa el número de OP');
      return;
    }
    if (!unidadesInput || !cajasInput) {
      showToast('error', `Completa unidades x ${tipoEmpaque} y cantidad de ${tipoEmpaque}s`);
      return;
    }
    const u = parseInt(unidadesInput, 10);
    const c = parseInt(cajasInput, 10);
    if (isNaN(u) || isNaN(c) || u < 0 || c < 0) {
      showToast('error', 'Valores de cantidad inválidos');
      return;
    }
    const sapActual = matchActual ? resolverModelo(matchActual, modeloSeleccionado).sap : '';

    const nuevoItem: ListaItem = {
      id: Date.now() + '-' + Math.random().toString(36).slice(2, 7),
      op: opInput.trim(),
      descripcion: descInput.trim(),
      sap: sapActual || '',
      cliente: clienteInput.trim(),
      observacion: observacionInput.trim(),
      tipoEmpaque,
      unidadesXCaja: u,
      cantidadCajas: c,
      totalUnidades: u * c,
      fecha: new Date().toISOString(),
    };

    setLista((prev) => [nuevoItem, ...prev]);
    showToast('ok', 'Agregado a la lista');
    limpiarFormulario();
  };

  const handleEliminarItem = (id: string) => {
    setLista((prev) => prev.filter((i) => i.id !== id));
  };

  const handleVaciarLista = () => {
    if (lista.length === 0) return;
    if (window.confirm('¿Vaciar toda la lista pendiente?')) {
      setLista([]);
      showToast('ok', 'Lista vaciada');
    }
  };

  // Excel import
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const nuevo = procesarWorkbookCatalogo(wb);
        if (nuevo) {
          await subirCatalogoASupabase(nuevo);
        } else {
          showToast('error', 'Debe tener columnas OP, Descripción y Cliente');
        }
      } catch {
        showToast('error', 'No se pudo leer el archivo Excel');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Sync catalog
  const handleActualizarCatalogo = async () => {
    setActualizando(true);
    try {
      const n = await sincronizarCatalogoSupabase();
      showToast('ok', n ? `Catálogo actualizado: ${n} OP nuevas/modificadas` : 'Catálogo al día');
    } catch {
      showToast('error', 'No se pudo actualizar el catálogo. Revisa tu conexión.');
    } finally {
      setActualizando(false);
    }
  };

  // Backup / Restore
  const handleAbrirRespaldo = () => {
    if (lista.length === 0) {
      showToast('error', 'La lista está vacía, no hay nada que respaldar');
      return;
    }
    setPanelRestaurarOpen(false);
    setTextoRespaldo(JSON.stringify(lista));
    setPanelRespaldoOpen(true);
  };

  const handleCopiarRespaldo = async () => {
    const ok = await copiarAlPortapapeles(textoRespaldo);
    showToast(ok ? 'ok' : 'error', ok ? 'Copiado al portapapeles' : 'No se pudo copiar automáticamente');
  };

  const handleAbrirRestaurar = () => {
    setPanelRespaldoOpen(false);
    setTextoRestaurar('');
    setPanelRestaurarOpen(true);
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
        showToast('error', 'El formato no corresponde a una lista');
        return;
      }
      const idsExistentes = new Set(lista.map((i) => i.id));
      let agregados = 0;
      const nuevosItems: ListaItem[] = [];
      datos.forEach((item: any) => {
        if (item && item.id && !idsExistentes.has(item.id)) {
          nuevosItems.push(item);
          idsExistentes.add(item.id);
          agregados++;
        }
      });
      setLista((prev) => [...nuevosItems, ...prev]);
      setPanelRestaurarOpen(false);
      showToast('ok', `${agregados} pieza(s) restaurada(s)`);
    } catch {
      showToast('error', 'Ese texto no es válido, revisa que lo copiaste completo');
    }
  };

  // Filtered catalog for Consultar
  const resultadosConsulta = useMemo(() => {
    const q = buscarQuery.trim().toLowerCase();
    if (!q) return catalogo;
    return catalogo.filter((c) => {
      if (
        c.op.toLowerCase().includes(q) ||
        c.descripcion.toLowerCase().includes(q) ||
        c.cliente.toLowerCase().includes(q) ||
        String(c.sap || '').toLowerCase().includes(q)
      ) {
        return true;
      }
      if (c.modelos && c.modelos.length > 0) {
        return c.modelos.some(
          (mo) =>
            (mo.descripcion || '').toLowerCase().includes(q) ||
            (mo.sap || '').toLowerCase().includes(q)
        );
      }
      return false;
    });
  }, [catalogo, buscarQuery]);

  const copiarInfoOp = async (r: CatalogoItem) => {
    const desc = r.modelos && r.modelos.length > 1 && r.descripcionBase ? r.descripcionBase : r.descripcion;
    const txt = `OP: ${r.op}\nDescripción: ${desc}\nCliente: ${r.cliente}`;
    const ok = await copiarAlPortapapeles(txt);
    showToast(ok ? 'ok' : 'error', ok ? `Copiado: ${r.op}` : 'No se pudo copiar');
  };

  const copiarModelosOp = async (r: CatalogoItem) => {
    if (!r.modelos) return;
    const base = r.descripcionBase || '';
    const txt =
      `OP: ${r.op}\nDescripción: ${base}\nCliente: ${r.cliente}\nCantidad: ${r.cantidadTotal || ''}\nMODELOS:\n` +
      r.modelos
        .map((mo) => {
          let nombre = mo.modelo;
          if (!nombre) {
            nombre = mo.descripcion || '';
            if (base && nombre.indexOf(base + ' - ') === 0) nombre = nombre.slice(base.length + 3);
          }
          return `- ${nombre}${mo.cantidadTotal ? ' - ' + mo.cantidadTotal : ''}`;
        })
        .join('\n');
    const ok = await copiarAlPortapapeles(txt);
    showToast(ok ? 'ok' : 'error', ok ? `Modelos copiados: ${r.op}` : 'No se pudo copiar');
  };

  return (
    <div className="space-y-2.5 sm:space-y-3.5">
      {/* Top action row */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-[#1A202C] border border-[#2D3748] rounded-xl p-2 sm:p-2.5 shadow-sm">
        <div className="flex items-center gap-2">
          <span className="text-xs text-[#A0AEC0]">
            <strong className="text-white font-semibold">{catalogo.length}</strong> OP en catálogo ·{' '}
            <strong className="text-[#F6AD55] font-semibold">{lista.length}</strong> en lista
          </span>
        </div>
        <div className="flex items-center gap-2 w-full sm:w-auto">
          <button
            type="button"
            onClick={handleActualizarCatalogo}
            disabled={actualizando}
            className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 bg-[#232A38] text-[#F7FAFC] border border-[#4A5568] py-1 px-2.5 rounded-lg text-xs font-semibold cursor-pointer hover:bg-[#2D3748] active:scale-95 transition-all"
          >
            <RefreshCw size={12} className={actualizando ? 'animate-spin' : ''} />
            <span>{actualizando ? 'Actualizando...' : 'Actualizar'}</span>
          </button>
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 bg-[#F2A93B] text-[#12161F] py-1 px-2.5 rounded-lg text-xs font-bold cursor-pointer hover:brightness-110 active:scale-95 transition-all shadow-sm shadow-[#F2A93B]/20"
          >
            <Upload size={12} />
            <span>Cargar Excel</span>
          </button>
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
          />
        </div>
      </div>

      {catalogo.length === 0 && (
        <div className="border border-dashed border-[#4A5568] bg-[#1A202C] rounded-xl p-4 text-center">
          <p className="text-xs font-semibold text-white mb-1">Aún no hay catálogo cargado</p>
          <p className="text-[11px] text-[#A0AEC0] max-w-[420px] mx-auto mb-2.5">
            Importa un archivo Excel con columnas <strong>OP</strong>, <strong>Descripción</strong>, <strong>Cliente</strong>,{' '}
            <strong>Cantidad</strong> y <strong>SAP</strong> para autocompletar automáticamente los datos.
          </p>
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="inline-flex gap-1.5 items-center bg-[#F2A93B] text-[#12161F] py-1.5 px-3 rounded-lg text-xs font-bold cursor-pointer hover:brightness-110 active:scale-95 transition-all"
          >
            <Upload size={13} />
            Cargar OP (Excel)
          </button>
        </div>
      )}

      {/* Subtab: AGREGAR */}
      {subTabInventario === 'agregar' && (
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-2.5 sm:gap-3.5 items-start">
          {/* New Piece Form (5 columns on desktop) */}
          <div className="lg:col-span-5 bg-[#1A202C] border border-[#2D3748] rounded-xl p-2.5 sm:p-3 shadow-sm">
            <div className="flex items-center justify-between mb-2 pb-1.5 border-b border-[#2D3748]">
              <h2 className="m-0 text-xs font-bold flex items-center gap-1.5 text-white font-['Oswald',sans-serif] tracking-wide uppercase">
                <Plus size={14} className="text-[#F2A93B]" />
                Registrar Pieza
              </h2>
              {opEncontrada && (
                <span className="text-[10px] font-semibold text-[#38A169] flex items-center gap-1">
                  <Check size={11} /> Catálogo
                </span>
              )}
            </div>

            {/* OP field (Solo números) */}
            <div className="mb-2 relative">
              <div className="flex items-center justify-between mb-0.5">
                <label className="block text-[10px] font-semibold text-[#CBD5E0] uppercase tracking-wider font-['Oswald',sans-serif]">
                  Número de OP <span className="text-[#F2A93B] font-normal text-[9.5px]">(solo números)</span>
                </label>
                {opInput && (
                  <span className="text-[10px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace]">
                    {opInput.length} dígito{opInput.length !== 1 ? 's' : ''}
                  </span>
                )}
              </div>
              <div className="relative">
                <input
                  type="text"
                  inputMode="numeric"
                  pattern="[0-9]*"
                  value={opInput}
                  onChange={handleOpChange}
                  onKeyDown={handleOpKeyDown}
                  onFocus={() => setShowSuggestions(true)}
                  onBlur={() => setTimeout(() => setShowSuggestions(false), 240)}
                  placeholder="Escribe números de OP (ej. 10234)"
                  autoComplete="off"
                  className={`w-full py-1.5 px-2.5 pr-8 rounded-lg border text-xs outline-none font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white transition-all ${
                    opEncontrada
                      ? 'border-[#38A169] focus:ring-1 focus:ring-[#38A169]/30'
                      : 'border-[#4A5568] focus:border-[#F2A93B] focus:ring-1 focus:ring-[#F2A93B]/30'
                  }`}
                />
                {opInput && (
                  <button
                    type="button"
                    onClick={() => {
                      setOpInput('');
                      setShowSuggestions(false);
                    }}
                    className="absolute right-2 top-1/2 -translate-y-1/2 text-[#A0AEC0] hover:text-white p-1 rounded hover:bg-[#1A202C] transition-colors"
                    title="Borrar OP"
                  >
                    <X size={13} />
                  </button>
                )}
              </div>

              {/* Suggestions dropdown with rich feedback */}
              {showSuggestions && opInput.trim() !== '' && (
                <div className="absolute z-40 left-0 right-0 mt-1 bg-[#1A202C] border border-[#4A5568] rounded-xl max-h-[250px] overflow-y-auto shadow-2xl overflow-hidden divide-y divide-[#2D3748]">
                  {suggestions.length > 0 ? (
                    <>
                      <div className="px-3 py-1.5 bg-[#232A38] flex items-center justify-between text-[10px] text-[#A0AEC0]">
                        <span className="font-semibold uppercase tracking-wider text-[#CBD5E0]">
                          Catálogo ({suggestions.length} resultado{suggestions.length !== 1 ? 's' : ''})
                        </span>
                        <span className="text-[9.5px]">Toca para autocompletar</span>
                      </div>
                      {suggestions.map((item) => (
                        <button
                          key={item.op}
                          type="button"
                          onMouseDown={(e) => {
                            e.preventDefault();
                            handleSelectOp(item);
                          }}
                          className="w-full text-left py-2 px-3 hover:bg-[#232A38] active:bg-[#2D3748] text-xs block text-white transition-colors cursor-pointer group"
                        >
                          <div className="flex items-center justify-between gap-1.5 mb-0.5">
                            <span className="font-bold text-[#F2A93B] group-hover:text-[#F6AD55] font-['IBM_Plex_Mono',monospace] text-xs px-1.5 py-0.5 rounded bg-[#232A38] border border-[#4A5568]/80">
                              {item.op}
                            </span>
                            <div className="flex items-center gap-1">
                              {item.modelos && item.modelos.length > 1 && (
                                <span className="text-[9.5px] px-1.5 py-0.2 rounded bg-[#4A5568]/40 text-[#CBD5E0]">
                                  {item.modelos.length} mod.
                                </span>
                              )}
                              <span className="text-[#A0AEC0] text-[10px] truncate max-w-[130px]">
                                {item.cliente || 'Sin cliente'}
                              </span>
                            </div>
                          </div>
                          <div className="text-[#CBD5E0] text-[11px] truncate">
                            {item.descripcionBase || item.descripcion}
                          </div>
                        </button>
                      ))}
                    </>
                  ) : (
                    <div className="p-3 text-center bg-[#1A202C]">
                      <div className="text-xs text-[#CBD5E0] font-medium">
                        OP <span className="text-[#F2A93B] font-mono font-bold">{opInput}</span> no está en el catálogo
                      </div>
                      <div className="text-[10.5px] text-[#A0AEC0] mt-0.5">
                        Puedes continuar e ingresar descripción y cliente manualmente.
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Model chips if multi-model */}
            {matchActual && matchActual.modelos && matchActual.modelos.length > 1 && (
              <div className="mb-2 p-2 bg-[#232A38] rounded-lg border border-[#2D3748]">
                <label className="block text-[10px] font-semibold text-[#CBD5E0] mb-1.5 uppercase tracking-wide">
                  Modelos de la OP ({matchActual.modelos.length})
                </label>
                <div className="flex gap-1 flex-wrap">
                  <button
                    type="button"
                    onClick={() => handleModeloChipClick(-1)}
                    className={`py-1 px-2 rounded-md text-[11px] font-semibold cursor-pointer border transition-all ${
                      modeloSeleccionado === -1
                        ? 'bg-[#F2A93B] text-[#12161F] border-[#F2A93B] shadow-sm'
                        : 'bg-[#1A202C] text-[#CBD5E0] border-[#4A5568] hover:text-white'
                    }`}
                  >
                    Todos (Base)
                  </button>
                  {matchActual.modelos.map((m, idx) => (
                    <button
                      key={idx}
                      type="button"
                      onClick={() => handleModeloChipClick(idx)}
                      className={`py-1 px-2 rounded-md text-[11px] font-semibold cursor-pointer border transition-all ${
                        modeloSeleccionado === idx
                          ? 'bg-[#F2A93B] text-[#12161F] border-[#F2A93B] shadow-sm'
                          : 'bg-[#1A202C] text-[#CBD5E0] border-[#4A5568] hover:text-white'
                      }`}
                    >
                      {m.descripcion || `Modelo ${idx + 1}`}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {/* Description & Cliente in 2-col on larger screens */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 mb-2">
              <div>
                <label className="block text-[10px] font-semibold text-[#CBD5E0] mb-0.5 uppercase tracking-wider font-['Oswald',sans-serif]">
                  Descripción
                </label>
                <input
                  type="text"
                  value={descInput}
                  onChange={(e) => setDescInput(e.target.value)}
                  placeholder="Se autocompleta con la OP"
                  className="w-full py-1.5 px-2.5 rounded-lg border border-[#4A5568] text-xs outline-none bg-[#232A38] text-white focus:border-[#F2A93B]"
                />
              </div>
              <div>
                <label className="block text-[10px] font-semibold text-[#CBD5E0] mb-0.5 uppercase tracking-wider font-['Oswald',sans-serif]">
                  Cliente
                </label>
                <input
                  type="text"
                  value={clienteInput}
                  onChange={(e) => setClienteInput(e.target.value)}
                  placeholder="Se autocompleta con la OP"
                  className="w-full py-1.5 px-2.5 rounded-lg border border-[#4A5568] text-xs outline-none bg-[#232A38] text-white focus:border-[#F2A93B]"
                />
              </div>
            </div>

            {/* Tipo de empaque segmented control */}
            <div className="mb-2">
              <label className="block text-[10px] font-semibold text-[#CBD5E0] mb-0.5 uppercase tracking-wider font-['Oswald',sans-serif]">
                Tipo de Empaque
              </label>
              <div className="grid grid-cols-2 gap-1.5 p-1 bg-[#232A38] rounded-lg border border-[#4A5568]">
                <button
                  type="button"
                  onClick={() => setTipoEmpaque('caja')}
                  className={`flex items-center justify-center gap-1.5 py-1 px-2 rounded-md text-xs font-bold cursor-pointer transition-all ${
                    tipoEmpaque === 'caja'
                      ? 'bg-[#F2A93B] text-[#12161F] shadow-sm'
                      : 'text-[#CBD5E0] hover:text-white'
                  }`}
                >
                  <span className={`w-1.5 h-1.5 rounded-full ${tipoEmpaque === 'caja' ? 'bg-[#12161F]' : 'bg-[#4A5568]'}`} />
                  Cajas
                </button>
                <button
                  type="button"
                  onClick={() => setTipoEmpaque('paquete')}
                  className={`flex items-center justify-center gap-1.5 py-1 px-2 rounded-md text-xs font-bold cursor-pointer transition-all ${
                    tipoEmpaque === 'paquete'
                      ? 'bg-[#F2A93B] text-[#12161F] shadow-sm'
                      : 'text-[#CBD5E0] hover:text-white'
                  }`}
                >
                  <span className={`w-1.5 h-1.5 rounded-full ${tipoEmpaque === 'paquete' ? 'bg-[#12161F]' : 'bg-[#4A5568]'}`} />
                  Paquetes
                </button>
              </div>
            </div>

            {/* Quantities */}
            <div className="grid grid-cols-2 gap-2 mb-2">
              <div>
                <label className="block text-[10px] font-semibold text-[#CBD5E0] mb-0.5 uppercase tracking-wider font-['Oswald',sans-serif]">
                  Unid. x {tipoEmpaque}
                </label>
                <input
                  id="inventario-unidades-input"
                  type="number"
                  inputMode="numeric"
                  min="0"
                  value={unidadesInput}
                  onChange={(e) => setUnidadesInput(e.target.value)}
                  placeholder="0"
                  className="w-full py-1.5 px-2.5 rounded-lg border border-[#4A5568] text-xs outline-none font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white focus:border-[#F2A93B]"
                />
              </div>
              <div>
                <label className="block text-[10px] font-semibold text-[#CBD5E0] mb-0.5 uppercase tracking-wider font-['Oswald',sans-serif]">
                  Cant. {tipoEmpaque}s
                </label>
                <input
                  type="number"
                  inputMode="numeric"
                  min="0"
                  value={cajasInput}
                  onChange={(e) => setCajasInput(e.target.value)}
                  placeholder="0"
                  className="w-full py-1.5 px-2.5 rounded-lg border border-[#4A5568] text-xs outline-none font-['IBM_Plex_Mono',monospace] bg-[#232A38] text-white focus:border-[#F2A93B]"
                />
              </div>
            </div>

            {/* Observación */}
            <div className="mb-2.5">
              <label className="block text-[10px] font-semibold text-[#CBD5E0] mb-0.5 uppercase tracking-wider font-['Oswald',sans-serif]">
                Observación (opcional)
              </label>
              <input
                type="text"
                value={observacionInput}
                onChange={(e) => setObservacionInput(e.target.value)}
                placeholder="Ej. caja abollada, lote especial..."
                className="w-full py-1.5 px-2.5 rounded-lg border border-[#4A5568] text-xs outline-none bg-[#232A38] text-white focus:border-[#F2A93B]"
              />
            </div>

            {/* Total box */}
            <div className="flex items-center justify-between bg-[#232A38] rounded-lg py-1.5 px-2.5 mb-2.5 border border-[#2D3748]">
              <span className="text-[11px] font-semibold text-[#CBD5E0] uppercase tracking-wide">Total Unidades:</span>
              <span className="text-base font-bold text-[#F2A93B] font-['IBM_Plex_Mono',monospace] tabular-nums">
                {totalCalculado.toLocaleString('es-PE')} u
              </span>
            </div>

            <button
              type="button"
              onClick={handleAgregarPieza}
              className="w-full flex items-center justify-center gap-1.5 bg-[#F2A93B] text-[#12161F] py-2 px-3 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer active:scale-[0.98] transition-all hover:brightness-105 shadow-sm shadow-[#F2A93B]/20"
            >
              <Plus size={15} />
              Agregar a la lista
            </button>
          </div>

          {/* Pending List Panel (7 columns on desktop) */}
          <div className="lg:col-span-7 space-y-2.5">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-1.5 border-b border-[#2D3748]">
              <h2 className="m-0 text-xs sm:text-sm flex items-center gap-1.5 text-white font-bold font-['Oswald',sans-serif] tracking-wide uppercase">
                Lista Pendiente de Exportar
              </h2>
              <div className="flex gap-1.5 items-center flex-wrap">
                <button
                  type="button"
                  onClick={handleAbrirRespaldo}
                  className="flex items-center gap-1 text-[11px] text-[#CBD5E0] hover:text-white bg-[#232A38] border border-[#4A5568] py-1 px-2 rounded-md cursor-pointer transition-colors"
                >
                  <Copy size={12} />
                  <span>Respaldar</span>
                </button>
                <button
                  type="button"
                  onClick={handleAbrirRestaurar}
                  className="flex items-center gap-1 text-[11px] text-[#CBD5E0] hover:text-white bg-[#232A38] border border-[#4A5568] py-1 px-2 rounded-md cursor-pointer transition-colors"
                >
                  <RotateCcw size={12} />
                  <span>Restaurar</span>
                </button>
                {lista.length > 0 && (
                  <button
                    type="button"
                    onClick={handleVaciarLista}
                    className="flex items-center gap-1 text-[11px] text-[#E53E3E] hover:bg-[#E53E3E]/20 bg-[#232A38] border border-[#E53E3E]/40 py-1 px-2 rounded-md cursor-pointer transition-colors"
                  >
                    <Trash2 size={12} />
                    <span>Vaciar</span>
                  </button>
                )}
              </div>
            </div>

            {/* Backup drawer */}
            {panelRespaldoOpen && (
              <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-3 shadow-xl">
                <p className="text-[11px] text-[#A0AEC0] m-0 mb-1.5">
                  Copia este texto y guárdalo en Notas o WhatsApp para respaldar tu inventario pendiente.
                </p>
                <textarea
                  readOnly
                  value={textoRespaldo}
                  className="w-full h-20 text-[11px] font-['IBM_Plex_Mono',monospace] p-2 rounded-lg border border-[#4A5568] bg-[#232A38] text-white"
                />
                <div className="flex gap-2 mt-2">
                  <button
                    type="button"
                    onClick={handleCopiarRespaldo}
                    className="flex-1 bg-[#F2A93B] text-[#12161F] py-1.5 px-2.5 rounded-lg text-xs font-bold cursor-pointer"
                  >
                    Copiar al portapapeles
                  </button>
                  <button
                    type="button"
                    onClick={() => setPanelRespaldoOpen(false)}
                    className="border border-[#4A5568] text-white py-1.5 px-3 rounded-lg text-xs cursor-pointer"
                  >
                    Cerrar
                  </button>
                </div>
              </div>
            )}

            {/* Restore drawer */}
            {panelRestaurarOpen && (
              <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-3 shadow-xl">
                <p className="text-[11px] text-[#A0AEC0] m-0 mb-1.5">
                  Pega aquí el texto JSON previamente respaldado.
                </p>
                <textarea
                  value={textoRestaurar}
                  onChange={(e) => setTextoRestaurar(e.target.value)}
                  placeholder="Pega aquí el texto de respaldo…"
                  className="w-full h-20 text-[11px] font-['IBM_Plex_Mono',monospace] p-2 rounded-lg border border-[#4A5568] bg-[#232A38] text-white"
                />
                <div className="flex gap-2 mt-2">
                  <button
                    type="button"
                    onClick={handleConfirmarRestaurar}
                    className="flex-1 bg-[#F2A93B] text-[#12161F] py-1.5 px-2.5 rounded-lg text-xs font-bold cursor-pointer"
                  >
                    Restaurar lista
                  </button>
                  <button
                    type="button"
                    onClick={() => setPanelRestaurarOpen(false)}
                    className="border border-[#4A5568] text-white py-1.5 px-3 rounded-lg text-xs cursor-pointer"
                  >
                    Cerrar
                  </button>
                </div>
              </div>
            )}

            {/* Odometer */}
            <Odometer label="Total unidades en lista" value={totalUnidadesLista} />

            {/* Items list */}
            {lista.length === 0 ? (
              <div className="border border-dashed border-[#4A5568] rounded-xl p-6 text-center text-xs text-[#CBD5E0] bg-[#1A202C]">
                Todavía no has agregado piezas a la lista pendiente.
              </div>
            ) : (
              <div className="space-y-2 max-h-[600px] overflow-y-auto pr-1">
                {lista.map((item) => {
                  const singular = item.tipoEmpaque === 'paquete' ? 'paquete' : 'caja';
                  const plural = item.tipoEmpaque === 'paquete' ? 'paquetes' : 'cajas';
                  const empaqueTxt = item.cantidadCajas === 1 ? singular : plural;

                  return (
                    <div
                      key={item.id}
                      className="bg-[#1A202C] border border-[#2D3748] hover:border-[#4A5568] rounded-xl p-2.5 sm:p-3 flex items-center justify-between gap-2.5 transition-colors shadow-sm"
                    >
                      <div className="min-w-0 flex-1">
                        <div className="flex items-center gap-1.5 flex-wrap mb-0.5">
                          <span className="text-[10px] font-bold bg-[#3A2E1B] text-[#F6AD55] py-0.5 px-1.5 rounded border border-[#F2A93B]/30 font-['IBM_Plex_Mono',monospace]">
                            {item.op}
                          </span>
                          <span className="text-[11px] text-[#CBD5E0] font-medium">{item.cliente}</span>
                          {item.sap && (
                            <span className="text-[10px] text-[#A0AEC0] font-['IBM_Plex_Mono',monospace]">
                              SAP: {item.sap}
                            </span>
                          )}
                        </div>
                        <div className="text-xs font-semibold text-white truncate">
                          {item.descripcion}
                        </div>
                        <div className="text-[11px] text-[#CBD5E0] flex items-center gap-1.5 flex-wrap mt-0.5">
                          <span>
                            {item.cantidadCajas} {empaqueTxt} × {item.unidadesXCaja} u =
                          </span>
                          <strong className="text-[#F2A93B] text-xs font-['IBM_Plex_Mono',monospace] tabular-nums font-bold">
                            {item.totalUnidades.toLocaleString('es-PE')} u
                          </strong>
                        </div>
                        {item.observacion && (
                          <div className="text-[10px] text-[#A0AEC0] italic mt-1 bg-[#232A38] px-2 py-0.5 rounded inline-block">
                            📝 {item.observacion}
                          </div>
                        )}
                      </div>
                      <button
                        type="button"
                        onClick={() => handleEliminarItem(item.id)}
                        className="shrink-0 p-1.5 text-[#E53E3E] hover:bg-[#E53E3E]/20 active:bg-[#E53E3E]/30 rounded-lg cursor-pointer transition-colors w-7 h-7 sm:w-8 sm:h-8 flex items-center justify-center"
                        title="Eliminar pieza"
                        aria-label="Eliminar pieza"
                      >
                        <Trash2 size={15} />
                      </button>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      )}

      {/* Subtab: CONSULTAR */}
      {subTabInventario === 'consultar' && (
        <div className="space-y-2.5">
          <div className="relative">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-[#CBD5E0]" />
            <input
              type="text"
              value={buscarQuery}
              onChange={(e) => setBuscarQuery(e.target.value)}
              placeholder="Buscar por OP, descripción, cliente o SAP…"
              className="w-full py-1.5 pl-9 pr-8 rounded-lg border border-[#4A5568] text-xs outline-none bg-[#1A202C] text-white focus:border-[#F2A93B] focus:ring-1 focus:ring-[#F2A93B]/30"
            />
            {buscarQuery && (
              <button
                type="button"
                onClick={() => setBuscarQuery('')}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-[#CBD5E0] hover:text-white p-0.5 flex items-center justify-center cursor-pointer"
              >
                <X size={14} />
              </button>
            )}
          </div>

          {catalogo.length === 0 ? (
            <p className="text-center py-6 text-[#CBD5E0] text-xs bg-[#1A202C] rounded-xl border border-[#2D3748]">
              Importa un catálogo para poder realizar consultas.
            </p>
          ) : (
            <>
              {/* MOBILE RESULTS VIEW (Touch Cards) */}
              <div className="block md:hidden space-y-2">
                {resultadosConsulta.slice(0, 100).map((r) => {
                  const hasMulti = r.modelos && r.modelos.length > 1;
                  return (
                    <div
                      key={r.op}
                      className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-2.5 space-y-1.5 shadow-sm"
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="font-['IBM_Plex_Mono',monospace] font-bold text-[#F6AD55] text-xs">
                          {r.op}
                        </span>
                        <div className="flex items-center gap-1">
                          <button
                            type="button"
                            onClick={() => copiarInfoOp(r)}
                            title="Copiar datos de OP"
                            className="bg-[#232A38] border border-[#4A5568] text-[#CBD5E0] active:text-white rounded-md p-1.5 flex items-center justify-center cursor-pointer hover:bg-[#2D3748]"
                          >
                            <Copy size={13} />
                          </button>
                          {hasMulti && (
                            <button
                              type="button"
                              onClick={() => copiarModelosOp(r)}
                              className="bg-[#232A38] border border-[#F2A93B] text-[#F6AD55] rounded-md px-2 py-0.5 text-[10px] font-bold flex items-center cursor-pointer hover:bg-[#F2A93B] hover:text-[#12161F]"
                            >
                              MODELOS
                            </button>
                          )}
                        </div>
                      </div>

                      <div className="text-xs font-semibold text-white truncate">
                        {r.descripcion}
                      </div>

                      <div className="flex items-center justify-between text-[11px] text-[#CBD5E0] pt-1 border-t border-[#2D3748]">
                        <span>Cliente: <strong className="text-white">{r.cliente}</strong></span>
                        {r.sap && (
                          <span className="font-['IBM_Plex_Mono',monospace] text-[#A0AEC0]">
                            SAP: {r.sap}
                          </span>
                        )}
                      </div>

                      {hasMulti && (
                        <div className="pt-1.5 border-t border-[#2D3748] bg-[#232A38]/50 p-2 rounded-lg space-y-1">
                          <span className="text-[10px] font-semibold text-[#F2A93B] uppercase tracking-wide block">
                            Modelos asociados ({r.modelos.length}):
                          </span>
                          <ul className="m-0 pl-3 list-disc text-[11px] text-[#CBD5E0] space-y-0.5">
                            {r.modelos.map((mo, i) => (
                              <li key={i}>
                                <span className="text-white">{mo.descripcion}</span>
                                {mo.sap && (
                                  <span className="font-['IBM_Plex_Mono',monospace] text-[#A0AEC0]"> · SAP {mo.sap}</span>
                                )}
                              </li>
                            ))}
                          </ul>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              {/* DESKTOP RESULTS VIEW (Table) */}
              <div className="hidden md:block border border-[#2D3748] rounded-xl overflow-x-auto bg-[#1A202C] shadow-sm">
                <table className="w-full min-w-[650px] text-xs border-collapse">
                  <thead>
                    <tr className="bg-[#232A38] border-b border-[#2D3748]">
                      <th className="text-left py-2 px-3 text-[11px] font-bold text-[#CBD5E0] uppercase tracking-wider font-['Oswald',sans-serif]">
                        OP
                      </th>
                      <th className="text-left py-2 px-3 text-[11px] font-bold text-[#CBD5E0] uppercase tracking-wider font-['Oswald',sans-serif]">
                        Descripción
                      </th>
                      <th className="text-left py-2 px-3 text-[11px] font-bold text-[#CBD5E0] uppercase tracking-wider font-['Oswald',sans-serif]">
                        Cliente
                      </th>
                      <th className="text-left py-2 px-3 text-[11px] font-bold text-[#CBD5E0] uppercase tracking-wider font-['Oswald',sans-serif]">
                        Cant. Total
                      </th>
                      <th className="text-left py-2 px-3 text-[11px] font-bold text-[#CBD5E0] uppercase tracking-wider font-['Oswald',sans-serif]">
                        SAP
                      </th>
                      <th className="py-2 px-3 text-right" />
                    </tr>
                  </thead>
                  <tbody>
                    {resultadosConsulta.slice(0, 200).map((r) => {
                      const hasMulti = r.modelos && r.modelos.length > 1;
                      return (
                        <React.Fragment key={r.op}>
                          <tr className="border-b border-[#2D3748] hover:bg-[#232A38] transition-colors">
                            <td className="py-2 px-3 font-['IBM_Plex_Mono',monospace] font-bold text-[#F6AD55]">
                              {r.op}
                            </td>
                            <td className="py-2 px-3 text-white font-medium">{r.descripcion}</td>
                            <td className="py-2 px-3 text-[#CBD5E0]">{r.cliente}</td>
                            <td className="py-2 px-3 font-['IBM_Plex_Mono',monospace] text-[#CBD5E0] tabular-nums">
                              {r.cantidadTotal || '—'}
                            </td>
                            <td className="py-2 px-3 font-['IBM_Plex_Mono',monospace] text-[#CBD5E0]">
                              {r.sap || '—'}
                            </td>
                            <td className="py-2 px-3 text-right">
                              <div className="flex gap-1 justify-end items-center">
                                <button
                                  type="button"
                                  onClick={() => copiarInfoOp(r)}
                                  title="Copiar OP, Descripción y Cliente"
                                  className="bg-[#232A38] border border-[#4A5568] text-[#CBD5E0] hover:text-white rounded-md p-1.5 cursor-pointer transition-colors"
                                >
                                  <Copy size={13} />
                                </button>
                                {hasMulti && (
                                  <button
                                    type="button"
                                    onClick={() => copiarModelosOp(r)}
                                    title="Copiar OP, Cliente y lista de modelos"
                                    className="bg-[#232A38] border border-[#F2A93B] text-[#F6AD55] rounded-md py-0.5 px-2 text-[10px] font-bold whitespace-nowrap cursor-pointer hover:bg-[#F2A93B] hover:text-[#12161F] transition-colors"
                                  >
                                    MODELOS
                                  </button>
                                )}
                              </div>
                            </td>
                          </tr>
                          {hasMulti && (
                            <tr className="border-b border-[#2D3748] bg-[#1A202C]/60">
                              <td />
                              <td colSpan={5} className="py-1.5 px-3 pb-2">
                                <ul className="m-0 pl-3 list-disc text-[11px] text-[#CBD5E0] space-y-0.5">
                                  {r.modelos.map((mo, i) => (
                                    <li key={i}>
                                      <span className="text-white font-medium">{mo.descripcion}</span>
                                      {mo.sap && (
                                        <>
                                          {' · SAP '}
                                          <span className="font-['IBM_Plex_Mono',monospace] text-[#F6AD55]">{mo.sap}</span>
                                        </>
                                      )}
                                      {mo.cantidadTotal && (
                                        <>
                                          {' · Cant. '}
                                          <span className="font-['IBM_Plex_Mono',monospace]">{mo.cantidadTotal}</span>
                                        </>
                                      )}
                                    </li>
                                  ))}
                                </ul>
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {resultadosConsulta.length === 0 && (
                <div className="text-center py-6 text-[#CBD5E0] text-xs bg-[#1A202C] rounded-xl border border-[#2D3748]">
                  Sin resultados {buscarQuery ? `para "${buscarQuery}"` : ''}
                </div>
              )}
              {resultadosConsulta.length > 200 && (
                <div className="text-center py-2.5 text-[#A0AEC0] text-[11px]">
                  Mostrando 200 de {resultadosConsulta.length} resultados. Afina tu búsqueda para ver más.
                </div>
              )}
            </>
          )}
        </div>
      )}

      {/* Floating Bottom Bar for Inventario: only displayed when there are items to export */}
      {lista.length > 0 && (
        <div className="fixed bottom-[52px] md:bottom-0 left-0 right-0 bg-[#1A202C]/95 backdrop-blur-md border-t border-[#F2A93B] z-30 shadow-2xl">
          <div className="max-w-[1440px] mx-auto px-3 sm:px-4 py-1.5 sm:py-2 flex items-center justify-between gap-3">
            <div className="text-xs text-[#CBD5E0]">
              <strong className="text-white font-semibold">{lista.length}</strong> pieza(s) ·{' '}
              <strong className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace] tabular-nums">
                {totalUnidadesLista.toLocaleString('es-PE')}
              </strong>{' '}
              u
            </div>
            <button
              type="button"
              onClick={() => exportarInventarioExcel(lista, catalogo)}
              className="flex items-center gap-1.5 bg-[#F2A93B] text-[#12161F] py-1 px-3 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer transition-all hover:brightness-105 active:scale-95 shadow-sm"
            >
              <Download size={13} />
              <span>Exportar Excel</span>
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

