import React, { useState, useRef, useMemo } from 'react';
import { useApp } from '../../context/AppContext';
import { Odometer } from '../common/Odometer';
import { Plus, RefreshCw, Upload, Check, Copy, Trash2, Search, X, RotateCcw, Download, Printer, Pencil, History, ArrowUpDown, FileText } from 'lucide-react';
import * as XLSX from 'xlsx';
import {
  procesarWorkbookCatalogo,
  buscarEnCatalogo,
  resolverModelo,
  exportarInventarioExcel,
  imprimirInventarioFormatoExcel,
} from '../../services/excel';
import { copiarAlPortapapeles } from '../../services/reports';
import { CatalogoItem, ListaItem } from '../../types';

export const InventarioView: React.FC = () => {
  const {
    catalogo,
    setCatalogo,
    lista,
    setLista,
    subTabInventario,
    setSubTabInventario,
    irAHistorialOp,
    showToast,
    session,
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

  // Search & Cleanup state for Consultar
  const [buscarQuery, setBuscarQuery] = useState('');
  const [campoBusqueda, setCampoBusqueda] = useState<'todos' | 'op' | 'descripcion' | 'cliente' | 'sap' | 'modelo'>('todos');
  const [filtroTipoOp, setFiltroTipoOp] = useState<'todos' | 'multi' | 'simple' | 'con_sap' | 'sin_sap'>('todos');
  const [filtroCliente, setFiltroCliente] = useState('');
  const [ordenCatalogo, setOrdenCatalogo] = useState<'desc' | 'asc' | 'cliente_asc' | 'cant_desc'>('desc');
  const [limiteResultados, setLimiteResultados] = useState(100);
  const [copiedKey, setCopiedKey] = useState<string | null>(null);

  // Inline edit state for pending list items
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editDescripcion, setEditDescripcion] = useState('');
  const [editTipoEmpaque, setEditTipoEmpaque] = useState<'caja' | 'paquete'>('caja');
  const [editUnidades, setEditUnidades] = useState('');
  const [editCajas, setEditCajas] = useState('');
  const [editObservacion, setEditObservacion] = useState('');

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Matched catalog item
  const matchActual = useMemo(() => {
    return buscarEnCatalogo(catalogo, opInput);
  }, [catalogo, opInput]);

  const opEncontrada = !!matchActual;

  const suggestions = useMemo(() => {
    const raw = opInput.trim();
    const rawLower = raw.toLowerCase();
    const digits = raw.replace(/\D/g, '');
    if (!digits && !raw) return [];
    return catalogo
      .filter((c) => {
        const cDigits = c.op.replace(/\D/g, '');
        return (digits && cDigits.includes(digits)) || c.op.toLowerCase().includes(rawLower);
      })
      .sort((a, b) => {
        const aLower = a.op.trim().toLowerCase();
        const bLower = b.op.trim().toLowerCase();
        const aExactStr = aLower === rawLower;
        const bExactStr = bLower === rawLower;
        if (aExactStr && !bExactStr) return -1;
        if (!aExactStr && bExactStr) return 1;

        const aDigits = a.op.replace(/\D/g, '');
        const bDigits = b.op.replace(/\D/g, '');
        const aExactDig = Boolean(digits && aDigits === digits);
        const bExactDig = Boolean(digits && bDigits === digits);
        if (aExactDig && !bExactDig) return -1;
        if (!aExactDig && bExactDig) return 1;

        const aStarts = digits ? aDigits.startsWith(digits) : aLower.startsWith(rawLower);
        const bStarts = digits ? bDigits.startsWith(digits) : bLower.startsWith(rawLower);
        if (aStarts && !bStarts) return -1;
        if (!aStarts && bStarts) return 1;
        return a.op.length - b.op.length;
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
    if (!digitsOnly) {
      setModeloSeleccionado(-1);
      setDescInput('');
      setClienteInput('');
      return;
    }
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
    if (isNaN(u) || isNaN(c) || u <= 0 || c <= 0) {
      showToast('error', `Las unidades y la cantidad de ${tipoEmpaque}s deben ser mayores a 0`);
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
    showToast('ok', 'Línea agregada a la lista');
    limpiarFormulario();
  };

  const handleExportarExcel = async () => {
    if (lista.length === 0) {
      showToast('error', 'No hay líneas en la lista para exportar');
      return;
    }
    try {
      await exportarInventarioExcel(lista, catalogo);
      showToast('ok', 'Excel de inventario exportado');
    } catch {
      showToast('error', 'No se pudo generar el Excel de inventario');
    }
  };

  const handleImprimirInventario = () => {
    if (lista.length === 0) {
      showToast('error', 'No hay líneas en la lista para imprimir');
      return;
    }
    const archivo = imprimirInventarioFormatoExcel(lista, catalogo);
    if (archivo) showToast('ok', `PDF generado: ${archivo}`);
  };

  const handleEliminarItem = (id: string) => {
    if (editingId === id) setEditingId(null);
    setLista((prev) => prev.filter((i) => i.id !== id));
  };

  const handleIniciarEdicion = (item: ListaItem) => {
    setEditingId(item.id);
    setEditDescripcion(item.descripcion || '');
    setEditTipoEmpaque(item.tipoEmpaque === 'paquete' ? 'paquete' : 'caja');
    setEditUnidades(String(item.unidadesXCaja || ''));
    setEditCajas(String(item.cantidadCajas || ''));
    setEditObservacion(item.observacion || '');
  };

  const handleCancelarEdicion = () => {
    setEditingId(null);
  };

  const handleGuardarEdicion = (id: string) => {
    const u = parseInt(editUnidades, 10);
    const c = parseInt(editCajas, 10);
    if (isNaN(u) || isNaN(c) || u <= 0 || c <= 0) {
      showToast('error', `Las unidades y la cantidad de ${editTipoEmpaque}s deben ser mayores a 0`);
      return;
    }
    setLista((prev) =>
      prev.map((item) =>
        item.id === id
          ? {
              ...item,
              descripcion: editDescripcion.trim() || item.descripcion,
              tipoEmpaque: editTipoEmpaque,
              unidadesXCaja: u,
              cantidadCajas: c,
              totalUnidades: u * c,
              observacion: editObservacion.trim(),
            }
          : item
      )
    );
    setEditingId(null);
    showToast('ok', 'Línea actualizada');
  };

  const cargarOpEnRegistro = (item: CatalogoItem, modeloIdx = -1) => {
    setOpInput(item.op);
    setModeloSeleccionado(modeloIdx);
    const mod = resolverModelo(item, modeloIdx);
    setDescInput(mod.descripcion);
    setClienteInput(item.cliente);
    setUnidadesInput('');
    setCajasInput('');
    setObservacionInput('');
    setShowSuggestions(false);
    setSubTabInventario('agregar');
    showToast('ok', `OP ${item.op} cargada para registrar`);
    setTimeout(() => {
      window.scrollTo({ top: 0, behavior: 'smooth' });
      document.getElementById('inventario-unidades-input')?.focus();
    }, 80);
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
      let nuevo: CatalogoItem[] | null = null;
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        nuevo = procesarWorkbookCatalogo(wb);
      } catch {
        showToast('error', 'No se pudo leer el archivo Excel');
        return;
      }

      if (!nuevo) {
        showToast('error', 'El Excel debe tener columnas OP, Descripción y Cliente');
        return;
      }

      if (session) {
        try {
          await subirCatalogoASupabase(nuevo);
        } catch (err: any) {
          showToast('error', err?.message || 'Error al subir el catálogo a la nube');
        }
      } else {
        setCatalogo((prev) => {
          const mapa = new Map<string, CatalogoItem>();
          prev.forEach((c) => mapa.set(c.op, c));
          nuevo!.forEach((c) => mapa.set(c.op, c));
          return Array.from(mapa.values());
        });
        showToast('ok', `Catálogo cargado localmente: ${nuevo.length} OP. Inicia sesión en Historial para sincronizar con los demás.`);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Sync catalog
  const handleActualizarCatalogo = async () => {
    setActualizando(true);
    try {
      const n = await sincronizarCatalogoSupabase(true);
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

  const normalizarBusqueda = (txt: any) =>
    String(txt || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .trim();

  const clientesUnicosCatalogo = useMemo(() => {
    const set = new Set<string>();
    catalogo.forEach((c) => {
      const cli = String(c.cliente || '').trim();
      if (cli) set.add(cli);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));
  }, [catalogo]);

  const tokensBusqueda = useMemo(() => {
    return normalizarBusqueda(buscarQuery).split(/\s+/).filter(Boolean);
  }, [buscarQuery]);

  const hayFiltrosActivos =
    buscarQuery.trim() !== '' ||
    campoBusqueda !== 'todos' ||
    filtroTipoOp !== 'todos' ||
    filtroCliente !== '';

  const limpiarFiltrosConsulta = () => {
    setBuscarQuery('');
    setCampoBusqueda('todos');
    setFiltroTipoOp('todos');
    setFiltroCliente('');
    setLimiteResultados(100);
  };

  // Filtered and cleanly sorted catalog for Consultar
  const resultadosConsulta = useMemo(() => {
    const qRaw = normalizarBusqueda(buscarQuery);
    const tokens = tokensBusqueda;

    const filtrados = catalogo.filter((c) => {
      // Filtro por cliente exacto
      if (filtroCliente && String(c.cliente || '').trim() !== filtroCliente) {
        return false;
      }

      // Filtro por tipo de OP / SAP
      const hasMulti = Boolean(c.modelos && c.modelos.length > 1);
      const hasSap =
        Boolean(String(c.sap || '').trim()) ||
        Boolean(c.modelos && c.modelos.some((mo) => Boolean(String(mo.sap || '').trim())));

      if (filtroTipoOp === 'multi' && !hasMulti) return false;
      if (filtroTipoOp === 'simple' && hasMulti) return false;
      if (filtroTipoOp === 'con_sap' && !hasSap) return false;
      if (filtroTipoOp === 'sin_sap' && hasSap) return false;

      if (tokens.length === 0) return true;

      const opTxt = normalizarBusqueda(c.op);
      const descTxt = normalizarBusqueda(
        [
          c.descripcion,
          c.descripcionBase,
          ...(c.modelos || []).map((mo) => `${mo.modelo || ''} ${mo.descripcion || ''}`),
        ].join(' ')
      );
      const cliTxt = normalizarBusqueda(c.cliente);
      const sapTxt = normalizarBusqueda(
        [c.sap, ...(c.modelos || []).map((mo) => mo.sap || '')].join(' ')
      );
      const modTxt = normalizarBusqueda(
        (c.modelos || []).map((mo) => `${mo.modelo || ''} ${mo.descripcion || ''} ${mo.sap || ''}`).join(' ')
      );

      let target = '';
      if (campoBusqueda === 'op') target = opTxt;
      else if (campoBusqueda === 'descripcion') target = descTxt;
      else if (campoBusqueda === 'cliente') target = cliTxt;
      else if (campoBusqueda === 'sap') target = sapTxt;
      else if (campoBusqueda === 'modelo') target = modTxt;
      else target = `${opTxt} ${descTxt} ${cliTxt} ${sapTxt}`;

      return tokens.every((tk) => target.includes(tk));
    });

    return filtrados.sort((a, b) => {
      if (qRaw && (campoBusqueda === 'todos' || campoBusqueda === 'op')) {
        const aOp = normalizarBusqueda(a.op);
        const bOp = normalizarBusqueda(b.op);
        const aExact = aOp === qRaw;
        const bExact = bOp === qRaw;
        if (aExact && !bExact) return -1;
        if (!aExact && bExact) return 1;
        const aStarts = aOp.startsWith(qRaw);
        const bStarts = bOp.startsWith(qRaw);
        if (aStarts && !bStarts) return -1;
        if (!aStarts && bStarts) return 1;
      }
      if (ordenCatalogo === 'cliente_asc') {
        const cCmp = String(a.cliente || '').localeCompare(String(b.cliente || ''), 'es', { sensitivity: 'base' });
        if (cCmp !== 0) return cCmp;
      } else if (ordenCatalogo === 'cant_desc') {
        const qA = Number(a.cantidadTotal) || 0;
        const qB = Number(b.cantidadTotal) || 0;
        if (qB !== qA) return qB - qA;
      }
      const cmp = String(a.op || '').localeCompare(String(b.op || ''), 'es', { numeric: true });
      return ordenCatalogo === 'asc' ? cmp : -cmp;
    });
  }, [catalogo, buscarQuery, tokensBusqueda, campoBusqueda, filtroTipoOp, filtroCliente, ordenCatalogo]);

  const modeloCoincideBusqueda = (mo: CatalogoItem['modelos'][number]) => {
    if (tokensBusqueda.length === 0) return false;
    if (campoBusqueda === 'op' || campoBusqueda === 'cliente') return false;
    const text = normalizarBusqueda(`${mo.modelo || ''} ${mo.descripcion || ''} ${mo.sap || ''}`);
    return tokensBusqueda.some((tk) => text.includes(tk));
  };

  const marcarCopiado = (key: string) => {
    setCopiedKey(key);
    window.setTimeout(() => {
      setCopiedKey((prev) => (prev === key ? null : prev));
    }, 1500);
  };

  const copiarInfoOp = async (r: CatalogoItem) => {
    const desc = r.modelos && r.modelos.length > 1 && r.descripcionBase ? r.descripcionBase : r.descripcion;
    const cantTxt =
      r.cantidadTotal && Number(r.cantidadTotal) > 0
        ? `\nCantidad Orden: ${Number(r.cantidadTotal).toLocaleString('es-PE')} u`
        : '';
    const sapTxt = r.sap ? `\nSAP: ${r.sap}` : '';
    const txt = `OP: ${r.op}\nDescripción: ${desc}\nCliente: ${r.cliente}${cantTxt}${sapTxt}`;
    const ok = await copiarAlPortapapeles(txt);
    if (ok) marcarCopiado(`op-${r.op}`);
    showToast(ok ? 'ok' : 'error', ok ? `Datos copiados: ${r.op}` : 'No se pudo copiar');
  };

  const copiarSoloDescripcionOp = async (r: CatalogoItem) => {
    const desc = (r.modelos && r.modelos.length > 1 && r.descripcionBase ? r.descripcionBase : r.descripcion) || '';
    if (!desc.trim()) {
      showToast('error', 'Esta OP no tiene descripción para copiar');
      return;
    }
    const ok = await copiarAlPortapapeles(desc.trim());
    if (ok) marcarCopiado(`desc-${r.op}`);
    showToast(ok ? 'ok' : 'error', ok ? `Descripción copiada: ${r.op}` : 'No se pudo copiar');
  };

  const obtenerNombreModelo = (r: CatalogoItem, mo: CatalogoItem['modelos'][number]) => {
    const base = r.descripcionBase || '';
    if (mo.modelo && mo.modelo.trim()) return mo.modelo.trim();
    let nombre = mo.descripcion || '';
    if (base && nombre.indexOf(base + ' - ') === 0) nombre = nombre.slice(base.length + 3);
    return nombre.trim();
  };

  const copiarModeloEspecifico = async (
    r: CatalogoItem,
    mo: CatalogoItem['modelos'][number],
    idx: number,
    soloDescripcion = false
  ) => {
    const descCompleta = (mo.descripcion || r.descripcion || '').trim();
    const nombreModelo = obtenerNombreModelo(r, mo);

    if (soloDescripcion) {
      const textoDesc = descCompleta || nombreModelo;
      if (!textoDesc) {
        showToast('error', 'Este modelo no tiene descripción para copiar');
        return;
      }
      const ok = await copiarAlPortapapeles(textoDesc);
      if (ok) marcarCopiado(`mod-desc-${r.op}-${idx}`);
      showToast(ok ? 'ok' : 'error', ok ? `Descripción de modelo copiada (${r.op})` : 'No se pudo copiar');
      return;
    }

    const modeloLinea =
      nombreModelo && descCompleta.toLowerCase() !== nombreModelo.toLowerCase()
        ? `\nModelo: ${nombreModelo}`
        : '';
    const cantMo =
      mo.cantidadTotal && Number(mo.cantidadTotal) > 0
        ? `\nCantidad Modelo: ${Number(mo.cantidadTotal).toLocaleString('es-PE')} u`
        : '';
    const sapVal = (mo.sap || r.sap || '').trim();
    const sapLinea = sapVal ? `\nSAP: ${sapVal}` : '';
    const txt = `OP: ${r.op}${modeloLinea}\nDescripción: ${descCompleta || nombreModelo}\nCliente: ${r.cliente}${cantMo}${sapLinea}`;
    const ok = await copiarAlPortapapeles(txt);
    if (ok) marcarCopiado(`mod-datos-${r.op}-${idx}`);
    showToast(ok ? 'ok' : 'error', ok ? `Modelo copiado: ${nombreModelo || r.op}` : 'No se pudo copiar');
  };

  const copiarModelosOp = async (r: CatalogoItem) => {
    if (!r.modelos) return;
    const base = r.descripcionBase || '';
    const cantTotalTxt =
      r.cantidadTotal && Number(r.cantidadTotal) > 0
        ? `${Number(r.cantidadTotal).toLocaleString('es-PE')} u`
        : '—';
    const txt =
      `OP: ${r.op}\nDescripción: ${base || r.descripcion}\nCliente: ${r.cliente}\nCantidad Total: ${cantTotalTxt}\nMODELOS:\n` +
      r.modelos
        .map((mo) => {
          const nombre = obtenerNombreModelo(r, mo);
          const cantMo =
            mo.cantidadTotal && Number(mo.cantidadTotal) > 0
              ? ` - ${Number(mo.cantidadTotal).toLocaleString('es-PE')} u`
              : '';
          const sapMo = mo.sap ? ` (SAP ${mo.sap})` : '';
          return `- ${nombre}${cantMo}${sapMo}`;
        })
        .join('\n');
    const ok = await copiarAlPortapapeles(txt);
    if (ok) marcarCopiado(`mods-${r.op}`);
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
                      setDescInput('');
                      setClienteInput('');
                      setModeloSeleccionado(-1);
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
                            <div className="flex items-center gap-1.5 flex-wrap">
                              <span className="font-bold text-[#F2A93B] group-hover:text-[#F6AD55] font-['IBM_Plex_Mono',monospace] text-xs px-1.5 py-0.5 rounded bg-[#232A38] border border-[#4A5568]/80">
                                {item.op}
                              </span>
                              {item.cantidadTotal && Number(item.cantidadTotal) > 0 ? (
                                <span className="text-[10px] font-bold text-[#F6AD55] font-['IBM_Plex_Mono',monospace] bg-[#1A202C] px-1.5 py-0.2 rounded border border-[#4A5568]">
                                  {Number(item.cantidadTotal).toLocaleString('es-PE')} u
                                </span>
                              ) : null}
                            </div>
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

            {/* Matched OP order info banner if found */}
            {matchActual && (() => {
              const modSel =
                modeloSeleccionado >= 0 && matchActual.modelos && matchActual.modelos[modeloSeleccionado]
                  ? matchActual.modelos[modeloSeleccionado]
                  : null;
              const cantMostrar =
                modSel && modSel.cantidadTotal
                  ? Number(modSel.cantidadTotal)
                  : matchActual.cantidadTotal
                  ? Number(matchActual.cantidadTotal)
                  : 0;
              if (!cantMostrar || cantMostrar <= 0) return null;
              return (
                <div className="mb-2 px-2.5 py-1.5 bg-[#232A38] rounded-lg border border-[#F2A93B]/30 flex items-center justify-between text-xs font-['IBM_Plex_Mono',monospace]">
                  <span className="text-[#CBD5E0] text-[11px]">
                    {modSel ? 'Cantidad del modelo:' : 'Cantidad total de la orden:'}
                  </span>
                  <span className="font-bold text-[#F6AD55]">{cantMostrar.toLocaleString('es-PE')} u</span>
                </div>
              );
            })()}

            {/* Model chips if multi-model */}
            {matchActual && matchActual.modelos && matchActual.modelos.length > 1 && (
              <div className="mb-2 p-2 bg-[#232A38] rounded-lg border border-[#2D3748]">
                <label className="block text-[10px] font-semibold text-[#F2A93B] mb-1.5 uppercase tracking-wide">
                  Modelos ({matchActual.modelos.length} — elige uno o Todos)
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
                          : 'bg-[#1A202C] text-[#CBD5E0] border-[#4A5568] hover:text-white hover:border-[#F2A93B]/60'
                      }`}
                    >
                      <span>{m.modelo || m.descripcion || `Modelo ${idx + 1}`}</span>
                      {m.cantidadTotal && Number(m.cantidadTotal) > 0 ? (
                        <span className="ml-1 text-[9.5px] opacity-90 font-['IBM_Plex_Mono',monospace] font-bold">
                          ({Number(m.cantidadTotal).toLocaleString('es-PE')} u)
                        </span>
                      ) : null}
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
                {lista.length > 0 && (
                  <button
                    type="button"
                    onClick={handleImprimirInventario}
                    className="flex items-center gap-1 text-[11px] text-white bg-[#2B6CB0] hover:bg-[#3182CE] py-1 px-2.5 rounded-md cursor-pointer transition-colors font-semibold"
                    title="Imprimir lista de inventario"
                  >
                    <Printer size={12} />
                    <span>Imprimir</span>
                  </button>
                )}
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
              <div className="space-y-2 max-h-[600px] overflow-y-auto pr-1 pb-12">
                {lista.map((item, idx) => {
                  const lineaNum = lista.length - idx;
                  const singular = item.tipoEmpaque === 'paquete' ? 'paquete' : 'caja';
                  const plural = item.tipoEmpaque === 'paquete' ? 'paquetes' : 'cajas';
                  const empaqueTxt = item.cantidadCajas === 1 ? singular : plural;
                  const estaEditando = editingId === item.id;
                  const editPreviewTotal =
                    (parseInt(editUnidades, 10) || 0) * (parseInt(editCajas, 10) || 0);

                  if (estaEditando) {
                    return (
                      <div
                        key={item.id}
                        className="bg-[#232A38] border border-[#F2A93B] rounded-xl p-2.5 sm:p-3 space-y-2 shadow-md"
                      >
                        <div className="flex items-center justify-between gap-2 pb-1 border-b border-[#2D3748]">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className="text-[10px] font-bold text-[#12161F] font-['IBM_Plex_Mono',monospace] bg-[#F2A93B] px-1.5 py-0.5 rounded">
                              #{lineaNum}
                            </span>
                            <span className="text-[10px] font-bold bg-[#3A2E1B] text-[#F6AD55] py-0.5 px-1.5 rounded border border-[#F2A93B]/30 font-['IBM_Plex_Mono',monospace]">
                              {item.op}
                            </span>
                            <span className="text-[11px] text-[#CBD5E0] font-medium truncate">
                              {item.cliente}
                            </span>
                          </div>
                          <span className="text-[10px] font-bold uppercase text-[#F6AD55] tracking-wider">
                            Editando línea
                          </span>
                        </div>

                        <div>
                          <label className="block text-[9.5px] font-semibold text-[#A0AEC0] uppercase mb-0.5">
                            Descripción / Modelo
                          </label>
                          <input
                            type="text"
                            value={editDescripcion}
                            onChange={(e) => setEditDescripcion(e.target.value)}
                            className="w-full py-1 px-2 rounded-lg border border-[#4A5568] bg-[#1A202C] text-xs text-white outline-none focus:border-[#F2A93B]"
                          />
                        </div>

                        <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                          <div>
                            <label className="block text-[9.5px] font-semibold text-[#A0AEC0] uppercase mb-0.5">
                              Empaque
                            </label>
                            <div className="grid grid-cols-2 gap-1 p-0.5 bg-[#1A202C] rounded-lg border border-[#4A5568]">
                              <button
                                type="button"
                                onClick={() => setEditTipoEmpaque('caja')}
                                className={`py-1 rounded text-[11px] font-bold cursor-pointer ${
                                  editTipoEmpaque === 'caja'
                                    ? 'bg-[#F2A93B] text-[#12161F]'
                                    : 'text-[#CBD5E0]'
                                }`}
                              >
                                Cajas
                              </button>
                              <button
                                type="button"
                                onClick={() => setEditTipoEmpaque('paquete')}
                                className={`py-1 rounded text-[11px] font-bold cursor-pointer ${
                                  editTipoEmpaque === 'paquete'
                                    ? 'bg-[#F2A93B] text-[#12161F]'
                                    : 'text-[#CBD5E0]'
                                }`}
                              >
                                Paquetes
                              </button>
                            </div>
                          </div>

                          <div>
                            <label className="block text-[9.5px] font-semibold text-[#A0AEC0] uppercase mb-0.5">
                              Unid. x {editTipoEmpaque}
                            </label>
                            <input
                              type="number"
                              inputMode="numeric"
                              min="1"
                              value={editUnidades}
                              onChange={(e) => setEditUnidades(e.target.value)}
                              className="w-full py-1 px-2 rounded-lg border border-[#4A5568] bg-[#1A202C] text-xs font-['IBM_Plex_Mono',monospace] text-white outline-none focus:border-[#F2A93B]"
                            />
                          </div>

                          <div>
                            <label className="block text-[9.5px] font-semibold text-[#A0AEC0] uppercase mb-0.5">
                              Cant. {editTipoEmpaque}s
                            </label>
                            <input
                              type="number"
                              inputMode="numeric"
                              min="1"
                              value={editCajas}
                              onChange={(e) => setEditCajas(e.target.value)}
                              className="w-full py-1 px-2 rounded-lg border border-[#4A5568] bg-[#1A202C] text-xs font-['IBM_Plex_Mono',monospace] text-white outline-none focus:border-[#F2A93B]"
                            />
                          </div>
                        </div>

                        <div>
                          <label className="block text-[9.5px] font-semibold text-[#A0AEC0] uppercase mb-0.5">
                            Observación
                          </label>
                          <input
                            type="text"
                            value={editObservacion}
                            onChange={(e) => setEditObservacion(e.target.value)}
                            placeholder="Opcional..."
                            className="w-full py-1 px-2 rounded-lg border border-[#4A5568] bg-[#1A202C] text-xs text-white outline-none focus:border-[#F2A93B]"
                          />
                        </div>

                        <div className="flex items-center justify-between gap-2 pt-1 border-t border-[#2D3748]">
                          <div className="text-xs text-[#CBD5E0]">
                            Nuevo total:{' '}
                            <strong className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace] tabular-nums">
                              {editPreviewTotal.toLocaleString('es-PE')} u
                            </strong>
                          </div>
                          <div className="flex items-center gap-1.5">
                            <button
                              type="button"
                              onClick={handleCancelarEdicion}
                              className="flex items-center gap-1 py-1 px-2.5 rounded-lg border border-[#4A5568] bg-[#1A202C] text-[#CBD5E0] hover:text-white text-xs font-semibold cursor-pointer"
                            >
                              <X size={13} />
                              <span>Cancelar</span>
                            </button>
                            <button
                              type="button"
                              onClick={() => handleGuardarEdicion(item.id)}
                              className="flex items-center gap-1 py-1 px-3 rounded-lg bg-[#38A169] hover:brightness-110 text-white text-xs font-bold cursor-pointer"
                            >
                              <Check size={13} />
                              <span>Guardar</span>
                            </button>
                          </div>
                        </div>
                      </div>
                    );
                  }

                  return (
                    <div
                      key={item.id}
                      className="bg-[#1A202C] border border-[#2D3748] hover:border-[#4A5568] rounded-xl p-2.5 sm:p-3 flex items-center justify-between gap-2.5 transition-colors shadow-sm"
                    >
                      <div className="min-w-0 flex-1">
                        <div className="flex items-center gap-1.5 flex-wrap mb-0.5">
                          <span className="text-[10px] font-bold text-[#A0AEC0] font-['IBM_Plex_Mono',monospace] bg-[#232A38] px-1.5 py-0.5 rounded border border-[#2D3748]">
                            #{lineaNum}
                          </span>
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
                      <div className="flex items-center gap-1 shrink-0">
                        <button
                          type="button"
                          onClick={() => handleIniciarEdicion(item)}
                          className="p-1.5 text-[#93C5FD] hover:bg-[#2B6CB0]/25 active:bg-[#2B6CB0]/40 rounded-lg cursor-pointer transition-colors w-7 h-7 sm:w-8 sm:h-8 flex items-center justify-center"
                          title="Editar línea"
                          aria-label="Editar línea"
                        >
                          <Pencil size={14} />
                        </button>
                        <button
                          type="button"
                          onClick={() => handleEliminarItem(item.id)}
                          className="p-1.5 text-[#E53E3E] hover:bg-[#E53E3E]/20 active:bg-[#E53E3E]/30 rounded-lg cursor-pointer transition-colors w-7 h-7 sm:w-8 sm:h-8 flex items-center justify-center"
                          title="Eliminar pieza"
                          aria-label="Eliminar pieza"
                        >
                          <Trash2 size={15} />
                        </button>
                      </div>
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
          {/* Buscador + Filtros */}
          <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-2.5 space-y-2 shadow-sm">
            <div className="flex flex-col sm:flex-row gap-2">
              <div className="flex flex-1 min-w-0 gap-1.5">
                <select
                  value={campoBusqueda}
                  onChange={(e) => {
                    setCampoBusqueda(e.target.value as any);
                    setLimiteResultados(100);
                  }}
                  aria-label="Campo de búsqueda"
                  className="py-1.5 px-2 rounded-lg border border-[#4A5568] bg-[#232A38] text-[#F6AD55] text-[11px] font-bold outline-none focus:border-[#F2A93B] cursor-pointer shrink-0"
                  title="Elegir en qué campo buscar"
                >
                  <option value="todos">Todo</option>
                  <option value="op">OP</option>
                  <option value="descripcion">Descripción</option>
                  <option value="modelo">Modelo</option>
                  <option value="cliente">Cliente</option>
                  <option value="sap">SAP</option>
                </select>
                <div className="relative flex-1 min-w-0">
                  <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-[#CBD5E0]" />
                  <input
                    type="text"
                    value={buscarQuery}
                    onChange={(e) => {
                      setBuscarQuery(e.target.value);
                      setLimiteResultados(100);
                    }}
                    placeholder={
                      campoBusqueda === 'op'
                        ? 'Buscar por número de OP…'
                        : campoBusqueda === 'modelo'
                        ? 'Buscar por nombre o modelo…'
                        : campoBusqueda === 'cliente'
                        ? 'Buscar por cliente…'
                        : campoBusqueda === 'sap'
                        ? 'Buscar por código SAP…'
                        : campoBusqueda === 'descripcion'
                        ? 'Buscar palabras en descripción…'
                        : 'Buscar por OP, modelo, descripción, cliente o SAP (acepta varias palabras)…'
                    }
                    className="w-full py-1.5 pl-9 pr-8 rounded-lg border border-[#4A5568] text-xs outline-none bg-[#232A38] text-white focus:border-[#F2A93B] focus:ring-1 focus:ring-[#F2A93B]/30"
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
              </div>

              <div className="flex items-center gap-1.5 shrink-0">
                <div className="flex items-center gap-1 py-1 px-2 rounded-lg border border-[#4A5568] bg-[#232A38]">
                  <ArrowUpDown size={13} className="text-[#F2A93B] shrink-0" />
                  <select
                    value={ordenCatalogo}
                    onChange={(e) => setOrdenCatalogo(e.target.value as any)}
                    aria-label="Orden del catálogo"
                    className="bg-transparent text-xs font-semibold text-[#CBD5E0] outline-none cursor-pointer"
                  >
                    <option value="desc" className="bg-[#1A202C] text-white">OP ↓ Recientes</option>
                    <option value="asc" className="bg-[#1A202C] text-white">OP ↑ Antiguas</option>
                    <option value="cliente_asc" className="bg-[#1A202C] text-white">Cliente A-Z</option>
                    <option value="cant_desc" className="bg-[#1A202C] text-white">Cantidad ↓ Mayor</option>
                  </select>
                </div>
              </div>
            </div>

            {/* Fila de filtros avanzados */}
            <div className="flex flex-wrap items-center justify-between gap-2 pt-1.5 border-t border-[#2D3748]">
              <div className="flex flex-wrap items-center gap-1.5">
                {(
                  [
                    { id: 'todos', label: 'Todas' },
                    { id: 'multi', label: 'Con modelos' },
                    { id: 'simple', label: 'Simples' },
                    { id: 'con_sap', label: 'Con SAP' },
                    { id: 'sin_sap', label: 'Sin SAP' },
                  ] as const
                ).map((t) => (
                  <button
                    key={t.id}
                    type="button"
                    onClick={() => {
                      setFiltroTipoOp(t.id);
                      setLimiteResultados(100);
                    }}
                    className={`py-0.5 px-2 rounded-md text-[11px] font-semibold border cursor-pointer transition-colors ${
                      filtroTipoOp === t.id
                        ? 'bg-[#F2A93B] text-[#12161F] border-[#F2A93B]'
                        : 'bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:text-white'
                    }`}
                  >
                    {t.label}
                  </button>
                ))}

                {clientesUnicosCatalogo.length > 0 && (
                  <select
                    value={filtroCliente}
                    onChange={(e) => {
                      setFiltroCliente(e.target.value);
                      setLimiteResultados(100);
                    }}
                    aria-label="Filtrar por cliente"
                    className={`py-0.5 px-2 rounded-md border text-[11px] font-semibold outline-none cursor-pointer max-w-[200px] ${
                      filtroCliente
                        ? 'bg-[#F2A93B]/20 border-[#F2A93B] text-[#F6AD55]'
                        : 'bg-[#232A38] border-[#4A5568] text-[#CBD5E0]'
                    }`}
                  >
                    <option value="" className="bg-[#1A202C] text-white">
                      Todos los clientes ({clientesUnicosCatalogo.length})
                    </option>
                    {clientesUnicosCatalogo.map((cli) => (
                      <option key={cli} value={cli} className="bg-[#1A202C] text-white">
                        {cli}
                      </option>
                    ))}
                  </select>
                )}
              </div>

              <div className="flex items-center gap-2 text-[11px] text-[#A0AEC0]">
                <span>
                  <strong className="text-white font-['IBM_Plex_Mono',monospace]">
                    {resultadosConsulta.length}
                  </strong>{' '}
                  de {catalogo.length} OP
                </span>
                {hayFiltrosActivos && (
                  <button
                    type="button"
                    onClick={limpiarFiltrosConsulta}
                    className="inline-flex items-center gap-1 text-[11px] text-[#F6AD55] hover:text-white bg-[#232A38] border border-[#F2A93B]/40 rounded px-1.5 py-0.5 cursor-pointer"
                  >
                    <X size={11} />
                    <span>Limpiar filtros</span>
                  </button>
                )}
              </div>
            </div>
          </div>

          {catalogo.length === 0 ? (
            <p className="text-center py-6 text-[#CBD5E0] text-xs bg-[#1A202C] rounded-xl border border-[#2D3748]">
              Importa un catálogo para poder realizar consultas.
            </p>
          ) : (
            <>
              {/* MOBILE RESULTS VIEW (Touch Cards) */}
              <div className="block md:hidden space-y-2">
                {resultadosConsulta.slice(0, limiteResultados).map((r) => {
                  const hasMulti = r.modelos && r.modelos.length > 1;
                  return (
                    <div
                      key={r.op}
                      className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-2.5 space-y-1.5 shadow-sm"
                    >
                      <div className="flex items-center justify-between gap-2">
                        <div className="flex items-center gap-2 flex-wrap">
                          <span className="font-['IBM_Plex_Mono',monospace] font-bold text-[#F6AD55] text-xs">
                            {r.op}
                          </span>
                          {r.cantidadTotal && Number(r.cantidadTotal) > 0 ? (
                            <span className="bg-[#F2A93B]/20 text-[#F6AD55] border border-[#F2A93B]/40 px-1.5 py-0.5 rounded text-[10.5px] font-bold font-['IBM_Plex_Mono',monospace]">
                              {Number(r.cantidadTotal).toLocaleString('es-PE')} u
                            </span>
                          ) : (
                            <span className="text-[#A0AEC0] text-[10px] font-['IBM_Plex_Mono',monospace]">
                              Sin cant.
                            </span>
                          )}
                        </div>
                        <div className="flex items-center gap-1 flex-wrap justify-end">
                          <button
                            type="button"
                            onClick={() => cargarOpEnRegistro(r)}
                            title="Cargar OP en Registrar / Agregar"
                            className="bg-[#F2A93B] text-[#12161F] rounded-md px-2 py-1 text-[10px] font-bold flex items-center gap-1 cursor-pointer hover:brightness-110 active:scale-95"
                          >
                            <Plus size={12} />
                            <span>Registrar</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => irAHistorialOp(r.op)}
                            title="Ver producción acumulada de esta OP en Historial"
                            className="bg-[#2B6CB0]/30 border border-[#63B3ED]/40 text-[#93C5FD] rounded-md px-2 py-1 text-[10px] font-bold flex items-center gap-1 cursor-pointer hover:bg-[#2B6CB0]/50 active:scale-95"
                          >
                            <History size={12} />
                            <span>Historial</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => copiarInfoOp(r)}
                            title="Copiar datos completos de la OP"
                            className={`border rounded-md px-1.5 py-1 text-[10px] font-semibold flex items-center gap-1 cursor-pointer ${
                              copiedKey === `op-${r.op}`
                                ? 'bg-[#38A169]/20 border-[#38A169] text-[#68D391]'
                                : 'bg-[#232A38] border-[#4A5568] text-[#CBD5E0] active:text-white hover:bg-[#2D3748]'
                            }`}
                          >
                            {copiedKey === `op-${r.op}` ? <Check size={12} /> : <Copy size={12} />}
                            <span>Datos</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => copiarSoloDescripcionOp(r)}
                            title="Copiar solo la descripción de la OP"
                            className={`border rounded-md px-1.5 py-1 text-[10px] font-semibold flex items-center gap-1 cursor-pointer ${
                              copiedKey === `desc-${r.op}`
                                ? 'bg-[#38A169]/20 border-[#38A169] text-[#68D391]'
                                : 'bg-[#232A38] border-[#4A5568] text-[#CBD5E0] active:text-white hover:bg-[#2D3748]'
                            }`}
                          >
                            {copiedKey === `desc-${r.op}` ? <Check size={12} /> : <FileText size={12} />}
                            <span>Desc.</span>
                          </button>
                          {hasMulti && (
                            <button
                              type="button"
                              onClick={() => copiarModelosOp(r)}
                              title="Copiar todos los modelos de la OP"
                              className={`border rounded-md px-2 py-1 text-[10px] font-bold flex items-center gap-1 cursor-pointer ${
                                copiedKey === `mods-${r.op}`
                                  ? 'bg-[#38A169]/20 border-[#38A169] text-[#68D391]'
                                  : 'bg-[#232A38] border-[#F2A93B] text-[#F6AD55] hover:bg-[#F2A93B] hover:text-[#12161F]'
                              }`}
                            >
                              {copiedKey === `mods-${r.op}` ? <Check size={11} /> : null}
                              <span>MODELOS</span>
                            </button>
                          )}
                        </div>
                      </div>

                      <div className="text-xs font-semibold text-white break-words">
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
                        <div className="pt-1.5 border-t border-[#2D3748] bg-[#232A38]/50 p-2 rounded-lg space-y-1.5">
                          <span className="text-[10px] font-semibold text-[#F2A93B] uppercase tracking-wide block">
                            Modelos asociados ({r.modelos.length}):
                          </span>
                          <div className="space-y-1.5">
                            {r.modelos.map((mo, i) => {
                              const resalta = modeloCoincideBusqueda(mo);
                              return (
                                <div
                                  key={i}
                                  className={`rounded-md p-1.5 border space-y-1 ${
                                    resalta
                                      ? 'bg-[#F2A93B]/10 border-[#F2A93B]/60'
                                      : 'bg-[#1A202C]/70 border-[#2D3748]'
                                  }`}
                                >
                                  <div className="text-[11px] text-white font-medium break-words">
                                    {mo.descripcion}
                                  </div>
                                  <div className="flex items-center justify-between gap-1.5 flex-wrap text-[10px]">
                                    <div className="flex items-center gap-1.5 flex-wrap">
                                      {mo.cantidadTotal && Number(mo.cantidadTotal) > 0 ? (
                                        <span className="text-[#F6AD55] font-['IBM_Plex_Mono',monospace] font-bold bg-[#1A202C] px-1.5 py-0.5 rounded border border-[#4A5568]">
                                          {Number(mo.cantidadTotal).toLocaleString('es-PE')} u
                                        </span>
                                      ) : null}
                                      {mo.sap && (
                                        <span className="font-['IBM_Plex_Mono',monospace] text-[#A0AEC0]">
                                          SAP {mo.sap}
                                        </span>
                                      )}
                                    </div>
                                    <div className="flex items-center gap-1">
                                      <button
                                        type="button"
                                        onClick={() => copiarModeloEspecifico(r, mo, i, false)}
                                        title={`Copiar datos completos de este modelo`}
                                        className={`inline-flex items-center gap-0.5 border rounded px-1.5 py-0.5 text-[9.5px] font-semibold cursor-pointer ${
                                          copiedKey === `mod-datos-${r.op}-${i}`
                                            ? 'bg-[#38A169]/20 border-[#38A169] text-[#68D391]'
                                            : 'bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:text-white'
                                        }`}
                                      >
                                        {copiedKey === `mod-datos-${r.op}-${i}` ? <Check size={10} /> : <Copy size={10} />}
                                        <span>Datos</span>
                                      </button>
                                      <button
                                        type="button"
                                        onClick={() => copiarModeloEspecifico(r, mo, i, true)}
                                        title={`Copiar solo descripción de este modelo`}
                                        className={`inline-flex items-center gap-0.5 border rounded px-1.5 py-0.5 text-[9.5px] font-semibold cursor-pointer ${
                                          copiedKey === `mod-desc-${r.op}-${i}`
                                            ? 'bg-[#38A169]/20 border-[#38A169] text-[#68D391]'
                                            : 'bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:text-white'
                                        }`}
                                      >
                                        {copiedKey === `mod-desc-${r.op}-${i}` ? <Check size={10} /> : <FileText size={10} />}
                                        <span>Desc.</span>
                                      </button>
                                      <button
                                        type="button"
                                        onClick={() => cargarOpEnRegistro(r, i)}
                                        title={`Registrar modelo: ${mo.descripcion}`}
                                        className="inline-flex items-center gap-0.5 bg-[#3A2E1B] text-[#F6AD55] border border-[#F2A93B]/40 rounded px-1.5 py-0.5 text-[9.5px] font-bold cursor-pointer hover:bg-[#F2A93B] hover:text-[#12161F]"
                                      >
                                        <Plus size={10} />
                                        <span>Reg.</span>
                                      </button>
                                    </div>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              {/* DESKTOP RESULTS VIEW (Table) */}
              <div className="hidden md:block border border-[#2D3748] rounded-xl overflow-x-auto bg-[#1A202C] shadow-sm">
                <table className="w-full min-w-[720px] text-xs border-collapse">
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
                    {resultadosConsulta.slice(0, limiteResultados).map((r) => {
                      const hasMulti = r.modelos && r.modelos.length > 1;
                      return (
                        <React.Fragment key={r.op}>
                          <tr className="border-b border-[#2D3748] hover:bg-[#232A38] transition-colors">
                            <td className="py-2 px-3 font-['IBM_Plex_Mono',monospace] font-bold text-[#F6AD55] whitespace-nowrap">
                              {r.op}
                            </td>
                            <td className="py-2 px-3 text-white font-medium">{r.descripcion}</td>
                            <td className="py-2 px-3 text-[#CBD5E0]">{r.cliente}</td>
                            <td className="py-2 px-3 font-['IBM_Plex_Mono',monospace] tabular-nums whitespace-nowrap">
                              {r.cantidadTotal && Number(r.cantidadTotal) > 0 ? (
                                <span className="font-bold text-[#F6AD55]">
                                  {Number(r.cantidadTotal).toLocaleString('es-PE')} u
                                </span>
                              ) : (
                                <span className="text-[#718096]">—</span>
                              )}
                            </td>
                            <td className="py-2 px-3 font-['IBM_Plex_Mono',monospace] text-[#CBD5E0] whitespace-nowrap">
                              {r.sap || '—'}
                            </td>
                            <td className="py-2 px-3 text-right">
                              <div className="flex gap-1 justify-end items-center">
                                <button
                                  type="button"
                                  onClick={() => cargarOpEnRegistro(r)}
                                  title="Cargar OP en Registrar / Agregar"
                                  className="bg-[#F2A93B] text-[#12161F] rounded-md py-1 px-2 text-[10.5px] font-bold flex items-center gap-1 whitespace-nowrap cursor-pointer hover:brightness-110 transition-colors"
                                >
                                  <Plus size={12} />
                                  <span>Registrar</span>
                                </button>
                                <button
                                  type="button"
                                  onClick={() => irAHistorialOp(r.op)}
                                  title="Ver historial de producción de esta OP"
                                  className="bg-[#2B6CB0]/25 border border-[#63B3ED]/40 text-[#93C5FD] rounded-md py-1 px-2 text-[10.5px] font-bold flex items-center gap-1 whitespace-nowrap cursor-pointer hover:bg-[#2B6CB0]/45 transition-colors"
                                >
                                  <History size={12} />
                                  <span>Historial</span>
                                </button>
                                <button
                                  type="button"
                                  onClick={() => copiarInfoOp(r)}
                                  title="Copiar OP, Descripción, Cliente, Cantidad y SAP"
                                  className={`border rounded-md py-1 px-1.5 text-[10px] font-semibold flex items-center gap-1 whitespace-nowrap cursor-pointer transition-colors ${
                                    copiedKey === `op-${r.op}`
                                      ? 'bg-[#38A169]/20 border-[#38A169] text-[#68D391]'
                                      : 'bg-[#232A38] border-[#4A5568] text-[#CBD5E0] hover:text-white'
                                  }`}
                                >
                                  {copiedKey === `op-${r.op}` ? <Check size={12} /> : <Copy size={12} />}
                                  <span>Datos</span>
                                </button>
                                <button
                                  type="button"
                                  onClick={() => copiarSoloDescripcionOp(r)}
                                  title="Copiar solo la descripción de la OP"
                                  className={`border rounded-md py-1 px-1.5 text-[10px] font-semibold flex items-center gap-1 whitespace-nowrap cursor-pointer transition-colors ${
                                    copiedKey === `desc-${r.op}`
                                      ? 'bg-[#38A169]/20 border-[#38A169] text-[#68D391]'
                                      : 'bg-[#232A38] border-[#4A5568] text-[#CBD5E0] hover:text-white'
                                  }`}
                                >
                                  {copiedKey === `desc-${r.op}` ? <Check size={12} /> : <FileText size={12} />}
                                  <span>Desc.</span>
                                </button>
                                {hasMulti && (
                                  <button
                                    type="button"
                                    onClick={() => copiarModelosOp(r)}
                                    title="Copiar OP, Cliente y todos los modelos"
                                    className={`border rounded-md py-1 px-2 text-[10px] font-bold whitespace-nowrap flex items-center gap-1 cursor-pointer transition-colors ${
                                      copiedKey === `mods-${r.op}`
                                        ? 'bg-[#38A169]/20 border-[#38A169] text-[#68D391]'
                                        : 'bg-[#232A38] border-[#F2A93B] text-[#F6AD55] hover:bg-[#F2A93B] hover:text-[#12161F]'
                                    }`}
                                  >
                                    {copiedKey === `mods-${r.op}` ? <Check size={11} /> : null}
                                    <span>MODELOS</span>
                                  </button>
                                )}
                              </div>
                            </td>
                          </tr>
                          {hasMulti && (
                            <tr className="border-b border-[#2D3748] bg-[#1A202C]/60">
                              <td />
                              <td colSpan={5} className="py-2 px-3">
                                <div className="space-y-1">
                                  {r.modelos.map((mo, i) => {
                                    const resalta = modeloCoincideBusqueda(mo);
                                    return (
                                      <div
                                        key={i}
                                        className={`flex items-center justify-between gap-2 px-2 py-1 rounded-md border text-[11px] ${
                                          resalta
                                            ? 'bg-[#F2A93B]/10 border-[#F2A93B]/60'
                                            : 'bg-[#232A38]/40 border-[#2D3748]/80'
                                        }`}
                                      >
                                        <div className="flex items-center gap-2 flex-wrap min-w-0">
                                          <span className="text-white font-medium">{mo.descripcion}</span>
                                          {mo.cantidadTotal && Number(mo.cantidadTotal) > 0 ? (
                                            <span className="font-['IBM_Plex_Mono',monospace] font-bold text-[#F6AD55] bg-[#1A202C] px-1.5 py-0.5 rounded border border-[#4A5568] text-[10px]">
                                              {Number(mo.cantidadTotal).toLocaleString('es-PE')} u
                                            </span>
                                          ) : null}
                                          {mo.sap && (
                                            <span className="font-['IBM_Plex_Mono',monospace] text-[#A0AEC0] text-[10px]">
                                              SAP {mo.sap}
                                            </span>
                                          )}
                                        </div>
                                        <div className="flex items-center gap-1 shrink-0">
                                          <button
                                            type="button"
                                            onClick={() => copiarModeloEspecifico(r, mo, i, false)}
                                            title="Copiar datos completos de este modelo"
                                            className={`inline-flex items-center gap-1 border rounded px-1.5 py-0.5 text-[10px] font-semibold cursor-pointer transition-colors ${
                                              copiedKey === `mod-datos-${r.op}-${i}`
                                                ? 'bg-[#38A169]/20 border-[#38A169] text-[#68D391]'
                                                : 'bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:text-white'
                                            }`}
                                          >
                                            {copiedKey === `mod-datos-${r.op}-${i}` ? <Check size={10} /> : <Copy size={10} />}
                                            <span>Datos</span>
                                          </button>
                                          <button
                                            type="button"
                                            onClick={() => copiarModeloEspecifico(r, mo, i, true)}
                                            title="Copiar solo la descripción de este modelo"
                                            className={`inline-flex items-center gap-1 border rounded px-1.5 py-0.5 text-[10px] font-semibold cursor-pointer transition-colors ${
                                              copiedKey === `mod-desc-${r.op}-${i}`
                                                ? 'bg-[#38A169]/20 border-[#38A169] text-[#68D391]'
                                                : 'bg-[#232A38] text-[#CBD5E0] border-[#4A5568] hover:text-white'
                                            }`}
                                          >
                                            {copiedKey === `mod-desc-${r.op}-${i}` ? <Check size={10} /> : <FileText size={10} />}
                                            <span>Desc.</span>
                                          </button>
                                          <button
                                            type="button"
                                            onClick={() => cargarOpEnRegistro(r, i)}
                                            title={`Registrar modelo: ${mo.descripcion}`}
                                            className="inline-flex items-center gap-0.5 bg-[#3A2E1B] text-[#F6AD55] border border-[#F2A93B]/40 rounded px-1.5 py-0.5 text-[10px] font-bold cursor-pointer hover:bg-[#F2A93B] hover:text-[#12161F] transition-colors"
                                          >
                                            <Plus size={10} />
                                            <span>Registrar modelo</span>
                                          </button>
                                        </div>
                                      </div>
                                    );
                                  })}
                                </div>
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
                <div className="text-center py-6 text-[#CBD5E0] text-xs bg-[#1A202C] rounded-xl border border-[#2D3748] space-y-2">
                  <div>Sin resultados para los filtros actuales</div>
                  {hayFiltrosActivos && (
                    <button
                      type="button"
                      onClick={limpiarFiltrosConsulta}
                      className="inline-flex items-center gap-1 py-1 px-2.5 rounded-lg bg-[#232A38] border border-[#4A5568] text-[#F6AD55] hover:text-white text-[11px] font-semibold cursor-pointer"
                    >
                      <X size={12} />
                      <span>Limpiar filtros</span>
                    </button>
                  )}
                </div>
              )}
              {resultadosConsulta.length > limiteResultados && (
                <div className="text-center py-2.5 flex items-center justify-center gap-3">
                  <span className="text-[#A0AEC0] text-[11px]">
                    Mostrando {limiteResultados} de {resultadosConsulta.length} resultados.
                  </span>
                  <button
                    type="button"
                    onClick={() => setLimiteResultados((prev) => prev + 100)}
                    className="py-1 px-3 rounded-lg bg-[#232A38] border border-[#4A5568] text-[#F6AD55] hover:text-white text-[11px] font-semibold cursor-pointer"
                  >
                    Mostrar 100 más
                  </button>
                </div>
              )}
            </>
          )}
        </div>
      )}

      {/* Floating Bottom Bar for Inventario: only displayed when there are items to export */}
      {lista.length > 0 && (
        <div className="fixed bottom-[52px] md:bottom-0 left-0 right-0 bg-[#1A202C]/95 backdrop-blur-md border-t border-[#F2A93B] z-30 shadow-2xl">
          <div className="max-w-[1440px] mx-auto px-3 sm:px-4 py-1.5 sm:py-2 flex items-center justify-between gap-2.5">
            <div className="text-xs text-[#CBD5E0]">
              <strong className="text-white font-semibold">{lista.length}</strong> línea(s) ·{' '}
              <strong className="text-[#F2A93B] font-['IBM_Plex_Mono',monospace] tabular-nums">
                {totalUnidadesLista.toLocaleString('es-PE')}
              </strong>{' '}
              u
            </div>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={handleImprimirInventario}
                className="flex items-center gap-1.5 bg-[#2B6CB0] hover:bg-[#3182CE] text-white py-1 px-2.5 sm:px-3 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer transition-all active:scale-95 shadow-sm"
                title="Imprimir lista de inventario"
              >
                <Printer size={13} />
                <span>Imprimir</span>
              </button>
              <button
                type="button"
                onClick={handleExportarExcel}
                className="flex items-center gap-1.5 bg-[#F2A93B] text-[#12161F] py-1 px-2.5 sm:px-3 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer transition-all hover:brightness-105 active:scale-95 shadow-sm"
              >
                <Download size={13} />
                <span>Exportar Excel</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

