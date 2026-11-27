import React, { useState, useMemo } from 'react';
import {
  Monitor,
  RefreshCw,
  ExternalLink,
  Settings,
  Tv,
  LayoutTemplate,
  Globe,
  AlertTriangle,
  Check,
  Maximize2,
  Minimize2,
} from 'lucide-react';
import { STORAGE_KEYS } from '../../services/storage';
import { useApp } from '../../context/AppContext';

type TareoEmbedMode = 'tv' | 'embed' | 'normal';

function normalizarUrlBase(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed) return '';
  if (/^https?:\/\//i.test(trimmed)) return trimmed;
  return `http://${trimmed}`;
}

function construirUrlTareo(baseRaw: string, mode: TareoEmbedMode): string {
  const normalized = normalizarUrlBase(baseRaw);
  if (!normalized) return '';
  try {
    const u = new URL(normalized);
    u.searchParams.delete('tv');
    u.searchParams.delete('embed');
    if (mode === 'tv') {
      u.searchParams.set('tv', '1');
    } else if (mode === 'embed') {
      u.searchParams.set('embed', 'live');
    }
    return u.toString();
  } catch {
    return '';
  }
}

export const TareoView: React.FC = () => {
  const { showToast } = useApp();

  const [savedUrl, setSavedUrl] = useState<string>(() => {
    try {
      return localStorage.getItem(STORAGE_KEYS.URL_MONITOR_TAREO) || '';
    } catch {
      return '';
    }
  });

  const [embedMode, setEmbedMode] = useState<TareoEmbedMode>(() => {
    try {
      const m = localStorage.getItem(STORAGE_KEYS.MODO_MONITOR_TAREO);
      if (m === 'embed' || m === 'normal' || m === 'tv') return m;
      return 'tv';
    } catch {
      return 'tv';
    }
  });

  const [urlInput, setUrlInput] = useState<string>(savedUrl);
  const [configAbierta, setConfigAbierta] = useState<boolean>(() => !savedUrl);
  const [reloadKey, setReloadKey] = useState<number>(0);
  const [expandido, setExpandido] = useState<boolean>(false);

  const iframeUrl = useMemo(() => construirUrlTareo(savedUrl, embedMode), [savedUrl, embedMode]);

  const mixedContentBlocked = useMemo(() => {
    if (!iframeUrl || typeof window === 'undefined') return false;
    try {
      const target = new URL(iframeUrl);
      const isLocal =
        target.hostname === 'localhost' ||
        target.hostname === '127.0.0.1';
      return window.location.protocol === 'https:' && target.protocol === 'http:' && !isLocal;
    } catch {
      return false;
    }
  }, [iframeUrl]);

  const handleGuardarUrl = (e: React.FormEvent) => {
    e.preventDefault();
    const normalizada = normalizarUrlBase(urlInput);
    if (!normalizada) {
      localStorage.removeItem(STORAGE_KEYS.URL_MONITOR_TAREO);
      setSavedUrl('');
      setUrlInput('');
      setConfigAbierta(true);
      showToast('ok', 'URL de Monitor Tareo borrada');
      return;
    }
    try {
      new URL(normalizada);
    } catch {
      showToast('error', 'Ingresa una dirección o IP válida (ej: 192.168.1.50:3000)');
      return;
    }
    localStorage.setItem(STORAGE_KEYS.URL_MONITOR_TAREO, normalizada);
    setSavedUrl(normalizada);
    setUrlInput(normalizada);
    setConfigAbierta(false);
    setReloadKey((k) => k + 1);
    showToast('ok', 'Monitor Tareo conectado');
  };

  const handleCambiarModo = (nuevoModo: TareoEmbedMode) => {
    setEmbedMode(nuevoModo);
    try {
      localStorage.setItem(STORAGE_KEYS.MODO_MONITOR_TAREO, nuevoModo);
    } catch {}
    setReloadKey((k) => k + 1);
  };

  return (
    <div className="space-y-3">
      {/* Compact Header & Controls Bar */}
      <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl p-2.5 sm:p-3 shadow-sm space-y-2.5">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-2 min-w-0">
            <div className="w-7 h-7 rounded-lg bg-[#F2A93B]/15 border border-[#F2A93B]/40 flex items-center justify-center text-[#F2A93B] shrink-0">
              <Monitor size={15} />
            </div>
            <div className="min-w-0">
              <h2 className="font-['Oswald',sans-serif] text-sm sm:text-base font-bold uppercase tracking-wide text-white leading-tight">
                Monitor Tareo en Vivo
              </h2>
              {iframeUrl && (
                <p className="text-[10.5px] font-['IBM_Plex_Mono',monospace] text-[#A0AEC0] truncate max-w-[240px] sm:max-w-[420px]">
                  {iframeUrl}
                </p>
              )}
            </div>
          </div>

          <div className="flex items-center gap-1.5 flex-wrap">
            {/* Mode Switch */}
            <div className="inline-flex p-0.5 bg-[#232A38] rounded-lg border border-[#4A5568]">
              <button
                type="button"
                onClick={() => handleCambiarModo('tv')}
                className={`flex items-center gap-1 py-1 px-2 rounded-md text-[11px] font-semibold cursor-pointer transition-colors ${
                  embedMode === 'tv' ? 'bg-[#F2A93B] text-[#12161F] font-bold' : 'text-[#CBD5E0] hover:text-white'
                }`}
                title="Solo tarjetas sin cabecera (?tv=1)"
              >
                <Tv size={12} />
                <span>TV</span>
              </button>
              <button
                type="button"
                onClick={() => handleCambiarModo('embed')}
                className={`flex items-center gap-1 py-1 px-2 rounded-md text-[11px] font-semibold cursor-pointer transition-colors ${
                  embedMode === 'embed' ? 'bg-[#F2A93B] text-[#12161F] font-bold' : 'text-[#CBD5E0] hover:text-white'
                }`}
                title="Monitor en vivo embebido (?embed=live)"
              >
                <LayoutTemplate size={12} />
                <span>Live</span>
              </button>
              <button
                type="button"
                onClick={() => handleCambiarModo('normal')}
                className={`flex items-center gap-1 py-1 px-2 rounded-md text-[11px] font-semibold cursor-pointer transition-colors ${
                  embedMode === 'normal' ? 'bg-[#F2A93B] text-[#12161F] font-bold' : 'text-[#CBD5E0] hover:text-white'
                }`}
                title="Vista completa de Monitor Tareo"
              >
                <Globe size={12} />
                <span>Completo</span>
              </button>
            </div>

            {iframeUrl && (
              <>
                <button
                  type="button"
                  onClick={() => setReloadKey((k) => k + 1)}
                  className="p-1.5 rounded-lg bg-[#232A38] border border-[#4A5568] text-[#CBD5E0] hover:text-white cursor-pointer"
                  title="Recargar vista"
                >
                  <RefreshCw size={14} />
                </button>
                <button
                  type="button"
                  onClick={() => setExpandido((prev) => !prev)}
                  className="hidden sm:inline-flex p-1.5 rounded-lg bg-[#232A38] border border-[#4A5568] text-[#CBD5E0] hover:text-white cursor-pointer"
                  title={expandido ? 'Altura estándar' : 'Expandir alto'}
                >
                  {expandido ? <Minimize2 size={14} /> : <Maximize2 size={14} />}
                </button>
                <a
                  href={iframeUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="flex items-center gap-1 py-1 px-2 rounded-lg bg-[#232A38] border border-[#4A5568] text-[#93C5FD] hover:text-white text-xs font-semibold"
                  title="Abrir Monitor Tareo en pestaña nueva"
                >
                  <ExternalLink size={12} />
                  <span className="hidden sm:inline">Abrir</span>
                </a>
              </>
            )}

            <button
              type="button"
              onClick={() => setConfigAbierta((prev) => !prev)}
              className={`flex items-center gap-1 py-1 px-2 rounded-lg border text-xs font-semibold cursor-pointer ${
                configAbierta
                  ? 'bg-[#2D3748] border-[#F2A93B]/60 text-[#F6AD55]'
                  : 'bg-[#232A38] border-[#4A5568] text-[#CBD5E0] hover:text-white'
              }`}
              title="Configurar URL / IP de Monitor Tareo"
            >
              <Settings size={13} />
              <span className="hidden sm:inline">Servidor</span>
            </button>
          </div>
        </div>

        {/* Collapsible URL Config Form */}
        {configAbierta && (
          <form onSubmit={handleGuardarUrl} className="pt-2.5 border-t border-[#2D3748] space-y-2">
            <label className="block text-[11px] font-semibold text-[#CBD5E0] uppercase tracking-wide">
              Dirección del servidor Monitor Tareo (IP local o URL HTTPS)
            </label>
            <div className="flex flex-col sm:flex-row gap-2">
              <input
                type="text"
                value={urlInput}
                onChange={(e) => setUrlInput(e.target.value)}
                placeholder="Ej: http://192.168.1.50:3000 o https://tu-monitor-tareo.run.app"
                className="flex-1 py-2 px-3 rounded-lg border border-[#4A5568] bg-[#12161F] text-white text-xs font-['IBM_Plex_Mono',monospace] outline-none focus:border-[#F2A93B]"
              />
              <button
                type="submit"
                className="flex items-center justify-center gap-1.5 bg-[#F2A93B] text-[#12161F] py-2 px-4 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif] tracking-wider cursor-pointer hover:brightness-105 shrink-0"
              >
                <Check size={14} />
                <span>Guardar y Conectar</span>
              </button>
            </div>
            <p className="text-[11px] text-[#A0AEC0] leading-snug">
              Modos disponibles: <strong className="text-[#CBD5E0]">TV</strong> agrega{' '}
              <code className="text-[#F6AD55]">?tv=1</code>, <strong className="text-[#CBD5E0]">Live</strong> agrega{' '}
              <code className="text-[#F6AD55]">?embed=live</code>, y <strong className="text-[#CBD5E0]">Completo</strong>{' '}
              carga la dirección sin parámetros extra.
            </p>
          </form>
        )}
      </div>

      {/* Mixed Content Warning when app is HTTPS and target is HTTP LAN IP */}
      {mixedContentBlocked && (
        <div className="bg-[#3A2E1B]/60 border border-[#F2A93B]/60 rounded-xl p-3 text-xs text-[#FBD38D] flex flex-col sm:flex-row sm:items-center justify-between gap-2">
          <div className="flex items-start gap-2">
            <AlertTriangle size={16} className="text-[#F2A93B] shrink-0 mt-0.5" />
            <div>
              <strong className="block text-white">Bloqueo HTTPS → HTTP del navegador</strong>
              <span>
                Estás abriendo Supervisión Planta por <code>https://</code> e intentando incrustar una IP local{' '}
                <code>http://</code>. Para verlo dentro del recuadro usa una URL <code>https://</code> en Monitor Tareo o abre ambas apps por red local <code>http://</code>.
              </span>
            </div>
          </div>
          <a
            href={iframeUrl}
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center justify-center gap-1.5 bg-[#F2A93B] text-[#12161F] py-1.5 px-3 rounded-lg text-xs font-bold uppercase font-['Oswald',sans-serif] shrink-0"
          >
            <ExternalLink size={12} />
            <span>Abrir aparte</span>
          </a>
        </div>
      )}

      {/* Embedded Frame or Empty Prompt */}
      {iframeUrl ? (
        <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl overflow-hidden shadow-lg">
          <iframe
            key={`${iframeUrl}-${reloadKey}`}
            src={iframeUrl}
            title="Monitor Tareo en Vivo"
            className={`w-full border-0 bg-[#12161F] ${
              expandido ? 'h-[calc(100vh-140px)] min-h-[640px]' : 'h-[calc(100vh-210px)] min-h-[520px]'
            }`}
            allow="clipboard-read; clipboard-write"
          />
        </div>
      ) : (
        <div className="border border-dashed border-[#4A5568] rounded-xl p-8 text-center bg-[#1A202C] space-y-2">
          <Monitor size={32} className="mx-auto text-[#F2A93B]/80" />
          <h3 className="font-['Oswald',sans-serif] text-base font-bold text-white uppercase tracking-wide">
            Conecta tu Monitor Tareo
          </h3>
          <p className="text-xs text-[#A0AEC0] max-w-md mx-auto leading-relaxed">
            Ingresa arriba la dirección del servidor de Monitor Tareo (por ejemplo{' '}
            <code className="text-[#F6AD55]">http://IP_DEL_SERVIDOR:3000</code> o la URL compartida) para ver el
            monitor en vivo directamente dentro de Supervisión de Planta.
          </p>
        </div>
      )}
    </div>
  );
};
