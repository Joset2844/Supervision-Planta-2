import React from 'react';
import { useApp } from '../../context/AppContext';
import { Box, Cpu, TrendingUp, FileText, History, Plus, Search, Sun, Moon, LucideIcon } from 'lucide-react';
import { TabType } from '../../types';

export const Header: React.FC = () => {
  const {
    activeTab,
    setActiveTab,
    subTabInventario,
    setSubTabInventario,
    catalogo,
    lista,
    maquinas,
    metaTurno,
  } = useApp();

  const activasCount = maquinas.filter((m) => m.activa).length;

  const desktopTabs: { id: TabType; label: string; icon: LucideIcon; badge?: number }[] = [
    { id: 'inventario', label: 'Inventario', icon: Box, badge: lista.length > 0 ? lista.length : undefined },
    { id: 'maquinas', label: 'Máquinas', icon: Cpu, badge: activasCount > 0 ? activasCount : undefined },
    { id: 'produccion', label: 'Producción', icon: TrendingUp },
    { id: 'resumen', label: 'Resumen', icon: FileText },
    { id: 'historialdb', label: 'Historial', icon: History },
  ];

  return (
    <header className="sticky top-0 z-30 bg-[#1A202C]/95 backdrop-blur-md border-b-[3px] border-[#F2A93B] shadow-md">
      {/* Top Bar */}
      <div className="max-w-[1440px] mx-auto px-3.5 sm:px-5 lg:px-8 py-2 md:py-3 flex items-center justify-between gap-4">
        {/* Zone 1: Brand & Plant Indicator */}
        <div className="flex items-center gap-2.5 shrink-0">
          <div className="w-7 h-7 sm:w-8 sm:h-8 rounded-lg bg-[#F2A93B] flex items-center justify-center shrink-0 shadow-sm shadow-[#F2A93B]/20">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#12161F" strokeWidth="2.5">
              <path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z" />
              <path d="m3.3 7 8.7 5 8.7-5M12 22V12" />
            </svg>
          </div>
          <div>
            <span className="text-white font-['Oswald',sans-serif] tracking-wide text-sm sm:text-base md:text-lg font-bold leading-tight block">
              Supervisión de Planta
            </span>
            <div className="flex items-center gap-2 text-[11px] text-[#A0AEC0] leading-none mt-0.5">
              <span>{catalogo.length} OP en catálogo</span>
              <span aria-hidden="true" className="text-[#4A5568]">·</span>
              <span className="text-[#F6AD55] font-medium">{activasCount} maq. activas</span>
            </div>
          </div>
        </div>

        {/* Zone 2 (Desktop): Main Navigation Tabs */}
        <nav className="hidden md:flex items-center gap-1.5 lg:gap-2">
          {desktopTabs.map((tab) => {
            const isActive = activeTab === tab.id;
            const IconComp = tab.icon;
            return (
              <button
                key={tab.id}
                type="button"
                onClick={() => setActiveTab(tab.id)}
                className={`relative flex items-center gap-2 py-2 px-3.5 lg:px-4 rounded-xl text-xs lg:text-sm font-['Oswald',sans-serif] tracking-wider uppercase whitespace-nowrap transition-all cursor-pointer ${
                  isActive
                    ? 'bg-[#F2A93B] text-[#12161F] font-bold shadow-sm'
                    : 'bg-[#232A38] text-[#CBD5E0] hover:bg-[#2D3748] hover:text-white'
                }`}
              >
                <IconComp size={15} />
                <span>{tab.label}</span>
                {typeof tab.badge === 'number' && (
                  <span
                    className={`min-w-[18px] h-[18px] px-1 rounded-full text-[10px] font-bold flex items-center justify-center ${
                      isActive ? 'bg-[#12161F] text-[#F2A93B]' : 'bg-[#4A5568] text-white'
                    }`}
                  >
                    {tab.badge}
                  </span>
                )}
              </button>
            );
          })}
        </nav>

        {/* Zone 3: Current Shift Info (or quick actions) */}
        <div className="flex items-center gap-2 shrink-0">
          <div className="flex items-center gap-1.5 px-2.5 py-1 sm:py-1.5 bg-[#232A38] border border-[#2D3748] rounded-lg text-xs text-[#CBD5E0]">
            {metaTurno.tipo === 'noche' ? (
              <>
                <Moon size={13} className="text-[#63B3ED]" />
                <span className="font-medium hidden xs:inline">Turno Noche</span>
              </>
            ) : (
              <>
                <Sun size={13} className="text-[#F2A93B]" />
                <span className="font-medium hidden xs:inline">Turno Día</span>
              </>
            )}
          </div>
        </div>
      </div>

      {/* Subtabs for Inventario (Visible on both Mobile and Desktop when Inventario is active) */}
      {activeTab === 'inventario' && (
        <div className="max-w-[1440px] mx-auto px-3.5 sm:px-5 lg:px-8 pb-2 pt-0 flex items-center gap-2">
          <div className="inline-flex p-1 bg-[#12161F] rounded-xl border border-[#2D3748]">
            <button
              type="button"
              onClick={() => setSubTabInventario('agregar')}
              className={`flex items-center gap-1.5 py-1.5 px-3.5 rounded-lg text-xs font-['Oswald',sans-serif] tracking-wider uppercase transition-all cursor-pointer min-h-[36px] ${
                subTabInventario === 'agregar'
                  ? 'bg-[#F2A93B] text-[#12161F] font-bold shadow-sm'
                  : 'text-[#A0AEC0] hover:text-white'
              }`}
            >
              <Plus size={14} />
              <span>Registrar / Agregar</span>
              {lista.length > 0 && (
                <span className="ml-1 px-1.5 py-0.2 rounded-full bg-[#12161F] text-[#F2A93B] text-[10px] font-bold">
                  {lista.length}
                </span>
              )}
            </button>
            <button
              type="button"
              onClick={() => setSubTabInventario('consultar')}
              className={`flex items-center gap-1.5 py-1.5 px-3.5 rounded-lg text-xs font-['Oswald',sans-serif] tracking-wider uppercase transition-all cursor-pointer min-h-[36px] ${
                subTabInventario === 'consultar'
                  ? 'bg-[#F2A93B] text-[#12161F] font-bold shadow-sm'
                  : 'text-[#A0AEC0] hover:text-white'
              }`}
            >
              <Search size={14} />
              <span>Consultar Catálogo</span>
            </button>
          </div>
        </div>
      )}
    </header>
  );
};
