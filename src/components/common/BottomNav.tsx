import React from 'react';
import { useApp } from '../../context/AppContext';
import { Box, Cpu, TrendingUp, FileText, History, LucideIcon } from 'lucide-react';
import { TabType } from '../../types';

interface NavItem {
  id: TabType;
  label: string;
  icon: LucideIcon;
  badge?: number;
}

export const BottomNav: React.FC = () => {
  const { activeTab, setActiveTab, lista, maquinas } = useApp();

  const activasCount = maquinas.filter((m) => m.activa).length;

  const navItems: NavItem[] = [
    { id: 'inventario', label: 'Inventario', icon: Box, badge: lista.length > 0 ? lista.length : undefined },
    { id: 'maquinas', label: 'Máquinas', icon: Cpu, badge: activasCount > 0 ? activasCount : undefined },
    { id: 'produccion', label: 'Producción', icon: TrendingUp },
    { id: 'resumen', label: 'Resumen', icon: FileText },
    { id: 'historialdb', label: 'Historial', icon: History },
  ];

  return (
    <nav
      aria-label="Navegación móvil"
      className="md:hidden fixed bottom-0 left-0 right-0 z-40 bg-[#1A202C]/95 backdrop-blur-md border-t border-[#2D3748] px-1 pb-[env(safe-area-inset-bottom,4px)] pt-0.5 shadow-2xl"
    >
      <div className="grid grid-cols-5 h-[50px] items-stretch">
        {navItems.map((item) => {
          const isActive = activeTab === item.id;
          const IconComponent = item.icon;

          return (
            <button
              key={item.id}
              type="button"
              onClick={() => setActiveTab(item.id)}
              className={`relative flex flex-col items-center justify-center gap-0.5 min-h-[44px] rounded-lg transition-all active:scale-95 cursor-pointer ${
                isActive
                  ? 'text-[#F2A93B]'
                  : 'text-[#A0AEC0] hover:text-[#CBD5E0] active:text-white'
              }`}
            >
              {/* Active top indicator bar */}
              {isActive && (
                <span className="absolute top-0 w-6 h-[2px] rounded-full bg-[#F2A93B]" />
              )}

              <div className="relative flex items-center justify-center">
                <IconComponent size={18} className={isActive ? 'stroke-[2.4]' : 'stroke-[1.8]'} />
                {typeof item.badge === 'number' && item.badge > 0 && (
                  <span className="absolute -top-1.5 -right-2 min-w-[14px] h-3.5 px-0.5 rounded-full bg-[#F2A93B] text-[#12161F] text-[9px] font-bold flex items-center justify-center leading-none">
                    {item.badge > 99 ? '99+' : item.badge}
                  </span>
                )}
              </div>

              <span
                className={`text-[9.5px] tracking-tight leading-none ${
                  isActive ? 'font-bold text-[#F6AD55]' : 'font-normal text-[#A0AEC0]'
                }`}
              >
                {item.label}
              </span>
            </button>
          );
        })}
      </div>
    </nav>
  );
};
