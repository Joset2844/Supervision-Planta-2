import React from 'react';
import { useApp } from './context/AppContext';
import { Header } from './components/common/Header';
import { BottomNav } from './components/common/BottomNav';
import { Toast } from './components/common/Toast';
import { InventarioView } from './components/inventario/InventarioView';
import { MaquinasView } from './components/maquinas/MaquinasView';
import { ProduccionView } from './components/produccion/ProduccionView';
import { ResumenView } from './components/resumen/ResumenView';
import { HistorialDBView } from './components/historial/HistorialDBView';

export const AppContent: React.FC = () => {
  const { activeTab } = useApp();

  return (
    <div className="min-h-screen bg-[#12161F] text-[#F7FAFC] flex flex-col antialiased selection:bg-[#F2A93B] selection:text-[#12161F]">
      <Header />
      <Toast />

      {/* Main container: pb-48 on mobile ensures any card (even a single active machine) can be scrolled completely above bottom action bars without getting covered! */}
      <main className="flex-1 w-full max-w-[1440px] mx-auto px-3 sm:px-4 lg:px-6 py-2.5 sm:py-3.5 pb-48 md:pb-16">
        {activeTab === 'inventario' && <InventarioView />}
        {activeTab === 'maquinas' && <MaquinasView />}
        {activeTab === 'produccion' && <ProduccionView />}
        {activeTab === 'resumen' && <ResumenView />}
        {activeTab === 'historialdb' && <HistorialDBView />}
      </main>

      {/* Mobile fixed bottom navigation bar (hidden on desktop md:) */}
      <BottomNav />
    </div>
  );
};

export default function App() {
  return <AppContent />;
}
