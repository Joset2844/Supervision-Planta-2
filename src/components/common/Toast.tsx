import React from 'react';
import { useApp } from '../../context/AppContext';

export const Toast: React.FC = () => {
  const { toasts, removeToast } = useApp();

  if (toasts.length === 0) return null;

  return (
    <div className="fixed top-4 left-1/2 -translate-x-1/2 z-50 flex flex-col gap-2 max-w-md w-full px-4 pointer-events-none">
      {toasts.map((t) => (
        <div
          key={t.id}
          onClick={() => removeToast(t.id)}
          className={`pointer-events-auto cursor-pointer shadow-lg py-2.5 px-4 rounded-lg text-sm font-medium flex items-center justify-between gap-2 text-white transition-all transform animate-in fade-in duration-200 ${
            t.type === 'ok' ? 'bg-[#38A169]' : 'bg-[#E53E3E]'
          }`}
        >
          <span>{t.text}</span>
          <button type="button" className="text-white/80 hover:text-white ml-2 text-xs">
            ✕
          </button>
        </div>
      ))}
    </div>
  );
};
