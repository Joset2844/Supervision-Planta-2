import React from 'react';

interface OdometerProps {
  label: string;
  value: number;
}

export const Odometer: React.FC<OdometerProps> = ({ label, value }) => {
  const digits = String(value || 0)
    .padStart(6, '0')
    .split('');

  return (
    <div className="bg-[#1A202C] border border-[#2D3748] rounded-xl py-2 px-3 sm:px-4 flex items-center justify-between gap-2 shadow-sm mb-2.5">
      <span className="text-[11px] sm:text-xs text-[#CBD5E0] font-['Oswald',sans-serif] tracking-wider uppercase font-semibold">
        {label}
      </span>
      <div className="flex gap-[2px] bg-[#0A0D14] rounded-md py-1 px-2 border border-[#2D3748] shadow-inner shrink-0">
        {digits.map((d, i) => (
          <span
            key={i}
            className="font-['IBM_Plex_Mono',monospace] w-[1.1ch] text-center text-base sm:text-lg font-bold text-[#F2A93B] tabular-nums"
          >
            {d}
          </span>
        ))}
      </div>
    </div>
  );
};
