import React from 'react';
import { CommunityMetric } from '../types';

interface CommunityCardItemProps {
  card: CommunityMetric;
  onEdit?: () => void;
  onDelete?: () => void;
  isEditMode?: boolean;
  onMemberChange?: (newVal: number) => void;
}

export const CommunityCardItem: React.FC<CommunityCardItemProps> = ({
  card,
  onEdit,
  onDelete,
  isEditMode = false,
  onMemberChange,
}) => {
  const isConnect =
    (card.roleBadge && card.roleBadge.toLowerCase().includes('connect') && !card.roleBadge.toLowerCase().includes('upconnect')) ||
    card.color === 'purple' ||
    card.color === 'amber';

  const isGreen =
    card.color === 'emerald' ||
    (card.roleBadge && card.roleBadge.toLowerCase().includes('capacitaci'));

  // Color schemes for badge & number
  let badgeClasses = 'bg-[#04244e] text-[#38bdf8] border border-[#0284c7]/40';
  let numberColor = 'text-[#00a6ff]';
  let borderClasses = 'border-[#0d2a58]';

  if (isConnect) {
    badgeClasses = 'bg-[#2a1708] text-[#fbbf24] border border-[#d97706]/40';
    numberColor = 'text-[#fbbf24]';
    borderClasses = 'border-[#3f2710]';
  } else if (isGreen) {
    badgeClasses = 'bg-[#06241a] text-[#34d399] border border-[#059669]/40';
    numberColor = 'text-[#34d399]';
    borderClasses = 'border-[#0b3829]';
  }

  return (
    <div
      className={`relative rounded-2xl bg-[#071329] border ${borderClasses} p-5 shadow-xl shadow-black/40 flex flex-col justify-between transition-all hover:border-slate-600 group`}
    >
      <div>
        {/* Top Header: Title and Role Badge */}
        <div className="flex items-center justify-between gap-3">
          <h3 className="text-base font-bold text-white tracking-tight truncate" title={card.name}>
            {card.name}
          </h3>
          <span className={`text-xs font-semibold px-2.5 py-1 rounded-md shrink-0 ${badgeClasses}`}>
            {card.roleBadge || (isConnect ? 'Distribuidor Connect' : 'Distribuidor Upconnect')}
          </span>
        </div>

        {/* Big Metric Number */}
        {isEditMode ? (
          <div className="my-2">
            <input
              type="number"
              min="0"
              value={card.members}
              onChange={(e) => onMemberChange && onMemberChange(parseInt(e.target.value, 10) || 0)}
              className={`w-full text-3xl font-black ${numberColor} bg-slate-950 border border-slate-700 rounded-lg px-2 py-0.5 tabular-nums focus:outline-none focus:border-sky-500`}
            />
          </div>
        ) : (
          <div className={`text-4xl font-extrabold ${numberColor} tabular-nums my-2.5 tracking-tight`}>
            {card.members}
          </div>
        )}

        {/* Subtitle */}
        <div className="text-xs text-slate-400 font-medium">
          {card.subtitle || 'Socios Registrados'}
        </div>
      </div>

      {/* Divider & Description */}
      <div>
        <div className="border-t border-slate-800/80 pt-3 mt-3" />
        <p className="text-xs text-slate-300 leading-relaxed font-normal">
          {card.description || 'Red de socios y aliados estratégicos.'}
        </p>
      </div>

      {/* Action buttons (Modificar / Eliminar) */}
      {(onEdit || onDelete) && (
        <div className="absolute top-2 right-2 opacity-0 group-hover:opacity-100 transition-opacity flex items-center gap-1 bg-[#071329]/95 p-1 rounded-lg border border-slate-800">
          {onEdit && (
            <button
              onClick={(e) => {
                e.stopPropagation();
                onEdit();
              }}
              className="p-1 rounded text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
              title="Modificar recuadro"
            >
              <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
              </svg>
            </button>
          )}
          {onDelete && (
            <button
              onClick={(e) => {
                e.stopPropagation();
                onDelete();
              }}
              className="p-1 rounded text-rose-400 hover:text-rose-300 hover:bg-rose-950/40 transition-colors"
              title="Eliminar recuadro"
            >
              <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M3 6h18" />
                <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6" />
                <path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
              </svg>
            </button>
          )}
        </div>
      )}
    </div>
  );
};
