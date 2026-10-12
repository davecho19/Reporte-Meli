import React, { useState, useEffect } from 'react';
import { CommunityMetric } from '../types';
import { X, Plus, Trash2, Check, Edit3 } from 'lucide-react';

interface CommunityCardsModalProps {
  isOpen: boolean;
  onClose: () => void;
  cards: CommunityMetric[];
  onSave: (cards: CommunityMetric[]) => void;
  mode: 'edit' | 'create';
  initialCardToEdit?: CommunityMetric | null;
}

export const CommunityCardsModal: React.FC<CommunityCardsModalProps> = ({
  isOpen,
  onClose,
  cards,
  onSave,
  mode,
  initialCardToEdit,
}) => {
  // State for Edit mode (all cards)
  const [editableCards, setEditableCards] = useState<CommunityMetric[]>([]);

  // State for Create mode (single new card)
  const [newCard, setNewCard] = useState<CommunityMetric>({
    id: '',
    name: 'UpConta Socios',
    roleBadge: 'Distribuidor Upconnect',
    members: 118,
    subtitle: 'Socios Registrados',
    description: 'Red principal de socios estratégicos y contadores vinculados a UpConnect.',
    color: 'sky',
  });

  useEffect(() => {
    if (isOpen) {
      setEditableCards(JSON.parse(JSON.stringify(cards)));
      if (mode === 'create') {
        setNewCard({
          id: `card_${Date.now()}`,
          name: '',
          roleBadge: 'Distribuidor Upconnect',
          members: 0,
          subtitle: 'Socios Registrados',
          description: '',
          color: 'sky',
        });
      }
    }
  }, [isOpen, cards, mode]);

  if (!isOpen) return null;

  const handleCardChange = (index: number, field: keyof CommunityMetric, value: any) => {
    setEditableCards((prev) => {
      const copy = [...prev];
      copy[index] = { ...copy[index], [field]: value };
      return copy;
    });
  };

  const handleRemoveCard = (index: number) => {
    setEditableCards((prev) => prev.filter((_, i) => i !== index));
  };

  const handleSaveEdit = () => {
    onSave(editableCards);
    onClose();
  };

  const handleSaveCreate = () => {
    if (!newCard.name.trim()) {
      alert('Por favor ingrese el nombre del recuadro.');
      return;
    }
    const cardToAdd: CommunityMetric = {
      ...newCard,
      id: newCard.id || `card_${Date.now()}`,
      color:
        newCard.roleBadge?.toLowerCase().includes('connect') && !newCard.roleBadge?.toLowerCase().includes('upconnect')
          ? 'amber'
          : newCard.color || 'sky',
    };
    onSave([...cards, cardToAdd]);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm no-print animate-in fade-in duration-200">
      <div className="w-full max-w-2xl bg-[#0e1626] border border-sky-500/40 rounded-2xl p-6 shadow-2xl flex flex-col max-h-[90vh] overflow-hidden">
        {/* Modal Header */}
        <div className="flex items-center justify-between pb-3 border-b border-slate-800">
          <div className="flex items-center gap-2">
            {mode === 'create' ? (
              <>
                <div className="w-8 h-8 rounded-lg bg-emerald-500/20 text-emerald-400 flex items-center justify-center font-bold">
                  <Plus className="w-4 h-4" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-white">Crear Nuevo Recuadro</h3>
                  <p className="text-xs text-slate-400">
                    Agregue un nuevo recuadro de comunidad/socio en la diapositiva 7
                  </p>
                </div>
              </>
            ) : (
              <>
                <div className="w-8 h-8 rounded-lg bg-sky-500/20 text-sky-400 flex items-center justify-center font-bold">
                  <Edit3 className="w-4 h-4" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-white">Modificar Recuadros</h3>
                  <p className="text-xs text-slate-400">
                    Cambie nombres, cantidades, roles y descripciones de los recuadros
                  </p>
                </div>
              </>
            )}
          </div>
          <button
            onClick={onClose}
            className="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="flex-1 overflow-y-auto py-4 space-y-4 pr-1">
          {mode === 'create' ? (
            <div className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-slate-300 mb-1">
                  Nombre del Recuadro *
                </label>
                <input
                  type="text"
                  placeholder="Ej: UpConta Socios, Franquicia VIP..."
                  value={newCard.name}
                  onChange={(e) => setNewCard({ ...newCard, name: e.target.value })}
                  className="w-full px-3 py-2 text-sm bg-slate-900 border border-slate-700 rounded-xl text-white placeholder-slate-500 focus:outline-none focus:border-sky-500"
                />
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-semibold text-slate-300 mb-1">
                    Rol / Etiqueta
                  </label>
                  <select
                    value={newCard.roleBadge}
                    onChange={(e) => setNewCard({ ...newCard, roleBadge: e.target.value })}
                    className="w-full px-3 py-2 text-sm bg-slate-900 border border-slate-700 rounded-xl text-white focus:outline-none focus:border-sky-500"
                  >
                    <option value="Distribuidor Upconnect">Distribuidor Upconnect</option>
                    <option value="Distribuidor Connect">Distribuidor Connect</option>
                    <option value="Capacitación">Capacitación</option>
                    <option value="Socio Estratégico">Socio Estratégico</option>
                    <option value="Franquicia VIP">Franquicia VIP</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-semibold text-slate-300 mb-1">
                    Cantidad (Número) *
                  </label>
                  <input
                    type="number"
                    min="0"
                    value={newCard.members}
                    onChange={(e) =>
                      setNewCard({ ...newCard, members: parseInt(e.target.value, 10) || 0 })
                    }
                    className="w-full px-3 py-2 text-sm bg-slate-900 border border-slate-700 rounded-xl text-white focus:outline-none focus:border-sky-500"
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-300 mb-1">
                  Subtítulo Inferior del Número
                </label>
                <input
                  type="text"
                  placeholder="Ej: Socios Registrados"
                  value={newCard.subtitle}
                  onChange={(e) => setNewCard({ ...newCard, subtitle: e.target.value })}
                  className="w-full px-3 py-2 text-sm bg-slate-900 border border-slate-700 rounded-xl text-white placeholder-slate-500 focus:outline-none focus:border-sky-500"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-300 mb-1">
                  Descripción
                </label>
                <textarea
                  rows={3}
                  placeholder="Ej: Red principal de socios estratégicos y contadores vinculados a UpConnect."
                  value={newCard.description}
                  onChange={(e) => setNewCard({ ...newCard, description: e.target.value })}
                  className="w-full px-3 py-2 text-sm bg-slate-900 border border-slate-700 rounded-xl text-white placeholder-slate-500 focus:outline-none focus:border-sky-500"
                />
              </div>

              {/* Vista previa en vivo del recuadro */}
              <div className="pt-2">
                <span className="text-[11px] font-semibold text-slate-400 block mb-2">
                  Vista Previa del Recuadro:
                </span>
                <div className="p-4 rounded-xl bg-[#071329] border border-[#0d2a58] shadow-lg">
                  <div className="flex items-center justify-between">
                    <h4 className="text-sm font-bold text-white">{newCard.name || 'Título'}</h4>
                    <span className="text-[10px] px-2 py-0.5 rounded bg-[#04244e] text-[#38bdf8] border border-[#0284c7]/40 font-semibold">
                      {newCard.roleBadge || 'Distribuidor Upconnect'}
                    </span>
                  </div>
                  <div className="text-3xl font-black text-[#00a6ff] tabular-nums my-1.5">
                    {newCard.members}
                  </div>
                  <div className="text-[11px] text-slate-400">{newCard.subtitle || 'Socios Registrados'}</div>
                  <div className="border-t border-slate-800 pt-2 mt-2 text-xs text-slate-300">
                    {newCard.description || 'Descripción del recuadro...'}
                  </div>
                </div>
              </div>
            </div>
          ) : (
            <div className="space-y-4">
              {editableCards.map((c, idx) => (
                <div
                  key={c.id || idx}
                  className="p-4 rounded-xl bg-slate-900/80 border border-slate-800 space-y-3 relative group"
                >
                  <div className="flex items-center justify-between pb-2 border-b border-slate-800/80">
                    <span className="text-xs font-bold text-sky-400">
                      Recuadro #{idx + 1}
                    </span>
                    {editableCards.length > 1 && (
                      <button
                        type="button"
                        onClick={() => handleRemoveCard(idx)}
                        className="text-xs text-rose-400 hover:text-rose-300 flex items-center gap-1 hover:bg-rose-950/30 px-2 py-1 rounded"
                        title="Eliminar este recuadro"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                        <span>Eliminar</span>
                      </button>
                    )}
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div>
                      <label className="block text-[11px] font-semibold text-slate-300 mb-1">
                        Nombre
                      </label>
                      <input
                        type="text"
                        value={c.name}
                        onChange={(e) => handleCardChange(idx, 'name', e.target.value)}
                        className="w-full px-2.5 py-1.5 text-xs bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none focus:border-sky-500 font-semibold"
                      />
                    </div>

                    <div>
                      <label className="block text-[11px] font-semibold text-slate-300 mb-1">
                        Cantidad
                      </label>
                      <input
                        type="number"
                        min="0"
                        value={c.members}
                        onChange={(e) =>
                          handleCardChange(idx, 'members', parseInt(e.target.value, 10) || 0)
                        }
                        className="w-full px-2.5 py-1.5 text-xs bg-slate-950 border border-slate-700 rounded-lg text-sky-400 focus:outline-none focus:border-sky-500 font-bold tabular-nums"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div>
                      <label className="block text-[11px] font-semibold text-slate-300 mb-1">
                        Rol / Etiqueta
                      </label>
                      <input
                        type="text"
                        value={c.roleBadge || ''}
                        placeholder="Distribuidor Upconnect / Connect"
                        onChange={(e) => handleCardChange(idx, 'roleBadge', e.target.value)}
                        className="w-full px-2.5 py-1.5 text-xs bg-slate-950 border border-slate-700 rounded-lg text-white focus:outline-none focus:border-sky-500"
                      />
                    </div>

                    <div>
                      <label className="block text-[11px] font-semibold text-slate-300 mb-1">
                        Subtítulo
                      </label>
                      <input
                        type="text"
                        value={c.subtitle || ''}
                        placeholder="Socios Registrados"
                        onChange={(e) => handleCardChange(idx, 'subtitle', e.target.value)}
                        className="w-full px-2.5 py-1.5 text-xs bg-slate-950 border border-slate-700 rounded-lg text-slate-300 focus:outline-none focus:border-sky-500"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-[11px] font-semibold text-slate-300 mb-1">
                      Descripción
                    </label>
                    <textarea
                      rows={2}
                      value={c.description || ''}
                      onChange={(e) => handleCardChange(idx, 'description', e.target.value)}
                      className="w-full px-2.5 py-1.5 text-xs bg-slate-950 border border-slate-700 rounded-lg text-slate-300 focus:outline-none focus:border-sky-500"
                    />
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-800">
          <button
            type="button"
            onClick={onClose}
            className="px-3 py-1.5 text-xs rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors"
          >
            Cancelar
          </button>
          <button
            type="button"
            onClick={mode === 'create' ? handleSaveCreate : handleSaveEdit}
            className="px-4 py-1.5 text-xs rounded-xl bg-sky-600 hover:bg-sky-500 text-white font-bold transition-colors shadow-lg shadow-sky-600/30 flex items-center gap-1.5"
          >
            <Check className="w-3.5 h-3.5" />
            <span>{mode === 'create' ? 'Crear Recuadro' : 'Guardar Cambios'}</span>
          </button>
        </div>
      </div>
    </div>
  );
};
