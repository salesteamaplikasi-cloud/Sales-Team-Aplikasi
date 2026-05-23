import React, { useState, useEffect } from "react";
import { X } from "lucide-react";

interface RewardModalProps {
  isOpen: boolean;
  type: 'merchant' | 'catalog';
  item?: any;
  onClose: () => void;
  onSave: (type: 'merchant' | 'catalog', item: any) => void;
  onDelete?: (type: 'merchant' | 'catalog', id: string) => void;
}

export const RewardModal: React.FC<RewardModalProps> = ({ isOpen, type, item, onClose, onSave, onDelete }) => {
  const [formData, setFormData] = useState<any>({ name: "", points: 0, sponsorOrDesc: "" });

  useEffect(() => {
    if (isOpen) {
      if (item) {
        setFormData({
          name: item.name,
          points: 'pointsRequired' in item ? item.pointsRequired : item.pointsValue,
          sponsorOrDesc: 'sponsor' in item ? item.sponsor : item.description,
        });
      } else {
        setFormData({ name: "", points: 0, sponsorOrDesc: "" });
      }
    }
  }, [isOpen, item]);

  if (!isOpen) return null;

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const updatedItem = {
      ...item,
      name: formData.name,
      [type === 'merchant' ? 'pointsRequired' : 'pointsValue']: Number(formData.points),
      [type === 'merchant' ? 'description' : 'sponsor']: formData.sponsorOrDesc,
    };
    onSave(type, updatedItem);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4">
      <div className="bg-white rounded-3xl p-6 w-full max-w-sm shadow-xl">
        <div className="flex justify-between items-center mb-4">
          <h2 className="text-lg font-black text-[#4A4A3C]">
            {item ? "Edit" : "Tambah"} {type === 'merchant' ? 'Reward Merchant' : 'Katalog Hadiah'}
          </h2>
          <button onClick={onClose}><X className="w-5 h-5" /></button>
        </div>
        <form onSubmit={handleSubmit} className="space-y-4">
          <input className="w-full border rounded-xl p-2" placeholder="Nama" value={formData.name} onChange={(e) => setFormData({...formData, name: e.target.value})} required />
          <input type="number" className="w-full border rounded-xl p-2" placeholder="Poin" value={formData.points} onChange={(e) => setFormData({...formData, points: e.target.value})} required />
          <input className="w-full border rounded-xl p-2" placeholder={type === 'merchant' ? 'Deskripsi' : 'Sponsor'} value={formData.sponsorOrDesc} onChange={(e) => setFormData({...formData, sponsorOrDesc: e.target.value})} required />
          
          <div className="flex gap-2">
            <button type="submit" className="flex-1 bg-emerald-700 text-white rounded-xl py-2 font-bold">Simpan</button>
            {item && onDelete && (
              <button type="button" onClick={() => onDelete(type, item.id)} className="flex-1 bg-rose-700 text-white rounded-xl py-2 font-bold">Hapus</button>
            )}
          </div>
        </form>
      </div>
    </div>
  );
};
