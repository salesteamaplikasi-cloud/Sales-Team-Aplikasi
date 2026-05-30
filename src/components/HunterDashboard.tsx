import React from 'react';
import { NooRecord } from '../types';
import { Target, Crosshair } from 'lucide-react';

interface HunterDashboardProps {
  nooRecords: NooRecord[];
}

export const HunterDashboard: React.FC<HunterDashboardProps> = ({ nooRecords }) => {
  const imamNoo = nooRecords.filter(n => n.salesmanName.toLowerCase().trim() === 'imam');
  
  const totalWarung = imamNoo.reduce((sum, n) => sum + n.warung, 0);
  const totalStore = imamNoo.reduce((sum, n) => sum + n.store, 0);
  const totalKiosk = imamNoo.reduce((sum, n) => sum + n.kiosk, 0);
  const totalWholesaler = imamNoo.reduce((sum, n) => sum + n.wholesaler, 0);
  const totalNooCount = totalWarung + totalStore + totalKiosk + totalWholesaler;

  return (
    <div className="p-6 bg-gray-50 rounded-lg h-full">
      <h2 className="text-2xl font-bold text-gray-900 mb-6">Pencari Pelanggan & NOO (Imam - Hunter)</h2>
      
      <div className="bg-white p-6 rounded-lg shadow-sm border border-gray-200">
        <div className="flex items-center gap-2 mb-4">
          <Crosshair className="text-rose-600" />
          <h3 className="text-lg font-semibold text-gray-900">Peraihan Outlet Baru (NOO)</h3>
        </div>
        
        <div className="grid grid-cols-2 gap-4 text-center mt-6">
          <div className="bg-rose-50 p-4 rounded-xl border border-rose-100">
            <span className="block text-xs font-bold text-rose-600 uppercase">Warung</span>
            <span className="text-2xl font-black text-gray-900">{totalWarung}</span>
          </div>
          <div className="bg-[#e2e8f0]/40 p-4 rounded-xl border border-[#e2e8f0]">
            <span className="block text-xs font-bold text-[#64748b] uppercase">Toko Modern</span>
            <span className="text-2xl font-black text-gray-900">{totalStore}</span>
          </div>
          <div className="bg-rose-50 p-4 rounded-xl border border-rose-100">
            <span className="block text-xs font-bold text-rose-600 uppercase">Kios</span>
            <span className="text-2xl font-black text-gray-900">{totalKiosk}</span>
          </div>
          <div className="bg-[#e2e8f0]/40 p-4 rounded-xl border border-[#e2e8f0]">
            <span className="block text-xs font-bold text-[#64748b] uppercase">Grosir</span>
            <span className="text-2xl font-black text-gray-900">{totalWholesaler}</span>
          </div>
        </div>

        <div className="mt-6 pt-6 border-t border-gray-100">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <Target className="text-gray-400 w-5 h-5" />
              <span className="text-sm font-bold text-gray-600 uppercase">Total NOO Tercapai</span>
            </div>
            <span className="text-3xl font-black text-rose-600">{totalNooCount}</span>
          </div>
        </div>
      </div>
    </div>
  );
};
