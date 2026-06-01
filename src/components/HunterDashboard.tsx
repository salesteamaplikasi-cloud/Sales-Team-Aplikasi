import React from 'react';
import { NooRecord, KpiReport } from '../types';
import { Target, Crosshair, TrendingUp, CheckCircle, AlertOctagon } from 'lucide-react';

interface HunterDashboardProps {
  nooRecords: NooRecord[];
  reports: KpiReport[];
}

export const HunterDashboard: React.FC<HunterDashboardProps> = ({ nooRecords, reports }) => {
  const imamNoo = nooRecords.filter(n => n.salesmanName.toLowerCase().trim() === 'imam');
  const imamReports = reports.filter(r => r.salesmanName && r.salesmanName.toLowerCase().trim() === 'imam');
  
  const totalWarung = imamNoo.reduce((sum, n) => sum + n.warung, 0);
  const totalStore = imamNoo.reduce((sum, n) => sum + n.store, 0);
  const totalKiosk = imamNoo.reduce((sum, n) => sum + n.kiosk, 0);
  const totalWholesaler = imamNoo.reduce((sum, n) => sum + n.wholesaler, 0);
  const totalNooCount = totalWarung + totalStore + totalKiosk + totalWholesaler;

  // KPI Calculations
  const cntDays = new Set(imamReports.map(r => r.date)).size || 1;
  const tcSum = imamReports.reduce((sum, r) => sum + r.tc, 0);
  const cpSum = imamReports.reduce((sum, r) => sum + r.cp, 0);
  const ecSum = imamReports.reduce((sum, r) => sum + r.ec, 0);
  const skuSum = imamReports.reduce((sum, r) => sum + r.skuTotal, 0);

  const cpPct = tcSum > 0 ? (cpSum / tcSum) * 100 : 0;
  const ecPct = cpSum > 0 ? (ecSum / cpSum) * 100 : 0;
  
  const targetSku = Math.round(ecSum * 15);
  const skuPct = targetSku > 0 ? (skuSum / targetSku) * 100 : 0;
  
  const isLayak = cpPct >= 80 && ecPct >= 40;

  return (
    <div className="p-6 bg-gray-50 rounded-lg h-full flex flex-col gap-6">
      <h2 className="text-2xl font-bold text-gray-900">Pencari Pelanggan & NOO (Imam - Hunter)</h2>
      
      {/* NOO CARD */}
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

      {/* KPI CARD */}
      <div className="bg-[#fcf8f2] border-t-4 border-rose-600 p-6 rounded-lg shadow-sm border border-[#e2e8f0]">
        <div className="flex flex-col gap-4">
          <div className="flex items-center justify-between border-b border-[#e2e8f0]/60 pb-3 mb-2">
            <h3 className="text-sm font-black text-[#0f172a] uppercase tracking-wider flex items-center gap-2">
              <TrendingUp className="w-5 h-5 text-rose-600" />
              Kinerja KPI Internal
            </h3>
            <span className="text-[10px] font-bold text-rose-700 bg-rose-100 px-2.5 py-1 rounded-md uppercase tracking-widest border border-rose-200">
              Total Hari Lapor: {imamReports.length > 0 ? cntDays : 0} Hari
            </span>
          </div>

          {/* Quick Metrics */}
          <div className="grid grid-cols-4 gap-2 mb-2">
            <div className="bg-white rounded-lg p-3 border border-[#e2e8f0]/60 text-center shadow-xs">
              <div className="text-[10px] font-extrabold text-[#64748b] uppercase tracking-wider mb-1">TC</div>
              <div className="text-lg font-black text-[#0f172a]">{tcSum}</div>
            </div>
            <div className="bg-white rounded-lg p-3 border border-[#e2e8f0]/60 text-center shadow-xs">
              <div className="text-[10px] font-extrabold text-[#64748b] uppercase tracking-wider mb-1">CP</div>
              <div className="text-lg font-black text-[#0f172a]">{cpSum}</div>
            </div>
            <div className="bg-white rounded-lg p-3 border border-[#e2e8f0]/60 text-center shadow-xs">
              <div className="text-[10px] font-extrabold text-[#64748b] uppercase tracking-wider mb-1">EC</div>
              <div className="text-lg font-black text-[#0f172a]">{ecSum}</div>
            </div>
            <div className="bg-white rounded-lg p-3 border border-[#e2e8f0]/60 text-center shadow-xs">
              <div className="text-[10px] font-extrabold text-[#64748b] uppercase tracking-wider mb-1">SKU</div>
              <div className="text-lg font-black text-[#2563eb]">{skuSum}</div>
            </div>
          </div>

          <div className="flex flex-col gap-3">
            {/* CP Progress */}
            <div>
              <div className="flex justify-between items-center mb-1">
                <span className="text-[10px] font-extrabold text-[#0f172a] uppercase tracking-wider">CP %</span>
                <span className="text-[10px] font-black">{cpPct.toFixed(1)}%</span>
              </div>
              <div className="w-full bg-[#e2e8f0]/60 rounded-full h-2">
                <div 
                  className={`h-2 rounded-full ${cpPct >= 80 ? 'bg-emerald-500' : 'bg-rose-500'}`}
                  style={{ width: `${Math.min(cpPct, 100)}%` }}
                />
              </div>
            </div>

            {/* EC Progress */}
            <div>
              <div className="flex justify-between items-center mb-1">
                <span className="text-[10px] font-extrabold text-[#0f172a] uppercase tracking-wider">EC %</span>
                <span className="text-[10px] font-black">{ecPct.toFixed(1)}%</span>
              </div>
              <div className="w-full bg-[#e2e8f0]/60 rounded-full h-2">
                <div 
                  className={`h-2 rounded-full ${ecPct >= 40 ? 'bg-emerald-500' : 'bg-rose-500'}`}
                  style={{ width: `${Math.min(ecPct, 100)}%` }}
                />
              </div>
            </div>
            
            {/* SKU Target */}
            <div>
              <div className="flex justify-between items-center mb-1">
                <span className="text-[10px] font-extrabold text-[#0f172a] uppercase tracking-wider flex gap-1">Target SKU: <span className="text-[#64748b]">{targetSku} SKU</span></span>
                <span className="text-[10px] font-black text-[#2563eb]">{skuPct.toFixed(1)}%</span>
              </div>
              <div className="w-full bg-[#e2e8f0]/50 rounded-full h-1.5 mt-1 border border-[#e2e8f0]">
                <div 
                  className="bg-[#2563eb] h-1.5 rounded-full" 
                  style={{ width: `${Math.min(skuPct, 100)}%` }}
                />
              </div>
            </div>
          </div>

          <div className="mt-2 pt-4 border-t border-[#e2e8f0]">
            <div className="flex items-center justify-between">
              <span className="text-[11px] font-extrabold text-[#64748b] uppercase tracking-wider text-center">Status Audit</span>
              {isLayak ? (
                <div className="flex items-center gap-1.5 bg-emerald-50 text-emerald-700 font-black text-xs uppercase px-3 py-1.5 rounded-md border border-emerald-200">
                  <CheckCircle className="w-3.5 h-3.5" /> MEMENUHI SYARAT
                </div>
              ) : (
                <div className="flex items-center gap-1.5 bg-rose-50 text-rose-700 font-black text-xs uppercase px-3 py-1.5 rounded-md border border-rose-200">
                  <AlertOctagon className="w-3.5 h-3.5" /> TIDAK LAYAK
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

    </div>
  );
};

