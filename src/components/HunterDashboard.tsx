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
  
  const isLayak = cpPct >= 80 && ecPct >= 40;

  return (
    <div className="p-6 bg-rose-50/50 rounded-xl border border-rose-100 flex flex-col h-full gap-6 shadow-sm">
      <div>
        <h2 className="text-lg font-serif font-black text-[#0f172a] uppercase tracking-tight pb-3 mb-2 flex items-center gap-2 border-b border-rose-200">
          <Target className="w-5 h-5 text-rose-600" />
          Pencari Pelanggan & NOO (Imam - Hunter)
        </h2>
        <p className="text-xs text-rose-800 font-medium leading-relaxed max-w-sm">
          Menganalisis hasil akuisisi outlet (NOO) dan performa kinerja lapangan Hunter.
        </p>
      </div>

      {/* NOO CARD */}
      <div className="bg-[#ffffff] p-5 rounded-2xl shadow-sm border border-[#e2e8f0]">
        <div className="flex items-center justify-between border-b border-[#e2e8f0]/60 pb-3 mb-4">
          <h3 className="text-sm font-black text-[#0f172a] uppercase tracking-wider flex items-center gap-2">
            <Crosshair className="w-4 h-4 text-rose-600" />
            Keberhasilan Akuisisi Outlet
          </h3>
        </div>
        
        <div className="grid grid-cols-2 gap-3 text-center mt-4">
          <div className="bg-rose-50 p-4 rounded-xl border border-rose-100">
            <span className="block text-[10px] font-bold text-rose-600 uppercase tracking-wider mb-1">Warung</span>
            <span className="text-3xl font-serif font-black text-[#0f172a]">{totalWarung}</span>
          </div>
          <div className="bg-[#f8fafc] p-4 rounded-xl border border-[#e2e8f0]">
            <span className="block text-[10px] font-bold text-[#64748b] uppercase tracking-wider mb-1">Toko Mod</span>
            <span className="text-3xl font-serif font-black text-[#0f172a]">{totalStore}</span>
          </div>
          <div className="bg-rose-50 p-4 rounded-xl border border-rose-100">
            <span className="block text-[10px] font-bold text-rose-600 uppercase tracking-wider mb-1">Kios</span>
            <span className="text-3xl font-serif font-black text-[#0f172a]">{totalKiosk}</span>
          </div>
          <div className="bg-[#f8fafc] p-4 rounded-xl border border-[#e2e8f0]">
            <span className="block text-[10px] font-bold text-[#64748b] uppercase tracking-wider mb-1">Grosir</span>
            <span className="text-3xl font-serif font-black text-[#0f172a]">{totalWholesaler}</span>
          </div>
        </div>

        <div className="mt-4 pt-4 border-t border-[#e2e8f0]/60">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <span className="text-xs font-bold text-[#64748b] uppercase">Total Outlet Baru Dibikin</span>
            </div>
            <span className="text-2xl font-black text-rose-600 font-serif">{totalNooCount} <span className="text-xs font-bold uppercase tracking-wider text-[#64748b] font-sans">Toko</span></span>
          </div>
        </div>
      </div>

      {/* KPI CARD */}
      <div className="bg-[#ffffff] p-5 rounded-2xl shadow-sm border border-[#e2e8f0] flex-1">
        <div className="flex flex-col gap-4">
          <div className="flex items-center justify-between border-b border-[#e2e8f0]/60 pb-3 mb-2">
            <h3 className="text-sm font-black text-[#0f172a] uppercase tracking-wider flex items-center gap-2">
              <TrendingUp className="w-4 h-4 text-rose-600" />
              Kinerja Lapangan Aktif
            </h3>
            <span className="text-[10px] font-bold text-rose-700 bg-rose-50 px-2.5 py-1 rounded-md uppercase tracking-widest border border-rose-100">
              {imamReports.length > 0 ? cntDays : 0} Hari Lapor
            </span>
          </div>

          {/* Quick Metrics */}
          <div className="grid grid-cols-4 gap-2 mb-2">
            <div className="bg-[#f8fafc] rounded-lg p-2.5 border border-[#e2e8f0]/60 text-center shadow-xs">
              <div className="text-[10px] font-extrabold text-[#64748b] uppercase tracking-wider mb-1">TC</div>
              <div className="text-lg font-black text-[#0f172a] font-serif">{tcSum}</div>
            </div>
            <div className="bg-[#f8fafc] rounded-lg p-2.5 border border-[#e2e8f0]/60 text-center shadow-xs">
              <div className="text-[10px] font-extrabold text-[#64748b] uppercase tracking-wider mb-1">CP</div>
              <div className="text-lg font-black text-[#0f172a] font-serif">{cpSum}</div>
            </div>
            <div className="bg-[#f8fafc] rounded-lg p-2.5 border border-[#e2e8f0]/60 text-center shadow-xs">
              <div className="text-[10px] font-extrabold text-[#64748b] uppercase tracking-wider mb-1">EC</div>
              <div className="text-lg font-black text-[#0f172a] font-serif">{ecSum}</div>
            </div>
            <div className="bg-rose-50 rounded-lg p-2.5 border border-rose-100 text-center shadow-xs">
              <div className="text-[10px] font-extrabold text-rose-800 uppercase tracking-wider mb-1">SKU</div>
              <div className="text-lg font-black text-rose-600 font-serif">{skuSum}</div>
            </div>
          </div>

          <div className="flex flex-col gap-3">
            {/* CP Progress */}
            <div>
              <div className="flex justify-between items-center mb-1">
                <span className="text-[10px] font-extrabold text-[#0f172a] uppercase tracking-wider">CP / TC</span>
                <span className="text-[10px] font-black">{cpPct.toFixed(1)}%</span>
              </div>
              <div className="w-full bg-[#e2e8f0]/60 rounded-full h-1.5 overflow-hidden">
                <div 
                  className={`h-full ${cpPct >= 80 ? 'bg-emerald-500' : 'bg-rose-500'}`}
                  style={{ width: `${Math.min(cpPct, 100)}%` }}
                />
              </div>
            </div>

            {/* EC Progress */}
            <div>
              <div className="flex justify-between items-center mb-1">
                <span className="text-[10px] font-extrabold text-[#0f172a] uppercase tracking-wider">EC / CP</span>
                <span className="text-[10px] font-black">{ecPct.toFixed(1)}%</span>
              </div>
              <div className="w-full bg-[#e2e8f0]/60 rounded-full h-1.5 overflow-hidden">
                <div 
                  className={`h-full ${ecPct >= 40 ? 'bg-emerald-500' : 'bg-rose-500'}`}
                  style={{ width: `${Math.min(ecPct, 100)}%` }}
                />
              </div>
            </div>
          </div>

          <div className="mt-2 pt-4 border-t border-[#e2e8f0]/60">
            <div className="flex items-center justify-between">
              <span className="text-[10px] font-extrabold text-[#64748b] uppercase tracking-wider text-center">Status Audit</span>
              {isLayak ? (
                <div className="flex items-center gap-1.5 bg-emerald-50/80 text-emerald-700 font-black text-[10px] uppercase px-3 py-1.5 rounded border border-emerald-200 shadow-xs">
                  <CheckCircle className="w-3.5 h-3.5" /> MEMENUHI LAYAK KPI
                </div>
              ) : (
                <div className="flex items-center gap-1.5 bg-rose-50/80 text-rose-700 font-black text-[10px] uppercase px-3 py-1.5 rounded border border-rose-200 shadow-xs">
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

