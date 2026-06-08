import React from 'react';
import { NooRecord, KpiReport, Salesman } from '../types';
import { Target, Crosshair, TrendingUp, CheckCircle, AlertOctagon } from 'lucide-react';

interface HunterDashboardProps {
  nooRecords: NooRecord[];
  reports: KpiReport[];
  hunters: Salesman[];
}

export const HunterDashboard: React.FC<HunterDashboardProps> = ({ nooRecords, reports, hunters }) => {
  if (hunters.length === 0) {
    return (
      <div className="p-6 bg-rose-50/50 rounded-xl border border-rose-100 flex flex-col h-full gap-6 shadow-sm">
        <h2 className="text-lg font-serif font-black text-[#0f172a] uppercase tracking-tight pb-3 border-b border-rose-200">
          <Target className="w-5 h-5 text-rose-600 inline-block mr-2" />
          Pencari Pelanggan & NOO (Hunters)
        </h2>
        <p className="text-sm text-rose-800">Belum ada salesman yang ditugaskan sebagai Hunter (NOO).</p>
      </div>
    );
  }

  return (
    <div className="p-6 bg-rose-50/50 rounded-xl border border-rose-100 flex flex-col h-full gap-6 shadow-sm">
      <div>
        <h2 className="text-lg font-serif font-black text-[#0f172a] uppercase tracking-tight pb-3 mb-2 flex items-center gap-2 border-b border-rose-200">
          <Target className="w-5 h-5 text-rose-600" />
          Pencari Pelanggan & NOO (Hunters)
        </h2>
        <p className="text-xs text-rose-800 font-medium leading-relaxed max-w-sm">
          Menganalisis hasil akuisisi outlet (NOO) dan performa kinerja lapangan untuk salesman bertugas NOO.
        </p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {hunters.map(hunter => {
          const hunterNoo = nooRecords.filter(n => n.salesmanId === hunter.id || n.salesmanName.toLowerCase().trim() === hunter.name.toLowerCase().trim());
          const hunterReports = reports.filter(r => r.salesmanId === hunter.id || (r.salesmanName && r.salesmanName.toLowerCase().trim() === hunter.name.toLowerCase().trim()));
          
          const totalGt = hunterNoo.reduce((sum, n) => sum + (n.gt || 0), 0);
          const totalMt = hunterNoo.reduce((sum, n) => sum + (n.mt || 0), 0);
          const totalNooCount = totalGt + totalMt;

          // KPI Calculations
          const cntDays = new Set(hunterReports.map(r => r.date)).size || 1;
          const tcSum = hunterReports.reduce((sum, r) => sum + r.tc, 0);
          const cpSum = hunterReports.reduce((sum, r) => sum + r.cp, 0);
          const ecSum = hunterReports.reduce((sum, r) => sum + r.ec, 0);
          const skuSum = hunterReports.reduce((sum, r) => sum + r.skuTotal, 0);

          const cpPct = tcSum > 0 ? (cpSum / tcSum) * 100 : 0;
          const ecPct = cpSum > 0 ? (ecSum / cpSum) * 100 : 0;
          
          const isLayak = cpPct >= 80 && ecPct >= 40;

          return (
            <div key={hunter.id} className="flex flex-col gap-4 border border-[#e2e8f0]/60 rounded-2xl p-4 bg-white">
              <h3 className="text-md font-bold text-[#0f172a] uppercase tracking-wider border-b border-[#e2e8f0] pb-2">{hunter.name}</h3>
              {/* NOO CARD */}
              <div className="bg-[#f8fafc] p-4 rounded-xl shadow-xs border border-[#e2e8f0]">
                <h4 className="text-xs font-black text-[#0f172a] uppercase tracking-wider flex items-center gap-2 mb-3">
                  <Crosshair className="w-3 h-3 text-rose-600" />
                  Akuisisi Outlet
                </h4>
                
                <div className="grid grid-cols-2 gap-2 text-center">
                  <div className="bg-rose-50 p-2 rounded-lg border border-rose-100">
                    <span className="block text-[9px] font-bold text-rose-600 uppercase tracking-wider mb-1">GT</span>
                    <span className="text-2xl font-serif font-black text-[#0f172a]">{totalGt}</span>
                  </div>
                  <div className="bg-white p-2 rounded-lg border border-[#e2e8f0]">
                    <span className="block text-[9px] font-bold text-[#64748b] uppercase tracking-wider mb-1">MT</span>
                    <span className="text-2xl font-serif font-black text-[#0f172a]">{totalMt}</span>
                  </div>
                </div>

                <div className="mt-3 pt-3 border-t border-[#e2e8f0]/60">
                  <div className="flex items-center justify-between">
                    <span className="text-[10px] font-bold text-[#64748b] uppercase">Total Outlet Baru</span>
                    <span className="text-xl font-black text-rose-600 font-serif">{totalNooCount}</span>
                  </div>
                </div>
              </div>

              {/* KPI CARD */}
              <div className="bg-[#f8fafc] p-4 rounded-xl shadow-xs border border-[#e2e8f0]">
                <div className="flex items-center justify-between border-b border-[#e2e8f0]/60 pb-2 mb-2">
                  <h4 className="text-xs font-black text-[#0f172a] uppercase tracking-wider flex items-center gap-2">
                    <TrendingUp className="w-3 h-3 text-rose-600" />
                    Kinerja Lapangan
                  </h4>
                  <span className="text-[9px] font-bold text-rose-700 bg-rose-50 px-2 py-0.5 rounded-md uppercase tracking-widest border border-rose-100">
                    {hunterReports.length > 0 ? cntDays : 0} Lapor
                  </span>
                </div>

                <div className="grid grid-cols-4 gap-1 mb-2">
                  <div className="bg-white rounded p-1.5 border border-[#e2e8f0] text-center">
                    <div className="text-[8px] font-bold text-[#64748b] uppercase">TC</div>
                    <div className="text-sm font-black text-[#0f172a]">{tcSum}</div>
                  </div>
                  <div className="bg-white rounded p-1.5 border border-[#e2e8f0] text-center">
                    <div className="text-[8px] font-bold text-[#64748b] uppercase">CP</div>
                    <div className="text-sm font-black text-[#0f172a]">{cpSum}</div>
                  </div>
                  <div className="bg-white rounded p-1.5 border border-[#e2e8f0] text-center">
                    <div className="text-[8px] font-bold text-[#64748b] uppercase">EC</div>
                    <div className="text-sm font-black text-[#0f172a]">{ecSum}</div>
                  </div>
                  <div className="bg-rose-50 rounded p-1.5 border border-rose-100 text-center">
                    <div className="text-[8px] font-bold text-rose-800 uppercase">SKU</div>
                    <div className="text-sm font-black text-rose-600">{skuSum}</div>
                  </div>
                </div>

                <div className="flex flex-col gap-2">
                  <div>
                    <div className="flex justify-between items-center mb-0.5">
                      <span className="text-[9px] font-extrabold text-[#0f172a] uppercase tracking-wider">CP / TC</span>
                      <span className="text-[9px] font-black">{cpPct.toFixed(0)}%</span>
                    </div>
                    <div className="w-full bg-[#e2e8f0]/60 rounded-full h-1">
                      <div className={`h-full ${cpPct >= 80 ? 'bg-emerald-500' : 'bg-rose-500'}`} style={{ width: `${Math.min(cpPct, 100)}%` }} />
                    </div>
                  </div>

                  <div>
                    <div className="flex justify-between items-center mb-0.5">
                      <span className="text-[9px] font-extrabold text-[#0f172a] uppercase tracking-wider">EC / CP</span>
                      <span className="text-[9px] font-black">{ecPct.toFixed(0)}%</span>
                    </div>
                    <div className="w-full bg-[#e2e8f0]/60 rounded-full h-1">
                      <div className={`h-full ${ecPct >= 40 ? 'bg-emerald-500' : 'bg-rose-500'}`} style={{ width: `${Math.min(ecPct, 100)}%` }} />
                    </div>
                  </div>
                </div>

                <div className="mt-3 pt-2 border-t border-[#e2e8f0]/60">
                  <div className="flex items-center justify-between">
                    <span className="text-[9px] font-extrabold text-[#64748b] uppercase tracking-wider">Status</span>
                    {isLayak ? (
                      <div className="flex items-center gap-1 bg-emerald-50 text-emerald-700 font-black text-[9px] uppercase px-2 py-1 rounded border border-emerald-200">
                        <CheckCircle className="w-3 h-3" /> LAYAK
                      </div>
                    ) : (
                      <div className="flex items-center gap-1 bg-rose-50 text-rose-700 font-black text-[9px] uppercase px-2 py-1 rounded border border-rose-200">
                        <AlertOctagon className="w-3 h-3" /> TDK LAYAK
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};

