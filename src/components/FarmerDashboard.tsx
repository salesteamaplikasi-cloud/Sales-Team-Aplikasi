import React from 'react';
import { KpiReport, NooRecord } from '../types';
import { TrendingUp, Users, AlertTriangle, Sprout, ShoppingCart, Target, ArrowRightCircle } from 'lucide-react';

interface FarmerDashboardProps {
  reports: KpiReport[];
  nooRecords: NooRecord[];
}

export const FarmerDashboard: React.FC<FarmerDashboardProps> = ({ reports, nooRecords }) => {
  const arisReports = reports.filter(r => r.salesmanName && r.salesmanName.toLowerCase() === 'aris');
  const imamNoo = nooRecords.filter(n => n.salesmanName.toLowerCase().trim() === 'imam');
  
  // Calculate farming potentials
  const totalImamWarung = imamNoo.reduce((sum, n) => sum + n.warung, 0);
  const totalImamStore = imamNoo.reduce((sum, n) => sum + n.store, 0);
  const totalImamKiosk = imamNoo.reduce((sum, n) => sum + n.kiosk, 0);
  const totalImamWholesaler = imamNoo.reduce((sum, n) => sum + n.wholesaler, 0);
  const totalImamNoo = totalImamWarung + totalImamStore + totalImamKiosk + totalImamWholesaler;

  // Let's pretend Aris converts some of Imam's NOOs into active CP
  // (In reality, we just show these as Farming Targets)
  const calculateTotal = (r: KpiReport) => r.billsReceived + (r.billsTransfer || 0);
  const avgPayment = arisReports.length > 0 ? arisReports.reduce((sum, r) => sum + calculateTotal(r), 0) / arisReports.length : 0;
  
  const potentialValue = totalImamNoo * (avgPayment || 500000); // estimated revenue per store

  return (
    <div className="p-6 bg-emerald-50/50 rounded-xl border border-emerald-100 flex flex-col h-full gap-6 shadow-sm">
      <div>
        <h2 className="text-lg font-serif font-black text-[#0f172a] uppercase tracking-tight pb-3 mb-2 flex items-center gap-2 border-b border-emerald-200">
          <Sprout className="w-5 h-5 text-emerald-600" />
          Pembinaan & Retensi Pelanggan (Aris - Farmer)
        </h2>
        <p className="text-xs text-emerald-800 font-medium leading-relaxed max-w-sm">
          Menganalisis ide strategi farming berdasarkan hasil akuisisi NOO dari Imam (Hunter).
        </p>
      </div>
      
      {/* Handover Pipeline (Ideasi Kreatif) */}
      <div className="bg-[#ffffff] p-5 rounded-2xl shadow-sm border border-[#e2e8f0]">
        <div className="flex items-center justify-between border-b border-[#e2e8f0]/60 pb-3 mb-4">
          <h3 className="text-sm font-black text-[#0f172a] uppercase tracking-wider flex items-center gap-2">
            <Target className="w-4 h-4 text-emerald-600" />
            Pipeline Serah Terima Pelanggan
          </h3>
        </div>
        
        <div className="flex items-center justify-between text-center gap-4 bg-[#f8fafc] p-4 rounded-xl border border-[#e2e8f0]">
          <div className="flex-1">
            <span className="block text-[10px] font-bold text-[#64748b] uppercase tracking-widest mb-1">NOO Terkumpul (Imam)</span>
            <span className="text-3xl font-serif font-black text-rose-600">{totalImamNoo} <span className="text-sm text-[#64748b] font-medium tracking-normal">Toko Baru</span></span>
          </div>
          <ArrowRightCircle className="w-8 h-8 text-emerald-300 md:shrink-0" />
          <div className="flex-1">
            <span className="block text-[10px] font-bold text-[#64748b] uppercase tracking-widest mb-1">Nilai Potensi Sales</span>
            <span className="text-xl font-serif font-black text-emerald-600">Rp {potentialValue.toLocaleString('id-ID')}</span>
          </div>
        </div>

        <div className="mt-4 grid grid-cols-2 lg:grid-cols-4 gap-2">
          <div className="bg-white border border-[#e2e8f0] rounded-lg p-3 text-center shadow-xs">
             <div className="text-[10px] font-bold text-[#64748b] uppercase tracking-wider">Warung</div>
             <div className="font-black text-[#0f172a]">{totalImamWarung}</div>
          </div>
          <div className="bg-white border border-[#e2e8f0] rounded-lg p-3 text-center shadow-xs">
             <div className="text-[10px] font-bold text-[#64748b] uppercase tracking-wider">Toko Mod</div>
             <div className="font-black text-[#0f172a]">{totalImamStore}</div>
          </div>
          <div className="bg-white border border-[#e2e8f0] rounded-lg p-3 text-center shadow-xs">
             <div className="text-[10px] font-bold text-[#64748b] uppercase tracking-wider">Kios</div>
             <div className="font-black text-[#0f172a]">{totalImamKiosk}</div>
          </div>
          <div className="bg-white border border-[#e2e8f0] rounded-lg p-3 text-center shadow-xs">
             <div className="text-[10px] font-bold text-[#64748b] uppercase tracking-wider">Grosir</div>
             <div className="font-black text-[#0f172a]">{totalImamWholesaler}</div>
          </div>
        </div>
      </div>

      <div className="bg-emerald-700 text-white p-5 rounded-2xl shadow-sm relative overflow-hidden flex-1">
        <div className="relative z-10">
          <h3 className="text-xs font-black uppercase tracking-wider mb-4 flex items-center gap-2 text-emerald-100">
            <ShoppingCart className="w-4 h-4 text-emerald-300" />
            Ideasi Penetrasi & Fokus KPI (ARIS)
          </h3>
          <ul className="text-xs space-y-3 font-medium text-emerald-50 leading-relaxed relative z-10">
            <li className="flex items-start gap-2.5">
              <span className="shrink-0 mt-0.5 w-1.5 h-1.5 bg-emerald-300 rounded-full"></span> 
              <span>Distribusikan SKU fast-moving prioritas ke <strong className="text-white bg-emerald-600 px-1 rounded">{totalImamWarung} Warung</strong> dan <strong className="text-white bg-emerald-600 px-1 rounded">{totalImamKiosk} Kios</strong> baru yang ditemukan dari database Imam bulan ini.</span>
            </li>
            <li className="flex items-start gap-2.5">
              <span className="shrink-0 mt-0.5 w-1.5 h-1.5 bg-emerald-300 rounded-full"></span> 
              <span>Tingkatkan rasio <strong className="text-white bg-emerald-600 px-1 rounded">Effective Call (EC/CP) &gt;= 50%</strong> untuk memvalidasi bahwa pelanggan baru dari Hunter sudah menempatkan order berulang secara rutin.</span>
            </li>
            <li className="flex items-start gap-2.5">
              <span className="shrink-0 mt-0.5 w-1.5 h-1.5 bg-emerald-300 rounded-full"></span> 
              <span>Targetkan kunjungan intensif (Weekly Call Plan) untuk menjaga <strong className="text-white bg-emerald-600 px-1 rounded">{totalImamWholesaler} Grosir</strong> demi mencapai volume transaksi target ISO maksimal.</span>
            </li>
          </ul>
        </div>
        <div className="absolute -bottom-8 -right-8 opacity-10 blur-xs">
          <Sprout className="w-48 h-48" strokeWidth={1} />
        </div>
      </div>

    </div>
  );
};

