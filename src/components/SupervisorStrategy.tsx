import React from 'react';
import { Target, Map, MapPin, Handshake, Gift } from 'lucide-react';

export const SupervisorStrategy: React.FC = () => {
  return (
    <div className="bg-[#ffffff] p-6 rounded-3xl shadow-sm border border-[#e2e8f0]">
      <h3 className="text-sm font-black text-[#0f172a] uppercase tracking-wider flex items-center gap-2 mb-4 border-b border-[#e2e8f0]/60 pb-3">
        <Map className="w-5 h-5 text-indigo-600" />
        Strategi NOO Area BARLINGMASCAKEB
      </h3>
      
      <p className="text-xs text-[#64748b] leading-relaxed mb-6 max-w-3xl">
        Standar strategi penetrasi pasar dan cara membuat New Outlet Opening (NOO) lebih menarik bagi calon warung, toko kelontong (GT), maupun minimarket (MT) di area Kabupaten Banjarnegara, Purbalingga, Banyumas, Cilacap, dan Kebumen.
      </p>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="bg-gradient-to-br from-indigo-50 to-white p-4 rounded-2xl border border-indigo-100 flex flex-col gap-2 relative overflow-hidden">
          <Gift className="w-6 h-6 text-indigo-600 mb-1" />
          <h4 className="text-xs font-bold text-[#0f172a] uppercase tracking-tight">1. Paket Perdana (Starter Pack) Plus Retur</h4>
          <p className="text-[10px] text-[#64748b] leading-snug">
            Warung di area daerah karesidenan Banyumas lebih suka rasa aman. Tawarkan bundling produk fast-moving dengan jaminan <strong>retur 100%</strong> jika tidak laku dalam 1 bulan pertama.
          </p>
        </div>

        <div className="bg-gradient-to-br from-emerald-50 to-white p-4 rounded-2xl border border-emerald-100 flex flex-col gap-2 relative overflow-hidden">
          <Handshake className="w-6 h-6 text-emerald-600 mb-1" />
          <h4 className="text-xs font-bold text-[#0f172a] uppercase tracking-tight">2. ToP (Term of Payment) Fleksibel</h4>
          <p className="text-[10px] text-[#64748b] leading-snug">
            Jangan paksakan pembayaran lunas (Cash before Delivery). Berikan dispensasi ToP Konsinyasi atau 14 Hari untuk nota NOO pertama dengan Plafon Nominal terbatas guna menumbuhkan <i>trust</i>.
          </p>
        </div>

        <div className="bg-gradient-to-br from-rose-50 to-white p-4 rounded-2xl border border-rose-100 flex flex-col gap-2 relative overflow-hidden">
          <Target className="w-6 h-6 text-rose-600 mb-1" />
          <h4 className="text-xs font-bold text-[#0f172a] uppercase tracking-tight">3. Reward & POS Material Langsung</h4>
          <p className="text-[10px] text-[#64748b] leading-snug">
            Sediakan POS Material (seperti Spanduk Toko gratis cetak nama warung) atau Rak/Display Kecil sebagai apresiasi langsung saat toko setuju mendaftar sebagai Outlet Baru.
          </p>
        </div>

        <div className="bg-gradient-to-br from-amber-50 to-white p-4 rounded-2xl border border-amber-100 flex flex-col gap-2 relative overflow-hidden">
          <MapPin className="w-6 h-6 text-amber-600 mb-1" />
          <h4 className="text-xs font-bold text-[#0f172a] uppercase tracking-tight">4. Sampling & Testimoni Lokal</h4>
          <p className="text-[10px] text-[#64748b] leading-snug">
            Siapkan produk khusus sampling (Taster). Sebutkan nama grosir atau toko modern besar (contoh: Moro, Rita, dsb di area Barlingmascakeb) yang sudah laku keras menjual produk kita.
          </p>
        </div>
      </div>
    </div>
  );
};
