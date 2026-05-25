import React, { useRef } from "react";
import * as XLSX from "xlsx";
import { motion, AnimatePresence } from "motion/react";
import { Crown, Sparkles, Check, Loader2, X, History, Gift, Upload } from "lucide-react";

interface CustomerLoyaltyPortalProps {
  customers: any[];
  setCustomers: React.Dispatch<React.SetStateAction<any[]>>;
  sheetsScriptUrl: string;
  salesmen: any[];
  handleSyncLoyaltyToSheets: (customCustomers?: any[], silent?: boolean) => Promise<void>;
  showToast: (msg: string, type?: "success" | "error" | "info" | "warning") => void;
  setCustomerMode: React.Dispatch<React.SetStateAction<boolean>>;
  customerActiveSubTab: "check" | "register";
  setCustomerActiveSubTab: React.Dispatch<React.SetStateAction<"check" | "register">>;
  customerSearchQuery: string;
  setCustomerSearchQuery: React.Dispatch<React.SetStateAction<string>>;
  matchedCustomer: any | null;
  setMatchedCustomer: React.Dispatch<React.SetStateAction<any | null>>;
  registerSuccessName: string;
  setRegisterSuccessName: React.Dispatch<React.SetStateAction<string>>;
  isCustomerSelfRegistering: boolean;
  setIsCustomerSelfRegistering: React.Dispatch<React.SetStateAction<boolean>>;
  selfRegForm: {
    name: string;
    address: string;
    salesmanName: string;
    area: string;
    jenisToko: string;
    estimatedOmzet: string;
    notesPerDay: string;
    storeAgeYears: string;
    ownership: string;
  };
  setSelfRegForm: React.Dispatch<React.SetStateAction<any>>;
  REWARDS_CATALOG: any[];
}

export const CustomerLoyaltyPortal: React.FC<CustomerLoyaltyPortalProps> = ({
  customers,
  setCustomers,
  sheetsScriptUrl,
  salesmen,
  handleSyncLoyaltyToSheets,
  showToast,
  setCustomerMode,
  customerActiveSubTab,
  setCustomerActiveSubTab,
  customerSearchQuery,
  setCustomerSearchQuery,
  matchedCustomer,
  setMatchedCustomer,
  registerSuccessName,
  setRegisterSuccessName,
  isCustomerSelfRegistering,
  setIsCustomerSelfRegistering,
  selfRegForm,
  setSelfRegForm,
  REWARDS_CATALOG,
}) => {
  const handleCustomerSelfRegister = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selfRegForm.name.trim()) {
      showToast("Nama toko wajib diisi!", "error");
      return;
    }
    if (!selfRegForm.address.trim()) {
      showToast("Alamat lengkap wajib diisi!", "error");
      return;
    }

    setIsCustomerSelfRegistering(true);
    let tier: "Platinum" | "Gold" | "Silver" | "Bronze" = "Bronze";
    const omzetValue = Number(selfRegForm.estimatedOmzet);
    if (omzetValue >= 15000000) tier = "Platinum";
    else if (omzetValue >= 8000000) tier = "Gold";
    else if (omzetValue >= 5000000) tier = "Silver";

    const customStore = {
      id: "c-self-" + Date.now(),
      name: selfRegForm.name,
      address: selfRegForm.address,
      salesmanName: selfRegForm.salesmanName || (salesmen[0]?.name || "RIZKY"),
      area: selfRegForm.area || (salesmen[0]?.area || "Semarang"),
      jenisToko: selfRegForm.jenisToko,
      estimatedOmzet: omzetValue,
      notesPerDay: Number(selfRegForm.notesPerDay),
      storeAgeYears: Number(selfRegForm.storeAgeYears),
      ownership: selfRegForm.ownership,
      points: 50, // Starter bonus points
      tier,
      createdAt: new Date().toISOString(),
      actionsLog: [
        {
          date: new Date().toISOString().split("T")[0],
          action: "Registrasi Mandiri",
          notes: "Pendaftaran customer profil mandiri via QR Code Portal DKR. Poin startup +50 diberikan.",
          status: "Selesai"
        }
      ]
    };

    const updated = [customStore, ...customers];
    setCustomers(updated);
    
    // Save to LocalStorage in parent environment
    try {
      localStorage.setItem("KPI_LOYALTY_CUSTOMERS", JSON.stringify(updated));
    } catch (_) {}

    // Clear self reg form
    setSelfRegForm({
      name: "",
      address: "",
      salesmanName: salesmen[0]?.name || "",
      area: salesmen[0]?.area || "",
      jenisToko: "Sembako",
      estimatedOmzet: "5000000",
      notesPerDay: "5",
      storeAgeYears: "2",
      ownership: "Milik Sendiri"
    });

    setRegisterSuccessName(customStore.name);
    setMatchedCustomer(customStore); // Set the card view directly so they can inspect their card immediately!
    setCustomerActiveSubTab("check"); // Switch to lookup/points card tab so they can see their beautiful card
    setIsCustomerSelfRegistering(false);
    showToast(`Registrasi sukses! Selamat bergabung ${customStore.name.toUpperCase()}!`, "success");

    // Live sync to Sheets!
    if (sheetsScriptUrl) {
      await handleSyncLoyaltyToSheets(updated, true);
    }
  };

  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleImportXLSX = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData: any[] = XLSX.utils.sheet_to_json(sheet);

      const newCustomers = jsonData.map((row: any) => ({
        id: "c-" + Date.now() + "-" + Math.random(),
        name: row['Nama Toko / Outlet Anda'] || "Tanpa Nama",
        address: row['Alamat Lengkap Outlet'] || "-",
        salesmanName: row['Pilih Sales Wilayah Anda'] || salesmen[0]?.name || "RIZKY",
        area: salesmen[0]?.area || "Semarang",
        jenisToko: row['Jenis & Sektor Toko'] || "Sembako",
        estimatedOmzet: Number(row['Estimasi Kas Belanja Outlet Per Bulan (Rupiah)']) || 0,
        notesPerDay: 0,
        storeAgeYears: Number(row['Lama Toko Berdiri (Tahun)']) || 0,
        ownership: row['Status Sewa Bangunan'] || "Milik Sendiri",
        points: 50,
        tier: "Bronze",
        createdAt: row['Tanggal Daftar'] || new Date().toISOString(),
        actionsLog: []
      }));

      const updated = [...newCustomers, ...customers];
      setCustomers(updated);
      try {
        localStorage.setItem("KPI_LOYALTY_CUSTOMERS", JSON.stringify(updated));
      } catch (_) {}
      showToast(`Berhasil mengimpor ${newCustomers.length} toko!`, "success");
    };
    reader.readAsArrayBuffer(file);
  };


  return (
    <div className="space-y-6 max-w-4xl mx-auto px-2 py-4">
      {/* Glowing Premium Customer Portal Banner */}
      <div className="bg-[#FAF9F6] border border-[#E5E5DF] p-6 rounded-3xl shadow-xl flex flex-col md:flex-row items-center justify-between gap-4 relative overflow-hidden">
        <div className="absolute top-0 right-0 w-24 h-24 bg-gradient-to-br from-amber-500/10 to-transparent rounded-bl-3xl pointer-events-none"></div>
        <div className="flex items-center gap-4 text-center md:text-left flex-col md:flex-row">
          <div className="w-14 h-14 bg-gradient-to-br from-amber-500 to-amber-600 rounded-2xl flex items-center justify-center shadow-md shadow-amber-500/20">
            <Crown className="w-8 h-8 text-white animate-pulse" />
          </div>
          <div>
            <h2 className="text-xl font-black text-[#5A5A40] uppercase tracking-wide flex items-center gap-2 justify-center md:justify-start">
              DKR Client Loyalty Portal
            </h2>
            <p className="text-xs text-[#8C8C70] font-sans mt-0.5 max-w-xl">
              Registrasi mandiri, cek tingkat tier, dan ketahui perolehan poin toko Anda untuk ditukarkan hadiah sponsor serta diskon belanja eksklusif.
            </p>
          </div>
        </div>

        <button
          onClick={() => {
            setCustomerMode(false);
            // remove query params
            try {
              const url = window.location.origin + window.location.pathname;
              window.history.pushState({}, "", url);
            } catch (_) {}
          }}
          className="px-4 py-2 bg-[#E5E5DF]/50 hover:bg-[#E5E5DF] text-[#4A4A3C] text-[10px] font-black uppercase tracking-wider rounded-xl transition cursor-pointer flex items-center gap-1.5 self-center"
          title="Kembali ke Konsol Admin"
        >
          🚪 Keluar Portal
        </button>
      </div>

      {/* Customer Portal Toggle buttons */}
      <div className="grid grid-cols-2 gap-3.5 max-w-sm mx-auto">
        <button
          onClick={() => setCustomerActiveSubTab("check")}
          className={`py-3 px-4 rounded-xl font-extrabold text-[#4A4A3C] text-xs uppercase tracking-wider transition-all duration-150 cursor-pointer text-center ${
            customerActiveSubTab === "check"
              ? "bg-[#FAF5E6] border border-amber-600/35 text-amber-950 shadow-md shadow-amber-500/5"
              : "bg-[#E5E5DF]/30 text-[#8C8C70] hover:bg-[#E5E5DF]/50 hover:text-[#4A4A3C]"
          }`}
        >
          🪙 Cek Kartu Poin Saya
        </button>
        <button
          onClick={() => setCustomerActiveSubTab("register")}
          className={`py-3 px-4 rounded-xl font-extrabold text-[#4A4A3C] text-xs uppercase tracking-wider transition-all duration-150 cursor-pointer text-center ${
            customerActiveSubTab === "register"
              ? "bg-[#FAF5E6] border border-amber-600/35 text-amber-950 shadow-md shadow-amber-500/5"
              : "bg-[#E5E5DF]/30 text-[#8C8C70] hover:bg-[#E5E5DF]/50 hover:text-[#4A4A3C]"
          }`}
        >
          ✨ Registrasi Outlet Baru
        </button>
      </div>

      {/* SUBTAB 1: CEK POIN / KARTU DIGITAL */}
      {customerActiveSubTab === "check" && (
        <div className="space-y-6 animate-fadeIn">
          <div className="bg-[#FAF9F6] p-6 rounded-3xl border border-[#E5E5DF] shadow-md max-w-xl mx-auto">
            <h3 className="text-xs font-black text-[#5A5A40] uppercase tracking-wider mb-2 flex items-center gap-2">
              🔍 Cari & Tampilkan Kartu Poin Toko
            </h3>
            <p className="text-[11px] text-[#8C8C70] mb-4">
              Masukkan nama toko Anda untuk melacak saldo poin dan status keanggotaan real-time.
            </p>

            <div className="relative space-y-2">
              <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                Ketik Nama Toko Anda:
              </label>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={customerSearchQuery}
                  onChange={(e) => {
                    setCustomerSearchQuery(e.target.value);
                    if (!e.target.value.trim()) setMatchedCustomer(null);
                  }}
                  placeholder="Contoh: Toko Makmur..."
                  className="flex-1 bg-white border border-[#E5E5DF] rounded-xl px-4 py-3 text-xs text-[#4A4A3C] font-semibold focus:outline-hidden focus:ring-1 focus:ring-[#5A5A40] shadow-xs"
                />
                <button
                  type="button"
                  onClick={() => {
                    if (!customerSearchQuery.trim()) {
                      showToast("Ketikkan nama toko anda untuk melacak!", "error");
                      return;
                    }
                    const match = customers.find(c => 
                      c.name.toLowerCase().includes(customerSearchQuery.toLowerCase())
                    );
                    if (match) {
                      setMatchedCustomer(match);
                      showToast(`Kartu Toko ${match.name.toUpperCase()} ditemukan!`, "success");
                    } else {
                      setMatchedCustomer(null);
                      showToast("Toko tidak ditemukan. Silakan cek ejaan nama toko Anda atau lakukan daftar mandiri!", "error");
                    }
                  }}
                  className="bg-[#5A5A40] hover:bg-[#4A4A3C] text-white px-5 py-3 rounded-xl text-xs font-bold uppercase tracking-wider transition cursor-pointer"
                >
                  Cari Kartu
                </button>
              </div>

              {/* Autocomplete suggestions dropdown lists */}
              {customerSearchQuery.trim() !== "" && !matchedCustomer && (() => {
                const filtered = customers.filter(c => 
                  c.name.toLowerCase().includes(customerSearchQuery.toLowerCase())
                ).slice(0, 5);
                if (filtered.length === 0) return null;
                return (
                  <div className="absolute left-0 right-0 z-20 bg-white border border-[#E5E5DF] rounded-xl mt-1.5 shadow-xl max-h-52 overflow-y-auto divide-y divide-[#E5E5DF]/40">
                    {filtered.map(c => (
                      <button
                        key={c.id}
                        onClick={() => {
                          setCustomerSearchQuery(c.name);
                          setMatchedCustomer(c);
                          showToast(`Toko ${c.name.toUpperCase()} berhasil dimuat!`, "success");
                        }}
                        className="w-full text-left px-4 py-2.5 hover:bg-[#E5E5DF]/20 transition flex items-center justify-between text-xs"
                      >
                        <div className="flex flex-col">
                          <span className="font-bold text-[#4A4A3C]">{c.name}</span>
                          <span className="text-[10px] text-[#8C8C70] truncate max-w-[280px]">📍 {c.address}</span>
                        </div>
                        <span className="text-[9px] font-black bg-amber-500/15 text-amber-800 px-1.5 py-0.5 rounded font-mono">
                          ★ Tier {c.tier}
                        </span>
                      </button>
                    ))}
                  </div>
                );
              })()}
            </div>
          </div>

          {/* VIP Animated / Glowing Metallic Loyalty Card */}
          {matchedCustomer && (() => {
            let cardGradient = "from-amber-950 via-neutral-900 to-amber-950 text-orange-100 shadow-orange-900/10";
            let ringColor = "ring-orange-500/35 border-orange-700/50";
            let benefitList = ["Poin awal starter, akses promosi umum, diskon berkala."];

            if (matchedCustomer.tier === "Platinum") {
              cardGradient = "from-[#0F0F1A] via-[#1F1435] to-[#0A0512] text-purple-100 shadow-purple-900/30";
              ringColor = "ring-purple-500/40 border-purple-500/30 ring-offset-black";
              benefitList = [
                "Prioritas Pengiriman SKU Utama (DKR Fast-Shipping)",
                "Vip potongan diskon khusus 5% untuk semua pembelanjaan SKU",
                "Points Multiplier x2.0 (Dua Kali Lebih Cepat Dapat Hadiah)",
                "Subsidi Tagihan dan Kredit Jatuh Tempo maksimal Rp 500.000/bulan"
              ];
            } else if (matchedCustomer.tier === "Gold") {
              cardGradient = "from-amber-600 via-amber-550 to-yellow-800 text-yellow-50 shadow-amber-600/20";
              ringColor = "ring-amber-500/50 border-amber-400/50";
              benefitList = [
                "Potongan diskon khusus 3% SKU utama (Sponsor Utama)",
                "Points Multiplier x1.5 (Poin Terkumpul Lebih Cepat)",
                "Voucher Free Premium Merchandise pada setiap tanggal 25"
              ];
            } else if (matchedCustomer.tier === "Silver") {
              cardGradient = "from-slate-400 via-zinc-500 to-zinc-700 text-slate-100 shadow-slate-500/20";
              ringColor = "ring-slate-400/45 border-slate-300/40";
              benefitList = [
                "Potongan diskon khusus 1% untuk produk SKU DKR reguler",
                "Prioritas antrean bongkar muatan logistik di Semarang",
                "Voucher Free Merchandise reguler pada setiap tanggal 25"
              ];
            } else {
              benefitList = [
                "Diberikan starter voucher bonus 50 poin langsung",
                "Prioritas info katalog promosi terbaru dari tim sales",
                "Konsultasi free penataan rak toko dan program diskon produsen"
              ];
            }

            return (
              <motion.div
                initial={{ scale: 0.95, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                className="max-w-xl mx-auto space-y-6"
              >
                {/* The Digital Card Element */}
                <div className={`aspect-[1.586/1] w-full rounded-3xl p-6 bg-gradient-to-br ${cardGradient} relative overflow-hidden ring-4 ${ringColor} shadow-2xl flex flex-col justify-between border select-none transition-all`}>
                  <div className="absolute inset-0 bg-white/5 pointer-events-none opacity-45 mix-blend-overlay"></div>
                  
                  {/* Top card metadata */}
                  <div className="flex items-start justify-between z-10">
                    <div className="flex items-center gap-1.5">
                      <Crown className="w-6 h-6 text-amber-500 fill-amber-500 shrink-0 animate-bounce" />
                      <span className="text-[10px] font-black uppercase tracking-widest font-mono text-white">
                        DKR SALES LOYALTY CARD
                      </span>
                    </div>
                    <span className="bg-white/15 backdrop-blur-md text-[8.5px] px-2.5 py-1 rounded-full uppercase font-black tracking-widest border border-white/20 text-white">
                      ★ {matchedCustomer.tier} MEMBER
                    </span>
                  </div>

                  {/* Card Chip Mock */}
                  <div className="w-9 h-7 rounded-md bg-gradient-to-br from-yellow-350 via-amber-400 to-yellow-550 border border-amber-600/35 relative overflow-hidden self-start ml-1 shadow-sm opacity-90 z-10">
                    <div className="absolute inset-x-0 top-1/2 -translate-y-1/2 h-px bg-amber-800/40"></div>
                    <div className="absolute inset-y-0 left-1/2 -translate-x-1/2 w-px bg-amber-800/40"></div>
                    <div className="absolute w-3 h-3 rounded-full border border-amber-800/30 left-1/2 top-1/2 -translate-x-1/2 -translate-y-1/2"></div>
                  </div>

                  {/* Card Middle: Store Name */}
                  <div className="flex flex-col gap-0.5 z-10 mt-2">
                    <span className="text-[8px] uppercase tracking-wider font-mono opacity-80 text-white/70">NAMA OUTLET</span>
                    <span className="text-xl font-black tracking-wide truncate drop-shadow-md text-white">
                      {matchedCustomer.name.toUpperCase()}
                    </span>
                  </div>

                  {/* Card Bottom: Agent Metadata & Points */}
                  <div className="flex justify-between items-end border-t border-white/10 pt-3 z-10 mt-1">
                    <div className="flex flex-col text-left">
                      <span className="text-[8px] uppercase tracking-wider opacity-65 text-white/50 font-black">WILAYAH / SALESMAN</span>
                      <span className="text-[10px] font-mono font-bold tracking-tight text-white">
                        {matchedCustomer.area.toUpperCase()} • {matchedCustomer.salesmanName}
                      </span>
                    </div>
                    <div className="flex flex-col items-end">
                      <span className="text-[8px] uppercase tracking-wider opacity-65 text-amber-300 font-black">SALDO POIN</span>
                      <span className="text-2xl font-black font-mono tracking-tighter text-amber-300 flex items-center gap-1 drop-shadow-sm">
                        🪙 {matchedCustomer.points || 0}
                      </span>
                    </div>
                  </div>
                </div>

                {/* Detailed Benefits List */}
                <div className="bg-[#FAF9F6] p-5 rounded-3xl border border-[#E5E5DF] shadow-xs space-y-3 text-left">
                  <h4 className="text-xs font-black text-[#4A4A3C] uppercase tracking-wider flex items-center gap-1.5 border-b border-[#E5E5DF] pb-2">
                    <Sparkles className="w-4 h-4 text-purple-600" />
                    Spesifikasi Hak & Keuntungan Keanggotaan Member Tier <strong>{matchedCustomer.tier}</strong>:
                  </h4>
                  <ul className="space-y-2">
                    {benefitList.map((ben, bIdx) => (
                      <li key={bIdx} className="text-xs text-[#4A4A3C] flex items-start gap-2 leading-relaxed">
                        <span className="text-emerald-600 mt-1 text-[13px] font-extrabold shrink-0">✓</span>
                        <span>{ben}</span>
                      </li>
                    ))}
                  </ul>
                  <p className="text-[10px] text-[#8C8C70] italic pt-1 text-center bg-[#E5E5DF]/10 p-2 rounded-xl border border-[#E5E5DF]/50">
                    * Poin dapat terkumpul ketika salesman kami membuat input kunjungan rutin & nota kirim ke Database.
                  </p>
                </div>

                {/* Timeline of Follow-Up Actions Audit from DKR */}
                <div className="bg-[#FAF9F6] p-5 rounded-3xl border border-[#E5E5DF] shadow-xs space-y-3.5 text-left">
                  <div>
                    <h4 className="text-xs font-black text-[#5A5A40] uppercase tracking-wider flex items-center gap-2">
                      <History className="w-4 h-4 text-[#8C8C70]" />
                      Riwayat Tindakan & Kunjungan Sales ({matchedCustomer.actionsLog?.length || 0})
                    </h4>
                    <p className="text-[10.5px] text-[#8C8C70] mt-0.5">Pantau agenda perbaikan & jadwal kunjungan dari tim Kepala Salesman.</p>
                  </div>

                  {matchedCustomer.actionsLog && matchedCustomer.actionsLog.length > 0 ? (
                    <div className="space-y-2">
                      {matchedCustomer.actionsLog.map((act: any, aIdx: number) => (
                        <div key={aIdx} className="bg-white p-2.5 rounded-xl border border-[#E5E5DF]/65 text-xs flex justify-between items-start gap-3">
                          <div className="flex flex-col gap-1 min-w-0">
                            <div className="flex items-center gap-1.5 flex-wrap">
                              <span className="font-mono text-[9px] bg-[#E5E5DF] text-[#4A4A3C] px-1.5 py-0.5 rounded font-bold">
                                {act.date}
                              </span>
                              <span className="font-black text-[#4A4A3C] uppercase tracking-wide text-[10px]">
                                {act.action}
                              </span>
                            </div>
                            <span className="text-[#8C8C70] italic">
                              "{act.notes}"
                            </span>
                          </div>
                          <span className="text-[9.5px] font-black bg-emerald-500/15 text-emerald-800 px-1.5 py-0.5 rounded shrink-0">
                            ✓ Selesai
                          </span>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <p className="text-[11px] text-[#8C8C70] italic text-center p-3 bg-[#E5E5DF]/10 border border-dashed border-[#E5E5DF] rounded-xl">
                      Belum ada riwayat follow-up. Tim kami akan mengecek status toko Anda.
                    </p>
                  )}
                </div>

                {/* Rewards Catalogue section for customer info */}
                <div className="bg-[#FAF9F6] p-5 rounded-3xl border border-[#E5E5DF] shadow-xs space-y-3 text-left">
                  <h4 className="text-xs font-black text-[#4A4A3C] uppercase tracking-wider flex items-center gap-2">
                    <Gift className="w-4 h-4 text-emerald-700" />
                    Pilihan Cinderamata Loyalti Sponsor Tersedia:
                  </h4>
                  <p className="text-[11px] text-[#8C8C70]">Tukarkan poin Anda langsung dengan menghubungi Sales Representative <strong>{matchedCustomer.salesmanName}</strong> ({matchedCustomer.area}) saat kunjungan!</p>
                  
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 mt-1">
                    {REWARDS_CATALOG.map(r => {
                      const canAfford = (matchedCustomer.points || 0) >= r.pointsCost;
                      return (
                        <div key={r.id} className="bg-white p-2 border border-[#E5E5DF] rounded-xl flex items-center justify-between text-xs gap-2">
                          <div className="flex flex-col min-w-0">
                            <span className="font-bold text-[#4A4A3C] truncate">{r.name}</span>
                            <span className="text-[9px] text-[#8C8C70] truncate">Sponsor: {r.sponsor}</span>
                          </div>
                          <span className={`text-[10px] font-mono px-2 py-1 rounded font-black shrink-0 ${
                            canAfford ? "bg-emerald-100 text-emerald-800" : "bg-gray-100 text-gray-500"
                          }`}>
                            {r.pointsCost} Poin
                          </span>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </motion.div>
            );
          })()}
        </div>
      )}

      {/* SUBTAB 2: OUTLET MANDIRI REGISTRATION */}
      {customerActiveSubTab === "register" && (
        <div className="max-w-xl mx-auto bg-[#FAF9F6] border border-[#E5E5DF] rounded-3xl p-6 shadow-md relative overflow-hidden text-left animate-fadeIn">
          <div className="absolute top-0 right-0 py-0.5 px-3 bg-[#5A5A40] text-[8px] font-black text-white rounded-bl-xl uppercase tracking-widest">
            FREE 50 POIN STARTER
          </div>
          
          <div className="mb-4">
            <div className="flex justify-between items-center">
              <h3 className="text-sm font-black text-[#4A4A3C] uppercase tracking-wider">
                Formulir Pendaftaran Profiling Mandiri Pelanggan
              </h3>
              <button
                onClick={() => fileInputRef.current?.click()}
                className="flex items-center gap-1.5 px-3 py-1.5 bg-[#5A5A40] text-white text-[10px] font-bold rounded-lg cursor-pointer hover:bg-[#4A4A3C]"
              >
                <Upload className="w-3 h-3" />
                Impor .xlsx
              </button>
              <input type="file" ref={fileInputRef} accept=".xlsx, .xls" onChange={handleImportXLSX} className="hidden" />
            </div>
            <p className="text-[11px] text-[#8C8C70] mt-0.5">
              Bergabunglah dengan program loyalitas pedagang retail DKR untuk mendapatkan akses potongan harga eksklusif dan instentif hadiah sponsor.
            </p>
          </div>

          <form onSubmit={handleCustomerSelfRegister} className="space-y-4">
            <div>
              <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                Nama Toko / Outlet Anda *:
              </label>
              <input
                type="text"
                required
                placeholder="Contoh: TOKO BERKAH ABADI"
                value={selfRegForm.name}
                onChange={(e) => setSelfRegForm({ ...selfRegForm, name: e.target.value })}
                className="w-full bg-white border border-[#E5E5DF] rounded-xl px-3 py-2 text-xs font-semibold focus:outline-hidden focus:ring-1 focus:ring-[#5A5A40]"
              />
            </div>

            <div>
              <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                Alamat Lengkap Outlet *:
              </label>
              <textarea
                required
                rows={2}
                placeholder="Contoh: Jl. Sudirman No. 45 Semarang Tengah"
                value={selfRegForm.address}
                onChange={(e) => setSelfRegForm({ ...selfRegForm, address: e.target.value })}
                className="w-full bg-white border border-[#E5E5DF] rounded-xl p-3 text-xs font-semibold focus:outline-hidden focus:ring-1 focus:ring-[#5A5A40]"
              />
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                  Pilih Sales Wilayah Anda:
                </label>
                <select
                  value={selfRegForm.salesmanName}
                  onChange={(e) => {
                    const matchedSales = salesmen.find(s => s.name === e.target.value);
                    setSelfRegForm({
                      ...selfRegForm,
                      salesmanName: e.target.value,
                      area: matchedSales?.area || selfRegForm.area
                    });
                  }}
                  className="w-full bg-white border border-[#E5E5DF] rounded-xl px-2.5 py-2 text-xs font-semibold focus:outline-hidden"
                >
                  <option value="">-- Pilih Salesman --</option>
                  {salesmen.map(s => (
                    <option key={s.id} value={s.name}>
                      {s.name} ({s.area})
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                  Jenis & Sektor Toko:
                </label>
                <select
                  value={selfRegForm.jenisToko}
                  onChange={(e) => setSelfRegForm({ ...selfRegForm, jenisToko: e.target.value })}
                  className="w-full bg-white border border-[#E5E5DF] rounded-xl px-2.5 py-2 text-xs font-semibold focus:outline-hidden"
                >
                  <option value="Sembako">Sembako</option>
                  <option value="Kelontong">Kelontong Kecil</option>
                  <option value="Kelontong Grosir">Kelontong Grosir Utama</option>
                  <option value="Pengecer">Pengecer Keliling</option>
                  <option value="Lainnya">Lainnya</option>
                </select>
              </div>
            </div>

            <div className="border-t border-dashed border-[#E5E5DF] pt-4">
              <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                Estimasi Kas Belanja Outlet Per Bulan (Rupiah):
              </label>
              
              <div className="flex items-center gap-3">
                <input
                  type="range"
                  min="2000000"
                  max="20000000"
                  step="500000"
                  value={selfRegForm.estimatedOmzet}
                  onChange={(e) => setSelfRegForm({ ...selfRegForm, estimatedOmzet: e.target.value })}
                  className="flex-1 accent-[#5A5A40] h-1.5 bg-[#E5E5DF] rounded-lg cursor-pointer"
                />
                <span className="font-mono text-xs font-black text-[#5A5A40] bg-[#FAF9F6] border border-[#E5E5DF] px-2.5 py-1 rounded-lg shrink-0">
                  Rp {Number(selfRegForm.estimatedOmzet).toLocaleString("id-ID")}
                </span>
              </div>

              {/* Live Formula Preview of Tier level based on estimated purchase omzet */}
              {(() => {
                const val = Number(selfRegForm.estimatedOmzet);
                let starsCount = "★ Bronze Tier";
                let tClr = "text-orange-900 bg-orange-50 border-orange-200";

                if (val >= 15000000) {
                  tClr = "text-purple-800 bg-purple-50 border-purple-200";
                  starsCount = "★★★ Platinum Tier";
                } else if (val >= 8000000) {
                  tClr = "text-amber-900 bg-amber-50 border-amber-200";
                  starsCount = "★★ Gold Tier";
                } else if (val >= 5000000) {
                  tClr = "text-blue-800 bg-blue-50 border-blue-200";
                  starsCount = "★ Silver Tier";
                }

                return (
                  <div className={`mt-2.5 border p-2.5 rounded-xl text-[11px] flex items-center justify-between ${tClr} transition-all duration-150`}>
                    <div>
                      Estimasi Tier Anda: <strong>{starsCount}</strong>
                    </div>
                    <span className="font-black bg-white/65 px-1.5 py-0.5 rounded uppercase text-[8.5px] tracking-wider border">
                      +50 Starter Poin
                    </span>
                  </div>
                );
              })()}
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                  Status Sewa Bangunan:
                </label>
                <select
                  value={selfRegForm.ownership}
                  onChange={(e) => setSelfRegForm({ ...selfRegForm, ownership: e.target.value })}
                  className="w-full bg-white border border-[#E5E5DF] rounded-xl px-2.5 py-1.5 text-xs text-[#4A4A3C]"
                >
                  <option value="Milik Sendiri">Milik Sendiri</option>
                  <option value="Sewa Kontrak">Sewa Kontrak Bulanan</option>
                  <option value="Milik Keluarga">Milik Keluarga/Warisan</option>
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                  Lama Toko Berdiri (Tahun):
                </label>
                <input
                  type="number"
                  min={1}
                  max={60}
                  value={selfRegForm.storeAgeYears}
                  onChange={(e) => setSelfRegForm({ ...selfRegForm, storeAgeYears: e.target.value })}
                  className="w-full bg-white border border-[#E5E5DF] rounded-xl px-2.5 py-1.5 text-xs text-[#4A4A3C]"
                />
              </div>
            </div>

            <button
              type="submit"
              disabled={isCustomerSelfRegistering}
              className="w-full bg-[#5A5A40] hover:bg-[#4A4A3C] text-white hover:text-white font-black text-xs uppercase tracking-wider py-4 rounded-xl transition cursor-pointer flex items-center justify-center gap-1.5 shadow-md shadow-[#5A5A40]/10"
            >
              {isCustomerSelfRegistering ? (
                <>
                  <Loader2 className="w-4 h-4 animate-spin shrink-0" />
                  Sedang Menyimpan...
                </>
              ) : (
                <>
                  🪙 DAFTAR OUTLET MANDIRI & HUBUNGKAN SPREADSHEET
                </>
              )}
            </button>
          </form>
        </div>
      )}

      {/* Customer-facing registration success dialog overlay */}
      {registerSuccessName && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-[#4A4A3C]/45 backdrop-blur-xs">
          <motion.div
            initial={{ scale: 0.95, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-white rounded-3xl p-6 w-full max-w-sm border border-emerald-300 shadow-2xl relative text-center"
          >
            <div className="w-16 h-16 bg-emerald-100 rounded-full flex items-center justify-center mx-auto mb-4 border border-emerald-300">
              <Check className="w-8 h-8 text-emerald-600 font-extrabold" />
            </div>
            <h4 className="text-base font-extrabold text-[#4A4A3C] uppercase tracking-wider">
              REGISTRASI SUKSES!
            </h4>
            <p className="text-xs text-[#8C8C70] mt-2 mb-4 leading-relaxed">
              Mitra Toko <strong>{registerSuccessName.toUpperCase()}</strong> berhasil tersimpan & tersinkronisasi ke Google Sheets secara real-time!
            </p>
            <div className="bg-emerald-50 text-emerald-950 border border-emerald-100 p-3 rounded-2xl mb-5 text-xs font-bold leading-relaxed flex flex-col gap-0.5 text-left">
              <span>🎁 Bonus Starter Voucher: 50 Poin DKR</span>
              <span className="text-[10px] text-emerald-800 font-normal">Gunakan poin ini untuk klaim hadiah sponsor kami saat sales mengunjungi toko Anda.</span>
            </div>
            <button
              onClick={() => setRegisterSuccessName("")}
              className="w-full bg-[#5A5A40] text-white hover:bg-[#4A4A3C] font-semibold text-xs py-3 rounded-xl transition uppercase cursor-pointer"
            >
              Lihat Kartu Poin Saya
            </button>
          </motion.div>
        </div>
      )}
    </div>
  );
};
