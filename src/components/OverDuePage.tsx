import React, { useState } from 'react';
import { motion } from 'motion/react';
import { Upload, FileText, Search, Download, Clock, MapPin, Phone, ShieldAlert, ArrowDownUp, TrendingUp, AlertTriangle, Users } from 'lucide-react';
import * as XLSX from 'xlsx';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, CartesianGrid } from 'recharts';

export interface OverdueInvoice {
  tanggal: string;
  nomor: string;
  jt: string;
  idPelanggan: string;
  pelanggan: string;
  alamat: string;
  kota: string;
  total: number;
  piutang: number;
  umur: number;
  salesFaktur: string;
  tanggalBayarTerakhir: string;
}

export function OverDuePage({ sheetsScriptUrl }: { sheetsScriptUrl?: string }) {
  const [invoices, setInvoices] = useState<OverdueInvoice[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [sortBy, setSortBy] = useState<'umurDesc' | 'umurAsc' | 'piutangDesc' | 'piutangAsc' | 'default'>('umurDesc');
  const [errorMsg, setErrorMsg] = useState('');
  const [isSyncing, setIsSyncing] = useState(false);
  const [isFetching, setIsFetching] = useState(false);
  const [syncStatus, setSyncStatus] = useState<'idle' | 'success' | 'error'>('idle');

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setErrorMsg('');
    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        const text = evt.target?.result as string;
        const lines = text.split('\n');

        let headerRowIndex = -1;

        // 1. Analyze Fixed Width columns from the header row
        let columnPositions: { name: string, start: number, end?: number }[] = [];

        for (let i = 0; i < Math.min(100, lines.length); i++) {
          const line = lines[i];
          const lower = line.toLowerCase();
          
          if (lower.includes('tanggal') && lower.includes('nomor #') && lower.includes('id pelanggan')) {
            headerRowIndex = i;
            
            // Extract column positions based on multiple spaces
            const colRegex = /([A-Za-z0-9#()\/ ]+?)(?:\s{2,}|$)/g;
            let match;
            while ((match = colRegex.exec(line)) !== null) {
              if (match[1].trim()) {
                columnPositions.push({
                   name: match[1].trim().toLowerCase(),
                   start: match.index
                });
              }
            }
            
            // Set end positions
            for (let j = 0; j < columnPositions.length - 1; j++) {
               columnPositions[j].end = columnPositions[j+1].start;
            }
            break;
          }
        }

        if (headerRowIndex === -1 || columnPositions.length === 0) {
          setErrorMsg('Format tabel TXT tidak dikenali. Gagal menemukan baris header (Nomor #, ID Pelanggan, dsb). Pastikan file yang diunggah adalah hasil export faktur overdue TXT Accurate yang valid.');
          if (e.target) e.target.value = '';
          return;
        }

        const parsedInvoices: OverdueInvoice[] = [];

        for (let i = headerRowIndex + 1; i < lines.length; i++) {
          const line = lines[i];
          if (!line.trim() || line.includes('====') || line.toLowerCase().includes('per tgl') || line.toLowerCase().includes('faktur overdue') || line.toLowerCase().includes('cabang :')) continue;

          // Helper to get value by fixed column width
          const getVal = (possibleNames: string[]) => {
            let col = columnPositions.find(c => possibleNames.includes(c.name));
            if (!col) {
              col = columnPositions.find(c => possibleNames.some(name => c.name.includes(name.toLowerCase())));
            }
            if (!col) return '';
            if (col.end) {
              return line.substring(col.start, col.end).trim();
            } else {
              return line.substring(col.start).trim();
            }
          };
          
          const parseNumForValue = (val: string) => {
             if (!val) return 0;
             const str = val.replace(/[^0-9,-]/g, '').replace(',', '.');
             return parseFloat(str) || 0;
          };

          const pelanggan = getVal(['pelanggan']);
          const totalVal = getVal(['total']);
          
          if (!pelanggan && !totalVal) continue; // Skip empty rows
          if (pelanggan.toLowerCase().includes('total')) continue; // Skip total row

          parsedInvoices.push({
            tanggal: getVal(['tanggal']) || '',
            nomor: getVal(['nomor #', 'nomor cpi', 'nomor']) || '',
            jt: getVal(['jt', 'jatuh tempo']) || '',
            idPelanggan: getVal(['id pelanggan']) || '',
            pelanggan: pelanggan,
            alamat: getVal(['alamat']) || '',
            kota: getVal(['kota']) || '',
            total: parseNumForValue(getVal(['total'])),
            piutang: parseNumForValue(getVal(['piutang'])),
            umur: parseNumForValue(getVal(['umur'])),
            salesFaktur: getVal(['sales faktur', 'salesman']) || '',
            tanggalBayarTerakhir: getVal(['tanggal bayar', 'tgl bayar']) || '',
          });
        }

        if (parsedInvoices.length === 0) {
          setErrorMsg('Berhasil membaca file TXT, namun tidak ada data faktur overdue yang ditemukan. Pastikan tabel di dalam file tersebut berisi data.');
        } else {
          setInvoices(parsedInvoices);
        }
      } catch (err) {
        console.error('Error importing:', err);
        setErrorMsg('Terjadi kesalahan sistem saat membaca file TXT. Format file mungkin tidak didukung atau corrupt.');
      }
      if (e.target) e.target.value = '';
    };

    reader.readAsText(file);
  };
  
  const generateDashboardData = () => {
    let totalOutstanding = 0;
    let totalOverdue = 0;

    let agingObj = {
      current: 0,
      d1_30: 0,
      d31_60: 0,
      d61_90: 0,
      d91_plus: 0
    };

    let salesMap: Record<string, { totalOD: number, count: number }> = {};
    let customerMap: Record<string, number> = {};

    invoices.forEach(inv => {
      totalOutstanding += inv.piutang;

      const isOD = inv.umur > 0;
      if (isOD) {
        totalOverdue += inv.piutang;
      }

      if (inv.umur <= 0) agingObj.current += inv.piutang;
      else if (inv.umur <= 30) agingObj.d1_30 += inv.piutang;
      else if (inv.umur <= 60) agingObj.d31_60 += inv.piutang;
      else if (inv.umur <= 90) agingObj.d61_90 += inv.piutang;
      else agingObj.d91_plus += inv.piutang;

      if (isOD) {
        const salesName = inv.salesFaktur || "Unknown";
        if (!salesMap[salesName]) salesMap[salesName] = { totalOD: 0, count: 0 };
        salesMap[salesName].totalOD += inv.piutang;
        salesMap[salesName].count += 1;

        const custName = inv.pelanggan;
        if (!customerMap[custName]) customerMap[custName] = 0;
        customerMap[custName] += inv.piutang;
      }
    });

    const agingData = [
      { name: "Current (Belum JT)", value: agingObj.current, color: "#64748b" },
      { name: "1-30 Hr OD", value: agingObj.d1_30, color: "#eab308" },
      { name: "31-60 Hr OD", value: agingObj.d31_60, color: "#f97316" },
      { name: "61-90 Hr OD", value: agingObj.d61_90, color: "#ef4444" },
      { name: ">90 Hr OD", value: agingObj.d91_plus, color: "#991b1b" }
    ];

    const topSales = Object.entries(salesMap)
      .map(([name, data]) => ({ name, ...data }))
      .sort((a, b) => b.totalOD - a.totalOD);

    const topCustomers = Object.entries(customerMap)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 10);

    const odRatio = totalOutstanding > 0 ? (totalOverdue / totalOutstanding) * 100 : 0;

    return { totalOutstanding, totalOverdue, odRatio, agingData, topSales, topCustomers };
  };

  const getSortedAndFilteredInvoices = () => {
     let result = [...invoices];
     
     if (searchTerm) {
        const lowerTerm = searchTerm.toLowerCase();
        result = result.filter(inv => 
           inv.pelanggan.toLowerCase().includes(lowerTerm) ||
           inv.idPelanggan.toLowerCase().includes(lowerTerm) ||
           inv.nomor.toLowerCase().includes(lowerTerm) ||
           inv.salesFaktur.toLowerCase().includes(lowerTerm)
        );
     }
     
     switch (sortBy) {
        case 'umurDesc':
           result.sort((a, b) => b.umur - a.umur);
           break;
        case 'umurAsc':
           result.sort((a, b) => a.umur - b.umur);
           break;
        case 'piutangDesc':
           result.sort((a, b) => b.piutang - a.piutang);
           break;
        case 'piutangAsc':
           result.sort((a, b) => a.piutang - b.piutang);
           break;
     }
     
     return result;
  };
  
  const handleSyncToSheets = async () => {
    if (!sheetsScriptUrl) {
      setErrorMsg("Tolong hubungkan dan masukkan URL Google Sheets Web App di pengaturan!");
      return;
    }
    setErrorMsg('');
    setIsSyncing(true);
    setSyncStatus('idle');
    try {
      const response = await fetch(sheetsScriptUrl, {
        method: "POST",
        mode: "cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "syncOverdue",
          data: invoices
        }),
      });
      const result = await response.json();
      if (result.success) {
        setSyncStatus('success');
        setTimeout(() => setSyncStatus('idle'), 3000);
      } else {
        setSyncStatus('error');
        setErrorMsg("Gagal sinkronisasi data ke Google Sheets.");
      }
    } catch (err) {
      console.error(err);
      setSyncStatus('error');
      setErrorMsg("Terjadi kesalahan jaringan.");
    } finally {
      setIsSyncing(false);
    }
  };

  const handleFetchFromSheets = async () => {
    if (!sheetsScriptUrl) {
      setErrorMsg("Tolong hubungkan dan masukkan URL Google Sheets Web App di pengaturan!");
      return;
    }
    setErrorMsg('');
    setIsFetching(true);
    try {
      const response = await fetch(sheetsScriptUrl, {
        method: "POST",
        mode: "cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "getOverdue"
        }),
      });
      const result = await response.json();
      if (result.success && result.data) {
        setInvoices(result.data);
      } else {
        setErrorMsg("Gagal mengambil data dari Google Sheets.");
      }
    } catch (err) {
      console.error(err);
      setErrorMsg("Terjadi kesalahan jaringan.");
    } finally {
      setIsFetching(false);
    }
  };

  const displayedInvoices = getSortedAndFilteredInvoices();
  const dashboardData = generateDashboardData();

  return (
    <div className="space-y-6">
      <div className="bg-[#ffffff] p-6 rounded-3xl border border-[#e2e8f0]/60 shadow-xs relative overflow-hidden flex flex-col md:flex-row md:items-center justify-between gap-6">
        <div className="absolute top-0 right-0 w-64 h-64 bg-rose-500/5 rounded-full blur-3xl -mr-10 -mt-10 pointer-events-none"></div>
        <div>
          <h2 className="text-xl md:text-2xl font-black text-[#0f172a] uppercase tracking-wider flex items-center gap-2">
            <Clock className="w-5 h-5 text-rose-600" />
            Laporan Overdue
          </h2>
          <p className="text-xs text-[#64748b] mt-2 max-w-xl pr-4">
            Impor data tagihan overdue dari hasil eksport Accurate Online (format .txt) untuk mengelola penagihan piutang dan umur faktur pelanggan.
          </p>
        </div>

        <div className="relative z-10 flex flex-wrap gap-3 mt-4 md:mt-0">
          <label className="bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold uppercase tracking-wider px-6 py-3.5 rounded-xl transition cursor-pointer shadow-sm hover:shadow-md flex items-center justify-center gap-2">
            <Upload className="w-4 h-4" />
            Impor File Accurate
            <input
              type="file"
              accept=".txt"
              className="hidden"
              onChange={handleFileUpload}
            />
          </label>
          {sheetsScriptUrl && invoices.length > 0 && (
            <button
              onClick={handleSyncToSheets}
              disabled={isSyncing}
              className="bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white text-xs font-bold uppercase tracking-wider px-6 py-3.5 rounded-xl transition shadow-sm hover:shadow-md flex items-center justify-center gap-2"
            >
              <FileText className="w-4 h-4" />
              {isSyncing ? "Menyimpan..." : "Simpan ke DB"}
            </button>
          )}
          {sheetsScriptUrl && (
             <button
              onClick={handleFetchFromSheets}
              disabled={isFetching}
              className="bg-white border border-[#e2e8f0] text-[#0f172a] hover:bg-gray-50 text-xs font-bold uppercase tracking-wider px-6 py-3.5 rounded-xl transition shadow-sm hover:shadow-md flex items-center justify-center gap-2"
             >
               <Download className="w-4 h-4 text-[#64748b]" />
               {isFetching ? "Menarik..." : "Tarik DB"}
             </button>
          )}
        </div>
      </div>
      
      {errorMsg && (
        <div className="bg-red-50 border border-red-200 text-red-700 p-4 rounded-2xl text-sm font-medium">
          {errorMsg}
        </div>
      )}
      
      {invoices.length > 0 && (
         <div className="mt-8">
           <h3 className="text-lg font-black text-[#0f172a] uppercase tracking-wider mb-4 flex items-center gap-2">
              <TrendingUp className="w-5 h-5 text-indigo-600" />
              Collection Health Dashboard
           </h3>
           
           <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
              <div className="bg-[#ffffff] p-5 rounded-2xl border border-[#e2e8f0]/60 shadow-xs flex flex-col gap-1 relative overflow-hidden">
                 <div className="absolute -right-3 -bottom-3 text-indigo-500/10">
                   <FileText className="w-20 h-20" />
                 </div>
                 <span className="text-[10px] items-center gap-1 text-[#64748b] font-bold uppercase tracking-wider z-10 flex">
                   Total Outstanding (Piutang)
                 </span>
                 <span className="text-xl font-black text-indigo-700 z-10 font-mono">
                   Rp {dashboardData.totalOutstanding.toLocaleString("id-ID")}
                 </span>
              </div>
              
              <div className="bg-[#ffffff] p-5 rounded-2xl border border-[#e2e8f0]/60 shadow-xs flex flex-col gap-1 relative overflow-hidden">
                 <div className="absolute -right-3 -bottom-3 text-rose-500/10">
                   <AlertTriangle className="w-20 h-20" />
                 </div>
                 <span className="text-[10px] items-center gap-1 text-[#64748b] font-bold uppercase tracking-wider z-10 flex">
                   Total Overdue (&gt; 0 Hari)
                 </span>
                 <span className="text-xl font-black text-rose-700 z-10 font-mono">
                   Rp {dashboardData.totalOverdue.toLocaleString("id-ID")}
                 </span>
              </div>

              <div className="bg-[#ffffff] p-5 rounded-2xl border border-[#e2e8f0]/60 shadow-xs flex flex-col gap-1 relative overflow-hidden">
                 <div className="absolute -right-3 -bottom-3 text-amber-500/10">
                   <TrendingUp className="w-20 h-20" />
                 </div>
                 <span className="text-[10px] items-center gap-1 text-[#64748b] font-bold uppercase tracking-wider z-10 flex">
                   Overdue Ratio
                 </span>
                 <span className={`text-xl font-black z-10 ${dashboardData.odRatio < 20 ? 'text-emerald-600' : dashboardData.odRatio < 35 ? 'text-amber-500' : 'text-rose-600'}`}>
                   {dashboardData.odRatio.toFixed(1)}%
                 </span>
                 <div className="w-full bg-gray-100 h-1.5 mt-2 rounded-full overflow-hidden z-10">
                    <div 
                       className={`h-full ${dashboardData.odRatio < 20 ? 'bg-emerald-500' : dashboardData.odRatio < 35 ? 'bg-amber-500' : 'bg-rose-500'}`} 
                       style={{ width: `${Math.min(dashboardData.odRatio, 100)}%` }}
                    />
                 </div>
              </div>
              
              <div className="bg-[#ffffff] p-5 rounded-2xl border border-[#e2e8f0]/60 shadow-xs flex flex-col gap-1 relative overflow-hidden">
                 <div className="absolute -right-3 -bottom-3 text-emerald-500/10">
                   <ShieldAlert className="w-20 h-20" />
                 </div>
                 <span className="text-[10px] items-center gap-1 text-[#64748b] font-bold uppercase tracking-wider z-10 flex">
                   Total Faktur Terdata
                 </span>
                 <span className="text-xl font-black text-[#0f172a] z-10">
                   {invoices.length} Faktur
                 </span>
              </div>
           </div>

           <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">
              {/* Chart Aging */}
              <div className="bg-[#ffffff] p-6 rounded-3xl border border-[#e2e8f0]/60 shadow-xs flex flex-col">
                 <h3 className="text-sm font-black text-[#0f172a] uppercase tracking-wider mb-6">Aging Piutang Chart</h3>
                 <div className="h-64 w-full min-h-[256px]">
                    <ResponsiveContainer width="100%" height={256}>
                       <BarChart data={dashboardData.agingData.filter(d => d.value > 0)}>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b', fontWeight: 'bold' }} />
                          <YAxis 
                             axisLine={false} 
                             tickLine={false} 
                             tick={{ fontSize: 10, fill: '#64748b' }}
                             tickFormatter={(val) => `Rp ${(val / 1000000).toFixed(0)} Jt`}
                          />
                          <Tooltip 
                             cursor={{ fill: '#f8fafc' }}
                             contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                             formatter={(value: any) => [`Rp ${Number(value).toLocaleString("id-ID")}`, "Total"]}
                          />
                          <Bar dataKey="value" radius={[4, 4, 0, 0]}>
                             {dashboardData.agingData.filter(d => d.value > 0).map((entry, index) => (
                                <Cell key={`cell-${index}`} fill={entry.color} />
                             ))}
                          </Bar>
                       </BarChart>
                    </ResponsiveContainer>
                 </div>
              </div>

              {/* Ranking Sales */}
              <div className="bg-[#ffffff] p-6 rounded-3xl border border-[#e2e8f0]/60 shadow-xs flex flex-col">
                 <h3 className="text-sm font-black text-[#0f172a] uppercase tracking-wider mb-4 flex items-center justify-between">
                    <span>Ranking Sales berdasarkan Overdue</span>
                    <Users className="w-4 h-4 text-[#64748b]" />
                 </h3>
                 <div className="flex-1 overflow-y-auto pr-2" style={{ maxHeight: '250px' }}>
                    <div className="space-y-3">
                       {dashboardData.topSales.map((sales, idx) => (
                          <div key={idx} className="flex justify-between items-center p-3 bg-[#f8fafc] rounded-xl border border-[#e2e8f0]/40">
                             <div className="flex items-center gap-3">
                                <div className={`w-6 h-6 rounded-full flex items-center justify-center text-[10px] font-bold ${idx < 3 ? 'bg-rose-100 text-rose-700' : 'bg-gray-200 text-gray-600'}`}>
                                   {idx + 1}
                                </div>
                                <div className="flex flex-col">
                                   <span className="text-xs font-black text-[#0f172a]">{sales.name}</span>
                                   <span className="text-[10px] text-[#64748b]">{sales.count} Faktur OD</span>
                                </div>
                             </div>
                             <span className="text-sm font-mono font-bold text-rose-600">
                                Rp {sales.totalOD.toLocaleString("id-ID")}
                             </span>
                          </div>
                       ))}
                       {dashboardData.topSales.length === 0 && (
                          <div className="text-xs text-[#64748b] py-4 text-center">Tidak ada data salesman.</div>
                       )}
                    </div>
                 </div>
              </div>
              
              {/* Top 10 Customer */}
              <div className="bg-[#ffffff] p-6 rounded-3xl border border-[#e2e8f0]/60 shadow-xs lg:col-span-2">
                 <h3 className="text-sm font-black text-[#0f172a] uppercase tracking-wider mb-4">Top 10 Customer Overdue</h3>
                 <div className="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-3">
                    {dashboardData.topCustomers.map((c, idx) => (
                       <div key={idx} className="flex justify-between items-center py-2 border-b border-[#e2e8f0]/40 last:border-0">
                          <div className="flex items-center gap-2 overflow-hidden">
                             <div className="w-5 h-5 rounded bg-gray-100 text-gray-500 flex items-center justify-center text-[9px] font-bold shrink-0">{idx + 1}</div>
                             <span className="text-xs font-bold text-[#0f172a] truncate">{c.name}</span>
                          </div>
                          <span className="text-xs font-mono font-bold text-rose-600 shrink-0">
                             Rp {c.value.toLocaleString("id-ID")}
                          </span>
                       </div>
                    ))}
                    {dashboardData.topCustomers.length === 0 && (
                       <div className="text-xs text-[#64748b] py-4">Tidak ada data customer.</div>
                    )}
                 </div>
              </div>
           </div>
           
           <h3 className="text-lg font-black text-[#0f172a] uppercase tracking-wider mb-4 flex items-center gap-2">
              <FileText className="w-5 h-5 text-indigo-600" />
              Detail Faktur
           </h3>
         </div>
      )}

      {invoices.length > 0 && (
         <div className="bg-[#ffffff] rounded-3xl border border-[#e2e8f0]/60 shadow-xs overflow-hidden flex flex-col">
           <div className="flex flex-col md:flex-row p-4 border-b border-[#e2e8f0]/60 gap-4 justify-between bg-[#f8fafc]">
              <div className="relative flex-1 max-w-sm">
                <Search className="absolute left-3 top-2.5 w-4 h-4 text-[#64748b]" />
                <input
                  type="text"
                  placeholder="Cari Pelanggan, Nomor Faktur, Salesman..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="w-full bg-[#ffffff] border border-[#e2e8f0] rounded-xl pl-9 pr-4 py-2 text-xs focus:outline-hidden focus:ring-2 focus:ring-[#2563eb] focus:bg-white text-[#0f172a] font-medium"
                />
              </div>
              
              <div className="flex gap-2 relative">
                <select
                  value={sortBy}
                  onChange={(e) => setSortBy(e.target.value as any)}
                  className="bg-[#ffffff] border border-[#e2e8f0] rounded-xl pl-3 pr-8 py-2 text-xs focus:outline-hidden focus:ring-2 focus:ring-[#2563eb] focus:bg-white text-[#0f172a] font-bold appearance-none relative z-20 cursor-pointer"
                >
                  <option value="default">Default Urutan</option>
                  <option value="umurDesc">Umur (Paling Lama)</option>
                  <option value="umurAsc">Umur (Paling Baru)</option>
                  <option value="piutangDesc">Piutang (Terbesar)</option>
                  <option value="piutangAsc">Piutang (Terkecil)</option>
                </select>
                <div className="absolute right-2 top-2.5 text-[#64748b] z-20 pointer-events-none">
                   <ArrowDownUp className="w-4 h-4" />
                </div>
              </div>
           </div>
           
           <div className="overflow-x-auto p-4 md:p-6 bg-gray-50/30">
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                 {displayedInvoices.map((inv, idx) => (
                    <motion.div 
                       initial={{ opacity: 0, y: 10 }}
                       animate={{ opacity: 1, y: 0 }}
                       transition={{ delay: idx * 0.02 }}
                       key={inv.nomor + idx}
                       className="bg-[#ffffff] border border-[#e2e8f0]/80 rounded-2xl p-4 shadow-sm hover:shadow-md transition-all flex flex-col group relative"
                    >
                       <div className="flex justify-between items-start mb-3">
                          <div className="flex flex-col">
                             <div className="flex items-center gap-1.5 mb-1">
                                <span className="text-[9px] font-bold tracking-wider text-[#64748b] bg-[#e2e8f0]/40 px-2 py-0.5 rounded-sm uppercase">
                                   {inv.idPelanggan}
                                </span>
                                <span className="text-[10px] font-mono font-bold text-[#0f172a]">
                                   {inv.nomor}
                                </span>
                             </div>
                             <h3 className="text-sm font-black text-[#0f172a] uppercase tracking-wide leading-tight group-hover:text-rose-600 transition-colors">
                                {inv.pelanggan}
                             </h3>
                          </div>
                          
                          <div className={`shrink-0 flex flex-col items-center justify-center rounded-xl p-2 min-w-14 items-center 
                             ${inv.umur > 45 ? 'bg-rose-50 text-rose-700' : inv.umur > 30 ? 'bg-amber-50 text-amber-700' : 'bg-blue-50 text-blue-700'}
                          `}>
                             <span className="text-[10px] font-bold uppercase tracking-wider mb-0.5 opacity-80">Umur</span>
                             <span className="text-lg font-black leading-none">{inv.umur}</span>
                             <span className="text-[9px] font-bold uppercase opacity-80 mt-0.5">Hari</span>
                          </div>
                       </div>
                       
                       <div className="flex gap-2 items-start mb-4 bg-gray-50 p-2.5 rounded-xl border border-gray-100">
                          <MapPin className="w-3.5 h-3.5 text-[#64748b] mt-0.5 shrink-0" />
                          <div className="flex flex-col">
                             <span className="text-[10px] text-[#0f172a] font-medium leading-relaxed">
                                {inv.alamat || "-"}
                             </span>
                             <span className="text-[9px] font-bold text-[#64748b] uppercase tracking-wider mt-0.5">
                                Kota: {inv.kota || "-"}
                             </span>
                          </div>
                       </div>
                       
                       <div className="mt-auto flex flex-col gap-2 pt-2 border-t border-[#e2e8f0]/60">
                          <div className="flex justify-between items-center text-[10px]">
                             <span className="text-[#64748b] font-bold uppercase tracking-wider">Tgl Faktur</span>
                             <span className="text-[#0f172a] font-medium">{inv.tanggal || "-"}</span>
                          </div>
                          <div className="flex justify-between items-center text-[10px]">
                             <span className="text-[#64748b] font-bold uppercase tracking-wider">Jatuh Tempo</span>
                             <span className="text-[#0f172a] font-medium">{inv.jt || "-"}</span>
                          </div>
                          <div className="flex justify-between items-center text-[10px]">
                             <span className="text-[#64748b] font-bold uppercase tracking-wider">Sales Factor</span>
                             <span className="text-[#0f172a] font-bold">{inv.salesFaktur || "-"}</span>
                          </div>
                          
                          <div className="mt-2 pt-2 border-t border-dashed border-[#e2e8f0]/60 flex justify-between items-end">
                             <span className="text-[10px] font-black text-rose-800 uppercase tracking-wider">Sisa Piutang</span>
                             <span className="text-lg font-mono font-black text-rose-700">Rp {inv.piutang.toLocaleString("id-ID")}</span>
                          </div>
                       </div>
                    </motion.div>
                 ))}
                 
                 {displayedInvoices.length === 0 && (
                    <div className="col-span-full py-16 text-center text-[#64748b] font-medium">
                       Tidak ada faktur yang ditemukan dengan filter saat ini.
                    </div>
                 )}
              </div>
           </div>
         </div>
      )}
      
      {invoices.length === 0 && (
         <div className="bg-[#ffffff] rounded-3xl border border-dashed border-[#e2e8f0] p-12 flex flex-col items-center justify-center text-center">
            <div className="w-16 h-16 bg-rose-50 rounded-full flex items-center justify-center text-rose-500 mb-4">
               <FileText className="w-8 h-8" />
            </div>
            <h3 className="text-[#0f172a] font-black text-lg mb-2">Belum Ada Data Overdue</h3>
            <p className="text-[#64748b] text-xs max-w-sm mb-6 leading-relaxed">
               Silakan impor ekspor file "Faktur Overdue" dari aplikasi Accurate Online (format .txt) untuk memvisualisasikan data piutang pelanggan.
            </p>
         </div>
      )}
    </div>
  );
}
