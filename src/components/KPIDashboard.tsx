import React, { useMemo } from 'react';
import { motion } from 'motion/react';
import { 
  LineChart, Line, BarChart as RechartsBarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, ComposedChart, Area
} from 'recharts';
import { LayoutDashboard, TrendingUp, AlertCircle, Calendar, BarChart2 } from 'lucide-react';
import { KpiReport } from '../types';

interface KPIDashboardProps {
  reports: KpiReport[];
}

export const KPIDashboard: React.FC<KPIDashboardProps> = ({ reports }) => {
  const chartData = useMemo(() => {
    // Group by date
    const grouped = reports.reduce((acc, curr) => {
      if (!acc[curr.date]) {
        acc[curr.date] = { date: curr.date, tc: 0, cp: 0, ec: 0, sku: 0 };
      }
      acc[curr.date].tc += curr.tc;
      acc[curr.date].cp += curr.cp;
      acc[curr.date].ec += curr.ec;
      acc[curr.date].sku += curr.skuTotal;
      return acc;
    }, {} as Record<string, { date: string; tc: number; cp: number; ec: number; sku: number }>);

    // Sort by date ascending
    const sorted = Object.values(grouped).sort((a: { date: string }, b: { date: string }) => new Date(a.date).getTime() - new Date(b.date).getTime());
    
    // Add calculated metrics: Call Plan Achievement (%), Effective Call Rate (%)
    return sorted.map((day: any) => {
      const cpPct = day.tc > 0 ? (day.cp / day.tc) * 100 : 0;
      const ecPct = day.cp > 0 ? (day.ec / day.cp) * 100 : 0;
      return {
        ...day,
        cpPct: parseFloat(cpPct.toFixed(1)),
        ecPct: parseFloat(ecPct.toFixed(1)),
        displayDate: new Date(day.date).toLocaleDateString('id-ID', { day: 'numeric', month: 'short' })
      };
    });
  }, [reports]);

  const latestStats = useMemo(() => {
    let tcSum = 0, cpSum = 0, ecSum = 0, skuSum = 0;
    reports.forEach(r => { tcSum += r.tc; cpSum += r.cp; ecSum += r.ec; skuSum += r.skuTotal; });
    const cpPct = tcSum > 0 ? (cpSum / tcSum) * 100 : 0;
    const ecPct = cpSum > 0 ? (ecSum / cpSum) * 100 : 0;
    const targetSku = Math.round(ecSum * 15);
    const skuPct = targetSku > 0 ? (skuSum / targetSku) * 100 : 0;
    
    const isCpLayak = cpPct >= 80;
    const isEcLayak = ecPct >= 40;
    const isSkuLayak = skuPct >= 100;
    
    const avgSkuPerSales = reports.length > 0 ? Math.round(skuSum / reports.length) : 0;
    
    return {
      tcSum, cpSum, ecSum, skuSum,
      cpPct: cpPct.toFixed(1),
      ecPct: ecPct.toFixed(1),
      skuPct: skuPct.toFixed(1),
      isCpLayak,
      isEcLayak,
      isSkuLayak,
      avgSkuPerSales,
      totalDays: chartData.length
    };
  }, [reports, chartData]);

  return (
    <motion.div
      initial={{ opacity: 0, y: 15 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -15 }}
      className="space-y-6"
    >
      <div className="bg-[#ffffff] rounded-3xl p-6 border border-[#e2e8f0] shadow-sm relative overflow-hidden">
        <h2 className="text-xl md:text-2xl font-serif italic font-extrabold text-[#0f172a] tracking-tight flex items-center gap-2 mb-2">
          <LayoutDashboard className="w-6 h-6 shrink-0 text-[#2563eb]" />
          EXECUTIVE DASHBOARD
        </h2>
        <p className="text-xs text-[#64748b] max-w-2xl leading-relaxed">
          Menampilkan sekilas tren kesehatan dan performa sales secara menyeluruh berbasis ISO standards. Indikasi dari metrik ini dirancang sebagai panduan fundamental bagi Supervisory dan Management.
        </p>

        {/* Global Summary Cards */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mt-6">
          <div className="bg-emerald-50 rounded-2xl p-4 border border-emerald-100">
            <span className="block text-[10px] font-bold text-emerald-800 uppercase tracking-widest mb-1">TOTAL HARI LAPOR</span>
            <span className="text-2xl font-serif font-black text-emerald-900">{latestStats.totalDays}</span>
          </div>
          <div className="bg-[#f8fafc] rounded-2xl p-4 border border-[#e2e8f0]">
            <span className="block text-[10px] font-bold text-[#64748b] uppercase tracking-widest mb-1">
              RATA-RATA SKU / SALES
            </span>
            <div className="flex items-end gap-2">
              <span className="text-2xl font-serif font-black text-[#0f172a]">
                {latestStats.avgSkuPerSales}
              </span>
              <span className="text-[10px] font-bold text-[#64748b] mb-1">
                (Daily)
              </span>
            </div>
          </div>
          <div className={`rounded-2xl p-4 border ${latestStats.isCpLayak ? 'bg-emerald-50 border-emerald-100' : 'bg-rose-50 border-rose-100'}`}>
            <span className={`block text-[10px] font-bold uppercase tracking-widest mb-1 ${latestStats.isCpLayak ? 'text-emerald-800' : 'text-rose-800'}`}>
              RATA-RATA CALL PLAN
            </span>
            <div className="flex items-end gap-2">
              <span className={`text-2xl font-serif font-black ${latestStats.isCpLayak ? 'text-emerald-900' : 'text-rose-900'}`}>
                {latestStats.cpPct}%
              </span>
              <span className={`text-[10px] font-bold mb-1 ${latestStats.isCpLayak ? 'text-emerald-700' : 'text-rose-700'}`}>
                (Std: {'>='} 80%)
              </span>
            </div>
          </div>
          <div className={`rounded-2xl p-4 border ${latestStats.isEcLayak ? 'bg-emerald-50 border-emerald-100' : 'bg-rose-50 border-rose-100'}`}>
            <span className={`block text-[10px] font-bold uppercase tracking-widest mb-1 ${latestStats.isEcLayak ? 'text-emerald-800' : 'text-rose-800'}`}>
              RATA-RATA EFFECTIVE CALL
            </span>
            <div className="flex items-end gap-2">
              <span className={`text-2xl font-serif font-black ${latestStats.isEcLayak ? 'text-emerald-900' : 'text-rose-900'}`}>
                {latestStats.ecPct}%
              </span>
              <span className={`text-[10px] font-bold mb-1 ${latestStats.isEcLayak ? 'text-emerald-700' : 'text-rose-700'}`}>
                (Std: {'>='} 40%)
              </span>
            </div>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* ISO Standard KPI Ratio Chart */}
        <div className="bg-[#ffffff] rounded-3xl p-6 border border-[#e2e8f0] shadow-sm">
          <div className="mb-6 flex justify-between items-center">
            <div>
              <h3 className="text-sm font-bold text-[#0f172a] uppercase tracking-wider flex items-center gap-2">
                <TrendingUp className="w-4 h-4 text-emerald-600" />
                Pertumbuhan Efisiensi KPI (%)
              </h3>
              <p className="text-[10px] text-[#64748b] mt-1 uppercase tracking-widest">Memantau Call Plan (CP/TC) & Effective Call (EC/CP)</p>
            </div>
          </div>
          
          <div className="h-[300px] w-full">
            <ResponsiveContainer width="100%" height={300}>
              <LineChart data={chartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis 
                  dataKey="displayDate" 
                  tick={{ fontSize: 10, fill: '#64748b', fontWeight: 600 }} 
                  axisLine={false} 
                  tickLine={false} 
                  dy={10}
                />
                <YAxis 
                  tick={{ fontSize: 10, fill: '#64748b' }} 
                  axisLine={false} 
                  tickLine={false} 
                  domain={[0, 100]}
                  tickFormatter={(val) => `${val}%`}
                />
                <Tooltip 
                  contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1)' }}
                  labelStyle={{ fontWeight: 'bold', color: '#0f172a', marginBottom: '4px' }}
                />
                <Legend iconType="circle" wrapperStyle={{ fontSize: '10px', fontWeight: 700, paddingTop: '20px' }} />
                <Line 
                  type="monotone" 
                  name="Call Plan (Std: 80%)" 
                  dataKey="cpPct" 
                  stroke="#10b981" 
                  strokeWidth={3} 
                  activeDot={{ r: 6 }} 
                />
                <Line 
                  type="monotone" 
                  name="Effective Call (Std: 40%)" 
                  dataKey="ecPct" 
                  stroke="#f43f5e" 
                  strokeWidth={3} 
                  activeDot={{ r: 6 }} 
                />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Volume Trend Chart */}
        <div className="bg-[#ffffff] rounded-3xl p-6 border border-[#e2e8f0] shadow-sm">
          <div className="mb-6 flex justify-between items-center">
            <div>
              <h3 className="text-sm font-bold text-[#0f172a] uppercase tracking-wider flex items-center gap-2">
                <BarChart2 className="w-4 h-4 text-[#2563eb]" />
                Volume Absolut Sales
              </h3>
              <p className="text-[10px] text-[#64748b] mt-1 uppercase tracking-widest">Tren Total Kunjungan (TC), Produktif (CP), Pesanan (EC)</p>
            </div>
          </div>
          
          <div className="h-[300px] w-full">
            <ResponsiveContainer width="100%" height={300}>
              <ComposedChart data={chartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis 
                  dataKey="displayDate" 
                  tick={{ fontSize: 10, fill: '#64748b', fontWeight: 600 }} 
                  axisLine={false} 
                  tickLine={false} 
                  dy={10}
                />
                <YAxis 
                  tick={{ fontSize: 10, fill: '#64748b' }} 
                  axisLine={false} 
                  tickLine={false} 
                />
                <Tooltip 
                  contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1)' }}
                  labelStyle={{ fontWeight: 'bold', color: '#0f172a', marginBottom: '4px' }}
                />
                <Legend iconType="circle" wrapperStyle={{ fontSize: '10px', fontWeight: 700, paddingTop: '20px' }} />
                <Bar dataKey="tc" name="Total Call (TC)" fill="#e2e8f0" radius={[4, 4, 0, 0]} />
                <Bar dataKey="cp" name="Call Plan (CP)" fill="#3b82f6" radius={[4, 4, 0, 0]} />
                <Line type="monotone" name="Effective Call (EC)" dataKey="ec" stroke="#f59e0b" strokeWidth={3} />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>
    </motion.div>
  );
};
