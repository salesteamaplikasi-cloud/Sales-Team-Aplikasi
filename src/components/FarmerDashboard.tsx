import React from 'react';
import { KpiReport, SalesmanRole } from '../types';
import { TrendingUp, Users, AlertTriangle } from 'lucide-react';

interface FarmerDashboardProps {
  reports: KpiReport[];
}

export const FarmerDashboard: React.FC<FarmerDashboardProps> = ({ reports }) => {
  // Logic to identify potential stores (reports)
  // Proxy store by ID Laporan (simplified)
  
  const potentialStores = reports.filter(r => r.billsReceived + (r.billsTransfer || 0) === 0 && r.tc > 0).slice(0, 5);
  
  // Calculate average total payment to identify declining stores
  const calculateTotal = (r: KpiReport) => r.billsReceived + (r.billsTransfer || 0);
  const totalReportsCount = reports.length;
  const avgPayment = reports.length > 0 ? reports.reduce((sum, r) => sum + calculateTotal(r), 0) / totalReportsCount : 0;
  
  const decliningPerformance = reports.filter(r => calculateTotal(r) < avgPayment * 0.8).slice(0, 5);

  return (
    <div className="p-6 bg-gray-50 rounded-lg">
      <h2 className="text-2xl font-bold text-gray-900 mb-6">Analisis & Performa Toko (Aris - Farmer)</h2>
      
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        <div className="bg-white p-6 rounded-lg shadow-sm border border-gray-200">
          <div className="flex items-center gap-2 mb-4">
            <TrendingUp className="text-green-600" />
            <h3 className="text-lg font-semibold text-gray-900">Potensi Peningkatan Transaksi</h3>
          </div>
          <div className="space-y-3">
            {potentialStores.map(report => (
              <div key={report.id} className="p-3 border border-gray-100 rounded flex justify-between items-center bg-gray-50">
                <div className="flex flex-col gap-0.5">
                  <span className="font-medium text-sm text-gray-800">Laporan: {report.id} ({report.date})</span>
                  <span className="text-xs text-gray-600">Sales: {report.salesmanName} | Total Tagihan Termasuk Transfer & Giro: <strong>Rp {(report.billsReceived + (report.billsTransfer || 0) + (report.billsGiro || 0)).toLocaleString('id-ID')}</strong></span>
                </div>
                <span className="font-bold text-green-700 text-sm">Potensial</span>
              </div>
            ))}
          </div>
        </div>
        
        <div className="bg-white p-6 rounded-lg shadow-sm border border-gray-200">
          <div className="flex items-center gap-2 mb-4">
            <AlertTriangle className="text-yellow-600" />
            <h3 className="text-lg font-semibold text-gray-900">Performa Perlu Perhatian</h3>
          </div>
          <div className="space-y-3">
             {decliningPerformance.map(report => (
              <div key={report.id} className="p-3 border border-gray-100 rounded flex justify-between items-center bg-gray-50">
                <div className="flex flex-col gap-0.5">
                  <span className="font-medium text-sm text-gray-800">Laporan: {report.id} ({report.date})</span>
                  <span className="text-xs text-gray-600">Sales: {report.salesmanName} | Total Tagihan Termasuk Transfer & Giro: <strong>Rp {(report.billsReceived + (report.billsTransfer || 0) + (report.billsGiro || 0)).toLocaleString('id-ID')}</strong></span>
                </div>
                <span className="font-bold text-yellow-700 text-sm">Turun</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};
