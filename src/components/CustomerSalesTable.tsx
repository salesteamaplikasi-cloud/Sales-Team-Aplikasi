import React, { useState } from "react";
import { CUSTOMER_DATA } from "../data/customerList";
import { Download, Search, Upload, Edit, X, Plus, Trash2 } from "lucide-react";
import type { VisitSchedule } from "../data/customerList";

export function CustomerSalesTable() {
  const [searchTerm, setSearchTerm] = useState("");
  const [customerData, setCustomerData] = useState(CUSTOMER_DATA);
  const [importMessage, setImportMessage] = useState<string | null>(null);

  // Edit State
  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [editFormData, setEditFormData] = useState<VisitSchedule | null>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const text = event.target?.result as string;
      const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

      let newData = [...customerData];
      const existingIds = new Set(newData.flatMap(v => v.customers.map(c => c.id)));
      let addedCount = 0;
      let duplicateCount = 0;

      lines.forEach(line => {
        const cols = line.split(/\t| {2,}/).map(c => c.trim()).filter(Boolean);
        
        if (cols.length >= 6) {
          const [id, nama, jalan, kota, salesman, ...kunjunganParts] = cols;
          const kunjungan = kunjunganParts.join(" ");

          if (id.toLowerCase().includes("id") || id.toLowerCase().includes("pelanggan")) return;

          if (!existingIds.has(id)) {
            existingIds.add(id);
            addedCount++;

            const newCustomer = { id, nama, jalan, kota };
            const existingScheduleIndex = newData.findIndex(v => v.salesman === salesman && v.kunjungan === kunjungan);

            if (existingScheduleIndex >= 0) {
              newData[existingScheduleIndex] = {
                ...newData[existingScheduleIndex],
                customers: [...newData[existingScheduleIndex].customers, newCustomer]
              };
            } else {
              newData.push({
                salesman,
                kunjungan,
                customers: [newCustomer]
              });
            }
          } else {
            duplicateCount++;
          }
        }
      });

      setCustomerData(newData);
      setImportMessage(`Impor selesai! Ditambahkan: ${addedCount} pelanggan baru. Diabaikan: ${duplicateCount} data duplikat.`);
      setTimeout(() => setImportMessage(null), 5000);
      e.target.value = "";
    };
    reader.readAsText(file);
  };

  const filteredData = customerData.map(visit => ({
    ...visit,
    customers: visit.customers.filter(c =>
      c.nama.toLowerCase().includes(searchTerm.toLowerCase()) ||
      c.id.toLowerCase().includes(searchTerm.toLowerCase())
    )
  })).filter(visit => visit.customers.length > 0);

  const exportToCSV = () => {
    const headers = ["Salesman", "Kunjungan", "ID Pelanggan", "Nama", "Jalan", "Kota"];
    const rows = customerData.flatMap(v => 
      v.customers.map(c => [v.salesman, v.kunjungan, c.id, c.nama, c.jalan, c.kota])
    );
    const csvContent = [headers, ...rows].map(r => r.join(",")).join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "data_pelanggan_sales.csv";
    link.click();
  };

  const handleSaveEdit = () => {
    if (editingIndex !== null && editFormData) {
      const newData = [...customerData];
      newData[editingIndex] = editFormData;
      setCustomerData(newData);
      setEditingIndex(null);
      setEditFormData(null);
    }
  };

  const handleCustomerEditChange = (customerIndex: number, field: string, value: string) => {
    if (editFormData) {
      const updatedCustomers = [...editFormData.customers];
      updatedCustomers[customerIndex] = { ...updatedCustomers[customerIndex], [field]: value };
      setEditFormData({ ...editFormData, customers: updatedCustomers });
    }
  };

  const handleDeleteCustomer = (customerIndex: number) => {
    if (editFormData) {
      const updatedCustomers = editFormData.customers.filter((_, i) => i !== customerIndex);
      setEditFormData({ ...editFormData, customers: updatedCustomers });
    }
  };

  return (
    <div className="relative p-6 bg-[#FAF9F6] rounded-xl border border-[#E5E5DF]">
      {/* Edit Modal */}
      {editingIndex !== null && editFormData && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden">
            <div className="p-4 border-b border-[#E5E5DF] flex justify-between items-center bg-[#FAF9F6]">
              <h3 className="text-xl font-bold font-serif text-[#4A4A3C]">Edit Data Kunjungan</h3>
              <button onClick={() => setEditingIndex(null)} className="text-[#8C8C70] hover:text-black">
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="p-6 overflow-y-auto flex-1">
              <div className="grid grid-cols-2 gap-4 mb-6">
                <div>
                  <label className="block text-xs font-bold text-[#8C8C70] mb-1">Salesman</label>
                  <input
                    type="text"
                    value={editFormData.salesman}
                    onChange={(e) => setEditFormData({ ...editFormData, salesman: e.target.value })}
                    className="w-full p-2 border border-[#E5E5DF] rounded focus:ring-2 focus:ring-[#5A5A40] outline-none"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-[#8C8C70] mb-1">Hari Kunjungan & Minggu</label>
                  <input
                    type="text"
                    value={editFormData.kunjungan}
                    onChange={(e) => setEditFormData({ ...editFormData, kunjungan: e.target.value })}
                    className="w-full p-2 border border-[#E5E5DF] rounded focus:ring-2 focus:ring-[#5A5A40] outline-none"
                  />
                </div>
              </div>

              <div className="mb-2 flex justify-between items-center">
                <h4 className="font-bold text-[#4A4A3C]">Daftar Pelanggan ({editFormData.customers.length})</h4>
                <button
                  onClick={() => setEditFormData({
                    ...editFormData,
                    customers: [...editFormData.customers, { id: "", nama: "", jalan: "", kota: "" }]
                  })}
                  className="flex items-center gap-1 text-xs bg-[#5A5A40] text-white px-2 py-1.5 rounded hover:bg-[#4A4A3C]"
                >
                  <Plus className="w-4 h-4" /> Tambah Pelanggan
                </button>
              </div>
              
              <div className="border border-[#E5E5DF] rounded-xl overflow-hidden">
                <table className="w-full text-left text-sm text-[#4A4A3C]">
                  <thead className="bg-[#E5E5DF]/20">
                    <tr>
                      <th className="p-2 w-1/4">ID</th>
                      <th className="p-2 w-1/4">Nama</th>
                      <th className="p-2 w-1/3">Jalan</th>
                      <th className="p-2">Kota</th>
                      <th className="p-2 w-10 text-center">Aksi</th>
                    </tr>
                  </thead>
                  <tbody>
                    {editFormData.customers.map((c, idx) => (
                      <tr key={idx} className="border-t border-[#E5E5DF]">
                        <td className="p-2">
                          <input type="text" value={c.id} onChange={(e) => handleCustomerEditChange(idx, 'id', e.target.value)} className="w-full p-1.5 border border-[#E5E5DF] rounded text-xs" />
                        </td>
                        <td className="p-2">
                          <input type="text" value={c.nama} onChange={(e) => handleCustomerEditChange(idx, 'nama', e.target.value)} className="w-full p-1.5 border border-[#E5E5DF] rounded text-xs" />
                        </td>
                        <td className="p-2">
                          <input type="text" value={c.jalan} onChange={(e) => handleCustomerEditChange(idx, 'jalan', e.target.value)} className="w-full p-1.5 border border-[#E5E5DF] rounded text-xs" />
                        </td>
                        <td className="p-2">
                          <input type="text" value={c.kota} onChange={(e) => handleCustomerEditChange(idx, 'kota', e.target.value)} className="w-full p-1.5 border border-[#E5E5DF] rounded text-xs" />
                        </td>
                        <td className="p-2 text-center">
                          <button onClick={() => handleDeleteCustomer(idx)} className="text-red-500 hover:text-red-700">
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="p-4 border-t border-[#E5E5DF] bg-[#FAF9F6] flex justify-end gap-3">
              <button 
                onClick={() => setEditingIndex(null)}
                className="px-4 py-2 rounded-lg font-bold text-[#8C8C70] hover:text-[#4A4A3C]"
              >
                Batal
              </button>
              <button 
                onClick={handleSaveEdit}
                className="px-4 py-2 bg-[#5A5A40] text-white rounded-lg font-bold hover:bg-[#4A4A3C]"
              >
                Simpan Perubahan
              </button>
            </div>
          </div>
        </div>
      )}

      {importMessage && (
        <div className="mb-4 bg-[#5A5A40] text-white p-3 rounded-lg text-sm flex items-center justify-between">
          <span>{importMessage}</span>
          <button onClick={() => setImportMessage(null)} className="font-bold hover:text-gray-300">×</button>
        </div>
      )}
      <div className="flex justify-between items-center mb-6">
        <h2 className="text-xl font-bold font-serif text-[#4A4A3C]">Daftar Pelanggan & Sales</h2>
        <div className="flex gap-2">
          <input 
            type="file" 
            id="import-upload" 
            className="hidden" 
            accept=".txt,.csv,.tsv" 
            onChange={handleFileUpload} 
          />
          <button 
            onClick={() => document.getElementById('import-upload')?.click()} 
            className="flex items-center gap-2 bg-[#FAF9F6] text-[#4A4A3C] border border-[#E5E5DF] hover:bg-[#E5E5DF]/50 px-4 py-2 rounded-lg text-sm font-semibold transition-colors"
          >
            <Upload className="w-4 h-4" /> Import TXT
          </button>
          
          <div className="relative">
            <Search className="absolute left-3 top-2.5 w-4 h-4 text-[#8C8C70]" />
            <input
              type="text"
              placeholder="Cari pelanggan..."
              className="pl-9 pr-4 py-2 rounded-lg border border-[#E5E5DF] text-sm"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          <button onClick={exportToCSV} className="flex items-center gap-2 bg-[#5A5A40] text-white px-4 py-2 rounded-lg text-sm font-semibold">
            <Download className="w-4 h-4" /> Export CSV
          </button>
        </div>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-left text-sm text-[#4A4A3C]">
          <thead className="bg-[#E5E5DF]/20 border-b border-[#E5E5DF]">
            <tr>
              <th className="p-3 w-1/4">Salesman</th>
              <th className="p-3 w-1/4">Kunjungan</th>
              <th className="p-3 w-1/2">Daftar Toko Dikunjungi (Dropdown)</th>
              <th className="p-3 w-16 text-center">Edit</th>
            </tr>
          </thead>
          <tbody>
            {filteredData.map((visit, i) => {
              // Find original index for editing
              const originalIndex = customerData.findIndex(v => v.salesman === visit.salesman && v.kunjungan === visit.kunjungan);

              return (
              <tr key={`${visit.salesman}-${visit.kunjungan}-${i}`} className="border-b border-[#E5E5DF]/50 hover:bg-[#E5E5DF]/10 align-top">
                <td className="p-3 font-semibold">{visit.salesman}</td>
                <td className="p-3">
                  <span className="bg-[#E5E5DF]/50 px-2 py-1 rounded text-xs font-bold text-[#4A4A3C]">
                    {visit.kunjungan}
                  </span>
                </td>
                <td className="p-3">
                  <select 
                    className="w-full p-2 bg-white border border-[#E5E5DF] rounded cursor-pointer outline-none focus:ring-2 focus:ring-[#5A5A40]/50 text-xs text-[#4A4A3C]"
                    defaultValue=""
                  >
                    <option value="" disabled>
                      ({visit.customers.length}) Lihat Toko...
                    </option>
                    {visit.customers.map((c) => (
                      <option key={c.id} value={c.id}>
                        [{c.id}] {c.nama} - {c.jalan}, {c.kota}
                      </option>
                    ))}
                  </select>
                </td>
                <td className="p-3 text-center">
                  <button 
                    onClick={() => {
                      setEditingIndex(originalIndex);
                      setEditFormData(JSON.parse(JSON.stringify(customerData[originalIndex])));
                    }}
                    className="p-1.5 bg-[#E5E5DF] text-[#4A4A3C] hover:bg-[#5A5A40] hover:text-white rounded transition-colors inline-flex justify-center items-center"
                    title="Edit Data"
                  >
                    <Edit className="w-4 h-4" />
                  </button>
                </td>
              </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
