import React, { useState } from "react";
import { CUSTOMER_DATA } from "../data/customerList";
import { Download, Search } from "lucide-react";

export function CustomerSalesTable() {
  const [searchTerm, setSearchTerm] = useState("");

  const filteredData = CUSTOMER_DATA.map(visit => ({
    ...visit,
    customers: visit.customers.filter(c =>
      c.nama.toLowerCase().includes(searchTerm.toLowerCase()) ||
      c.id.toLowerCase().includes(searchTerm.toLowerCase())
    )
  })).filter(visit => visit.customers.length > 0);

  const exportToCSV = () => {
    const headers = ["Salesman", "Kunjungan", "ID Pelanggan", "Nama", "Jalan", "Kota"];
    const rows = CUSTOMER_DATA.flatMap(v => 
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

  return (
    <div className="p-6 bg-[#FAF9F6] rounded-xl border border-[#E5E5DF]">
      <div className="flex justify-between items-center mb-6">
        <h2 className="text-xl font-bold font-serif text-[#4A4A3C]">Daftar Pelanggan & Sales</h2>
        <div className="flex gap-2">
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
            </tr>
          </thead>
          <tbody>
            {filteredData.map((visit, i) => (
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
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
