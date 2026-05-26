/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { Salesman, Product, KpiReport } from "./types";

export const INITIAL_SALESMEN: Salesman[] = [];

export const INITIAL_PRODUCTS: Product[] = [
  { id: "p-1", name: "COTTON BUD YO PIPI", category: "Baby Care", skuCode: "CB-YPP", isActive: true },
  { id: "p-2", name: "TISSUE YO PIPI POP UP", category: "Tissue", skuCode: "TJ-YPP-PU", isActive: true },
  { id: "p-3", name: "TISSUE YO PIPI 250S", category: "Tissue", skuCode: "TJ-YPP-250", isActive: true },
  { id: "p-4", name: "LED MADRIM TOPAZ", category: "Electrical/Lighting", skuCode: "LED-MT", isActive: true },
  { id: "p-5", name: "LED MADRIM SOLAR TITAN", category: "Electrical/Lighting", skuCode: "LED-MST", isActive: true },
  { id: "p-6", name: "LED MADRIM RUBY", category: "Electrical/Lighting", skuCode: "LED-MR", isActive: true }
];

export const STANDARD_CYCLES = [
  "Senin Ganjil",
  "Selasa Ganjil",
  "Rabu Ganjil",
  "Kamis Ganjil",
  "Jumat Ganjil",
  "Sabtu Ganjil",
  "Senin Genap",
  "Selasa Genap",
  "Rabu Genap",
  "Kamis Genap",
  "Jumat Genap",
  "Sabtu Genap"
];

export const INITIAL_REPORTS: KpiReport[] = [
  {
    id: "rep-1",
    salesmanId: "s-1",
    salesmanName: "RENY",
    date: "2026-05-22",
    cycle: "Rabu Genap",
    tc: 14,
    cp: 14,
    ec: 10,
    skuTotal: 24,
    operationalCost: 25000,
    billsReceived: 1850000,
    notes: "Semua toko dikunjungi, kendala cuaca hujan rintik-rintik.",
    createdAt: "2026-05-22T17:00:00Z"
  },
  {
    id: "rep-2",
    salesmanId: "s-2",
    salesmanName: "BUDI",
    date: "2026-05-22",
    cycle: "Rabu Genap",
    tc: 12,
    cp: 10,
    ec: 7,
    skuTotal: 15,
    operationalCost: 20000,
    billsReceived: 1200000,
    notes: "Toko Berkah tidak bisa dikunjungi karena tutup.",
    createdAt: "2026-05-22T17:15:00Z"
  },
  {
    id: "rep-rino-1",
    salesmanId: "s-6",
    salesmanName: "Rino",
    date: "2026-05-21",
    cycle: "Kamis Ganjil",
    tc: 16,
    cp: 16,
    ec: 8,
    skuTotal: 90,
    operationalCost: 15000,
    billsReceived: 4000000,
    notes: "Hari pertama rekap bulanan area Cilongok.",
    createdAt: "2026-05-21T17:00:00Z"
  },
  {
    id: "rep-rino-2",
    salesmanId: "s-6",
    salesmanName: "Rino",
    date: "2026-05-22",
    cycle: "Jumat Ganjil",
    tc: 16,
    cp: 16,
    ec: 7,
    skuTotal: 88,
    operationalCost: 10000,
    billsReceived: 3853585,
    notes: "Hari kedua penyelesaian audit penagihan pasar Cilongok.",
    createdAt: "2026-05-22T17:10:00Z"
  }
];

/**
 * Gets the current day status for auto-detect (e.g. "Sabtu Ganjil" or "Sabtu Genap")
 */
export function autoDetectCycle(dateString: string): string {
  const d = new Date(dateString);
  const days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  const dayName = days[d.getDay()];
  
  // Calculate is it Odd (Ganjil) or Even (Genap) week of the year
  // Simple check: get week of year
  const firstDayOfYear = new Date(d.getFullYear(), 0, 1);
  const pastDaysOfYear = (d.getTime() - firstDayOfYear.getTime()) / 86400000;
  const weekNumber = Math.ceil((pastDaysOfYear + firstDayOfYear.getDay() + 1) / 7);
  const parity = weekNumber % 2 === 0 ? "Genap" : "Ganjil";
  
  if (dayName === "Minggu") {
    return "Senin Ganjil"; // fallback
  }
  
  return `${dayName} ${parity}`;
}

/**
 * Parser utility to read plain text copier from Group Chats/Excel and extract form values.
 * Extremely flexible support for WA forms!
 */
export function parseKpiText(text: string, salesmen: Salesman[]): any {
  const result: any = {
    salesmanId: "",
    salesmanName: "",
    cycle: "",
    tc: 0,
    cp: 0,
    ec: 0,
    skuTotal: 0,
    operationalCost: 0,
    billsReceived: 0,
    notes: "",
    warnings: []
  };

  if (!text || text.trim() === "") return result;

  const lines = text.split("\n");
  
  // Regular expressions for clean semantic detection
  const tcRegex = /(?:tc|amplop|call)\s*[:|=|-]?\s*(\d+)/i;
  const cpRegex = /(?:cp|kunjungan|visit)\s*[:|=|-]?\s*(\d+)/i;
  const ecRegex = /(?:ec|order|effective)\s*[:|=|-]?\s*(\d+)/i;
  const skuRegex = /(?:sku|sku total|barang)\s*[:|=|-]?\s*(\d+)/i;
  
  // Operational cost regex (supporting Rp, dots like 25.000, etc.)
  const oprRegex = /(?:operasional|ops|b\.opr|opr|biaya)\s*[:|=|-]?\s*(?:rp\.?\s*)?([\d.,]+)/i;
  
  // Bills / collection regex (Rp, dots, etc.)
  const billRegex = /(?:tagihan|tgh|bayar|collection|didapat|rp)\s*[:|=|-]?\s*(?:rp\.?\s*)?([\d.,]+)/i;

  const cycleRegex = /(?:siklus|hari|kunjungan)\s*[:|=|-]?\s*([a-zA-Z\s]+(?:ganjil|genap|ganjil\/genap)?)/i;
  const salesmanRegex = /(?:salesman|nama|sales|salesperson)\s*[:|=|-]?\s*([a-zA-Z0-9\s_]+)/i;

  // Let's sweep line by line
  let rawNotesList: string[] = [];
  
  for (let line of lines) {
    const trimmedLine = line.trim();
    if (!trimmedLine) continue;

    // Try matches
    const tcMatch = trimmedLine.match(tcRegex);
    if (tcMatch) {
      result.tc = parseInt(tcMatch[1], 10);
      continue;
    }

    const cpMatch = trimmedLine.match(cpRegex);
    if (cpMatch) {
      result.cp = parseInt(cpMatch[1], 10);
      continue;
    }

    const ecMatch = trimmedLine.match(ecRegex);
    if (ecMatch) {
      result.ec = parseInt(ecMatch[1], 10);
      continue;
    }

    const skuMatch = trimmedLine.match(skuRegex);
    if (skuMatch) {
      result.skuTotal = parseInt(skuMatch[1], 10);
      continue;
    }

    const oprMatch = trimmedLine.match(oprRegex);
    if (oprMatch) {
      // Clean periods/commas in currency
      const cleanVal = oprMatch[1].replace(/[.,]/g, "");
      result.operationalCost = parseInt(cleanVal, 10) || 0;
      continue;
    }

    const billMatch = trimmedLine.match(billRegex);
    if (billMatch) {
      const cleanVal = billMatch[1].replace(/[.,]/g, "");
      result.billsReceived = parseInt(cleanVal, 10) || 0;
      continue;
    }

    const cycleMatch = trimmedLine.match(cycleRegex);
    if (cycleMatch) {
      result.cycle = cycleMatch[1].trim();
      continue;
    }

    const salesMatch = trimmedLine.match(salesmanRegex);
    if (salesMatch) {
      const candidateName = salesMatch[1].trim();
      // Look up in our salesman database
      const matched = salesmen.find(
        s => s.name.toLowerCase() === candidateName.toLowerCase() ||
             candidateName.toLowerCase().includes(s.name.toLowerCase()) ||
             s.name.toLowerCase().includes(candidateName.toLowerCase())
      );
      if (matched) {
        result.salesmanId = matched.id;
        result.salesmanName = matched.name;
      } else {
        result.salesmanName = candidateName; // Temporary name
        result.warnings.push(`Salesman "${candidateName}" tidak ditemukan di database. Kami menyarankannya sebagai nama baru.`);
      }
      continue;
    }

    // If it doesn't match standard lines, it might be inline format like "RENY | Rabu Genap | TC: 12 | CP: 12"
    if (trimmedLine.includes("|")) {
      const parts = trimmedLine.split("|");
      for (let part of parts) {
        const pt = part.trim();
        // Check inline parts
        const tcM = pt.match(/(?:tc|amplop)\s*[:|=]?\s*(\d+)/i);
        if (tcM) result.tc = parseInt(tcM[1], 10);
        
        const cpM = pt.match(/(?:cp|kunjungan)\s*[:|=]?\s*(\d+)/i);
        if (cpM) result.cp = parseInt(cpM[1], 10);
        
        const ecM = pt.match(/(?:ec|order)\s*[:|=]?\s*(\d+)/i);
        if (ecM) result.ec = parseInt(ecM[1], 10);
        
        const skuM = pt.match(/(?:sku)\s*[:|=]?\s*(\d+)/i);
        if (skuM) result.skuTotal = parseInt(skuM[1], 10);
        
        const oprM = pt.match(/(?:opr|operasional)\s*[:|=]?\s*([\d.]+)/i);
        if (oprM) result.operationalCost = parseInt(oprM[1].replace(/[.]/g, ""), 10) || 0;
        
        const billM = pt.match(/(?:tgh|tagihan)\s*[:|=]?\s*([\d.]+)/i);
        if (billM) result.billsReceived = parseInt(billM[1].replace(/[.]/g, ""), 10) || 0;
      }
      continue;
    }

    // Otherwise, collect as user notes (exclude headers)
    if (!trimmedLine.toLowerCase().includes("laporan sales") && 
        !trimmedLine.toLowerCase().includes("audit kpi")) {
      rawNotesList.push(trimmedLine);
    }
  }

  // If we haven't matched salesman, check if first line is just a capital name
  if (!result.salesmanId && lines.length > 0) {
    const firstLine = lines[0].trim().toUpperCase();
    if (firstLine && firstLine.length < 25 && !firstLine.includes("LAPORAN") && !firstLine.includes("TC") && !firstLine.includes("CP")) {
      const matched = salesmen.find(s => s.name.toUpperCase() === firstLine);
      if (matched) {
        result.salesmanId = matched.id;
        result.salesmanName = matched.name;
      }
    }
  }

  if (rawNotesList.length > 0) {
    result.notes = rawNotesList.join(" ");
  }

  return result;
}
