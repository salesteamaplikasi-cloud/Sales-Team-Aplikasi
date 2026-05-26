/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface RewardMerchant {
  id: string;
  name: string;
  description: string;
  pointsRequired: number;
}

export interface CatalogHadiah {
  id: string;
  name: string;
  sponsor: string;
  pointsValue: number;
}

export interface Salesman {
  id: string;
  name: string;
  phone?: string;
  area?: string;
  isActive: boolean;
  createdAt: string;
}

export interface Product {
  id: string;
  name: string;
  category?: string;
  skuCode?: string;
  isActive: boolean;
}

export interface KpiReportProduct {
  productId: string;
  productName: string;
}

export interface KpiReport {
  id: string;
  salesmanId: string;
  salesmanName: string;
  date: string; // YYYY-MM-DD
  cycle: string; // e.g., "Senin Ganjil", "Rabu Genap", etc.
  tc: number; // TC (Amplop)
  cp: number; // CP (Kunjungan)
  ec: number; // EC (Order)
  skuTotal: number; // SKU Total
  operationalCost: number; // Biaya Operasional
  billsReceived: number; // Tagihan Bayar Tunai
  billsTransfer?: number; // Tagihan Bayar Transfer
  billsGiro?: number; // Tagihan Giro
  productsDetail?: KpiReportProduct[]; // Optional dynamic product details
  notes?: string;
  createdAt: string;
}

export interface ImportParsingResult {
  success: boolean;
  matchedSalesmanName?: string;
  matchedSalesmanId?: string;
  cycle?: string;
  tc?: number;
  cp?: number;
  ec?: number;
  skuTotal?: number;
  operationalCost?: number;
  billsReceived?: number;
  billsTransfer?: number;
  billsGiro?: number;
  rawTextUsed: string;
  warnings: string[];
}
