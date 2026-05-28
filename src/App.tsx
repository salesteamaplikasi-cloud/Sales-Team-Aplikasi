/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from "react";
import {
  User,
  Users,
  Plus,
  Send,
  Edit2,
  Trash2,
  UserPlus,
  FileText,
  Check,
  Clipboard,
  AlertTriangle,
  Calendar,
  Sparkles,
  Database,
  History,
  TrendingUp,
  X,
  Search,
  CheckCircle2,
  Download,
  Copy,
  Info,
  DollarSign,
  Briefcase,
  Layers,
  ShoppingBag,
  Menu,
  ChevronLeft,
  ChevronRight,
  Crown,
  Gift,
  Award,
  MapPin,
  Boxes,
  RefreshCw,
  ExternalLink,
  Clock,
  CalendarDays,
  CalendarRange,
  LogOut,
  Target
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";

import { RewardModal } from "./components/RewardModal";
import { Salesman, Product, KpiReport, ImportParsingResult, RewardMerchant, CatalogHadiah, SalesmanGoal, NooRecord } from "./types";
import * as XLSX from "xlsx";
import {
  INITIAL_SALESMEN,
  INITIAL_PRODUCTS,
  STANDARD_CYCLES,
  INITIAL_REPORTS,
  autoDetectCycle,
  parseKpiText
} from "./data";

import { CustomerLoyaltyPortal } from "./components/CustomerLoyaltyPortal";
import { CustomerSalesTable } from "./components/CustomerSalesTable";
import { FarmerDashboard } from "./components/FarmerDashboard";

const APPS_SCRIPT_CODE_STENCIL = `function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // --- HELPER UNTUK MENDAPATKAN ATAU MEMBUAT SHEET + STYLING HEADER ---
    function getOrCreateGlobalSheet(ss, name, headers, headerColor) {
      var s = ss.getSheetByName(name);
      if (!s) {
        s = ss.insertSheet(name);
      }
      if (s.getLastRow() === 0) {
        s.appendRow(headers);
        var headerRange = s.getRange(1, 1, 1, headers.length);
        headerRange.setFontWeight("bold");
        headerRange.setBackground(headerColor || "#5A5A40");
        headerRange.setFontColor("#FAF9F6");
        headerRange.setHorizontalAlignment("center");
        s.setRowHeight(1, 28);
      }
      return s;
    }

    function autoResizeColumns(sheet, maxCols) {
      for (var col = 1; col <= maxCols; col++) {
        sheet.autoResizeColumn(col);
      }
    }

    // AKSI 1: Tes Koneksi
    if (data.action === "test") {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        message: "Koneksi ke Google Sheets via Portal Audit berhasil diverifikasi!" 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // AKSI 2: Tambah Laporan KPI Tunggal
    if (data.action === "addReport") {
      var sheet = getOrCreateGlobalSheet(ss, "Laporan KPI Sales", [
        "ID Laporan", "Tanggal KPI", "Nama Salesman", "Siklus", "Area",
        "TC (Amplop)", "CP (Kunjungan)", "EC (Order)", "SKU Total", 
        "Tagihan Bayar Tunai", "Tagihan Bayar Transfer", "Tagihan Giro", "Biaya Operasional (Rp)", "Catatan", 
        "Tanggal Dibuat", "Rincian SKU Produk"
      ], "#5A5A40");
      
      var rep = data.report;
      var productsStr = "";
      if (rep.productsDetail && rep.productsDetail.length > 0) {
        productsStr = rep.productsDetail.map(function(p) {
          return p.productName;
        }).join(", ");
      }
      
      sheet.appendRow([
        rep.id, rep.date, rep.salesmanName, rep.cycle, rep.area,
        rep.tc, rep.cp, rep.ec, rep.skuTotal,
        rep.billsReceived, rep.billsTransfer || 0, rep.billsGiro || 0, rep.operationalCost, rep.notes || "",
        rep.createdAt, productsStr
      ]);
      
      autoResizeColumns(sheet, 15);
      
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        message: "Laporan harian " + rep.salesmanName + " ("+ rep.date +") sukses dimasukkan!" 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // AKSI 3: Sinkronisasi Laporan KPI Masal (Bulk)
    if (data.action === "syncAll") {
      var headers = [
        "ID Laporan", "Tanggal KPI", "Nama Salesman", "Siklus", "Area",
        "TC (Amplop)", "CP (Kunjungan)", "EC (Order)", "SKU Total", 
        "Tagihan Bayar Tunai", "Tagihan Bayar Transfer", "Tagihan Giro", "Biaya Operasional (Rp)", "Catatan", 
        "Tanggal Dibuat", "Rincian SKU Produk"
      ];
      var sheet = getOrCreateGlobalSheet(ss, "Laporan KPI Sales", headers, "#5A5A40");
      
      // Update header baris pertama menyelaraskan dengan kolom terbaru
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      var reps = data.reports;
      if (sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
      
      for (var i = 0; i < reps.length; i++) {
        var rep = reps[i];
        var productsStr = "";
        if (rep.productsDetail && rep.productsDetail.length > 0) {
          productsStr = rep.productsDetail.map(function(p) {
            return p.productName;
          }).join(", ");
        }
        
        sheet.appendRow([
          rep.id, rep.date, rep.salesmanName, rep.cycle,
          rep.tc, rep.cp, rep.ec, rep.skuTotal,
          rep.billsReceived, rep.billsTransfer || 0, rep.billsGiro || 0, rep.operationalCost, rep.notes || "",
          rep.createdAt, productsStr
        ]);
      }
      
      autoResizeColumns(sheet, 15);
      
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        message: "Sukses menyinkronkan " + reps.length + " data laporan audit KPI!" 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // AKSI 4: Sinkronisasi Database Salesman
    if (data.action === "syncSalesmen") {
      var sheet = getOrCreateGlobalSheet(ss, "Daftar Salesman", [
        "ID Salesman", "Nama Salesman", "Area Wilayah", "No. HP / Telepon"
      ], "#4A4A3C");
      
      var salesmen = data.salesmen;
      if (sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
      
      for (var i = 0; i < salesmen.length; i++) {
        var s = salesmen[i];
        sheet.appendRow([
          s.id, s.name, s.area, s.phone || "-"
        ]);
      }
      
      autoResizeColumns(sheet, 4);
      
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        message: "Berhasil menyinkronkan " + salesmen.length + " data Salesman ke tab '" + sheet.getName() + "'"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // AKSI 5: Sinkronisasi Database SKU Produk
    if (data.action === "syncProducts") {
      var sheet = getOrCreateGlobalSheet(ss, "Daftar Produk SKU", [
        "IDsku", "Nama SKU Produk", "Kategori", "SKU Code"
      ], "#3B3D2A");
      
      var products = data.products;
      if (sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
      
      for (var i = 0; i < products.length; i++) {
        var p = products[i];
        sheet.appendRow([
          p.id, p.name, p.category, p.skuCode || "-"
        ]);
      }
      
      autoResizeColumns(sheet, 4);
      
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        message: "Berhasil menyinkronkan " + products.length + " data SKU Produk ke tab '" + sheet.getName() + "'"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // AKSI 6: Sinkronisasi Program Loyalti & Profiling Toko
    if (data.action === "syncLoyalty") {
      // Setup satu Sheet tunggal sesuai tab manual yang di buat user
      var sheetLoyalty = getOrCreateGlobalSheet(ss, "Program Loyalti Toko & Klaim", [
        "Tipe Data", "ID / ID Log", "Nama Toko / Outlet", "Alamat / Detail Item", 
        "Sales PJ / Tanggal", "Area / Poin Terpakai", "Jenis / Status Kerja", 
        "Estimasi Omzet (Rp)", "Nota per Hari", "Umur Gabung (Th)", 
        "Bangunan", "Poin Tersisa", "Tingkat Tier", "Tanggal Profiling"
      ], "#4A4238");

      var customers = data.customers;
      var redeems = data.redeems;

      // Bersihkan baris data lama (kecuali header ke-1)
      if (sheetLoyalty.getLastRow() > 1) {
        sheetLoyalty.deleteRows(2, sheetLoyalty.getLastRow() - 1);
      }

      var totalRowsAdded = 0;

      // 1. Ekspor Profiling Toko + Tindakan Follow-Up nya
      for (var i = 0; i < customers.length; i++) {
        var c = customers[i];
        
        // Baris Profiling Toko
        sheetLoyalty.appendRow([
          "PROFILING TOKO",
          c.id,
          c.name,
          c.address || "-",
          c.salesmanName,
          c.area,
          c.jenisToko,
          c.estimatedOmzet,
          c.notesPerDay || 0,
          c.storeAgeYears || 0,
          c.ownership,
          c.points || 0,
          c.tier,
          c.createdAt || "-"
        ]);
        totalRowsAdded++;

        // Baris-baris Follow up dari toko ini
        if (c.actionsLog && c.actionsLog.length > 0) {
          for (var j = 0; j < c.actionsLog.length; j++) {
            var log = c.actionsLog[j];
            sheetLoyalty.appendRow([
              "TINDAKAN FOLLOW UP",
              "-",
              c.name,
              log.notes || "-",
              log.date,
              "-",
              log.action,
              "-",
              "-",
              "-",
              "-",
              "-",
              "-",
              "-"
            ]);
            totalRowsAdded++;
          }
        }
      }

      // 2. Ekspor Log Klaim Reward Merchant
      for (var i = 0; i < redeems.length; i++) {
        var r = redeems[i];
        sheetLoyalty.appendRow([
          "KLAIM REWARD",
          r.id,
          r.customerName,
          r.rewardName,
          r.date,
          r.pointsSpent,
          r.status || "Berhasil",
          "-",
          "-",
          "-",
          "-",
          "-",
          "-",
          "-"
        ]);
        totalRowsAdded++;
      }

      autoResizeColumns(sheetLoyalty, 14);

      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        message: "Berhasil menyinkronkan " + customers.length + " toko ke tab '" + sheetLoyalty.getName() + "' (Total: " + totalRowsAdded + " baris data terkirim!)"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // AKSI 7: Tarik Data Laporan KPI Sales untuk halaman KPI Sales
    if (data.action === "getReports") {
      var sheet = ss.getSheetByName("Laporan KPI Sales");
      if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({ 
          success: true, 
          reports: [] 
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ 
          success: true, 
          reports: [] 
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var rawData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      var reportsList = [];
      for (var i = 0; i < rawData.length; i++) {
        var rowVal = rawData[i];
        var item = {};
        for (var j = 0; j < headers.length; j++) {
          var headerKey = headers[j];
          var cellVal = rowVal[j];
          if (cellVal instanceof Date) {
            var yr = cellVal.getFullYear();
            var mo = ("0" + (cellVal.getMonth() + 1)).slice(-2);
            var dy = ("0" + cellVal.getDate()).slice(-2);
            item[headerKey] = yr + "-" + mo + "-" + dy;
          } else {
            item[headerKey] = cellVal;
          }
        }
        reportsList.push(item);
      }
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        reports: reportsList 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // AKSI 8: Impor & Sinkronisasi Respon Google Form
    // (V3) Update matching: Fokus pada header mengandung 'toko' atau 'outlet' saja agar terhindar dari bentrok dengan 'nama pemilik'
    if (data.action === "importFromGoogleForm") {
      var formSheet = ss.getSheetByName("Jawaban Formulir 1") || ss.getSheetByName("Form Responses 1") || ss.getSheetByName("Form Responses");
      if (!formSheet) {
        return ContentService.createTextOutput(JSON.stringify({ 
          success: false, 
          message: "Sheet respon Google Form ('Jawaban Formulir 1' atau 'Form Responses 1') tidak ditemukan di Spreadsheet Anda. Pastikan Form sudah di-link-kan ke Sheets!" 
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      var lastRow = formSheet.getLastRow();
      if (lastRow <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ 
          success: true, 
          responses: [],
          message: "Sheet terdeteksi, namun belum ada respon pengisian dari pelanggan masuk." 
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      var headers = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
      var rawData = formSheet.getRange(2, 1, lastRow - 1, formSheet.getLastColumn()).getValues();
      var importedList = [];
      
      for (var i = 0; i < rawData.length; i++) {
        var rowVal = rawData[i];
        var item = {
          name: "",
          address: "",
          salesmanName: "",
          jenisToko: "Sembako",
          estimatedOmzet: 5000000,
          ownership: "Milik Sendiri",
          storeAgeYears: 2,
          timestamp: ""
        };
        
        // Mengambil data secara statis berdasarkan urutan kolom (A=0, B=1, C=2, D=3... dll)
        item.timestamp = (rowVal[0] !== undefined && rowVal[0] !== null) ? rowVal[0].toString() : "";
        item.name = (rowVal[1] !== undefined && rowVal[1] !== null) ? rowVal[1].toString().trim() : "";
        item.address = (rowVal[2] !== undefined && rowVal[2] !== null) ? rowVal[2].toString().trim() : "";
        item.salesmanName = (rowVal[3] !== undefined && rowVal[3] !== null) ? rowVal[3].toString().trim() : "";
        
        item.jenisToko = (rowVal[4] !== undefined && rowVal[4] !== null) ? rowVal[4].toString().trim() : "Sembako";
        
        var omzetVal = (rowVal[5] !== undefined && rowVal[5] !== null) ? rowVal[5].toString() : "";
        var num = parseFloat(omzetVal.replace(/[^0-9]/g, ""));
        item.estimatedOmzet = isNaN(num) ? 5000000 : num;
        
        var ageVal = (rowVal[6] !== undefined && rowVal[6] !== null) ? rowVal[6].toString() : "";
        var sage = parseInt(ageVal.replace(/[^0-9]/g, ""), 10);
        item.storeAgeYears = isNaN(sage) ? 2 : sage;
        
        item.ownership = (rowVal[7] !== undefined && rowVal[7] !== null) ? rowVal[7].toString().trim() : "Milik Sendiri";
        
        if (item.name) {
          importedList.push(item);
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        responses: importedList,
        message: "Berhasil memuat " + importedList.length + " respon pendaftaran dari Google Form!" 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // AKSI 9: Tarik Data Salesman
    if (data.action === "getSalesmen") {
      var sheet = ss.getSheetByName("Daftar Salesman");
      if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({ success: true, salesmen: [] })).setMimeType(ContentService.MimeType.JSON);
      }
      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ success: true, salesmen: [] })).setMimeType(ContentService.MimeType.JSON);
      }
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var rawData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      var salesmenList = [];
      for (var i = 0; i < rawData.length; i++) {
        var rowVal = rawData[i];
        var item = {};
        for (var j = 0; j < headers.length; j++) {
          item[headers[j]] = rowVal[j];
        }
        salesmenList.push(item);
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, salesmen: salesmenList })).setMimeType(ContentService.MimeType.JSON);
    }

    // AKSI 10: Tarik Data Produk
    if (data.action === "getProducts") {
      var sheet = ss.getSheetByName("Daftar Produk SKU") || ss.getSheetByName("Daftar Produk");
      if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({ success: true, products: [] })).setMimeType(ContentService.MimeType.JSON);
      }
      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ success: true, products: [] })).setMimeType(ContentService.MimeType.JSON);
      }
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var rawData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      var productsList = [];
      for (var i = 0; i < rawData.length; i++) {
        var rowVal = rawData[i];
        var item = {};
        for (var j = 0; j < headers.length; j++) {
          item[headers[j]] = rowVal[j];
        }
        productsList.push(item);
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, products: productsList })).setMimeType(ContentService.MimeType.JSON);
    }

    // AKSI 11: Sinkronisasi NOO Salesman (New Outlet Opening)
    if (data.action === "syncNoo") {
      var sheetNoo = getOrCreateGlobalSheet(ss, "Log NOO Salesman", [
        "ID Log", "Tanggal Log", "Nama Salesman", "ID Salesman", 
        "Warung / Toko Kelontong", "Store / Toko Modern", 
        "Kios Atap", "Grosir / Wholesaler", "Total Outlet Baru"
      ], "#C01C42");
      
      var nooLogs = data.nooLogs || [];
      
      if (sheetNoo.getLastRow() > 1) {
        sheetNoo.deleteRows(2, sheetNoo.getLastRow() - 1);
      }
      
      for (var i = 0; i < nooLogs.length; i++) {
        var n = nooLogs[i];
        var warungVal = Number(n.warung || 0);
        var storeVal = Number(n.store || 0);
        var kioskVal = Number(n.kiosk || 0);
        var wholesalerVal = Number(n.wholesaler || 0);
        var totalHarian = warungVal + storeVal + kioskVal + wholesalerVal;
        
        sheetNoo.appendRow([
          n.id,
          n.date || "",
          n.salesmanName || "",
          n.salesmanId || "",
          warungVal,
          storeVal,
          kioskVal,
          wholesalerVal,
          totalHarian
        ]);
      }
      
      autoResizeColumns(sheetNoo, 9);
      
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        message: "Berhasil menyinkronkan " + nooLogs.length + " data log New Outlet Opening (NOO) ke tab '" + sheetNoo.getName() + "'"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // AKSI 12: Tarik Data NOO Salesman
    if (data.action === "getNoo") {
      var sheetNoo = ss.getSheetByName("Log NOO Salesman");
      if (!sheetNoo) {
        return ContentService.createTextOutput(JSON.stringify({ success: true, nooLogs: [] })).setMimeType(ContentService.MimeType.JSON);
      }
      var lastRow = sheetNoo.getLastRow();
      if (lastRow <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ success: true, nooLogs: [] })).setMimeType(ContentService.MimeType.JSON);
      }
      var headers = sheetNoo.getRange(1, 1, 1, sheetNoo.getLastColumn()).getValues()[0];
      var rawData = sheetNoo.getRange(2, 1, lastRow - 1, sheetNoo.getLastColumn()).getValues();
      var nooList = [];
      for (var i = 0; i < rawData.length; i++) {
        var rowVal = rawData[i];
        var item = {};
        for (var j = 0; j < headers.length; j++) {
          var headerKey = headers[j];
          var cellVal = rowVal[j];
          if (cellVal instanceof Date) {
            var yr = cellVal.getFullYear();
            var mo = ("0" + (cellVal.getMonth() + 1)).slice(-2);
            var dy = ("0" + cellVal.getDate()).slice(-2);
            item[headerKey] = yr + "-" + mo + "-" + dy;
          } else {
            item[headerKey] = cellVal;
          }
        }
        nooList.push(item);
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, nooLogs: nooList })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      message: "Operasi atau aksi '" + data.action + "' tidak dikenali." 
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      message: "Terjadi kesalahan di server Apps Script: " + err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}`;

export default function App() {
  // --- AUTHENTICATION STATE ---
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(() => {
    return localStorage.getItem("KPI_IS_AUTHENTICATED") === "true";
  });
  const [loginEmail, setLoginEmail] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");

  // --- STATE LIST MANAGEMENT ---
  const [salesmen, setSalesmen] = useState<Salesman[]>([]);
  const [salesmanGoals, setSalesmanGoals] = useState<SalesmanGoal[]>([]);
  const [products, setProducts] = useState<Product[]>([]);
  const [reports, setReports] = useState<KpiReport[]>([]);
  
  // Collapsible Sidebar & Mobile navigation states
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState<boolean>(() => {
    return localStorage.getItem("KPI_SIDEBAR_COLLAPSED") === "true";
  });
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState<boolean>(false);

  // Active Tab
  const [activeTab, setActiveTab] = useState<"form" | "salesmen" | "products" | "reports" | "sheets" | "loyalty" | "kpisales" | "claims" | "customerData">(() => {
    try {
      const p = new URLSearchParams(window.location.search);
      if (p.get("mode") === "customer-loyalty" || p.get("view") === "loyalty") {
        return "loyalty";
      }
    } catch (_) {}
    return "form";
  });

  // State for fetched reports from Google Sheets to show on KPI Sales page
  const [fetchedReports, setFetchedReports] = useState<any[]>(() => {
    try {
      const stored = localStorage.getItem("KPI_FETCHED_REPORTS_FROM_SHEETS");
      if (stored) return JSON.parse(stored);
    } catch (_) {}
    return [];
  });
  const [isFetchingReports, setIsFetchingReports] = useState<boolean>(false);
  const [lastFetchTime, setLastFetchTime] = useState<string>(() => {
    return localStorage.getItem("KPI_LAST_FETCH_TIME") || "";
  });

  // KPI Sales page filters
  const [kpiTabMode, setKpiTabMode] = useState<"dashboard" | "goals">("dashboard");
  const [goalSelectedMonth, setGoalSelectedMonth] = useState<string>(new Date().toISOString().substring(0, 7));
  const [kpiSalesFilter, setKpiSalesFilter] = useState<string>("ALL");
  const [kpiCycleFilter, setKpiCycleFilter] = useState<string>("ALL");
  const [kpiTimeFrame, setKpiTimeFrame] = useState<"all" | "daily" | "weekly" | "monthly">("all");
  const [kpiSelectedDate, setKpiSelectedDate] = useState<string>("ALL");
  const [kpiSelectedWeek, setKpiSelectedWeek] = useState<string>("ALL");
  const [kpiSelectedMonth, setKpiSelectedMonth] = useState<string>("ALL");

  // New states for tracking specific salesman tabs (all, Aris, Imam) and NOO
  const [kpiSalesmanTab, setKpiSalesmanTab] = useState<"all" | "Aris" | "Imam">("all");
  const [nooRecords, setNooRecords] = useState<NooRecord[]>(() => {
    try {
      const stored = localStorage.getItem("KPI_DB_NOO_LOGS");
      if (stored) return JSON.parse(stored);
    } catch (_) {}
    return [
      {
        id: "noo-aris-1",
        salesmanId: "s-7",
        salesmanName: "Aris",
        date: "2026-05-25",
        warung: 3,
        store: 1,
        kiosk: 2,
        wholesaler: 0
      },
      {
        id: "noo-aris-2",
        salesmanId: "s-7",
        salesmanName: "Aris",
        date: "2026-05-26",
        warung: 2,
        store: 2,
        kiosk: 1,
        wholesaler: 1
      },
      {
        id: "noo-imam-1",
        salesmanId: "s-8",
        salesmanName: "Imam",
        date: "2026-05-25",
        warung: 1,
        store: 1,
        kiosk: 3,
        wholesaler: 0
      },
      {
        id: "noo-imam-2",
        salesmanId: "s-8",
        salesmanName: "Imam",
        date: "2026-05-26",
        warung: 4,
        store: 0,
        kiosk: 0,
        wholesaler: 1
      }
    ];
  });

  // Local input states for daily NOO logging form
  const [newNooDate, setNewNooDate] = useState<string>(new Date().toISOString().substring(0, 10));
  const [newNooWarung, setNewNooWarung] = useState<number>(0);
  const [newNooStore, setNewNooStore] = useState<number>(0);
  const [newNooKiosk, setNewNooKiosk] = useState<number>(0);
  const [newNooWholesaler, setNewNooWholesaler] = useState<number>(0);

  // Synchronize NOO with local storage
  useEffect(() => {
    localStorage.setItem("KPI_DB_NOO_LOGS", JSON.stringify(nooRecords));
  }, [nooRecords]);
  
  // Riwayat Audit filters
  const [auditFilterStartDate, setAuditFilterStartDate] = useState<string>("");
  const [auditFilterEndDate, setAuditFilterEndDate] = useState<string>("");
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);
  const [deleteConfirmProductId, setDeleteConfirmProductId] = useState<string | null>(null);
  const [deleteConfirmCustomerId, setDeleteConfirmCustomerId] = useState<string | null>(null);

  // Customer Mode (Public self-service Portal)
  const [customerMode, setCustomerMode] = useState<boolean>(() => {
    try {
      const p = new URLSearchParams(window.location.search);
      return p.get("mode") === "customer-loyalty" || p.get("view") === "loyalty";
    } catch (_) {
      return false;
    }
  });

  const [customerSearchQuery, setCustomerSearchQuery] = useState<string>("");
  const [matchedCustomer, setMatchedCustomer] = useState<any | null>(null);
  const [registerSuccessName, setRegisterSuccessName] = useState<string>("");
  const [isCustomerSelfRegistering, setIsCustomerSelfRegistering] = useState<boolean>(false);
  const [customerActiveSubTab, setCustomerActiveSubTab] = useState<"check" | "register">("check");
  const [selfRegForm, setSelfRegForm] = useState({
    name: "",
    address: "",
    salesmanName: "",
    area: "",
    jenisToko: "Sembako",
    estimatedOmzet: "5000000",
    notesPerDay: "5",
    storeAgeYears: "2",
    ownership: "Milik Sendiri"
  });

  // Loyalty Program: Store Profiling database
  const [customers, setCustomers] = useState<any[]>(() => {
    try {
      const stored = localStorage.getItem("KPI_LOYALTY_CUSTOMERS");
      if (stored) return JSON.parse(stored);
    } catch (_) {}
    return [
      {
        id: "c-1",
        name: "Toko Makmur Jaya",
        address: "Jl. Raya Bobotsari No.123",
        salesmanName: "RIZKY",
        area: "Bobotsari",
        jenisToko: "Sembako",
        estimatedOmzet: 8500000,
        notesPerDay: 12,
        storeAgeYears: 4,
        ownership: "Milik Sendiri",
        points: 380,
        tier: "Gold",
        createdAt: "2026-05-10T10:00:00Z",
        actionsLog: [
          { date: "2026-05-15", action: "Kunjungan", notes: "Cek stok & tawarkan promo paket sembako", status: "Proses" }
        ]
      },
      {
        id: "c-2",
        name: "Toko Sejahtera",
        address: "Poros Utama Karangreja Blok B",
        salesmanName: "RENY",
        area: "Karangreja",
        jenisToko: "ATK",
        estimatedOmzet: 6200000,
        notesPerDay: 8,
        storeAgeYears: 2,
        ownership: "Sewa",
        points: 210,
        tier: "Silver",
        createdAt: "2026-05-12T11:00:00Z",
        actionsLog: [
          { date: "2026-05-18", action: "Telepon", notes: "Ingatkan limit jatuh tempo pembayaran nota lama", status: "Selesai" }
        ]
      },
      {
        id: "c-3",
        name: "UD Cahaya",
        address: "Kawasan Pasar Induk Purbalingga No.4",
        salesmanName: "BUDI",
        area: "Purbalingga",
        jenisToko: "Kelontong",
        estimatedOmzet: 25000000,
        notesPerDay: 22,
        storeAgeYears: 8,
        ownership: "Milik Sendiri",
        points: 1240,
        tier: "Platinum",
        createdAt: "2026-05-08T09:00:00Z",
        actionsLog: [
          { date: "2026-05-20", action: "Kasih Promo", notes: "Terapkan tier bonus diskon loyalty kuartalan", status: "Selesai" }
        ]
      },
      {
        id: "c-4",
        name: "Toko Rahayu",
        address: "Kios Bukateja Pojok Timur No.12",
        salesmanName: "RIZKY",
        area: "Bukateja",
        jenisToko: "Sembako / Kosmetik",
        estimatedOmzet: 4300000,
        notesPerDay: 4,
        storeAgeYears: 1,
        ownership: "Milik Sendiri",
        points: 85,
        tier: "Bronze",
        createdAt: "2026-05-14T14:00:00Z",
        actionsLog: []
      }
    ];
  });

  const [rewardMerchants, setRewardMerchants] = useState<RewardMerchant[]>(() => {
    try {
      const stored = localStorage.getItem("KPI_REWARD_MERCHANTS");
      if (stored) return JSON.parse(stored);
    } catch (_) {}
    return [
      { id: "r-1", name: "Voucher Belanja 50rb", description: "Voucher diskon di mitra DKR", pointsRequired: 500 },
      { id: "r-2", name: "Paket Sembako Starter", description: "Paket basic kebutuhan pokok", pointsRequired: 1000 }
    ];
  });

  const [katalogHadiah, setKatalogHadiah] = useState<CatalogHadiah[]>(() => {
    try {
      const stored = localStorage.getItem("KPI_KATALOG_HADIAH");
      if (stored) return JSON.parse(stored);
    } catch (_) {}
    return [
      { id: "k-1", name: "Kipas Angin Cosmos", sponsor: "Cosmos", pointsValue: 5000 },
      { id: "k-2", name: "Kompor Gas", sponsor: "Rinnai", pointsValue: 8000 }
    ];
  });

  const REWARDS_CATALOG = katalogHadiah.map(k => ({
    id: k.id,
    name: k.name,
    sponsor: k.sponsor,
    pointsCost: k.pointsValue
  }));

  // --- REWARDS CRUD HANDLERS ---
  const [rewardModal, setRewardModal] = useState<{ isOpen: boolean; type: 'merchant' | 'catalog'; item?: any }>({ isOpen: false, type: 'merchant' });

  const handleSaveReward = (type: 'merchant' | 'catalog', item: any) => {
    if (type === 'merchant') {
      if (item.id) {
        setRewardMerchants(prev => prev.map(r => r.id === item.id ? item : r));
      } else {
        setRewardMerchants(prev => [...prev, { ...item, id: 'r-' + Date.now() }]);
      }
    } else {
      if (item.id) {
        setKatalogHadiah(prev => prev.map(r => r.id === item.id ? item : r));
      } else {
        setKatalogHadiah(prev => [...prev, { ...item, id: 'k-' + Date.now() }]);
      }
    }
    setRewardModal({ isOpen: false, type: 'merchant' });
    showToast("Reward berhasil disimpan!", "success");
    if (sheetsScriptUrl) handleSyncLoyaltyToSheets(customers, loyaltyRedeemHistory, true);
  };

  const handleDeleteReward = (type: 'merchant' | 'catalog', id: string) => {
    if (type === 'merchant') {
      setRewardMerchants(prev => prev.filter(r => r.id !== id));
    } else {
      setKatalogHadiah(prev => prev.filter(r => r.id !== id));
    }
    setRewardModal({ isOpen: false, type: 'merchant' });
    showToast("Reward berhasil dihapus!", "success");
    if (sheetsScriptUrl) handleSyncLoyaltyToSheets(customers, loyaltyRedeemHistory, true);
  };
    
  // Loyalty Redeem rewards listing state
  const [loyaltyRedeemHistory, setLoyaltyRedeemHistory] = useState<any[]>(() => {
    try {
      const stored = localStorage.getItem("KPI_LOYALTY_REDEEMS");
      return stored ? JSON.parse(stored) : [];
    } catch (_) {
      return [];
    }
  });

  // New customer modal state
  const [isCustomerModalOpen, setIsCustomerModalOpen] = useState<boolean>(false);
  const [newCustomer, setNewCustomer] = useState({
    name: "",
    address: "",
    salesmanName: "",
    area: "",
    jenisToko: "Sembako",
    estimatedOmzet: 5000000,
    notesPerDay: 5,
    storeAgeYears: 2,
    ownership: "Milik Sendiri"
  });

  // Simulator states for loyalty page
  const [selectedCustomerIdForAction, setSelectedCustomerIdForAction] = useState<string>("c-1");
  const [followUpAction, setFollowUpAction] = useState<string>("Kunjungan");
  const [followUpNotes, setFollowUpNotes] = useState<string>("");

  const [selectedCustomerIdForRedeem, setSelectedCustomerIdForRedeem] = useState<string>("c-1");
  const [selectedRewardId, setSelectedRewardId] = useState<string>("r-1");

  // Notification state
  const [toast, setToast] = useState<{ message: string; type: "success" | "error" | "info" } | null>(null);
  const [isGoogleFormGuideOpen, setIsGoogleFormGuideOpen] = useState<boolean>(false);

  // --- GOOGLE SHEETS SYNC SYSTEM ---
  const [sheetsScriptUrl, setSheetsScriptUrl] = useState<string>(() => {
    const stored = localStorage.getItem("KPI_SHEETS_SCRIPT_URL");
    const targetDefault = "https://script.google.com/macros/s/AKfycbz2a1vnNlVwQl6z77-3b5LYgXZAhprIgH3a_JAl0El5GPPI8xiZZGPdGsapa_mM6S0DQA/exec";
    const oldDefaults = [
      "https://script.google.com/macros/s/AKfycbwXR5ztGzNXMc7wxOyfNIgczvlfZCf83t_SYEvWGPS67DnhLYn8uS9BZVxoSU5ygtY8iA/exec",
      "https://script.google.com/macros/s/AKfycbzqPNZ-p4BzDtrpx6qRqvYthMMM2wrQGmaGiZDFtUiv5zkDg_G5AWf9gnJd_a8GjiwzZA/exec",
      "https://script.google.com/macros/s/AKfycbxPp41Z06VMS5s7xyFpAjunHk0LV0cBXlBJIkQ9-jJYL9T5xSkrhxsk4Uh1hCFQnE7qow/exec",
      "https://script.google.com/macros/s/AKfycbwx-OO-AGrU1zjePdDN8Uo1QlECRPbC6WItFHZhzWvvMVVRz61KbpYgtHjpJe7ttpFGrA/exec",
      "https://script.google.com/macros/s/AKfycbxqkcVj_HP1lNvMXlFxM3sntzkfgsJRdo2xBNo-WgpOID3pzkNRMzDAnzXO33HjcLlQgQ/exec"
    ];
    if (!stored || stored.trim() === "" || stored.includes("placeholder") || oldDefaults.includes(stored.trim())) {
      localStorage.setItem("KPI_SHEETS_SCRIPT_URL", targetDefault);
      return targetDefault;
    }
    return stored;
  });
  const [sheetsLibraryUrl, setSheetsLibraryUrl] = useState<string>(() => {
    const stored = localStorage.getItem("KPI_SHEETS_LIBRARY_URL");
    const targetDefault = "https://script.google.com/macros/library/d/1y8VqApcuL-QYGyNyZFB0kF3MI7T9qwCtptmzUpQbNL34H-Gz0LCJBxkt/10";
    const oldDefaults = [
      "https://script.google.com/macros/library/d/1y8VqApcuL-QYGyNyZFB0kF3MI7T9qwCtptmzUpQbNL34H-Gz0LCJBxkt/1",
      "https://script.google.com/macros/library/d/1y8VqApcuL-QYGyNyZFB0kF3MI7T9qwCtptmzUpQbNL34H-Gz0LCJBxkt/3",
      "https://script.google.com/macros/library/d/1y8VqApcuL-QYGyNyZFB0kF3MI7T9qwCtptmzUpQbNL34H-Gz0LCJBxkt/4",
      "https://script.google.com/macros/library/d/1y8VqApcuL-QYGyNyZFB0kF3MI7T9qwCtptmzUpQbNL34H-Gz0LCJBxkt/7",
      "https://script.google.com/macros/library/d/1y8VqApcuL-QYGyNyZFB0kF3MI7T9qwCtptmzUpQbNL34H-Gz0LCJBxkt/8"
    ];
    if (!stored || stored.trim() === "" || oldDefaults.includes(stored.trim())) {
      localStorage.setItem("KPI_SHEETS_LIBRARY_URL", targetDefault);
      return targetDefault;
    }
    return stored;
  });
  const [isAutoSyncEnabled, setIsAutoSyncEnabled] = useState<boolean>(() => {
    return localStorage.getItem("KPI_SHEETS_AUTO_SYNC") === "true";
  });
  const [syncedReports, setSyncedReports] = useState<{ [reportId: string]: boolean }>(() => {
    try {
      const stored = localStorage.getItem("KPI_SHEETS_SYNCED_REPORTS");
      return stored ? JSON.parse(stored) : {};
    } catch {
      return {};
    }
  });
  const [isSyncingAll, setIsSyncingAll] = useState<boolean>(false);
  const [isSyncingSalesmen, setIsSyncingSalesmen] = useState<boolean>(false);
  const [isSyncingProducts, setIsSyncingProducts] = useState<boolean>(false);
  const [isFetchingProducts, setIsFetchingProducts] = useState<boolean>(false);
  const [isSyncingLoyalty, setIsSyncingLoyalty] = useState<boolean>(false);
  const [isSyncingNoo, setIsSyncingNoo] = useState<boolean>(false);
  const [isFetchingNoo, setIsFetchingNoo] = useState<boolean>(false);
  const [isSubmittingReport, setIsSubmittingReport] = useState<boolean>(false);
  const [testConnectionStatus, setTestConnectionStatus] = useState<"idle" | "testing" | "success" | "error">("idle");
  const [kpiDataSource, setKpiDataSource] = useState<"sheets" | "local">("sheets");

  // --- FORM STATE ---
  const [selectedSalesmanId, setSelectedSalesmanId] = useState<string>("");
  const [selectedCycle, setSelectedCycle] = useState<string>(() => autoDetectCycle(new Date().toISOString().split("T")[0]));
  const [area, setArea] = useState<string>("Banyumas");
  const [availableAreas, setAvailableAreas] = useState<string[]>(["Banyumas", "Banjarnegara", "Cilacap", "Purbalingga", "Purwokerto", "Kebumen"]);

  useEffect(() => {
    if (selectedSalesmanId && selectedCycle && salesmen.length > 0) {
      const salesman = salesmen.find(s => s.id === selectedSalesmanId);
      if (salesman) {
        try {
          const savedCustomerDB = localStorage.getItem("customerDatabase");
          if (savedCustomerDB) {
            const customerDbParsed = JSON.parse(savedCustomerDB);
            const matchingVisit = customerDbParsed.find(
              (v: any) => v.salesman.trim().toUpperCase() === salesman.name.trim().toUpperCase() && v.kunjungan.trim().toUpperCase() === selectedCycle.trim().toUpperCase()
            );
            if (matchingVisit && matchingVisit.customers && matchingVisit.customers.length > 0) {
              const uniqueKotas = Array.from(new Set(matchingVisit.customers.map((c: any) => c.kota))).filter(Boolean) as string[];
              if (uniqueKotas.length > 0) {
                setAvailableAreas(uniqueKotas);
                if (!uniqueKotas.includes(area)) {
                  setArea(uniqueKotas[0]);
                }
                return;
              }
            }
          }
        } catch (e) {
          console.error("Error reading customer DB", e);
        }
      }
    }
    
    // Fallback to standard defaults if no specific customer data matches
    const defaults = ["Banyumas", "Banjarnegara", "Cilacap", "Purbalingga", "Purwokerto", "Kebumen"];
    setAvailableAreas(defaults);
    if (!defaults.includes(area) && defaults.length > 0) {
      setArea(defaults[0]);
    }
  }, [selectedSalesmanId, selectedCycle, salesmen]);

  const [tc, setTc] = useState<number>(0);
  const [cp, setCp] = useState<number>(0);
  const [ec, setEc] = useState<number>(0);
  const [skuTotal, setSkuTotal] = useState<number>(0);
  const [operationalCost, setOperationalCost] = useState<number>(0);
  const [billsReceived, setBillsReceived] = useState<number>(0);
  const [billsTransfer, setBillsTransfer] = useState<number>(0);
  const [billsGiro, setBillsGiro] = useState<number>(0);
  const [reportDate, setReportDate] = useState<string>(new Date().toISOString().split("T")[0]);
  const [notes, setNotes] = useState<string>("");

  // Product Selection Details (for helping compute SKUs and simulating a real cart)
  const [selectedProducts, setSelectedProducts] = useState<{ [productId: string]: number }>({});

  // Floating Product Dropdown UI State
  const [productSearch, setProductSearch] = useState<string>("");
  const [isProductPickerOpen, setIsProductPickerOpen] = useState<boolean>(false);
  const productPickerRef = useRef<HTMLDivElement>(null);

  // Salesman Dropdown autocomplete UI State
  const [salesmanSearch, setSalesmanSearch] = useState<string>("");
  const [isSalesmanDropdownOpen, setIsSalesmanDropdownOpen] = useState<boolean>(false);
  const salesmanDropdownRef = useRef<HTMLDivElement>(null);

  // --- DATABASE EDIT MODALS ---
  const [salesmanModal, setSalesmanModal] = useState<{ isOpen: boolean; id?: string; name: string; area: string; phone: string }>({
    isOpen: false,
    name: "",
    area: "",
    phone: ""
  });

  const [productModal, setProductModal] = useState<{ isOpen: boolean; id?: string; name: string; category: string; skuCode: string }>({
    isOpen: false,
    name: "",
    category: "",
    skuCode: ""
  });

  // --- RAW COPY PASTE PARSER PANEL STATE ---
  const [rawPasteText, setRawPasteText] = useState<string>("");
  const [isParserOpen, setIsParserOpen] = useState<boolean>(false);
  const [parseResult, setParseResult] = useState<ImportParsingResult | null>(null);

  // --- LOAD INITIAL DATA FROM LOCALSTORAGE ---
  useEffect(() => {
    // 1. Salesmen
    const storedSalesmen = localStorage.getItem("KPI_DB_SALESMEN");
    if (storedSalesmen && JSON.parse(storedSalesmen).length > 0) {
      setSalesmen(JSON.parse(storedSalesmen));
    } else {
      localStorage.setItem("KPI_DB_SALESMEN", JSON.stringify(INITIAL_SALESMEN));
      setSalesmen(INITIAL_SALESMEN);
    }

    // 2. Products
    const storedProducts = localStorage.getItem("KPI_DB_PRODUCTS");
    if (storedProducts && JSON.parse(storedProducts).some((p: any) => p.id.startsWith("F-"))) {
      setProducts(JSON.parse(storedProducts));
    } else {
      localStorage.setItem("KPI_DB_PRODUCTS", JSON.stringify(INITIAL_PRODUCTS));
      setProducts(INITIAL_PRODUCTS);
    }

    // 3. Reports
    const storedReports = localStorage.getItem("KPI_DB_REPORTS");
    if (storedReports) {
      setReports(JSON.parse(storedReports));
    } else {
      localStorage.setItem("KPI_DB_REPORTS", JSON.stringify(INITIAL_REPORTS));
      setReports(INITIAL_REPORTS);
    }
    
    // 4. Salesman Goals
    const storedGoals = localStorage.getItem("KPI_DB_SALESMAN_GOALS");
    if (storedGoals) {
      setSalesmanGoals(JSON.parse(storedGoals));
    }
  }, []);

  useEffect(() => {
    localStorage.setItem("KPI_REWARD_MERCHANTS", JSON.stringify(rewardMerchants));
  }, [rewardMerchants]);

  useEffect(() => {
    localStorage.setItem("KPI_KATALOG_HADIAH", JSON.stringify(katalogHadiah));
  }, [katalogHadiah]);

  useEffect(() => {
    setSelectedCycle(autoDetectCycle(reportDate));
  }, [reportDate]);

  // Auto-fetch master data (Products SKU & Salesmen) whenever the URL configuration changes or app boots
  useEffect(() => {
    if (sheetsScriptUrl) {
      handleFetchProductsFromSheets(true); // quietly pull background master SKU products data
      handleFetchSalesmenFromSheets(true); // quietly pull background master salesmen data
    }
  }, [sheetsScriptUrl]);

  // Auto-fetch spreadsheet reports and NOO logs when switching tabs
  useEffect(() => {
    if (sheetsScriptUrl) {
      if (activeTab === "kpisales" || activeTab === "reports" || activeTab === "sheets") {
        handleFetchReportsFromSheets(true); // quietly pull background spreadsheet data
        handleFetchNooFromSheets(true); // quietly pull background NOO logs
      } else if (activeTab === "products") {
        handleFetchProductsFromSheets(true); // quietly pull background products SKU data to stay fresh
      }
    }
  }, [activeTab, sheetsScriptUrl]);

  const saveToLocalStorage = (key: string, data: any) => {
    localStorage.setItem(key, JSON.stringify(data));
  };

  // Toast Helper
  const showToast = (message: string, type: "success" | "error" | "info" = "success") => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 4000);
  };

  // Close helper for floating UI elements when clicking outside
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (productPickerRef.current && !productPickerRef.current.contains(event.target as Node)) {
        setIsProductPickerOpen(false);
      }
      if (salesmanDropdownRef.current && !salesmanDropdownRef.current.contains(event.target as Node)) {
        setIsSalesmanDropdownOpen(false);
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  // Set default Selected Salesman Name in header display
  const activeSalesman = salesmen.find(s => s.id === selectedSalesmanId);

  // --- HANDLERS FOR THE INPUT FORM ---

  const handleAutoDetectCycle = () => {
    const detected = autoDetectCycle(reportDate);
    setSelectedCycle(detected);
    showToast(`Siklus otomatis terdeteksi: ${detected}`, "info");
  };

  // When a product is selected in the picker
  const handleSelectProduct = (product: Product) => {
    // Treat presence in object as selected
    const nextSelected = {
      ...selectedProducts,
      [product.id]: 1
    };
    setSelectedProducts(nextSelected);

    // Compute SKU Total: simply the number of unique focus products selected
    const newSkuTotalCount = Object.keys(nextSelected).length;
    setSkuTotal(newSkuTotalCount);

    showToast(`Produk "${product.name}" ditambahkan ke rincian penjualan.`, "success");
    setIsProductPickerOpen(false);
    setProductSearch("");
  };

  const handleUpdateProductQty = (productId: string, qty: number) => {
    const nextSelected = { ...selectedProducts };
    if (qty <= 0) {
      delete nextSelected[productId];
    } else {
      nextSelected[productId] = 1; // force max 1 just in case
    }
    setSelectedProducts(nextSelected);

    // Compute new SKU Total
    const newSkuTotalCount = Object.keys(nextSelected).length;
    setSkuTotal(newSkuTotalCount);
  };

  // Submit report to local database and sync with Google Sheets
  const handleSubmitReport = async (e: React.FormEvent) => {
    e.preventDefault();

    let matchedSales: Salesman | null | undefined = undefined;

    if (!selectedSalesmanId) {
      showToast("Gagal mendeteksi Salesman. Silakan pilih salesman terlebih dahulu!", "error");
      return;
    }

    matchedSales = salesmen.find(s => s.id === selectedSalesmanId) || activeSalesman || (salesmen && salesmen.length > 0 ? salesmen[0] : null);
    if (!matchedSales) {
      showToast("Salesman tidak terdaftar di database! Harap tambahkan salesman terlebih dahulu.", "error");
      return;
    }

    setIsSubmittingReport(true);

    // Build productsDetail
    const productsDetail = Object.entries(selectedProducts).map(([pId, _]) => {
      const prod = products.find(p => p.id === pId);
      return {
        productId: pId,
        productName: prod ? prod.name : "Produk tidak dikenal"
      };
    });

    const newReport: KpiReport = {
      id: "rep-" + Date.now(),
      salesmanId: selectedSalesmanId,
      salesmanName: matchedSales.name,
      date: reportDate,
      cycle: selectedCycle,
      area: area,
      tc,
      cp,
      ec,
      skuTotal,
      operationalCost,
      billsReceived,
      billsTransfer,
      billsGiro,
      notes,
      productsDetail,
      createdAt: new Date().toISOString()
    };

    const updated = [newReport, ...reports];
    setReports(updated);
    saveToLocalStorage("KPI_DB_REPORTS", updated);

    // Dynamic NOO calculation & entry for Aris & Imam
    const slNameLower = matchedSales.name.toLowerCase().trim();
    const isArisOrImam = slNameLower === "aris" || slNameLower === "imam";
    let savedNooRecord: NooRecord | null = null;

    if (isArisOrImam && (newNooWarung > 0 || newNooStore > 0 || newNooKiosk > 0 || newNooWholesaler > 0)) {
      savedNooRecord = {
        id: "noo-" + Date.now(),
        salesmanId: selectedSalesmanId,
        salesmanName: matchedSales.name,
        date: reportDate,
        warung: newNooWarung,
        store: newNooStore,
        kiosk: newNooKiosk,
        wholesaler: newNooWholesaler
      };
      
      const updatedNoo = [savedNooRecord, ...nooRecords];
      setNooRecords(updatedNoo);
      saveToLocalStorage("KPI_DB_NOO_LOGS", updatedNoo);
    }

    // Connected to Sheets directly
    if (sheetsScriptUrl) {
      showToast("Sedang menyimpan laporan ke Google Sheets...", "info");
      const ok = await syncSingleReportToSheets(newReport, true);
      
      if (savedNooRecord) {
        // Re-construct matching array to sync to Sheets in real-time
        const updatedNoo = [savedNooRecord, ...nooRecords];
        await handleSyncNooToSheets(updatedNoo, true);
      }

      if (ok) {
        if (savedNooRecord) {
          showToast(`Laporan & Log NOO ${matchedSales.name} berhasil disimpan ke Database & Google Sheets!`, "success");
        } else {
          showToast(`Laporan ${matchedSales.name} berhasil disimpan ke Database & Google Sheets!`, "success");
        }
      } else {
        showToast(`Laporan tersimpan di Lokal, namun gagal terkirim ke Google Sheets. Silakan cek koneksi/URL!`, "error");
      }
    } else {
      if (savedNooRecord) {
        showToast(`Laporan harian KPI & Log NOO untuntuk ${matchedSales.name} berhasil tersimpan ke database lokal!`, "success");
      } else {
        showToast(`Laporan harian KPI untuk ${matchedSales.name} berhasil tersimpan ke database lokal!`, "success");
      }
    }

    // Reset Form
    setSelectedSalesmanId("");
    setSalesmanSearch("");
    setTc(0);
    setCp(0);
    setEc(0);
    setSkuTotal(0);
    setSelectedProducts({});
    setOperationalCost(0);
    setBillsReceived(0);
    setBillsTransfer(0);
    setBillsGiro(0);
    setNotes("");
    
    // Reset NOO fields
    setNewNooWarung(0);
    setNewNooStore(0);
    setNewNooKiosk(0);
    setNewNooWholesaler(0);

    setIsSubmittingReport(false);
  };

  // --- GOOGLE SHEETS SYNC SYSTEM FUNCTIONS ---

  const syncSingleReportToSheets = async (rep: KpiReport, silent = false) => {
    const url = localStorage.getItem("KPI_SHEETS_SCRIPT_URL") || sheetsScriptUrl;
    if (!url) {
      if (!silent) showToast("Google Apps Script URL belum diatur!", "error");
      return false;
    }
    try {
      const response = await fetch(url, {
        method: "POST",
        mode: "cors",
        headers: {
          "Content-Type": "text/plain",
        },
        body: JSON.stringify({
          action: "addReport",
          report: rep
        })
      });
      const text = await response.text();
      let resData;
      try {
        resData = JSON.parse(text);
      } catch {
        resData = { success: true, message: "Koneksi berhasil" };
      }

      if (resData.success) {
        const nextSynced = { ...syncedReports, [rep.id]: true };
        setSyncedReports(nextSynced);
        localStorage.setItem("KPI_SHEETS_SYNCED_REPORTS", JSON.stringify(nextSynced));
        if (!silent) {
          showToast(`Laporan ${rep.salesmanName} sukses disinkronkan ke Google Sheets!`, "success");
        }
        return true;
      } else {
        if (!silent) {
          showToast(`Sinkronisasi gagal: ${resData.message || "kesalahan internal"}`, "error");
        }
        return false;
      }
    } catch (err: any) {
      console.error("Sheets sync error:", err);
      if (!silent) {
        showToast(`Gagal menghubungi Google Apps Script. Silakan periksa URL Anda!`, "error");
      }
      return false;
    }
  };

  const handleTestConnection = async () => {
    if (!sheetsScriptUrl) {
      showToast("Tolong masukkan URL Google Apps Script Web App terlebih dahulu!", "error");
      return;
    }
    setTestConnectionStatus("testing");
    try {
      const response = await fetch(sheetsScriptUrl, {
        method: "POST",
        mode: "cors",
        headers: {
          "Content-Type": "text/plain",
        },
        body: JSON.stringify({
          action: "test"
        })
      });
      const text = await response.text();
      const resData = JSON.parse(text);
      if (resData.success) {
        setTestConnectionStatus("success");
        showToast("Koneksi ke Google Sheets berhasil!", "success");
      } else {
        setTestConnectionStatus("error");
        showToast(`Uji koneksi gagal: ${resData.message}`, "error");
      }
    } catch (err) {
      console.error("Test connection error:", err);
      setTestConnectionStatus("error");
      showToast("Gagal menghubungi URL. Minta deploy ulang Apps Script dengan akses ke 'Anyone'!", "error");
    }
  };

  const handleFetchReportsFromSheets = async (silent = false) => {
    if (!sheetsScriptUrl) {
      if (!silent) {
        showToast("Tolong hubungkan dan masukkan URL Google Sheets Web App terlebih dahulu di tab 'Google Sheets'!", "error");
      }
      return [];
    }
    setIsFetchingReports(true);
    try {
      const response = await fetch(sheetsScriptUrl, {
        method: "POST",
        mode: "cors",
        headers: {
          "Content-Type": "text/plain",
        },
        body: JSON.stringify({
          action: "getReports"
        })
      });
      const text = await response.text();
      const resData = JSON.parse(text);
      if (resData.success && resData.reports) {
        const rawList = resData.reports;
        const normalized: KpiReport[] = rawList.map((item: any, index: number) => {
          return {
            id: String(item["ID Laporan"] || item.id || `fetched-${index}`),
            date: String(item["Tanggal KPI"] || item.date || ""),
            salesmanName: String(item["Nama Salesman"] || item.salesmanName || "SALES"),
            salesmanId: String(item["ID Salesman"] || item.salesmanId || "s-unknown"),
            cycle: String(item["Siklus"] || item.cycle || ""),
            area: String(item["Area"] || item.area || "Banyumas"),
            tc: Number(item["TC (Amplop)"] || item.tc || 0),
            cp: Number(item["CP (Kunjungan)"] || item.cp || 0),
            ec: Number(item["EC (Order)"] || item.ec || 0),
            skuTotal: Number(item["SKU Total"] || item.skuTotal || 0),
            operationalCost: parseInt(String(item["Biaya Operasional (Rp)"] || item.operationalCost || 0).replace(/[^0-9]/g, ""), 10) || 0,
            billsReceived: Number(item["Tagihan Bayar Tunai"] || item["Tagihan Didapat (Rp)"] || item.billsReceived || 0),
            billsTransfer: Number(item["Tagihan Bayar Transfer"] || item.billsTransfer || 0),
            billsGiro: Number(item["Tagihan Giro"] || item.billsGiro || 0),
            notes: String(item["Catatan"] || item.notes || ""),
            createdAt: String(item["Tanggal Dibuat"] || item.createdAt || ""),
            productsDetail: [] // empty placeholders
          };
        }).filter((r: any) => r.salesmanName && r.salesmanName !== "Nama Salesman");

        setFetchedReports(normalized);
        setKpiDataSource("sheets");
        localStorage.setItem("KPI_FETCHED_REPORTS_FROM_SHEETS", JSON.stringify(normalized));
        const nowStr = new Date().toLocaleString("id-ID");
        setLastFetchTime(nowStr);
        localStorage.setItem("KPI_LAST_FETCH_TIME", nowStr);
        if (!silent) {
          showToast(`Berhasil menarik ${normalized.length} data laporan audit dari Google Sheets!`, "success");
        }
        return normalized;
      } else {
        if (!silent) {
          showToast(`Gagal menarik data: ${resData.message || "Tolong periksa Apps Script Anda."}`, "error");
        }
        return [];
      }
    } catch (err) {
      console.error("Fetch reports error:", err);
      if (!silent) {
        showToast("Gagal menghubungi Google Sheet. Verifikasi koneksi & deploy ulang Apps Script!", "error");
      }
      return [];
    } finally {
      setIsFetchingReports(false);
    }
  };

  const handleFetchSalesmenFromSheets = async (silent = false) => {
    if (!sheetsScriptUrl) return;
    try {
      const response = await fetch(sheetsScriptUrl, {
        method: "POST",
        mode: "cors",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({ action: "getSalesmen" })
      });
      const text = await response.text();
      const resData = JSON.parse(text);
      if (resData.success && resData.salesmen) {
        const normalized: Salesman[] = resData.salesmen.map((item: any) => ({
          id: String(item["ID Salesman"] || item.id),
          name: String(item["Nama Salesman"] || item.name),
          area: String(item["Area Wilayah"] || item.area),
          phone: String(item["No. HP / Telepon"] || item.phone || "-"),
          isActive: true
        }));
        setSalesmen(normalized);
        localStorage.setItem("KPI_DB_SALESMEN", JSON.stringify(normalized));
        if (!silent) {
          showToast("Berhasil menarik data dari tab 'Daftar Salesman' di Google Sheets!", "success");
        }
      }
    } catch (err) {
      console.error("Fetch salesmen error:", err);
      if (!silent) {
        showToast("Gagal menarik data Salesman. Pastikan tab sheet bernama 'Daftar Salesman' dan script aktif.", "error");
      }
    }
  };

  const handleFetchProductsFromSheets = async (silent = false) => {
    if (!sheetsScriptUrl) {
      if (!silent) showToast("Tolong masukkan URL Google Sheets Web App terlebih dahulu di tab 'Google Sheets'!", "error");
      return;
    }
    setIsFetchingProducts(true);
    try {
      const response = await fetch(sheetsScriptUrl, {
        method: "POST",
        mode: "cors",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({ action: "getProducts" })
      });
      const text = await response.text();
      const resData = JSON.parse(text);
      if (resData.success && resData.products) {
        const normalized: Product[] = resData.products.map((item: any) => ({
          id: String(item["IDsku"] || item["ID Produk"] || item.id || ""),
          name: String(item["Nama SKU Produk"] || item["Nama Produk"] || item.name || ""),
          category: String(item["Kategori"] || item["Kategori Produk"] || item.category || "-"),
          skuCode: String(item["SKU Code"] || item.skuCode || item["IDsku"] || ""),
          isActive: true
        })).filter((p: any) => p.name && p.name !== "Nama SKU Produk" && p.id && p.id !== "IDsku");
        
        setProducts(normalized);
        localStorage.setItem("KPI_DB_PRODUCTS", JSON.stringify(normalized));
        if (!silent) {
          showToast(`Berhasil menarik ${normalized.length} data Produk dari Google Sheets!`, "success");
        }
      } else {
        if (!silent) {
          showToast(`Gagal menarik data: ${resData.message || "Pastikan tab sheet bernama 'Daftar Produk SKU' atau 'Daftar Produk' sudah ada."}`, "error");
        }
      }
    } catch (err) {
      console.error("Fetch products error:", err);
      if (!silent) {
        showToast("Gagal menghubungi Google Apps Script untuk menarik data produk.", "error");
      }
    } finally {
      setIsFetchingProducts(false);
    }
  };

  const handleSyncAllToSheets = async () => {
    if (!sheetsScriptUrl) {
      showToast("Tolong masukkan URL Google Apps Script Web App terlebih dahulu!", "error");
      return;
    }
    if (reports.length === 0) {
      showToast("Belum ada data laporan untuk disinkronkan!", "error");
      return;
    }
    
    setIsSyncingAll(true);
    try {
      const response = await fetch(sheetsScriptUrl, {
        method: "POST",
        mode: "cors",
        headers: {
          "Content-Type": "text/plain",
        },
        body: JSON.stringify({
          action: "syncAll",
          reports: reports
        })
      });
      const text = await response.text();
      const resData = JSON.parse(text);
      if (resData.success) {
        const nextSynced: { [id: string]: boolean } = {};
        reports.forEach(r => {
          nextSynced[r.id] = true;
        });
        setSyncedReports(nextSynced);
        localStorage.setItem("KPI_SHEETS_SYNCED_REPORTS", JSON.stringify(nextSynced));
        showToast(`Sukses menyinkronkan ${reports.length} laporan ke Google Sheets!`, "success");
      } else {
        showToast(`Gagal sinkronisasi masal: ${resData.message}`, "error");
      }
    } catch (err) {
      console.error("Bulk sync error:", err);
      showToast("Gagal menyinkronkan seluruh data. Cek ulang konfigurasi!", "error");
    } finally {
      setIsSyncingAll(false);
    }
  };

  const handleSyncSalesmenToSheets = async () => {
    if (!sheetsScriptUrl) {
      showToast("Tolong masukkan URL Google Apps Script Web App terlebih dahulu!", "error");
      return;
    }
    if (salesmen.length === 0) {
      showToast("Belum ada data Salesman untuk disinkronkan!", "error");
      return;
    }
    
    setIsSyncingSalesmen(true);
    try {
      const response = await fetch(sheetsScriptUrl, {
        method: "POST",
        mode: "cors",
        headers: {
          "Content-Type": "text/plain",
        },
        body: JSON.stringify({
          action: "syncSalesmen",
          salesmen: salesmen
        })
      });
      const text = await response.text();
      const resData = JSON.parse(text);
      if (resData.success) {
        showToast(resData.message || "Sukses menyinkronkan data Salesman!", "success");
      } else {
        showToast(`Gagal sinkronisasi Salesman: ${resData.message}`, "error");
      }
    } catch (err) {
      console.error("Salesmen sync error:", err);
      showToast("Gagal menyinkronkan data Salesman. Cek koneksi & deploy ulang!", "error");
    } finally {
      setIsSyncingSalesmen(false);
    }
  };

  const handleSyncProductsToSheets = async (customProducts?: Product[], silent = false) => {
    if (!sheetsScriptUrl) {
      if (!silent) showToast("Tolong masukkan URL Google Apps Script Web App terlebih dahulu!", "error");
      return;
    }
    const targetProducts = customProducts || products;
    if (targetProducts.length === 0) {
      if (!silent) showToast("Belum ada data SKU Produk untuk disinkronkan!", "error");
      return;
    }
    
    setIsSyncingProducts(true);
    try {
      const response = await fetch(sheetsScriptUrl, {
        method: "POST",
        mode: "cors",
        headers: {
          "Content-Type": "text/plain",
        },
        body: JSON.stringify({
          action: "syncProducts",
          products: targetProducts
        })
      });
      const text = await response.text();
      const resData = JSON.parse(text);
      if (resData.success) {
        if (!silent) {
          showToast(resData.message || "Sukses menyinkronkan SKU Produk!", "success");
        }
      } else {
        if (!silent) {
          showToast(`Gagal sinkronisasi SKU Produk: ${resData.message}`, "error");
        }
      }
    } catch (err) {
      console.error("Products sync error:", err);
      if (!silent) {
        showToast("Gagal menyinkronkan data SKU Produk. Cek koneksi & deploy ulang!", "error");
      }
    } finally {
      setIsSyncingProducts(false);
    }
  };

  const handleSyncLoyaltyToSheets = async (customCustomers?: any[], customRedeems?: any[], silent: boolean = false) => {
    if (!sheetsScriptUrl) {
      if (!silent) showToast("Tolong masukkan URL Google Apps Script Web App terlebih dahulu!", "error");
      return;
    }
    const targetCustomers = customCustomers || customers;
    const targetRedeems = customRedeems || loyaltyRedeemHistory;
    if (targetCustomers.length === 0) {
      if (!silent) showToast("Belum ada data Profiling Loyalti Toko untuk disinkronkan!", "error");
      return;
    }
    
    setIsSyncingLoyalty(true);
    try {
      const response = await fetch(sheetsScriptUrl, {
        method: "POST",
        mode: "cors",
        headers: {
          "Content-Type": "text/plain",
        },
        body: JSON.stringify({
          action: "syncLoyalty",
          customers: targetCustomers,
          redeems: targetRedeems
        })
      });
      const text = await response.text();
      const resData = JSON.parse(text);
      if (resData.success) {
        if (!silent) showToast(resData.message || "Sukses menyinkronkan database Loyalti!", "success");
      } else {
        if (!silent) showToast(`Gagal sinkronisasi Loyalti: ${resData.message}`, "error");
      }
    } catch (err) {
      console.error("Loyalty sync error:", err);
      if (!silent) showToast("Gagal menyinkronkan data Loyalti & Profiling. Cek koneksi & deploy ulang!", "error");
    } finally {
      setIsSyncingLoyalty(false);
    }
  };

  const handleSyncNooToSheets = async (customNoo?: NooRecord[], silent = false) => {
    if (!sheetsScriptUrl) {
      if (!silent) showToast("Tolong hubungkan dan masukkan URL Google Sheets Web App terlebih dahulu di tab 'Google Sheets'!", "error");
      return;
    }
    const targetNoo = customNoo || nooRecords;
    setIsSyncingNoo(true);
    try {
      const response = await fetch(sheetsScriptUrl, {
        method: "POST",
        mode: "cors",
        headers: {
          "Content-Type": "text/plain",
        },
        body: JSON.stringify({
          action: "syncNoo",
          nooLogs: targetNoo
        })
      });
      const text = await response.text();
      const resData = JSON.parse(text);
      if (resData.success) {
        if (!silent) {
          showToast(resData.message || "Sukses menyinkronkan database NOO!", "success");
        }
      } else {
        if (!silent) {
          showToast(`Gagal sinkronisasi NOO: ${resData.message}`, "error");
        }
      }
    } catch (err) {
      console.error("NOO sync error:", err);
      if (!silent) {
        showToast("Gagal menyinkronkan data NOO ke Google Sheets. Cek koneksi & deploy ulang!", "error");
      }
    } finally {
      setIsSyncingNoo(false);
    }
  };

  const handleFetchNooFromSheets = async (silent = false) => {
    if (!sheetsScriptUrl) {
      if (!silent) {
         showToast("Tolong hubungkan dan masukkan URL Google Sheets Web App terlebih dahulu di tab 'Google Sheets'!", "error");
      }
      return [];
    }
    setIsFetchingNoo(true);
    try {
      const response = await fetch(sheetsScriptUrl, {
        method: "POST",
        mode: "cors",
        headers: {
          "Content-Type": "text/plain",
        },
        body: JSON.stringify({
          action: "getNoo"
        })
      });
      const text = await response.text();
      const resData = JSON.parse(text);
      if (resData.success && resData.nooLogs) {
        const rawList = resData.nooLogs;
        const normalized: NooRecord[] = rawList.map((item: any, index: number) => {
          return {
            id: String(item["ID Log"] || item.id || `fetched-noo-${index}`),
            date: String(item["Tanggal Log"] || item.date || ""),
            salesmanName: String(item["Nama Salesman"] || item.salesmanName || ""),
            salesmanId: String(item["ID Salesman"] || item.salesmanId || ""),
            warung: Number(item["Warung / Toko Kelontong"] || item.warung || 0),
            store: Number(item["Store / Toko Modern"] || item.store || 0),
            kiosk: Number(item["Kios Atap"] || item.kiosk || 0),
            wholesaler: Number(item["Grosir / Wholesaler"] || item.wholesaler || 0)
          };
        }).filter((n: any) => n.salesmanName && n.salesmanName !== "Nama Salesman");

        setNooRecords(normalized);
        localStorage.setItem("KPI_DB_NOO_LOGS", JSON.stringify(normalized));
        if (!silent) {
          showToast(`Berhasil menarik ${normalized.length} data log harian NOO dari Google Sheets!`, "success");
        }
        return normalized;
      } else {
        if (!silent) {
          showToast(`Gagal menarik data NOO: ${resData.message || "Tolong periksa Apps Script Anda."}`, "error");
        }
        return [];
      }
    } catch (err) {
      console.error("Fetch NOO error:", err);
      if (!silent) {
        showToast("Gagal menghubungi Google Sheet untuk menarik data NOO. Verifikasi koneksi & deploy ulang Apps Script!", "error");
      }
      return [];
    } finally {
      setIsFetchingNoo(false);
    }
  };

  const [isImportingGoogleForm, setIsImportingGoogleForm] = useState<boolean>(false);

  const handleImportFromGoogleForm = async () => {
    if (!sheetsScriptUrl) {
      showToast("Tolong masukkan URL Google Apps Script Web App terlebih dahulu!", "error");
      return;
    }
    
    setIsImportingGoogleForm(true);
    try {
      const response = await fetch(sheetsScriptUrl, {
        method: "POST",
        mode: "cors",
        headers: {
          "Content-Type": "text/plain",
        },
        body: JSON.stringify({
          action: "importFromGoogleForm"
        })
      });
      const text = await response.text();
      const resData = JSON.parse(text);
      if (resData.success) {
        if (!resData.responses || resData.responses.length === 0) {
          showToast("Koneksi berhasil, namun belum ada respon pendaftaran baru di Google Form Anda.", "info");
          return;
        }

        const imported = resData.responses;
        let addedCount = 0;
        const updatedCustomers = [...customers];

        imported.forEach((item: any) => {
          const isDuplicate = updatedCustomers.some(
            (c) => c.name.toLowerCase().trim() === item.name.toLowerCase().trim()
          );
          if (!isDuplicate) {
            const matchedSalesman = salesmen.find(
              (s) => s.name.toLowerCase().trim() === (item.salesmanName || "").toLowerCase().trim()
            );
            
            let finalTier: "Platinum" | "Gold" | "Silver" | "Bronze" = "Bronze";
            const omzetValue = Number(item.estimatedOmzet) || 5000000;
            if (omzetValue >= 15000000) finalTier = "Platinum";
            else if (omzetValue >= 8000000) finalTier = "Gold";
            else if (omzetValue >= 5000000) finalTier = "Silver";

            const newCust = {
              id: "c-form-" + Date.now() + "_" + Math.floor(Math.random() * 1000),
              name: item.name,
              address: item.address || "Belum diisi",
              salesmanName: matchedSalesman ? matchedSalesman.name : (salesmen[0]?.name || "RIZKY"),
              area: matchedSalesman ? matchedSalesman.area : (salesmen[0]?.area || "Semarang"),
              jenisToko: item.jenisToko || "Sembako",
              estimatedOmzet: omzetValue,
              notesPerDay: 5,
              storeAgeYears: Number(item.storeAgeYears) || 2,
              ownership: item.ownership || "Milik Sendiri",
              points: 50, // bonus starter points
              tier: finalTier,
              createdAt: item.timestamp || new Date().toISOString(),
              actionsLog: [
                {
                  date: new Date().toISOString().split("T")[0],
                  action: "Sinkronisasi Google Form",
                  notes: `Daftar mandiri via respon Google Form. Nilai omzet Rp ${omzetValue.toLocaleString("id-ID")}. +50 Poin diberikan.`,
                  status: "Selesai"
                }
              ]
            };
            updatedCustomers.unshift(newCust);
            addedCount++;
          }
        });

        if (addedCount > 0) {
          setCustomers(updatedCustomers);
          saveToLocalStorage("KPI_LOYALTY_CUSTOMERS", updatedCustomers);
          
          await handleSyncLoyaltyToSheets(updatedCustomers, undefined, true);
          showToast(`Berhasil mengimpor ${addedCount} mitra toko baru dari Google Form!`, "success");
        } else {
          showToast("Tidak ada mitra toko baru yang unik untuk diimpor (semua sudah terdaftar sebelumnya).", "info");
        }
      } else {
        if (resData.message && resData.message.includes("tidak dikenali")) {
          showToast(`Gagal: Apps Script versi lama! Buka Menu Akun -> Setup Google Sheets -> Salin kode terbaru -> Lakukan 'New Deployment' di Google Apps Script.`, "error");
        } else {
          showToast(`Gagal import Google Form: ${resData.message}`, "error");
        }
      }
    } catch (err) {
      console.error("Import google form error:", err);
      showToast("Gagal melakukan impor Google Form. Periksa URL Apps Script & pastikan sudah di-link ke Sheets!", "error");
    } finally {
      setIsImportingGoogleForm(false);
    }
  };

  // --- TEXT WA COPY PASTE PARSER ---

  const handleParseText = () => {
    if (!rawPasteText || rawPasteText.trim() === "") {
      showToast("Tempel teks laporan terlebih dahulu!", "error");
      return;
    }

    const result = parseKpiText(rawPasteText, salesmen);
    
    if (result.salesmanId) {
      setSelectedSalesmanId(result.salesmanId);
      setSalesmanSearch(result.salesmanName);
    } else if (result.salesmanName) {
      // Prompt user to select/add this name
      setSalesmanSearch(result.salesmanName);
      setSelectedSalesmanId("");
    }

    if (result.cycle) {
      // Find matches in STANDARD_CYCLES
      const cycleMatch = STANDARD_CYCLES.find(
        c => c.toLowerCase() === result.cycle.toLowerCase() ||
             result.cycle.toLowerCase().includes(c.toLowerCase())
      );
      if (cycleMatch) {
         setSelectedCycle(cycleMatch);
      }
    }

    setTc(result.tc);
    setCp(result.cp);
    setEc(result.ec);
    
    // Set SKU Total
    setSkuTotal(result.skuTotal);
    
    setOperationalCost(result.operationalCost);
    setBillsReceived(result.billsReceived);
    setBillsTransfer(result.billsTransfer || 0);
    setBillsGiro(result.billsGiro || 0);
    if (result.notes) {
      setNotes(result.notes);
    }

    // Compose feedback
    let warningMsg = "";
    if (result.warnings.length > 0) {
      warningMsg = `. Catatan: ${result.warnings.join(", ")}`;
      showToast(`Teks terurai tetapi butuh verifikasi${warningMsg}`, "info");
    } else {
      showToast("Laporan berhasil diuraikan & Form terisi otomatis!", "success");
    }

    // Reset parser tab and close it
    setIsParserOpen(false);
  };

  // Format form to WhatsApp report style
  const handleCopyCurrentFormAsWA = () => {
    const slName = activeSalesman ? activeSalesman.name : "(Belum Pilih)";
    const waText = `AUDIT KPI SALES FORCE\n` +
      `-----------------------------\n` +
      `Salesman     : ${slName}\n` +
      `Tanggal      : ${reportDate}\n` +
      `Siklus Hari  : ${selectedCycle}\n` +
      `-----------------------------\n` +
      `TC (AMPLOP)  : ${tc}\n` +
      `CP (KUNJUNGAN): ${cp}\n` +
      `EC (ORDER)   : ${ec}\n` +
      `SKU TOTAL    : ${skuTotal}\n` +
      `-----------------------------\n` +
      `Operasional  : Rp ${operationalCost.toLocaleString("id-ID")}\n` +
      `Tagihan (Rp) : Rp ${billsReceived.toLocaleString("id-ID")}\n` +
      `Catatan      : ${notes || "-"}\n` +
      `-----------------------------\n` +
      `*Dikuasai oleh Auditor Portal*`;

    navigator.clipboard.writeText(waText);
    showToast("Format laporan berhasil disalin ke clipboard Anda!", "success");
  };

  // --- DB SALESMAN MANAGEMENT ---

  const handleAddSalesman = () => {
    setSalesmanModal({
      isOpen: true,
      name: "",
      area: "",
      phone: ""
    });
  };

  const handleEditSalesman = (s: Salesman) => {
    setSalesmanModal({
      isOpen: true,
      id: s.id,
      name: s.name,
      area: s.area || "",
      phone: s.phone || ""
    });
  };

  const handleSaveSalesman = (e: React.FormEvent) => {
    e.preventDefault();
    if (!salesmanModal.name.trim()) {
      showToast("Nama Salesman harus diisi!", "error");
      return;
    }

    if (salesmanModal.id) {
      // Edit
      const updated = salesmen.map(s => {
        if (s.id === salesmanModal.id) {
          return {
            ...s,
            name: salesmanModal.name.toUpperCase(),
            area: salesmanModal.area,
            phone: salesmanModal.phone
          };
        }
        return s;
      });
      setSalesmen(updated);
      saveToLocalStorage("KPI_DB_SALESMEN", updated);
      showToast(`Salesman ${salesmanModal.name.toUpperCase()} diperbarui!`);
    } else {
      // Create new
      const newS: Salesman = {
        id: "s-" + Date.now(),
        name: salesmanModal.name.toUpperCase(),
        area: salesmanModal.area,
        phone: salesmanModal.phone,
        isActive: true,
        createdAt: new Date().toISOString()
      };
      const updated = [...salesmen, newS];
      setSalesmen(updated);
      saveToLocalStorage("KPI_DB_SALESMEN", updated);
      showToast(`Salesman ${newS.name} berhasil ditambahkan ke database!`);
    }

    setSalesmanModal({ isOpen: false, name: "", area: "", phone: "" });
  };

  const handleDeleteSalesman = (id: string, name: string) => {
    if (confirm(`Apakah Anda yakin ingin menghapus salesman ${name} dari database?`)) {
      const updated = salesmen.filter(s => s.id !== id);
      setSalesmen(updated);
      saveToLocalStorage("KPI_DB_SALESMEN", updated);
      showToast(`Salesman ${name} telah dihapus.`, "info");
    }
  };

  const handleUpdateGoal = (salesmanId: string, monthString: string, field: keyof SalesmanGoal, value: number) => {
    setSalesmanGoals(prev => {
      const existingIdx = prev.findIndex(g => g.salesmanId === salesmanId && g.monthString === monthString);
      let newGoals = [...prev];
      if (existingIdx >= 0) {
        newGoals[existingIdx] = { ...newGoals[existingIdx], [field]: value };
      } else {
        newGoals.push({
          salesmanId,
          monthString,
          tcTarget: field === "tcTarget" ? value : 0,
          cpTarget: field === "cpTarget" ? value : 0,
          ecTarget: field === "ecTarget" ? value : 0,
          skuTarget: field === "skuTarget" ? value : 0,
        });
      }
      saveToLocalStorage("KPI_DB_SALESMAN_GOALS", newGoals);
      return newGoals;
    });
  };


  // --- DB PRODUCT MANAGEMENT ---

  const handleAddProduct = () => {
    setProductModal({
      isOpen: true,
      name: "",
      category: "",
      skuCode: ""
    });
  };

  const handleEditProduct = (p: Product) => {
    setProductModal({
      isOpen: true,
      id: p.id,
      name: p.name,
      category: p.category || "",
      skuCode: p.skuCode || ""
    });
  };

  const handleSaveProduct = (e: React.FormEvent) => {
    e.preventDefault();
    if (!productModal.name.trim()) {
      showToast("Nama produk harus diisi!", "error");
      return;
    }

    let updated: Product[] = [];
    if (productModal.id) {
      // Edit
      updated = products.map(p => {
        if (p.id === productModal.id) {
          return {
            ...p,
            name: productModal.name.toUpperCase(),
            category: productModal.category,
            skuCode: productModal.skuCode
          };
        }
        return p;
      });
      setProducts(updated);
      saveToLocalStorage("KPI_DB_PRODUCTS", updated);
      showToast(`Produk ${productModal.name.toUpperCase()} diperbarui!`);
    } else {
      // Create new
      let nextNum = 1;
      products.forEach(p => {
        const match = p.id.match(/^F-(\d+)$/i);
        if (match) {
          const num = parseInt(match[1], 10);
          if (num >= nextNum) nextNum = num + 1;
        }
      });
      const newP: Product = {
        id: "F-" + nextNum,
        name: productModal.name.toUpperCase(),
        category: productModal.category,
        skuCode: productModal.skuCode || "F-" + nextNum,
        isActive: true
      };
      updated = [...products, newP];
      setProducts(updated);
      saveToLocalStorage("KPI_DB_PRODUCTS", updated);
      showToast(`Produk ${newP.name} berhasil ditambahkan ke database!`);
    }

    setProductModal({ isOpen: false, name: "", category: "", skuCode: "" });

    // Live sync to Sheets!
    if (sheetsScriptUrl) {
      handleSyncProductsToSheets(updated, true);
    }
  };

  const handleDeleteProduct = (id: string, name: string) => {
    setDeleteConfirmProductId(id);
  };

  const confirmDeleteProduct = () => {
    if (deleteConfirmProductId) {
      const p = products.find(prod => prod.id === deleteConfirmProductId);
      if (p) {
        const updated = products.filter(prod => prod.id !== deleteConfirmProductId);
        setProducts(updated);
        saveToLocalStorage("KPI_DB_PRODUCTS", updated);
        showToast(`Produk ${p.name.toUpperCase()} telah dihapus dari database.`, "info");

        // Live sync to Sheets!
        if (sheetsScriptUrl) {
          handleSyncProductsToSheets(updated, true);
        }
      }
      setDeleteConfirmProductId(null);
    }
  };

  // Delete Audit Report from records
  const handleDeleteReport = (id: string) => {
    setDeleteConfirmId(id);
  };

  const confirmDeleteReport = () => {
    if (deleteConfirmId) {
      const updated = reports.filter(r => r.id !== deleteConfirmId);
      setReports(updated);
      saveToLocalStorage("KPI_DB_REPORTS", updated);
      showToast("Laporan telah dihapus dari riwayat.", "info");
      setDeleteConfirmId(null);
    }
  };

  // Filtered Products for Autocomplete Dropdown list
  const filteredProducts = products.filter(p =>
    p.name.toLowerCase().includes(productSearch.toLowerCase()) ||
    (p.skuCode && p.skuCode.toLowerCase().includes(productSearch.toLowerCase()))
  );

  // Filtered Salesmen for Autocomplete Dropdown list
  const filteredSalesmen = salesmen.filter(s =>
    s.name.toLowerCase().includes(salesmanSearch.toLowerCase()) ||
    (s.area && s.area.toLowerCase().includes(salesmanSearch.toLowerCase()))
  );

  // --- PROGRAM LOYALTI HANDLERS ---
  const handleSaveCustomerProfile = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newCustomer.name.trim()) {
      showToast("Nama toko harus diisi!", "error");
      return;
    }

    // Determine Loyalty Tier based on monthly volume estimation
    let tier: "Platinum" | "Gold" | "Silver" | "Bronze" = "Bronze";
    const omzetValue = Number(newCustomer.estimatedOmzet);
    if (omzetValue >= 15000000) tier = "Platinum";
    else if (omzetValue >= 8000000) tier = "Gold";
    else if (omzetValue >= 5000000) tier = "Silver";

    const customStore = {
      id: "c-" + Date.now(),
      name: newCustomer.name,
      address: newCustomer.address || "Alamat tidak spesifik",
      salesmanName: newCustomer.salesmanName || (salesmen[0]?.name || "RIZKY"),
      area: newCustomer.area || (salesmen[0]?.area || "Semarang"),
      jenisToko: newCustomer.jenisToko,
      estimatedOmzet: omzetValue,
      notesPerDay: Number(newCustomer.notesPerDay),
      storeAgeYears: Number(newCustomer.storeAgeYears),
      ownership: newCustomer.ownership,
      points: 50, // Starter voucher reward points
      tier,
      createdAt: new Date().toISOString(),
      actionsLog: []
    };

    const updated = [customStore, ...customers];
    setCustomers(updated);
    saveToLocalStorage("KPI_LOYALTY_CUSTOMERS", updated);
    setIsCustomerModalOpen(false);
    
    // Clear state
    setNewCustomer({
      name: "",
      address: "",
      salesmanName: "",
      area: "",
      jenisToko: "Sembako",
      estimatedOmzet: 5000000,
      notesPerDay: 5,
      storeAgeYears: 2,
      ownership: "Milik Sendiri"
    });
    showToast(`Toko ${customStore.name.toUpperCase()} berhasil diprofiling!`);

    // Dynamic Spreadsheet Auto Synchronization
    if (sheetsScriptUrl) {
      await handleSyncLoyaltyToSheets(updated, undefined, true);
    }
  };

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
    saveToLocalStorage("KPI_LOYALTY_CUSTOMERS", updated);
    
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
      await handleSyncLoyaltyToSheets(updated, undefined, true);
    }
  };

  const handleAddFollowUpAction = () => {
    if (!followUpNotes.trim()) {
      showToast("Tolong tuliskan rincian tindakan terlebih dahulu!", "error");
      return;
    }

    const updated = customers.map(c => {
      if (c.id === selectedCustomerIdForAction) {
        const logs = c.actionsLog || [];
        const newLog = {
          date: new Date().toISOString().split("T")[0],
          action: followUpAction,
          notes: followUpNotes,
          status: "Selesai"
        };
        // Give +15 Points as dynamic loyalty reward for completing DKR system actions
        return {
          ...c,
          points: c.points + 15,
          actionsLog: [newLog, ...logs]
        };
      }
      return c;
    });

    setCustomers(updated);
    saveToLocalStorage("KPI_LOYALTY_CUSTOMERS", updated);
    setFollowUpNotes("");
    showToast("Tindakan disimpan! Toko ini berhasil dianugerahi +15 Poin Loyalitas.");
  };

  const handleRedeemReward = async () => {
    const cust = customers.find(c => c.id === selectedCustomerIdForRedeem);
    const reward = [...rewardMerchants, ...katalogHadiah].find(r => r.id === selectedRewardId);

    if (!cust || !reward) {
      showToast("Data redeem tidak lengkap atau salah.", "error");
      return;
    }

    const pointsCost = 'pointsRequired' in reward ? reward.pointsRequired : (reward as CatalogHadiah).pointsValue;
    if ((cust.points || 0) < pointsCost) {
      showToast(`Poin ${cust.name} tidak mencukupi (${cust.points || 0}/${pointsCost} Poin)!`, "error");
      return;
    }

    const updatedCustomers = customers.map(c => {
      if (c.id === selectedCustomerIdForRedeem) {
        return { ...c, points: (c.points || 0) - pointsCost };
      }
      return c;
    });

    const redeemItem = {
      id: "rd-" + Date.now(),
      customerName: cust.name,
      rewardName: reward.name,
      pointsSpent: pointsCost,
      date: new Date().toISOString().split("T")[0],
      status: "Berhasil Diproses"
    };

    const updatedRedeems = [redeemItem, ...loyaltyRedeemHistory];

    setCustomers(updatedCustomers);
    saveToLocalStorage("KPI_LOYALTY_CUSTOMERS", updatedCustomers);
    setLoyaltyRedeemHistory(updatedRedeems);
    saveToLocalStorage("KPI_LOYALTY_REDEEMS", updatedRedeems);

    showToast(`Sukses klaim hadiah! Poin Toko ${cust.name} dikurangi ${pointsCost}.`);

    // Hubungkan & Sinkronkan otomatis ke Google Sheets!
    if (sheetsScriptUrl) {
      await handleSyncLoyaltyToSheets(updatedCustomers, updatedRedeems, true);
    }
  };

  const handleDeleteCustomer = (id: string, name: string) => {
    setDeleteConfirmCustomerId(id);
  };

  const confirmDeleteCustomer = async () => {
    if (deleteConfirmCustomerId) {
      const custToDel = customers.find(c => c.id === deleteConfirmCustomerId);
      if (custToDel) {
        const updated = customers.filter(c => c.id !== deleteConfirmCustomerId);
        setCustomers(updated);
        saveToLocalStorage("KPI_LOYALTY_CUSTOMERS", updated);
        showToast(`Profil Toko ${custToDel.name.toUpperCase()} berhasil dihapus!`, "info");
        
        // Sync with Google Sheets in background if URL is active
        if (sheetsScriptUrl) {
          await handleSyncLoyaltyToSheets(updated, undefined, true);
        }
      }
      setDeleteConfirmCustomerId(null);
    }
  };

  // --- CALCULATION HELPER STATS FOR TAB DISPLAYS ---
  const targetDatasetForKpi = kpiDataSource === "sheets" ? fetchedReports : reports;
  const filteredAuditReports = targetDatasetForKpi.filter((rep: any) => {
    let pass = true;
    if (auditFilterStartDate && rep.date < auditFilterStartDate) pass = false;
    if (auditFilterEndDate && rep.date > auditFilterEndDate) pass = false;
    return pass;
  });
  const totalAuditSaved = targetDatasetForKpi.length;
  const grandTotalCollection = targetDatasetForKpi.reduce((sum, r) => sum + r.billsReceived, 0);
  const avgSkuPerVisit = targetDatasetForKpi.length > 0 ? (targetDatasetForKpi.reduce((sum, r) => sum + r.skuTotal, 0) / targetDatasetForKpi.length).toFixed(1) : "0";

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (loginEmail === "sales@gmail.com" && loginPassword === "Sales#123") {
      setIsAuthenticated(true);
      localStorage.setItem("KPI_IS_AUTHENTICATED", "true");
      setLoginError("");
    } else {
      setLoginError("Email atau password tidak valid.");
    }
  };

  if (!isAuthenticated && !customerMode) {
    return (
      <div className="min-h-screen bg-[#FAF9F6] text-[#4A4A3C] font-sans flex items-center justify-center p-4">
        <div className="bg-white p-8 rounded-3xl shadow-xl border border-[#E5E5DF] max-w-sm w-full">
          <div className="flex justify-center mb-6">
            <div className="bg-[#E5E5DF]/30 p-4 rounded-2xl border border-[#E5E5DF]/50">
              <User className="w-10 h-10 text-[#5A5A40]" />
            </div>
          </div>
          <h1 className="text-2xl font-bold text-center uppercase tracking-tight text-zinc-900 mb-2">Login Portal</h1>
          <p className="text-center text-[#8C8C70] text-xs font-semibold mb-8">Silakan masuk untuk kelola Audit KPI Sales</p>
          
          {loginError && (
            <div className="bg-rose-50 text-rose-600 p-3 rounded-xl text-xs font-bold mb-6 text-center border border-rose-100">
              {loginError}
            </div>
          )}

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-2">Email</label>
              <input 
                type="email" 
                value={loginEmail}
                onChange={(e) => setLoginEmail(e.target.value)}
                className="w-full bg-[#FAF9F6] border border-[#E5E5DF] px-4 py-3 rounded-xl text-sm font-semibold text-[#4A4A3C] focus:border-[#5A5A40] focus:outline-hidden transition"
                placeholder="Masukkan email"
                required
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-2">Password</label>
              <input 
                type="password" 
                value={loginPassword}
                onChange={(e) => setLoginPassword(e.target.value)}
                className="w-full bg-[#FAF9F6] border border-[#E5E5DF] px-4 py-3 rounded-xl text-sm font-semibold text-[#4A4A3C] focus:border-[#5A5A40] focus:outline-hidden transition"
                placeholder="Masukkan password"
                required
              />
            </div>
            <button 
              type="submit"
              className="w-full bg-[#5A5A40] hover:bg-[#4A4A3C] text-white font-bold uppercase tracking-widest text-xs py-3.5 rounded-xl cursor-pointer transition shadow-md mt-6"
            >
              Masuk Sekarang
            </button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#FAF9F6] text-[#4A4A3C] font-sans selection:bg-[#5A5A40] selection:text-white flex flex-col md:flex-row">
      
      {/* Toast Notification pop up */}
      <AnimatePresence>
        {toast && (
          <motion.div
            initial={{ opacity: 0, y: -20, scale: 0.95 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: -10, scale: 0.95 }}
            className={`fixed top-4 left-1/2 -translate-x-1/2 z-50 flex items-center gap-3 px-5 py-3 rounded-xl shadow-xl border text-sm font-semibold max-w-md ${
              toast.type === "success"
                ? "bg-[#FAF9F6] text-[#5A5A40] border-[#8C8C70]/30 animate-pulse"
                : toast.type === "error"
                ? "bg-[#FAF9F6] text-rose-800 border-rose-200"
                : "bg-[#FAF9F6] text-[#8C8C70] border-[#E5E5DF]"
            }`}
          >
            {toast.type === "success" && <CheckCircle2 className="w-5 h-5 text-[#5A5A40]" />}
            {toast.type === "error" && <AlertTriangle className="w-5 h-5 text-rose-600" />}
            {toast.type === "info" && <Info className="w-5 h-5 text-[#8C8C70]" />}
            <span>{toast.message}</span>
          </motion.div>
        )}
      </AnimatePresence>

      {/* 1. COLLAPSIBLE SIDEBAR: ONLY SHOWS ON DESKTOP/LAPTOP SCREEN MODES */}
      {!customerMode && (
        <aside 
          className={`hidden md:flex flex-col bg-[#FAF9F6] border-r border-[#E5E5DF] h-screen sticky top-0 shrink-0 transition-all duration-300 z-40 ${
            isSidebarCollapsed ? "w-[76px]" : "w-64"
          } select-none`}
        >
        {/* Sidebar Brand & Collapse Toggle Button */}
        <div className="p-4 border-b border-[#E5E5DF] flex items-center justify-between gap-2 overflow-hidden">
          {!isSidebarCollapsed && (
            <div className="flex flex-col">
          <span className="text-base font-semibold tracking-tight text-zinc-900 uppercase whitespace-nowrap">
                PORTAL KPI
              </span>
              <span className="text-[9px] font-bold text-[#8C8C70] uppercase tracking-wider leading-none">
                Auditor Desk System
              </span>
            </div>
          )}
          {isSidebarCollapsed && (
            <div className="mx-auto p-1 bg-zinc-900 text-white rounded-lg text-xs font-bold">
              AKP
            </div>
          )}
          
          <button
            onClick={() => {
              const newVal = !isSidebarCollapsed;
              setIsSidebarCollapsed(newVal);
              localStorage.setItem("KPI_SIDEBAR_COLLAPSED", newVal ? "true" : "false");
            }}
            className="p-1.5 hover:bg-[#E5E5DF]/50 text-[#8C8C70] hover:text-[#4A4A3C] rounded-lg transition shrink-0 cursor-pointer"
            title={isSidebarCollapsed ? "Expand Sidebar" : "Minimize Sidebar"}
          >
            {isSidebarCollapsed ? <ChevronRight className="w-4 h-4" /> : <ChevronLeft className="w-4 h-4" />}
          </button>
        </div>

        {/* Auditor Profile Header Section (Only if Expanded) */}
        {!isSidebarCollapsed && (
          <div className="p-4 bg-[#E5E5DF]/20 border-b border-[#E5E5DF]/40 flex items-center gap-3">
            <div className="w-9 h-9 rounded-full bg-[#5A5A40] text-xs font-serif italic text-white flex items-center justify-center font-bold">
              AR
            </div>
            <div className="min-w-0">
              <p className="text-xs font-black text-[#4A4A3C] truncate uppercase">Aris</p>
              <p className="text-[10px] text-[#8C8C70] font-semibold uppercase tracking-wider">Kepala Sales</p>
            </div>
          </div>
        )}

        {/* Sidebar Navigation Items Vertical List */}
        <nav className="flex-1 p-3 space-y-1.5 mt-2">
          {/* Item 1: Input KPI */}
          <button
            onClick={() => setActiveTab("form")}
            className={`w-full flex items-center gap-3 px-3 py-2.5 text-xs font-extrabold uppercase tracking-wider rounded-xl transition-all ${
              activeTab === "form"
                ? "bg-[#5A5A40] text-[#FAF9F6] shadow-sm font-bold"
                : "text-[#8C8C70] hover:text-[#4A4A3C] hover:bg-[#E5E5DF]/30"
            }`}
            title="Input Laporan KPI"
          >
            <FileText className="w-4 h-4 shrink-0" />
            {!isSidebarCollapsed && <span className="truncate">Input KPI Portal</span>}
          </button>

          {/* Item 6 - NEW: Program Loyalti */}
          <button
            onClick={() => setActiveTab("loyalty")}
            className={`w-full flex items-center justify-between px-3 py-2.5 text-xs font-extrabold uppercase tracking-wider rounded-xl transition-all ${
              activeTab === "loyalty"
                ? "bg-[#5A5A40] text-[#FAF9F6] shadow-sm font-bold"
                : "text-[#8C8C70] hover:text-[#4A4A3C] hover:bg-[#E5E5DF]/30"
            }`}
            title="Program Loyalti Toko"
          >
            <div className="flex items-center gap-3 min-w-0">
              <Crown className="w-4 h-4 shrink-0 text-amber-500" />
              {!isSidebarCollapsed && <span className="truncate select-none">Program Loyalti</span>}
            </div>
            {!isSidebarCollapsed && (
              <span className="text-[9px] bg-amber-500/20 text-amber-800 font-black px-1.5 py-0.5 rounded-md">
                Baru
              </span>
            )}
          </button>

          {/* Item 8: Klaim Hadiah */}
          <button
            onClick={() => setActiveTab("claims")}
            className={`w-full flex items-center justify-between px-3 py-2.5 text-xs font-extrabold uppercase tracking-wider rounded-xl transition-all ${
              activeTab === "claims"
                ? "bg-[#5A5A40] text-[#FAF9F6] shadow-sm font-bold"
                : "text-[#8C8C70] hover:text-[#4A4A3C] hover:bg-[#E5E5DF]/30"
            }`}
            title="Klaim Hadiah"
          >
            <div className="flex items-center gap-3 min-w-0">
              <Gift className="w-4 h-4 shrink-0 text-emerald-600" />
              {!isSidebarCollapsed && <span className="truncate select-none">Klaim Hadiah</span>}
            </div>
            {!isSidebarCollapsed && (
              <span className="text-[9px] bg-emerald-500/20 text-emerald-800 font-black px-1.5 py-0.5 rounded-md">
                Klaim
              </span>
            )}
          </button>

          {/* Item 7 - NEW: KPI Sales Dashboard */}
          <button
            onClick={() => setActiveTab("kpisales")}
            className={`w-full flex items-center justify-between px-3 py-2.5 text-xs font-extrabold uppercase tracking-wider rounded-xl transition-all ${
              activeTab === "kpisales"
                ? "bg-[#5A5A40] text-[#FAF9F6] shadow-sm font-bold"
                : "text-[#8C8C70] hover:text-[#4A4A3C] hover:bg-[#E5E5DF]/30"
            }`}
            title="KPI Sales Tab"
          >
            <div className="flex items-center gap-3 min-w-0">
              <TrendingUp className="w-4 h-4 shrink-0 text-rose-600" />
              {!isSidebarCollapsed && <span className="truncate">KPI Sales</span>}
            </div>
            {!isSidebarCollapsed && (
              <span className="text-[9px] bg-rose-500/20 text-rose-800 font-black px-1.5 py-0.5 rounded-md">
                Target
              </span>
            )}
          </button>

          {/* Item 2: Database Salesman */}
          <button
            onClick={() => setActiveTab("salesmen")}
            className={`w-full flex items-center gap-3 px-3 py-2.5 text-xs font-extrabold uppercase tracking-wider rounded-xl transition-all ${
              activeTab === "salesmen"
                ? "bg-[#5A5A40] text-[#FAF9F6] shadow-sm font-bold"
                : "text-[#8C8C70] hover:text-[#4A4A3C] hover:bg-[#E5E5DF]/30"
            }`}
            title="Database Salesman"
          >
            <User className="w-4 h-4 shrink-0" />
            {!isSidebarCollapsed && <span className="truncate">Data Salesman</span>}
          </button>

          {/* Item 2.1: Data Pelanggan */}
          <button
            onClick={() => setActiveTab("customerData")}
            className={`w-full flex items-center gap-3 px-3 py-2.5 text-xs font-extrabold uppercase tracking-wider rounded-xl transition-all ${
              activeTab === "customerData"
                ? "bg-[#5A5A40] text-[#FAF9F6] shadow-sm font-bold"
                : "text-[#8C8C70] hover:text-[#4A4A3C] hover:bg-[#E5E5DF]/30"
            }`}
            title="Data Pelanggan"
          >
            <Users className="w-4 h-4 shrink-0" />
            {!isSidebarCollapsed && <span className="truncate">Data Pelanggan</span>}
          </button>

          {/* Item 3: Database Produk */}
          <button
            onClick={() => setActiveTab("products")}
            className={`w-full flex items-center gap-3 px-3 py-2.5 text-xs font-extrabold uppercase tracking-wider rounded-xl transition-all ${
              activeTab === "products"
                ? "bg-[#5A5A40] text-[#FAF9F6] shadow-sm font-bold"
                : "text-[#8C8C70] hover:text-[#4A4A3C] hover:bg-[#E5E5DF]/30"
            }`}
            title="Database SKU Produk"
          >
            <ShoppingBag className="w-4 h-4 shrink-0" />
            {!isSidebarCollapsed && <span className="truncate">Data Produk</span>}
          </button>

          {/* Item 4: Riwayat Audit */}
          <button
            onClick={() => setActiveTab("reports")}
            className={`w-full flex items-center justify-between px-3 py-2.5 text-xs font-extrabold uppercase tracking-wider rounded-xl transition-all ${
              activeTab === "reports"
                ? "bg-[#5A5A40] text-[#FAF9F6] shadow-sm font-bold"
                : "text-[#8C8C70] hover:text-[#4A4A3C] hover:bg-[#E5E5DF]/30"
            }`}
            title="Riwayat Laporan Audit"
          >
            <div className="flex items-center gap-3 min-w-0">
              <History className="w-4 h-4 shrink-0" />
              {!isSidebarCollapsed && <span className="truncate">Riwayat Audit</span>}
            </div>
            {!isSidebarCollapsed && (
              <span className="text-[10px] bg-[#4A4A3C]/10 text-[#4A4A3C] px-2 py-0.5 rounded-md font-mono">
                {targetDatasetForKpi.length}
              </span>
            )}
          </button>

          {/* Item 5: Google Sheets Sync */}
          <button
            onClick={() => setActiveTab("sheets")}
            className={`w-full flex items-center gap-3 px-3 py-2.5 text-xs font-extrabold uppercase tracking-wider rounded-xl transition-all ${
              activeTab === "sheets"
                ? "bg-[#5A5A40] text-[#FAF9F6] shadow-sm font-bold"
                : "text-[#8C8C70] hover:text-[#4A4A3C] hover:bg-[#E5E5DF]/30"
            }`}
            title="Google Sheets Sync"
          >
            <Layers className="w-4 h-4 shrink-0" />
            {!isSidebarCollapsed && <span className="truncate">Google Sheets</span>}
          </button>

          {/* Logout Button */}
          <div className="border-t border-[#E5E5DF] my-2 pt-2">
            <button
              onClick={() => {
                setIsAuthenticated(false);
                localStorage.setItem("KPI_IS_AUTHENTICATED", "false");
                setLoginPassword("");
              }}
              className="w-full flex items-center gap-3 px-3 py-2.5 text-xs font-extrabold uppercase tracking-wider rounded-xl transition-all text-rose-600 hover:bg-rose-50"
              title="Keluar / Logout"
            >
              <LogOut className="w-4 h-4 shrink-0 cursor-pointer" />
              {!isSidebarCollapsed && <span className="truncate">Keluar</span>}
            </button>
          </div>
        </nav>

        {/* Sidebar Footer System Badge */}
        <div className="p-4 border-t border-[#E5E5DF] text-center text-[10px] text-[#8C8C70] shrink-0">
          {!isSidebarCollapsed && (
            <div className="flex flex-col gap-0.5">
              <p className="font-bold text-[#5A5A40]">DKR SALES SYSTEM</p>
              <p className="font-mono text-[9px]">v1.4.0 • Live</p>
            </div>
          )}
          {isSidebarCollapsed && (
            <span className="font-mono font-bold text-[#5A5A40] text-[10px]">1.4</span>
          )}
        </div>
      </aside>
      )}

      {/* 2. THE MAIN PORTAL SCREEN AREA & RESPONSIVE COLUMN */}
      <div className="flex-1 flex flex-col min-w-0">
        
        {/* RESPONSIVE HEADER BANNER: INTERACTS MAINLY ON SMARTPHONES */}
        {!customerMode && (
          <header className="bg-[#FAF9F6] border-b border-[#E5E5DF] md:hidden sticky top-0 z-30 shadow-xs">
          <div className="px-4 py-3 flex items-center justify-between">
            {/* Branding */}
            <div>
              <span className="text-lg font-serif italic font-extrabold text-[#4A4A3C] uppercase">
                PORTAL KPI
              </span>
              <p className="text-[8px] font-bold text-[#8C8C70] uppercase tracking-widest leading-none mt-0.5">
                Auditor Portal
              </p>
            </div>

            {/* Hamburger Toggle menu button for smartphone */}
            <button
              onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
              className="p-2 hover:bg-[#E5E5DF]/50 text-[#5A5A40] rounded-xl transition cursor-pointer border border-[#E5E5DF]"
              aria-label="Toggle Mobile Menu"
            >
              <Menu className="w-5 h-5" />
            </button>
          </div>

          {/* 
            MOBILE NAVIGATION PANEL (Stacks dropdown buttons vertically downwards 
            guaranteeing "terapkan vertical kebawah" on smartphone devices)
          */}
          <AnimatePresence>
            {isMobileMenuOpen && (
              <motion.nav 
                initial={{ height: 0, opacity: 0 }}
                animate={{ height: "auto", opacity: 1 }}
                exit={{ height: 0, opacity: 0 }}
                className="border-t border-[#E5E5DF]/60 bg-[#FAF9F6] px-4 py-3 flex flex-col gap-2 shadow-inner overflow-hidden"
              >
                {/* 1. Input KPI Tab Button */}
                <button
                  onClick={() => {
                    setActiveTab("form");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full py-2.5 px-4 rounded-xl text-left text-xs font-black uppercase tracking-wider flex items-center gap-3 ${
                    activeTab === "form" ? "bg-[#5A5A40] text-[#FAF9F6]" : "bg-[#E5E5DF]/20 text-[#8C8C70]"
                  }`}
                >
                  <FileText className="w-4 h-4" />
                  Input KPI Portal
                </button>

                {/* 2. Program Loyalti Tab Button (New!) */}
                <button
                  onClick={() => {
                    setActiveTab("loyalty");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full py-2.5 px-4 rounded-xl text-left text-xs font-black uppercase tracking-wider flex items-center justify-between ${
                    activeTab === "loyalty" ? "bg-[#5A5A40] text-[#FAF9F6] border border-[#5A5A40]" : "bg-gradient-to-r from-amber-500/10 to-[#E5E5DF]/20 text-amber-900"
                  }`}
                >
                  <div className="flex items-center gap-3">
                    <Crown className="w-4 h-4 text-amber-500" />
                    <span>Program Loyalti</span>
                  </div>
                  <span className="text-[8px] bg-amber-500 text-white font-extrabold px-1.5 py-0.5 rounded uppercase">
                    New Loyalty
                  </span>
                </button>

                {/* 2b. Klaim Hadiah Tab Button */}
                <button
                  onClick={() => {
                    setActiveTab("claims");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full py-2.5 px-4 rounded-xl text-left text-xs font-black uppercase tracking-wider flex items-center justify-between ${
                    activeTab === "claims" ? "bg-[#5A5A40] text-[#FAF9F6] border border-[#5A5A40]" : "bg-gradient-to-r from-emerald-500/10 to-[#E5E5DF]/20 text-emerald-900"
                  }`}
                >
                  <div className="flex items-center gap-3">
                    <Gift className="w-4 h-4 text-emerald-600" />
                    <span>Klaim Hadiah</span>
                  </div>
                  <span className="text-[8px] bg-emerald-600 text-white font-extrabold px-1.5 py-0.5 rounded uppercase">
                    Klaim
                  </span>
                </button>

                {/* KPI Sales Tab Button */}
                <button
                  onClick={() => {
                    setActiveTab("kpisales");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full py-2.5 px-4 rounded-xl text-left text-xs font-black uppercase tracking-wider flex items-center justify-between ${
                    activeTab === "kpisales" ? "bg-[#5A5A40] text-[#FAF9F6]" : "bg-gradient-to-r from-rose-500/10 to-[#E5E5DF]/20 text-rose-900"
                  }`}
                >
                  <div className="flex items-center gap-3">
                    <TrendingUp className="w-4 h-4 text-rose-600" />
                    <span>KPI Sales</span>
                  </div>
                  <span className="text-[8px] bg-rose-500 text-white font-extrabold px-1.5 py-0.5 rounded uppercase">
                    Tarif
                  </span>
                </button>

                {/* 3. Database Salesman */}
                <button
                  onClick={() => {
                    setActiveTab("salesmen");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full py-2.5 px-4 rounded-xl text-left text-xs font-black uppercase tracking-wider flex items-center gap-3 ${
                    activeTab === "salesmen" ? "bg-[#5A5A40] text-[#FAF9F6]" : "bg-[#E5E5DF]/20 text-[#8C8C70]"
                  }`}
                >
                  <User className="w-4 h-4" />
                  Database Salesman
                </button>

                {/* 3.1. Data Pelanggan */}
                <button
                  onClick={() => {
                    setActiveTab("customerData");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full py-2.5 px-4 rounded-xl text-left text-xs font-black uppercase tracking-wider flex items-center gap-3 ${
                    activeTab === "customerData" ? "bg-[#5A5A40] text-[#FAF9F6]" : "bg-[#E5E5DF]/20 text-[#8C8C70]"
                  }`}
                >
                  <Users className="w-4 h-4" />
                  Data Pelanggan
                </button>

                {/* 4. Database Produk */}
                <button
                  onClick={() => {
                    setActiveTab("products");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full py-2.5 px-4 rounded-xl text-left text-xs font-black uppercase tracking-wider flex items-center gap-3 ${
                    activeTab === "products" ? "bg-[#5A5A40] text-[#FAF9F6]" : "bg-[#E5E5DF]/20 text-[#8C8C70]"
                  }`}
                >
                  <ShoppingBag className="w-4 h-4" />
                  Database Produk
                </button>

                {/* 5. Riwayat Audit */}
                <button
                  onClick={() => {
                    setActiveTab("reports");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full py-2.5 px-4 rounded-xl text-left text-xs font-black uppercase tracking-wider flex items-center justify-between ${
                    activeTab === "reports" ? "bg-[#5A5A40] text-[#FAF9F6]" : "bg-[#E5E5DF]/20 text-[#8C8C70]"
                  }`}
                >
                  <div className="flex items-center gap-3">
                    <History className="w-4 h-4" />
                    <span>Riwayat Audit</span>
                  </div>
                  <span className="text-[10px] bg-[#4A4A3C]/15 text-[#4A4A3C] px-2 py-0.5 rounded font-mono font-bold">
                    {targetDatasetForKpi.length} Records
                  </span>
                </button>

                {/* 6. Google Sheets */}
                <button
                  onClick={() => {
                    setActiveTab("sheets");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full py-2.5 px-4 rounded-xl text-left text-xs font-black uppercase tracking-wider flex items-center gap-3 ${
                    activeTab === "sheets" ? "bg-[#5A5A40] text-[#FAF9F6]" : "bg-[#E5E5DF]/20 text-[#8C8C70]"
                  }`}
                >
                  <Layers className="w-4 h-4" />
                  Google Sheets Integration
                </button>

                {/* Logout Button */}
                <div className="border-t border-[#E5E5DF]/50 my-1 pt-1">
                  <button
                    onClick={() => {
                      setIsAuthenticated(false);
                      localStorage.setItem("KPI_IS_AUTHENTICATED", "false");
                      setLoginPassword("");
                      setIsMobileMenuOpen(false);
                    }}
                    className="w-full py-2.5 px-4 rounded-xl text-left text-xs font-black uppercase tracking-wider flex items-center gap-3 text-rose-600 hover:bg-rose-50 cursor-pointer"
                  >
                    <LogOut className="w-4 h-4 cursor-pointer" />
                    Keluar / Logout
                  </button>
                </div>
              </motion.nav>
            )}
          </AnimatePresence>
        </header>
        )}

        {/* Global Page Layout Panel and Wrapper Body */}
        <div className="max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6">
        
        {/* UPPER KPI SUMMARY STRIP */}
        <section className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
          <div className="bg-[#FAF9F6] p-4 rounded-2xl border border-[#E5E5DF] shadow-xs">
            <span className="text-xs font-bold text-[#8C8C70] block uppercase tracking-wider">Total Laporan</span>
            <span className="text-xl font-extrabold text-[#4A4A3C] block mt-1">{totalAuditSaved} Dokumen</span>
            <span className="text-[10px] text-[#8C8C70]/80">Tersimpan di LocalStorage</span>
          </div>
          <div className="bg-[#FAF9F6] p-4 rounded-2xl border border-[#E5E5DF] shadow-xs">
            <span className="text-xs font-bold text-[#8C8C70] block uppercase tracking-wider">Total Dana Tagihan</span>
            <span className="text-xl font-extrabold text-[#5A5A40] block mt-1">Rp {grandTotalCollection.toLocaleString("id-ID")}</span>
            <span className="text-[10px] text-emerald-700/80 font-semibold">Berhasil ditagih sales</span>
          </div>
          <div className="bg-[#FAF9F6] p-4 rounded-2xl border border-[#E5E5DF] shadow-xs">
            <span className="text-xs font-bold text-[#8C8C70] block uppercase tracking-wider">Estimasi Rata SKU</span>
            <span className="text-xl font-extrabold text-[#8C8C70] block mt-1">{avgSkuPerVisit} SKU</span>
            <span className="text-[10px] text-[#8C8C70]/80">Kapasitas penawaran efektif</span>
          </div>
          <div className="bg-gradient-to-br from-[#5A5A40] to-[#4A4A3C] p-4 rounded-2xl text-[#FAF9F6] shadow-sm">
            <span className="text-xs font-bold block uppercase tracking-wider text-[#FAF9F6]/80">Database Salesman</span>
            <span className="text-xl font-extrabold block mt-1">{salesmen.length} Anggota</span>
            <span className="text-[10px] text-[#FAF9F6]/70 italic">Pilihan dropdown otomatis</span>
          </div>
        </section>

        {/* CONTROLLER SECTION */}
        <AnimatePresence mode="wait">
          
          {/* TAB 1: FORM INPUT KPI (MIRRORS SCREENSHOT EXACTLY) */}
          {activeTab === "form" && (
            <motion.div
              layout
              key="form-tab"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.25 }}
              className="max-w-4xl mx-auto w-full flex flex-col gap-6"
            >
              {/* THE MAIN INPUT FORM SECTION */}
              <div className="w-full flex flex-col gap-6">
                
                {/* NATURAL TONES HEADER CARD COMPONENT - EXACTLY "BANTU INPUT LAPORAN" */}
                <div className="bg-gradient-to-r from-[#5A5A40] via-[#8C8C70] to-[#5A5A40] text-[#FAF9F6] rounded-3xl p-6 shadow-sm flex items-center justify-between border border-[#E5E5DF]">
                  <div className="flex items-center gap-4">
                    <div className="w-14 h-14 rounded-2xl bg-white/10 backdrop-blur-md flex items-center justify-center border border-white/20">
                      <User className="w-7 h-7 text-[#FAF9F6]" />
                    </div>
                    <div>
                      <h2 className="text-xl font-serif italic font-bold tracking-tight uppercase">
                        Bantu Input Laporan
                      </h2>
                      <p className="text-[10px] text-[#FAF9F6]/85 font-semibold tracking-widest uppercase mt-0.5">
                        SALESMAN YANG DIINPUT: <span className="text-[#FAF9F6] font-bold underline decoration-dotted decoration-2 underline-offset-4 transition-all">{activeSalesman ? activeSalesman.name : "BELUM DIPILIH"}</span>
                      </p>
                    </div>
                  </div>
                </div>

                {/* THE PORTAL FORM */}
                <form onSubmit={handleSubmitReport} className="bg-[#FAF9F6] rounded-3xl p-6 border border-[#E5E5DF] shadow-xs flex flex-col gap-6">
                  
                  {/* ROW 1: SELECT DATE */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-widest mb-1.5">
                        Tanggal Penjualan
                      </label>
                      <input
                        type="date"
                        value={reportDate}
                        onChange={(e) => setReportDate(e.target.value)}
                        className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl px-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] focus:bg-white text-[#4A4A3C] font-medium transition"
                        required
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-widest mb-1.5 flex items-center justify-between">
                        <span>Siklus Hari Kunjungan *</span>
                        <button
                          type="button"
                          onClick={handleAutoDetectCycle}
                          className="text-[10px] font-bold text-[#5A5A40] bg-[#E5E5DF]/55 hover:bg-[#E5E5DF] px-2.5 py-1 rounded-md flex items-center gap-1 uppercase transition"
                        >
                          <Sparkles className="w-3 h-3 text-[#5A5A40]" />
                          Auto-Detect
                        </button>
                      </label>
                      <select
                        value={selectedCycle}
                        onChange={(e) => setSelectedCycle(e.target.value)}
                        className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl px-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] focus:bg-white text-[#4A4A3C] font-semibold transition"
                        required
                      >
                        {STANDARD_CYCLES.map(c => (
                          <option key={c} value={c}>{c}</option>
                        ))}
                      </select>
                    </div>

                    <div className="w-full">
                      <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-widest mb-1.5 mt-4">
                        AREA *
                      </label>
                      <select
                        value={area}
                        onChange={(e) => setArea(e.target.value)}
                        className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl px-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] focus:bg-white text-[#4A4A3C] font-semibold transition"
                        required
                      >
                        {availableAreas.map(a => (
                          <option key={a} value={a}>{a}</option>
                        ))}
                      </select>
                    </div>
                  </div>

                  {/* ROW 2: SEARCHABLE DROPDOWN CHANNELS FOR SALESMAN (AS DEMANDED) */}
                  <div className="relative" ref={salesmanDropdownRef}>
                    <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-widest mb-1.5">
                      PILIH SALESMAN (BANTU INPUT) *
                    </label>
                    <div
                      onClick={() => setIsSalesmanDropdownOpen(true)}
                      className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-2xl p-4 flex items-center justify-between cursor-pointer hover:border-[#8C8C70] hover:bg-[#FAF9F6]/50 transition duration-150"
                    >
                      <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-xl bg-[#E5E5DF]/40 flex items-center justify-center">
                          <User className="w-5 h-5 text-[#5A5A40]" />
                        </div>
                        <div>
                          <p className="text-sm font-bold text-[#4A4A3C]">
                            {activeSalesman ? activeSalesman.name : "Klik untuk memilih salesman..."}
                          </p>
                          <p className="text-xs text-[#8C8C70]">
                            {activeSalesman ? `${activeSalesman.area || "Tanpa Area"} • ${activeSalesman.phone || "No HP -"}` : "Menampakkan dropdown database sederhana"}
                          </p>
                        </div>
                      </div>
                      <div className="text-[#8C8C70] font-bold text-xs pointer-events-none">▼</div>
                    </div>

                    {/* RENDER FLOATING SALESMAN AUTOCOMPLETE PANEL */}
                    <AnimatePresence>
                      {isSalesmanDropdownOpen && (
                        <motion.div
                          initial={{ opacity: 0, y: 10 }}
                          animate={{ opacity: 1, y: 0 }}
                          exit={{ opacity: 0, y: 5 }}
                          className="absolute left-0 right-0 mt-2 bg-[#FAF9F6] border border-[#E5E5DF] rounded-2xl shadow-xl z-40 max-h-72 overflow-y-auto"
                        >
                          <div className="p-3 border-b border-[#E5E5DF] flex items-center gap-2 sticky top-0 bg-[#FAF9F6] z-10">
                            <Search className="w-4 h-4 text-[#8C8C70]" />
                            <input
                              type="text"
                              placeholder="Ketik/Cari nama atau area sales..."
                              value={salesmanSearch}
                              onChange={(e) => setSalesmanSearch(e.target.value)}
                              className="w-full text-xs focus:outline-hidden text-[#4A4A3C] bg-transparent py-1 font-semibold"
                              autoFocus
                            />
                            {salesmanSearch && (
                              <button type="button" onClick={() => setSalesmanSearch("")} className="text-xs text-[#8C8C70] hover:text-[#4A4A3C]">
                                <X className="w-4 h-4" />
                              </button>
                            )}
                          </div>
                          
                          <div className="divide-y divide-[#E5E5DF]">
                            {filteredSalesmen.map((sl) => (
                              <div
                                key={sl.id}
                                onClick={() => {
                                  setSelectedSalesmanId(sl.id);
                                  setIsSalesmanDropdownOpen(false);
                                }}
                                className={`p-3.5 hover:bg-[#E5E5DF]/40 cursor-pointer flex items-center justify-between transition ${
                                  selectedSalesmanId === sl.id ? "bg-[#E5E5DF]/60" : ""
                                }`}
                              >
                                <div>
                                  <p className="text-sm font-bold text-[#4A4A3C]">{sl.name}</p>
                                  <p className="text-[11px] text-[#8C8C70] font-medium">Area: {sl.area || "-"} • Tlp: {sl.phone || "-"}</p>
                                </div>
                                {selectedSalesmanId === sl.id && (
                                  <span className="w-2.5 h-2.5 rounded-full bg-[#5A5A40]"></span>
                                )}
                              </div>
                            ))}

                            {filteredSalesmen.length === 0 && (
                              <div className="p-6 text-center">
                                <p className="text-xs text-[#8C8C70] font-semibold uppercase">Salesman tidak ditemukan</p>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setActiveTab("salesmen");
                                    setIsSalesmanDropdownOpen(false);
                                    handleAddSalesman();
                                  }}
                                  className="mt-2 text-xs text-[#5A5A40] hover:underline font-bold flex items-center justify-center gap-1 mx-auto"
                                >
                                  <Plus className="w-3.5 h-3.5" /> Tambah Salesman Baru ke DB
                                </button>
                              </div>
                            )}
                          </div>
                        </motion.div>
                      )}
                    </AnimatePresence>
                  </div>

                  {/* ROW 3: FLOATING DECORATIVE PRODUCT REGISTRY (FROM PILIH SALESFORCE / CARI PRODUK DROPDOWN ON THE RIGHT OF PICTURE) */}
                  <div className="bg-[#FAF9F6] rounded-2xl p-4 border border-[#E5E5DF]/60">
                    <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 mb-3">
                      <div>
                        <span className="text-xs font-bold text-[#8C8C70] uppercase tracking-widest block">
                          Uraian Detail Penjualan Produk (Bantuan Kalkulator SKU)
                        </span>
                        <p className="text-[10px] text-[#8C8C70]/85">
                          Pilih produk yang terjual untuk otomatis menghitung nilai SKU Total Anda
                        </p>
                      </div>

                      {/* Dropdown Cari Produk Trigger */}
                      <div className="relative w-full sm:w-auto" ref={productPickerRef}>
                        <button
                          type="button"
                          onClick={() => setIsProductPickerOpen(!isProductPickerOpen)}
                          className="w-full sm:w-auto px-4 py-2 bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl text-xs font-bold text-[#5A5A40] hover:bg-[#E5E5DF]/30 flex items-center justify-center gap-2 shadow-xs cursor-pointer transition"
                        >
                          <Search className="w-3.5 h-3.5 text-[#5A5A40]" />
                          Cari & Pilih Produk...
                        </button>

                        {/* Floating Product Autocomplete exact mirror from screenshot */}
                        <AnimatePresence>
                          {isProductPickerOpen && (
                            <motion.div
                              initial={{ opacity: 0, scale: 0.95, y: 10 }}
                              animate={{ opacity: 1, scale: 1, y: 0 }}
                              exit={{ opacity: 0, scale: 0.95, y: 5 }}
                              className="absolute right-0 mt-2 w-80 bg-[#FAF9F6] border border-[#E5E5DF] rounded-2xl shadow-xl z-50 overflow-hidden"
                            >
                              {/* Autocomplete Input Search */}
                              <div className="p-3 border-b border-[#E5E5DF] bg-[#FAF9F6] flex items-center gap-1.5">
                                <Search className="w-4 h-4 text-[#8C8C70]" />
                                <input
                                  type="text"
                                  placeholder="Cari Produk..."
                                  value={productSearch}
                                  onChange={(e) => setProductSearch(e.target.value)}
                                  className="w-full bg-transparent text-xs focus:outline-hidden text-[#4A4A3C] font-semibold placeholder-[#8C8C70]/70"
                                  autoFocus
                                />
                                {productSearch && (
                                  <button type="button" onClick={() => setProductSearch("")}>
                                    <X className="w-3.5 h-3.5 text-[#8C8C70] hover:text-[#4A4A3C]" />
                                  </button>
                                )}
                              </div>

                              {/* Suggestion list */}
                              <div className="max-h-60 overflow-y-auto divide-y divide-[#E5E5DF]/45">
                                {filteredProducts.map((p) => (
                                  <button
                                    key={p.id}
                                    type="button"
                                    onClick={() => handleSelectProduct(p)}
                                    className="w-full text-left px-4 py-3 hover:bg-[#E5E5DF]/40 hover:text-[#4A4A3C] transition flex items-center justify-between text-xs"
                                  >
                                    <div>
                                      <p className="font-bold text-[#4A4A3C] uppercase tracking-tight">{p.name}</p>
                                      {p.category && (
                                        <p className="text-[10px] text-[#8C8C70]">{p.category} • {p.skuCode || "-"}</p>
                                      )}
                                    </div>
                                  </button>
                                ))}

                                {filteredProducts.length === 0 && (
                                  <div className="p-4 text-center text-[#8C8C70] text-xs">
                                    Tidak ada produk yang cocok
                                  </div>
                                )}
                              </div>
                            </motion.div>
                          )}
                        </AnimatePresence>
                      </div>
                    </div>

                    {/* Show chosen products details */}
                    {Object.keys(selectedProducts).length > 0 ? (
                      <div className="flex flex-col gap-2 max-h-48 overflow-y-auto">
                        {Object.entries(selectedProducts).map(([pId, qty]) => {
                          const p = products.find(prod => prod.id === pId);
                          if (!p) return null;
                          return (
                            <div key={pId} className="flex items-center justify-between bg-[#FAF9F6] px-3.5 py-2.5 rounded-xl border border-[#E5E5DF]">
                              <div>
                                <h4 className="text-xs font-bold text-[#4A4A3C] uppercase">{p.name}</h4>
                              </div>
                              
                              <div className="flex items-center gap-2">
                                <button
                                  type="button"
                                  onClick={() => handleUpdateProductQty(pId, 0)}
                                  className="w-8 h-8 rounded-md bg-rose-50 text-rose-600 flex items-center justify-center hover:bg-rose-100 transition"
                                >
                                  <X className="w-4 h-4" />
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    ) : (
                      <div className="p-4 bg-[#FAF9F6] rounded-xl border border-[#E5E5DF] text-center text-[#8C8C70]/70 text-xs italic">
                        Belum ada produk rincian yang dipilih. Nilai SKU Total bisa diisi manual atau otomatis di bawah.
                      </div>
                    )}
                  </div>

                  {/* ROW 4: KPI METRICS (TC, CP, EC, SKU TOTAL) WITH EXACT COLOR ACCENTS AND PROPORTIONS */}
                  <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
                    
                    {/* TC */}
                    <div className="bg-[#FAF9F6] p-4 rounded-2xl border border-[#E5E5DF]">
                      <label className="block text-xs font-bold text-[#4A4A3C] uppercase tracking-wider mb-2">
                        TC (AMPLOP) *
                      </label>
                      <input
                        type="number"
                        min="0"
                        onFocus={(e) => e.target.value === "0" && e.target.select()}
                        value={tc}
                        onChange={(e) => setTc(parseInt(e.target.value, 10) || 0)}
                        className="w-full bg-[#FAF9F6] text-center rounded-xl py-3 text-lg font-mono font-black text-[#5A5A40] border border-[#E5E5DF] focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40]"
                        required
                      />
                    </div>

                    {/* CP */}
                    <div className="bg-[#FAF9F6] p-4 rounded-2xl border border-[#E5E5DF]">
                      <label className="block text-xs font-bold text-[#4A4A3C] uppercase tracking-wider mb-2">
                        CP (KUNJUNGAN) *
                      </label>
                      <input
                        type="number"
                        min="0"
                        onFocus={(e) => e.target.value === "0" && e.target.select()}
                        value={cp}
                        onChange={(e) => setCp(parseInt(e.target.value, 10) || 0)}
                        className="w-full bg-[#FAF9F6] text-center rounded-xl py-3 text-lg font-mono font-black text-[#5A5A40] border border-[#E5E5DF] focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40]"
                        required
                      />
                    </div>

                    {/* EC (ORDER) */}
                    <div className="bg-[#E5E5DF]/40 p-4 rounded-2xl border border-[#8C8C70]/50">
                      <label className="block text-xs font-bold text-[#4A4A3C] uppercase tracking-wider mb-2">
                        EC (ORDER) *
                      </label>
                      <input
                        type="number"
                        min="0"
                        onFocus={(e) => e.target.value === "0" && e.target.select()}
                        value={ec}
                        onChange={(e) => setEc(parseInt(e.target.value, 10) || 0)}
                        className="w-full bg-[#FAF9F6] text-center rounded-xl py-3 text-lg font-mono font-black text-[#5A5A40] border border-[#E5E5DF] focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40]"
                        required
                      />
                    </div>

                    {/* SKU TOTAL */}
                    <div className="bg-[#E5E5DF]/40 p-4 rounded-2xl border border-[#8C8C70]/50">
                      <label className="block text-xs font-bold text-[#4A4A3C] uppercase tracking-wider mb-2">
                        SKU TOTAL *
                      </label>
                      <input
                        type="number"
                        min="0"
                        onFocus={(e) => e.target.value === "0" && e.target.select()}
                        value={skuTotal}
                        onChange={(e) => setSkuTotal(parseInt(e.target.value, 10) || 0)}
                        className="w-full bg-[#FAF9F6] text-center rounded-xl py-3 text-lg font-mono font-black text-[#5A5A40] border border-[#E5E5DF] focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40]"
                        required
                      />
                    </div>

                  </div>

                  {/* CONDITIONAL NOO INPUTS FOR IMAM (NOO) */}
                  {activeSalesman && (activeSalesman.name.toLowerCase().trim() === "imam") && (
                    <div className="bg-rose-50/50 rounded-2xl p-5 border border-rose-100 flex flex-col gap-4 animate-in fade-in slide-in-from-top-3 duration-200">
                      <div className="flex items-center justify-between border-b border-rose-200/60 pb-2">
                        <div className="flex items-center gap-2">
                          <span className="text-base">📌</span>
                          <div>
                            <span className="text-xs font-black text-rose-800 uppercase tracking-widest block">
                              Log New Outlet Opening (NOO) - {activeSalesman.name}
                            </span>
                            <span className="text-[10px] text-rose-700/80 font-bold block">
                              Pilihan otomatis terdeteksi! Silakan input outlet baru yang berhasil dibuka hari ini
                            </span>
                          </div>
                        </div>
                        <span className="text-[9px] font-black uppercase bg-rose-200 text-rose-800 px-2 py-0.5 rounded-md">
                          Auto-Mapping Tab
                        </span>
                      </div>
                      
                      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                        {/* WARUNG */}
                        <div className="bg-white p-3 rounded-xl border border-rose-200/50">
                          <label className="block text-[10px] font-bold text-rose-900 uppercase tracking-wider mb-1.5 text-center">
                            Warung NOO
                          </label>
                          <input
                            type="number"
                            min="0"
                            onFocus={(e) => e.target.value === "0" && e.target.select()}
                            value={newNooWarung}
                            onChange={(e) => setNewNooWarung(parseInt(e.target.value, 10) || 0)}
                            className="w-full text-center rounded-lg py-2 text-base font-mono font-black text-rose-900 bg-rose-50/20 border border-rose-200 focus:outline-[#E5E5DF] focus:ring-1 focus:ring-rose-400 focus:bg-white"
                          />
                        </div>

                        {/* TOKO */}
                        <div className="bg-white p-3 rounded-xl border border-rose-200/50">
                          <label className="block text-[10px] font-bold text-rose-900 uppercase tracking-wider mb-1.5 text-center">
                            Toko NOO
                          </label>
                          <input
                            type="number"
                            min="0"
                            onFocus={(e) => e.target.value === "0" && e.target.select()}
                            value={newNooStore}
                            onChange={(e) => setNewNooStore(parseInt(e.target.value, 10) || 0)}
                            className="w-full text-center rounded-lg py-2 text-base font-mono font-black text-rose-900 bg-rose-50/20 border border-rose-200 focus:outline-[#E5E5DF] focus:ring-1 focus:ring-rose-400 focus:bg-white"
                          />
                        </div>

                        {/* KIOS */}
                        <div className="bg-white p-3 rounded-xl border border-rose-200/50">
                          <label className="block text-[10px] font-bold text-rose-900 uppercase tracking-wider mb-1.5 text-center">
                            Kios NOO
                          </label>
                          <input
                            type="number"
                            min="0"
                            onFocus={(e) => e.target.value === "0" && e.target.select()}
                            value={newNooKiosk}
                            onChange={(e) => setNewNooKiosk(parseInt(e.target.value, 10) || 0)}
                            className="w-full text-center rounded-lg py-2 text-base font-mono font-black text-rose-900 bg-rose-50/20 border border-rose-200 focus:outline-[#E5E5DF] focus:ring-1 focus:ring-rose-400 focus:bg-white"
                          />
                        </div>

                        {/* GROSIR */}
                        <div className="bg-white p-3 rounded-xl border border-rose-200/50">
                          <label className="block text-[10px] font-bold text-rose-900 uppercase tracking-wider mb-1.5 text-center">
                            Grosir NOO
                          </label>
                          <input
                            type="number"
                            min="0"
                            onFocus={(e) => e.target.value === "0" && e.target.select()}
                            value={newNooWholesaler}
                            onChange={(e) => setNewNooWholesaler(parseInt(e.target.value, 10) || 0)}
                            className="w-full text-center rounded-lg py-2 text-base font-mono font-black text-rose-900 bg-rose-50/20 border border-rose-200 focus:outline-[#E5E5DF] focus:ring-1 focus:ring-rose-400 focus:bg-white"
                          />
                        </div>
                      </div>
                    </div>
                  )}

                  {/* ROW 5: FINANCIAL METRICS */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    
                    {/* BIAYA OPERASIONAL */}
                    <div>
                      <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-widest mb-1.5">
                        BIAYA OPERASIONAL *
                      </label>
                      <div className="relative">
                        <span className="absolute left-4 top-3.5 text-xs text-[#8C8C70] font-bold">Rp</span>
                        <input
                          type="number"
                          min="0"
                          onFocus={(e) => e.target.value === "0" && e.target.select()}
                          value={operationalCost}
                          onChange={(e) => setOperationalCost(parseInt(e.target.value, 10) || 0)}
                          className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl pl-10 pr-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] focus:bg-white text-[#4A4A3C] font-semibold"
                          required
                        />
                      </div>
                      <div className="flex gap-1.5 mt-1.5">
                        {[0, 10000, 15000, 20000, 25000].map(val => (
                          <button
                            key={val}
                            type="button"
                            onClick={() => setOperationalCost(val)}
                            className="bg-[#E5E5DF]/50 hover:bg-[#E5E5DF] text-[#4A4A3C] text-[10px] px-2.5 py-1 rounded font-bold transition"
                          >
                            Rp {val.toLocaleString("id-ID")}
                          </button>
                        ))}
                      </div>
                    </div>

                    {/* TAGIHAN BAYAR TUNAI */}
                    <div>
                      <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-widest mb-1.5">
                        Tagihan Bayar Tunai *
                      </label>
                      <div className="relative">
                        <span className="absolute left-4 top-3.5 text-xs text-[#8C8C70] font-bold">Rp</span>
                        <input
                          type="number"
                          min="0"
                          onFocus={(e) => e.target.value === "0" && e.target.select()}
                          value={billsReceived}
                          onChange={(e) => setBillsReceived(parseInt(e.target.value, 10) || 0)}
                          className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl pl-10 pr-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] focus:bg-white text-[#4A4A3C] font-bold"
                          required
                        />
                      </div>
                      <p className="text-[10px] text-[#8C8C70] mt-1">
                        Terbilang: <span className="text-[#4A4A3C]/75 font-semibold italic">Rp {billsReceived.toLocaleString("id-ID")} Rupiah</span>
                      </p>
                    </div>

                    {/* TAGIHAN BAYAR TRANSFER */}
                    <div>
                      <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-widest mb-1.5">
                        Tagihan Bayar Transfer
                      </label>
                      <div className="relative">
                        <span className="absolute left-4 top-3.5 text-xs text-[#8C8C70] font-bold">Rp</span>
                        <input
                          type="number"
                          min="0"
                          onFocus={(e) => e.target.value === "0" && e.target.select()}
                          value={billsTransfer}
                          onChange={(e) => setBillsTransfer(parseInt(e.target.value, 10) || 0)}
                          className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl pl-10 pr-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] focus:bg-white text-[#4A4A3C] font-bold"
                        />
                      </div>
                    </div>

                    {/* TAGIHAN GIRO */}
                    <div>
                      <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-widest mb-1.5">
                        Tagihan Giro
                      </label>
                      <div className="relative">
                        <span className="absolute left-4 top-3.5 text-xs text-[#8C8C70] font-bold">Rp</span>
                        <input
                          type="number"
                          min="0"
                          onFocus={(e) => e.target.value === "0" && e.target.select()}
                          value={billsGiro}
                          onChange={(e) => setBillsGiro(parseInt(e.target.value, 10) || 0)}
                          className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl pl-10 pr-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] focus:bg-white text-[#4A4A3C] font-bold"
                        />
                      </div>
                    </div>

                  </div>

                  {/* NOTES/ADDITIONAL INFO */}
                  <div>
                    <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-widest mb-1.5">
                      Catatan Temuan / Kendala (Opsional)
                    </label>
                    <textarea
                      rows={2}
                      value={notes}
                      onChange={(e) => setNotes(e.target.value)}
                      placeholder="Contoh: Kendala armada bocor, hujan lebat, toko berkah tutup sementara, atau kendala tagihan..."
                      className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl px-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] focus:bg-white text-[#4A4A3C] transition"
                    />
                  </div>

                  {/* FORM TRIGGER NATURAL TONES BUTTON */}
                  <div className="flex flex-col sm:flex-row gap-3 pt-2">
                    <button
                      type="submit"
                      id="btn-submit-form"
                      disabled={isSubmittingReport}
                      className="w-full bg-[#5A5A40] hover:bg-[#4A4A3C] text-[#FAF9F6] font-extrabold text-sm uppercase tracking-wider py-4 px-6 rounded-2xl transition duration-150 flex items-center justify-center gap-2.5 shadow-md cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                      {isSubmittingReport ? (
                        <>
                          <span className="w-5 h-5 border-2 border-[#FAF9F6]/30 border-t-[#FAF9F6] rounded-full animate-spin" />
                          Menyimpan ke Database...
                        </>
                      ) : (
                        <>
                          <FileText className="w-5 h-5 text-[#FAF9F6]" />
                          SIMPAN KE DATABASE
                        </>
                      )}
                    </button>
                  </div>

                </form>

              </div>
            </motion.div>
          )}

          {/* TAB 2: SALESMAN DATABASE MANAGEMENT (ANSWERING USER'S DETAILED DB QUESTION) */}
          {activeTab === "customerData" && (
            <motion.div
              layout
              key="customer-data-tab"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.25 }}
              className="max-w-6xl mx-auto w-full"
            >
              <CustomerSalesTable />
            </motion.div>
          )}

          {activeTab === "salesmen" && (
            <motion.div
              layout
              key="salesmen-tab"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.25 }}
              className="flex flex-col gap-6"
            >
              {/* TOP DB INSTRUCTION SECTION */}
              <div className="bg-[#FAF9F6] rounded-3xl p-6 border border-[#E5E5DF] shadow-xs">
                <div className="flex flex-col md:flex-row gap-6 items-start md:items-center justify-between">
                  <div className="max-w-2xl">
                    <h2 className="text-lg font-bold text-[#4A4A3C] uppercase flex items-center gap-2 font-serif italic">
                      <UserPlus className="w-5 h-5 text-[#5A5A40]" />
                      Kelola Database Salesman
                    </h2>
                    <p className="text-xs text-[#8C8C70] mt-1 leading-relaxed">
                      Tambahkan, rubah, atau hapus nama salesman harian. Semua nama yang terdaftar di sini akan muncul langsung di autocomplete dropdown input KPI. <strong>Ini adalah pembuktian database fungsional!</strong>
                    </p>
                  </div>
                  
                  <button
                    onClick={handleAddSalesman}
                    className="w-full md:w-auto px-5 py-3 bg-[#5A5A40] hover:bg-[#4A4A3C] text-[#FAF9F6] font-[#FAF9F6] font-bold text-xs uppercase tracking-wider rounded-xl transition duration-150 flex items-center justify-center gap-2 cursor-pointer shadow-xs"
                  >
                    <UserPlus className="w-4 h-4" />
                    Tambah Salesman Baru
                  </button>
                </div>

                {/* DB QUESTION EXPLANATION COMPONENT */}
                <div className="mt-6 bg-[#E5E5DF]/20 p-5 rounded-2xl border border-[#E5E5DF]">
                  <h4 className="text-xs font-black text-[#5A5A40] uppercase tracking-widest mb-1.5 flex items-center gap-2">
                    <Info className="w-4 h-4 text-[#5A5A40]" />
                    "Apa Saja Yang Dibutuhkan Agar Pilihan Nama Sales Bisa Dropdown?" (Pertanyaan Anda)
                  </h4>
                  <p className="text-xs text-[#4A4A3C] leading-relaxed mb-3">
                    Untuk mengubah input manual tradisional menjadi <strong>Dropdown Cerdas berbasis Database</strong>, sistem memerlukan beberapa elemen arsitektur dasar berikut:
                  </p>
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                    <div className="bg-[#FAF9F6] p-3.5 rounded-xl border border-[#E5E5DF]">
                      <span className="font-bold text-[#5A5A40] block mb-1">1. Model Data (Schema)</span>
                      <span className="text-[#8C8C70]">Sebuah spesifikasi struktur data Salesmen yang jelas (contoh: ID, Nama, Area, Nomor HP, dan Status Keaktifan).</span>
                    </div>
                    <div className="bg-[#FAF9F6] p-3.5 rounded-xl border border-[#E5E5DF]">
                      <span className="font-bold text-[#5A5A40] block mb-1">2. Penyimpanan (Database)</span>
                      <span className="text-[#8C8C70]">Sistem penyimpanan persisten. Untuk local client kita menggunakan <strong>LocalStorage Cache</strong>. Untuk online-multiuser, kita butuh database web seperti <strong>Cloud Firestore (Firebase)</strong>.</span>
                    </div>
                    <div className="bg-[#FAF9F6] p-3.5 rounded-xl border border-[#E5E5DF]">
                      <span className="font-bold text-[#5A5A40] block mb-1">3. Sinkronisasi UI Form</span>
                      <span className="text-[#8C8C70]">Sebuah kueri data list di React (state loop) untuk memetakan nama sales terdaftar ke dalam tag HTML <code>&lt;select&gt;</code> atau dropdown autocomplete modal.</span>
                    </div>
                  </div>
                </div>
              </div>

              {/* SALESMAN DATA TABLE */}
              <div className="bg-[#FAF9F6] rounded-3xl border border-[#E5E5DF] shadow-xs overflow-hidden">
                <div className="p-5 border-b border-[#E5E5DF] flex items-center justify-between bg-[#E5E5DF]/10">
                  <span className="text-xs font-bold text-[#8C8C70] uppercase tracking-wider">
                    Daftar Anggota Salesman Aktif ({salesmen.length})
                  </span>
                  <button 
                    onClick={handleFetchSalesmenFromSheets}
                    className="text-[10px] bg-white border border-[#E5E5DF] px-2 py-1 rounded-md hover:bg-gray-50 uppercase font-bold text-[#4A4A3C]"
                  >
                    Sync Data
                  </button>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm whitespace-nowrap">
                    <thead className="bg-[#E5E5DF]/30 text-[11px] font-bold text-[#8C8C70] uppercase tracking-wider border-b border-[#E5E5DF]">
                      <tr>
                        <th className="px-6 py-4">Nama Salesman</th>
                        <th className="px-6 py-4">Wilayah / Area</th>
                        <th className="px-6 py-4">Nomor HP</th>
                        <th className="px-6 py-4 text-center">Tindakan</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E5DF]/40">
                      {salesmen.map((s) => (
                        <tr key={s.id} className="hover:bg-[#E5E5DF]/20 transition">
                          <td className="px-6 py-4">
                            <div className="flex items-center gap-3">
                              <div className="w-9 h-9 rounded-lg bg-[#E5E5DF] text-[#5A5A40] font-bold flex items-center justify-center text-xs">
                                {s.name.charAt(0)}
                              </div>
                              <div>
                                <span className="font-bold text-[#4A4A3C] uppercase block">{s.name}</span>
                                <span className="text-[10px] text-[#8C8C70] font-mono">ID: {s.id}</span>
                              </div>
                            </div>
                          </td>
                          <td className="px-6 py-4">
                            <span className="px-2.5 py-1 bg-[#E5E5DF]/50 rounded-md text-[#4A4A3C] text-xs font-semibold">
                              {s.area || "Semua Wilayah"}
                            </span>
                          </td>
                          <td className="px-6 py-4 font-mono text-xs text-[#8C8C70]">
                            {s.phone || "Tidak ada data"}
                          </td>
                          <td className="px-6 py-4">
                            <div className="flex items-center justify-center gap-3">
                              <button
                                onClick={() => handleEditSalesman(s)}
                                className="p-2 text-[#8C8C70] hover:text-[#5A5A40] hover:bg-[#E5E5DF]/40 rounded-lg transition"
                                title="Edit Salesman"
                              >
                                <Edit2 className="w-4 h-4" />
                              </button>
                              <button
                                onClick={() => handleDeleteSalesman(s.id, s.name)}
                                className="p-2 text-[#8C8C70] hover:text-rose-600 hover:bg-rose-50/50 rounded-lg transition"
                                title="Hapus Salesman"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </motion.div>
          )}

          {/* TAB 3: PRODUCT DATABASE MANAGEMENT */}
          {activeTab === "products" && (
            <motion.div
              layout
              key="products-tab"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.25 }}
              className="flex flex-col gap-6"
            >
              {/* TOP PRODUCTS HEADER */}
              <div className="bg-[#FAF9F6] rounded-3xl p-6 border border-[#E5E5DF] shadow-xs">
                <div className="flex flex-col md:flex-row gap-6 items-start md:items-center justify-between">
                  <div>
                    <h2 className="text-lg font-bold text-[#4A4A3C] uppercase flex items-center gap-2 font-serif italic">
                      <ShoppingBag className="w-5 h-5 text-[#5A5A40]" />
                      Kelola Database Produk
                    </h2>
                    <p className="text-xs text-[#8C8C70] mt-1 leading-relaxed">
                      Visualisasi katalog produk distributor. Produk yang didaftarkan di sini akan muncul langsung di floating product search input form KPI untuk mempermudah perhitungan SKU total.
                    </p>
                  </div>
                  
                  <button
                    onClick={handleAddProduct}
                    className="w-full md:w-auto px-5 py-3 bg-[#5A5A40] hover:bg-[#4A4A3C] text-[#FAF9F6] font-[#FAF9F6] font-bold text-xs uppercase tracking-wider rounded-xl transition duration-150 flex items-center justify-center gap-2 cursor-pointer shadow-xs"
                  >
                    <Plus className="w-4 h-4" />
                    Tambah Produk Baru
                  </button>
                </div>
              </div>

              {/* SPREADSHEET SKU SYNC BAR */}
              {sheetsScriptUrl ? (
                <div className="bg-[#FAF9F6] border border-[#E5E5DF] rounded-3xl p-5 flex flex-col sm:flex-row items-center justify-between gap-4 shadow-xs">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-xl bg-emerald-50 text-emerald-700 flex items-center justify-center font-bold shrink-0">
                      <span className="text-lg">📦</span>
                    </div>
                    <div>
                      <h4 className="text-xs font-black text-emerald-800 uppercase tracking-wider">Master Data Produk Fokus Terhubung</h4>
                      <p className="text-[10px] text-[#8C8C70] font-bold mt-1">
                        Auto-sync aktif: Setiap penambahan, pengubahan, atau penghapusan SKU Produk langsung terupdate di Google Sheets (Tab 'Daftar Produk SKU')!
                      </p>
                    </div>
                  </div>
                  <div className="flex gap-2 w-full sm:w-auto">
                    <button
                      type="button"
                      onClick={() => handleFetchProductsFromSheets()}
                      disabled={isFetchingProducts}
                      className="flex-1 sm:flex-none bg-white hover:bg-gray-50 border border-[#E5E5DF] text-[#5A5A40] text-xs font-black px-4 py-2.5 rounded-2xl flex items-center justify-center gap-2 transition cursor-pointer disabled:opacity-50"
                    >
                      <RefreshCw className={`w-3.5 h-3.5 ${isFetchingProducts ? "animate-spin" : ""}`} />
                      Tarik SKU Sheets
                    </button>
                    <button
                      type="button"
                      onClick={() => handleSyncProductsToSheets(products, false)}
                      disabled={isSyncingProducts}
                      className="flex-1 sm:flex-none bg-rose-600 hover:bg-rose-700 text-white text-xs font-black px-4 py-2.5 rounded-2xl flex items-center justify-center gap-2 transition cursor-pointer disabled:opacity-50 shadow-sm"
                    >
                      <Send className={`w-3.5 h-3.5 ${isSyncingProducts ? "animate-spin" : ""}`} />
                      Kirim Manual
                    </button>
                  </div>
                </div>
              ) : (
                <div className="bg-amber-50 border border-amber-200 rounded-3xl p-5 flex items-center gap-3">
                  <span className="text-xl">⚠️</span>
                  <div>
                    <h4 className="text-xs font-black text-amber-800 uppercase">Spreadsheet Belum Terhubung</h4>
                    <p className="text-[10px] text-amber-900/80">Silakan hubungkan URL Google Apps Script Anda di tab <strong className="text-rose-700">Google Sheets Linker</strong> untuk sinkronisasi otomatis harian SKU Produk.</p>
                  </div>
                </div>
              )}

              {/* PRODUCT CARDS LIST */}
              <div className="bg-[#FAF9F6] rounded-3xl border border-[#E5E5DF] shadow-xs overflow-hidden">
                <div className="p-5 border-b border-[#E5E5DF] flex items-center justify-between bg-[#E5E5DF]/10">
                  <span className="text-xs font-bold text-[#8C8C70] uppercase tracking-wider">
                    Daftar Inventaris Produk ({products.length})
                  </span>
                  <button 
                    onClick={handleFetchProductsFromSheets}
                    className="text-[10px] bg-white border border-[#E5E5DF] px-2 py-1 rounded-md hover:bg-gray-50 uppercase font-bold text-[#4A4A3C]"
                  >
                    Sync Data
                  </button>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm whitespace-nowrap">
                    <thead className="bg-[#E5E5DF]/30 text-[11px] font-bold text-[#8C8C70] uppercase tracking-wider border-b border-[#E5E5DF]">
                      <tr>
                        <th className="px-6 py-4">Kode SKU / Nama Produk</th>
                        <th className="px-6 py-4">Kategori Produk</th>
                        <th className="px-6 py-4 text-center">Tindakan</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E5DF]/40">
                      {products.map((p) => (
                        <tr key={p.id} className="hover:bg-[#E5E5DF]/20 transition">
                          <td className="px-6 py-4">
                            <div>
                              <span className="font-bold text-[#4A4A3C] uppercase block">{p.name}</span>
                              <span className="bg-[#E5E5DF] text-[#4A4A3C] px-1.5 py-0.5 rounded text-[10px] font-bold inline-block mt-1 uppercase">
                                SKU: {p.skuCode || "Tanpa Kode"}
                              </span>
                            </div>
                          </td>
                          <td className="px-6 py-4 text-xs font-semibold text-[#8C8C70]">
                            {p.category || "Umum"}
                          </td>
                          <td className="px-6 py-4">
                            <div className="flex items-center justify-center gap-3">
                              <button
                                onClick={() => handleEditProduct(p)}
                                className="p-2 text-[#8C8C70] hover:text-[#5A5A40] hover:bg-[#E5E5DF]/35 rounded-lg transition"
                                title="Edit Produk"
                              >
                                <Edit2 className="w-4 h-4" />
                              </button>
                              <button
                                onClick={() => handleDeleteProduct(p.id, p.name)}
                                className="p-2 text-[#8C8C70] hover:text-rose-600 hover:bg-rose-50/50 rounded-lg transition"
                                title="Hapus Produk"
                              >
                                <Trash2 className="w-4 h-4 text-[#8C8C70]" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </motion.div>
          )}

          {/* TAB 4: HISTORIC SUBMITTED AUDIT REPORTS (RECAPS) */}
          {activeTab === "reports" && (
            <motion.div
              layout
              key="reports-tab"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.25 }}
              className="flex flex-col gap-6"
            >
              {/* HISTORIC STRIP HEADER */}
              <div className="bg-[#FAF9F6] rounded-3xl p-6 border border-[#E5E5DF] shadow-xs flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                  <h2 className="text-lg font-bold text-[#4A4A3C] uppercase font-serif italic">
                    Arsip Riwayat Laporan KPI Sales
                  </h2>
                  <p className="text-xs text-[#8C8C70] mt-1 leading-relaxed">
                    Halaman verifikasi dan audit auditor portal. Berisi riwayat input data harian yang dikirim oleh admin KPI.
                  </p>
                </div>
                
                {filteredAuditReports.length > 0 && (
                  <button
                    onClick={() => {
                      const wsData = filteredAuditReports.map((r: any, index: number) => ({
                        "No": index + 1,
                        "Tanggal": r.date,
                        "Nama Salesman": r.salesmanName,
                        "Siklus": r.cycle,
                        "TC (Total Call)": r.tc,
                        "CP (Effective Call)": r.cp,
                        "EC (Effective Customer)": r.ec,
                        "Total SKU": r.skuTotal,
                        "Biaya Operasional (Rp)": r.operationalCost,
                        "Tagihan Tunai": r.billsReceived,
                        "Tagihan Transfer": r.billsTransfer || 0,
                        "Tagihan Giro": r.billsGiro || 0,
                        "Catatan / Kendala": r.notes || ""
                      }));

                      // Create worksheet
                      const ws = XLSX.utils.json_to_sheet(wsData);
                      
                      // Auto-fit column widths
                      ws["!cols"] = [
                        { wch: 6 },   // No
                        { wch: 13 },  // Tanggal
                        { wch: 20 },  // Nama Salesman
                        { wch: 15 },  // Siklus
                        { wch: 16 },  // TC
                        { wch: 16 },  // CP
                        { wch: 16 },  // EC
                        { wch: 12 },  // SKU
                        { wch: 22 },  // Biaya Operasional
                        { wch: 20 },  // Tagihan Tunai
                        { wch: 20 },  // Tagihan Transfer
                        { wch: 20 },  // Tagihan Giro
                        { wch: 45 }   // Catatan
                      ];

                      // Create workbook
                      const wb = XLSX.utils.book_new();
                      XLSX.utils.book_append_sheet(wb, ws, "Laporan Audited KPI");

                      // Download XLSX
                      XLSX.writeFile(wb, `LAPORAN_AUDIT_KPI_REKAP_${new Date().toISOString().split("T")[0]}.xlsx`);
                      showToast("Laporan berhasil diekspor ke format Excel (.xlsx)!", "success");
                    }}
                    className="px-4 py-2 bg-[#5A5A40] hover:bg-[#4A4A3C] text-[#FAF9F6] text-xs font-bold uppercase tracking-wider rounded-xl transition flex items-center gap-1.5 self-start cursor-pointer md:self-auto shadow-sm"
                  >
                    <Download className="w-4 h-4" />
                    Ekspor EXCEL (.XLSX)
                  </button>
                )}
              </div>

              {/* Filter Area */}
              <div className="bg-[#FAF9F6] rounded-2xl p-4 border border-[#E5E5DF] shadow-xs flex flex-col md:flex-row gap-4">
                <div className="flex-1">
                  <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                    Tanggal Mulai
                  </label>
                  <input
                    type="date"
                    value={auditFilterStartDate}
                    onChange={(e) => setAuditFilterStartDate(e.target.value)}
                    className="w-full bg-white border border-[#E5E5DF] text-[#4A4A3C] font-mono text-sm px-3 py-2 rounded-xl focus:outline-hidden focus:border-[#8C8C70]"
                  />
                </div>
                <div className="flex-1">
                  <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                    Tanggal Akhir
                  </label>
                  <input
                    type="date"
                    value={auditFilterEndDate}
                    onChange={(e) => setAuditFilterEndDate(e.target.value)}
                    className="w-full bg-white border border-[#E5E5DF] text-[#4A4A3C] font-mono text-sm px-3 py-2 rounded-xl focus:outline-hidden focus:border-[#8C8C70]"
                  />
                </div>
                <div className="flex items-end">
                  <button
                    onClick={() => {
                      setAuditFilterStartDate("");
                      setAuditFilterEndDate("");
                    }}
                    className="px-4 py-2 hover:bg-[#E5E5DF]/50 text-[#8C8C70] text-xs font-bold uppercase tracking-wider rounded-xl transition cursor-pointer h-[38px] border border-transparent hover:border-[#E5E5DF]"
                  >
                    Reset Filter
                  </button>
                </div>
              </div>

              {/* REPORT CARDS VIEW */}
              {filteredAuditReports.length > 0 ? (
                <div className="grid grid-cols-1 gap-4">
                  {filteredAuditReports.map((rep: any) => (
                    <div
                      key={rep.id}
                      className="bg-[#FAF9F6] rounded-3xl p-6 border border-[#E5E5DF] shadow-xs hover:border-[#8C8C70] transition-all duration-150 flex flex-col lg:flex-row lg:items-center justify-between gap-6"
                    >
                      {/* Left: General Info */}
                      <div className="flex-1">
                        <div className="flex items-center gap-3 mb-2 flex-wrap">
                          <span className="px-3 py-1 bg-[#E5E5DF] text-[#4A4A3C] text-xs font-bold uppercase rounded-lg">
                            {rep.salesmanName}
                          </span>
                          <span className="text-[#8C8C70] font-bold text-xs flex items-center gap-1">
                            <Calendar className="w-3.5 h-3.5 text-[#8C8C70]" />
                            {rep.date}
                          </span>
                          <span className="px-2.5 py-0.5 bg-[#E5E5DF]/50 rounded-md text-[#4A4A3C] text-xs font-semibold">
                            {rep.cycle}
                          </span>
                        </div>

                        {/* KPI Metrics List Badge */}
                        <div className="grid grid-cols-4 gap-2 max-w-lg mb-3 mt-3">
                          <div className="bg-[#E5E5DF]/30 p-2 rounded-lg text-center">
                            <span className="text-[10px] text-[#8C8C70] font-bold block">TC</span>
                            <span className="text-sm font-mono font-black text-[#4A4A3C]">{rep.tc}</span>
                          </div>
                          <div className="bg-[#E5E5DF]/30 p-2 rounded-lg text-center">
                            <span className="text-[10px] text-[#8C8C70] font-bold block">CP</span>
                            <span className="text-sm font-mono font-black text-[#4A4A3C]">{rep.cp}</span>
                          </div>
                          <div className="bg-[#E5E5DF]/60 p-2 rounded-lg text-center">
                            <span className="text-[10px] text-[#5A5A40] font-bold block">EC</span>
                            <span className="text-sm font-mono font-black text-[#5A5A40]">{rep.ec}</span>
                          </div>
                          <div className="bg-[#8C8C70]/20 p-2 rounded-lg text-center">
                            <span className="text-[10px] text-[#5A5A40] font-bold block">SKU</span>
                            <span className="text-sm font-mono font-black text-[#5A5A40]">{rep.skuTotal}</span>
                          </div>
                        </div>

                        {rep.notes && (
                          <div className="text-xs text-[#8C8C70] mt-2 bg-[#E5E5DF]/20 p-2 rounded-xl italic border-l-2 border-[#8C8C70]">
                            " {rep.notes} "
                          </div>
                        )}

                        {/* Rincian produk yang diinput */}
                        {rep.productsDetail && rep.productsDetail.length > 0 && (
                          <div className="mt-3 flex items-center gap-1.5 flex-wrap">
                            <span className="text-[10px] font-bold text-[#8C8C70] uppercase">Rincian SKU Fokus:</span>
                            {rep.productsDetail.map((prod, i) => (
                              <span key={i} className="bg-[#E5E5DF]/40 py-0.5 px-2 rounded-md text-[10px] text-[#4A4A3C] font-semibold uppercase">
                                {prod.productName}
                              </span>
                            ))}
                          </div>
                        )}
                      </div>

                      {/* Right: Operational Values */}
                      <div className="flex flex-col sm:flex-row lg:flex-col items-start sm:items-center lg:items-end justify-between lg:justify-center gap-4 lg:text-right">
                        <div>
                          <span className="text-[10px] font-bold text-[#8C8C70] block uppercase">Bayar Tunai</span>
                          <span className="text-lg font-mono font-black text-[#5A5A40] block">
                            Rp {rep.billsReceived.toLocaleString("id-ID")}
                          </span>
                        </div>
                        {((rep.billsTransfer || 0) > 0 || (rep.billsGiro || 0) > 0) && (
                          <div className="flex gap-4">
                            {(rep.billsTransfer || 0) > 0 && (
                              <div>
                                <span className="text-[10px] font-bold text-[#8C8C70] block uppercase">Transfer</span>
                                <span className="text-sm font-mono font-black text-blue-700 block mt-0.5">
                                  Rp {(rep.billsTransfer || 0).toLocaleString("id-ID")}
                                </span>
                              </div>
                            )}
                            {(rep.billsGiro || 0) > 0 && (
                              <div>
                                <span className="text-[10px] font-bold text-[#8C8C70] block uppercase">Giro</span>
                                <span className="text-sm font-mono font-black text-amber-600 block mt-0.5">
                                  Rp {(rep.billsGiro || 0).toLocaleString("id-ID")}
                                </span>
                              </div>
                            )}
                          </div>
                        )}
                        
                        <div>
                          <span className="text-[10px] font-bold text-[#8C8C70] block uppercase">Operational Cost</span>
                          <span className="text-sm font-mono font-bold text-[#8C8C70] block">
                            Rp {((typeof rep.operationalCost === 'number' && !isNaN(rep.operationalCost)) ? rep.operationalCost : 0).toLocaleString("id-ID")}
                          </span>
                        </div>

                        {/* Copy WA or Delete from list */}
                        <div className="flex items-center gap-2">
                          <button
                            onClick={() => {
                              const detailText = `LAPORAN SALES FORCE\n` +
                                `Salesman: ${rep.salesmanName}\n` +
                                `Tanggal: ${rep.date}\n` +
                                `Siklus: ${rep.cycle}\n` +
                                `TC: ${rep.tc} | CP: ${rep.cp} | EC: ${rep.ec} | SKU: ${rep.skuTotal}\n` +
                                `Ops: Rp ${((typeof rep.operationalCost === 'number' && !isNaN(rep.operationalCost)) ? rep.operationalCost : 0).toLocaleString("id-ID")}\n` +
                                `Tunai: Rp ${rep.billsReceived.toLocaleString("id-ID")}\n` +
                                `Transfer: Rp ${(rep.billsTransfer || 0).toLocaleString("id-ID")}\n` +
                                `Giro: Rp ${(rep.billsGiro || 0).toLocaleString("id-ID")}\n` +
                                `Catatan: ${rep.notes || "-"}`;
                              navigator.clipboard.writeText(detailText);
                              showToast("Rekap laporan disalin ke clipboard!", "success");
                            }}
                            className="text-xs text-[#4A4A3C] bg-[#E5E5DF] hover:bg-[#E5E5DF]/80 py-1.5 px-3 rounded-lg font-bold flex items-center gap-1 transition"
                            title="Format Kirim WA"
                          >
                            <Copy className="w-3.5 h-3.5" /> WA
                          </button>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="bg-[#FAF9F6] rounded-3xl p-12 border border-[#E5E5DF] text-center text-[#8C8C70] max-w-sm mx-auto shadow-xs">
                  <History className="w-12 h-12 text-[#E5E5DF] mx-auto mb-4" />
                  <p className="text-sm font-bold text-[#4A4A3C] uppercase">Belum ada Laporan Audit</p>
                  <p className="text-xs text-[#8C8C70] mt-1">Silakan mulai dengan mengirim Laporan Audit KPI di Tab Utama!</p>
                  <button
                    onClick={() => setActiveTab("form")}
                    className="mt-4 px-4 py-2 bg-[#5A5A40] text-[#FAF9F6] rounded-lg text-xs font-bold transition cursor-pointer"
                  >
                    Buka Form Input
                  </button>
                </div>
              )}
            </motion.div>
          )}

          {activeTab === "sheets" && (
            <motion.div
              key="sheets-tab"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              className="grid grid-cols-1 lg:grid-cols-3 gap-6"
            >
              {/* Left Column: Connection & Bulk Sync Settings */}
              <div className="lg:col-span-1 flex flex-col gap-6">
                
                {/* 1. Configuration Panel */}
                <div className="bg-[#FAF9F6] rounded-3xl p-6 border border-[#E5E5DF] shadow-xs">
                  <div className="flex items-center gap-2.5 mb-4">
                    <div className="p-2 bg-[#5A5A40]/10 text-[#5A5A40] rounded-xl">
                      <Layers className="w-5 h-5 text-[#5A5A40]" />
                    </div>
                    <div>
                      <h3 className="text-sm font-extrabold text-[#4A4A3C] uppercase tracking-wider">
                        Koneksi Spreadsheet
                      </h3>
                      <p className="text-[10px] text-[#8C8C70]">Hubungkan dengan Google Sheets API</p>
                    </div>
                  </div>

                  <hr className="border-[#E5E5DF] mb-4" />

                    <div className="space-y-4">
                      {/* Web App URL Input */}
                      <div>
                        <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1.5 flex items-center justify-between">
                          <span>URL Deploy Web App Apps Script</span>
                          {sheetsScriptUrl !== "https://script.google.com/macros/s/AKfycbz2a1vnNlVwQl6z77-3b5LYgXZAhprIgH3a_JAl0El5GPPI8xiZZGPdGsapa_mM6S0DQA/exec" && (
                            <button
                              type="button"
                              onClick={() => {
                                const targetDefault = "https://script.google.com/macros/s/AKfycbz2a1vnNlVwQl6z77-3b5LYgXZAhprIgH3a_JAl0El5GPPI8xiZZGPdGsapa_mM6S0DQA/exec";
                                setSheetsScriptUrl(targetDefault);
                                localStorage.setItem("KPI_SHEETS_SCRIPT_URL", targetDefault);
                                showToast("URL Web App default Anda berhasil dipulihkan!", "success");
                              }}
                              className="text-[9px] font-black text-[#5A5A40] hover:underline cursor-pointer"
                            >
                              Set Default Web App
                            </button>
                          )}
                        </label>
                        <input
                          type="url"
                          placeholder="https://script.google.com/macros/s/.../exec"
                          value={sheetsScriptUrl}
                          onChange={(e) => {
                            setSheetsScriptUrl(e.target.value);
                            localStorage.setItem("KPI_SHEETS_SCRIPT_URL", e.target.value);
                          }}
                          className="w-full px-3 py-2 bg-[#E5E5DF]/20 border border-[#E5E5DF] rounded-xl text-xs text-[#4A4A3C] focus:outline-none focus:border-[#5A5A40] transition placeholder:text-[#8C8C70]/50 font-mono"
                        />
                      </div>

                      {/* Apps Script Library URL Input */}
                      <div>
                        <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1.5 flex items-center justify-between">
                          <span>URL Library Google Apps Script</span>
                          {sheetsLibraryUrl !== "https://script.google.com/macros/library/d/1y8VqApcuL-QYGyNyZFB0kF3MI7T9qwCtptmzUpQbNL34H-Gz0LCJBxkt/10" && (
                            <button
                              type="button"
                              onClick={() => {
                                const targetDefault = "https://script.google.com/macros/library/d/1y8VqApcuL-QYGyNyZFB0kF3MI7T9qwCtptmzUpQbNL34H-Gz0LCJBxkt/10";
                                setSheetsLibraryUrl(targetDefault);
                                localStorage.setItem("KPI_SHEETS_LIBRARY_URL", targetDefault);
                                showToast("URL Library default berhasil dipulihkan!", "success");
                              }}
                              className="text-[9px] font-black text-[#5A5A40] hover:underline cursor-pointer"
                            >
                              Set Default Library
                            </button>
                          )}
                        </label>
                        <input
                          type="url"
                          placeholder="https://script.google.com/macros/library/d/... atau versi terakhir"
                          value={sheetsLibraryUrl}
                          onChange={(e) => {
                            setSheetsLibraryUrl(e.target.value);
                            localStorage.setItem("KPI_SHEETS_LIBRARY_URL", e.target.value);
                          }}
                          className="w-full px-3 py-2 bg-[#E5E5DF]/20 border border-[#E5E5DF] rounded-xl text-xs text-[#4A4A3C] focus:outline-none focus:border-[#5A5A40] transition placeholder:text-[#8C8C70]/50 font-mono"
                        />
                      </div>
                      
                      {/* Info Panel Deployment Milik User */}
                      <div className="p-2.5 bg-[#5A5A40]/10 rounded-xl border border-[#E5E5DF]/50 space-y-1 mt-1">
                        <div className="flex flex-col gap-0.5">
                          <span className="text-[10px] text-[#8C8C70] font-semibold">Deployment ID Aktif:</span>
                          <span className="font-mono text-[10px] text-[#4A4A3C] break-all bg-white/60 p-1 rounded-sm border border-[#E5E5DF] select-all">
                            {(() => {
                              const match = sheetsScriptUrl.match(/\/macros\/s\/([^\/]+)/);
                              return match ? match[1] : (sheetsScriptUrl || "Tidak Ada");
                            })()}
                          </span>
                        </div>
                        <div className="flex items-center justify-between text-[10px] mt-1 pt-1 border-t border-[#E5E5DF]/60">
                          <span className="text-[#8C8C70] font-semibold">Spreadsheet Library:</span>
                          <a 
                            href={sheetsLibraryUrl || "https://script.google.com/macros/library/d/1y8VqApcuL-QYGyNyZFB0kF3MI7T9qwCtptmzUpQbNL34H-Gz0LCJBxkt/10"} 
                            target="_blank" 
                            rel="noreferrer"
                            className="font-mono text-[10px] text-[#5A5A40] underline hover:text-[#4A4A3C] truncate max-w-[150px]"
                            title="Buka Spreadsheet Library"
                          >
                            Buka Library ↗
                          </a>
                        </div>
                      </div>

                    <div className="flex items-center justify-between p-3 bg-[#E5E5DF]/20 rounded-xl border border-[#E5E5DF]/40">
                      <div>
                        <span className="text-xs font-bold text-[#4A4A3C] block">Sinkronisasi Otomatis</span>
                        <span className="text-[10px] text-[#8C8C70]">Kirim data real-time saat disubmit</span>
                      </div>
                      <label className="relative inline-flex items-center cursor-pointer">
                        <input
                          type="checkbox"
                          checked={isAutoSyncEnabled}
                          onChange={(e) => {
                            setIsAutoSyncEnabled(e.target.checked);
                            localStorage.setItem("KPI_SHEETS_AUTO_SYNC", e.target.checked ? "true" : "false");
                            showToast(
                              e.target.checked 
                                ? "Sinkronisasi otomatis diaktifkan!" 
                                : "Sinkronisasi otomatis dinonaktifkan.",
                              "info"
                            );
                          }}
                          className="sr-only peer"
                        />
                        <div className="w-9 h-5 bg-[#E5E5DF] rounded-full peer peer-focus:ring-0 peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-0.5 after:left-[2px] after:bg-white after:border-[#E5E5DF] after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-[#5A5A40]"></div>
                      </label>
                    </div>

                    {/* Test Button */}
                    <button
                      onClick={handleTestConnection}
                      disabled={testConnectionStatus === "testing"}
                      className="w-full py-2.5 bg-[#FAF9F6] text-[#5A5A40] border border-[#5A5A40] hover:bg-[#5A5A40] hover:text-[#FAF9F6] rounded-xl text-xs font-bold transition flex items-center justify-center gap-2 cursor-pointer disabled:opacity-50"
                    >
                      {testConnectionStatus === "testing" ? (
                        <>
                          <span className="w-3.5 h-3.5 border-2 border-[#5A5A40] border-t-transparent rounded-full animate-spin"></span>
                          Menguji Koneksi...
                        </>
                      ) : (
                        <>
                          Uji Hubungkan Koneksi
                        </>
                      )}
                    </button>

                    {testConnectionStatus !== "idle" && (
                      <div className={`p-3 rounded-xl border text-xs flex items-start gap-2 ${
                        testConnectionStatus === "success"
                          ? "bg-emerald-50 text-emerald-800 border-emerald-200"
                          : "bg-rose-50 text-rose-800 border-rose-200"
                      }`}>
                        {testConnectionStatus === "success" ? (
                          <>
                            <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                            <div>
                              <p className="font-bold">Koneksi Berhasil!</p>
                              <p className="text-[10px] text-emerald-700">Aplikasi siap mengirim data secara berkala atau otomatis.</p>
                            </div>
                          </>
                        ) : (
                          <>
                            <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                            <div>
                              <p className="font-bold">Koneksi Gagal</p>
                              <p className="text-[10px] text-rose-700">Verifikasi URL Anda dan pastikan deploy Anda diatur ke hak akses 'Siapa saja (Anyone)'.</p>
                            </div>
                          </>
                        )}
                      </div>
                    )}
                  </div>
                </div>

                {/* 2. Bulk Sync Card */}
                <div className="bg-[#FAF9F6] rounded-3xl p-6 border border-[#E5E5DF] shadow-xs">
                  <div className="flex items-center gap-2.5 mb-4">
                    <div className="p-2 bg-[#5A5A40]/10 text-[#5A5A40] rounded-xl">
                      <Database className="w-5 h-5 text-[#5A5A40]" />
                    </div>
                    <div>
                      <h3 className="text-sm font-extrabold text-[#4A4A3C] uppercase tracking-wider">
                        Upload Data Masal
                      </h3>
                      <p className="text-[10px] text-[#8C8C70]">Kirim data riwayat lokal sekaligus</p>
                    </div>
                  </div>

                  <hr className="border-[#E5E5DF] mb-4" />

                  <div className="space-y-4">
                    <p className="text-[10px] text-[#8C8C70] leading-relaxed pb-1">
                      Klik tombol di bawah ini untuk mengekspor atau menyinkronkan data lokal dari setiap menu halaman ke tab Spreadsheet Anda secara instan:
                    </p>

                    {/* 1. Laporan KPI */}
                    <div className="p-3 bg-white border border-[#E5E5DF] rounded-2xl flex flex-col gap-2.5 shadow-2xs">
                      <div className="flex items-center justify-between">
                        <div>
                          <span className="text-xs font-bold text-[#4A4A3C] block">1. Laporan Harian KPI</span>
                          <span className="text-[10px] text-[#8C8C70]">Record lokal: <strong>{reports.length} Laporan</strong></span>
                        </div>
                        <span className="text-[10px] font-mono px-2 py-0.5 rounded-sm bg-amber-500/10 text-amber-900 border border-amber-500/20 font-bold">
                          Tab: Laporan KPI Sales
                        </span>
                      </div>
                      <button
                        onClick={handleSyncAllToSheets}
                        disabled={isSyncingAll || reports.length === 0}
                        className="w-full py-2 bg-[#5A5A40] text-[#FAF9F6] hover:bg-[#4A4A3C] disabled:bg-[#8C8C70]/30 disabled:cursor-not-allowed rounded-xl text-xs font-bold transition flex items-center justify-center gap-2 cursor-pointer shadow-3xs"
                      >
                        {isSyncingAll ? (
                          <>
                            <span className="w-3 h-3 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                            Sinkronisasi Laporan...
                          </>
                        ) : (
                          <>
                            <Download className="w-3.5 h-3.5" />
                            Sinkronkan Laporan KPI ({reports.length})
                          </>
                        )}
                      </button>
                    </div>

                    {/* 2. Database Salesman */}
                    <div className="p-3 bg-white border border-[#E5E5DF] rounded-2xl flex flex-col gap-2.5 shadow-2xs">
                      <div className="flex items-center justify-between">
                        <div>
                          <span className="text-xs font-bold text-[#4A4A3C] block">2. Database Salesman</span>
                          <span className="text-[10px] text-[#8C8C70]">Record lokal: <strong>{salesmen.length} Anggota</strong></span>
                        </div>
                        <span className="text-[10px] font-mono px-2 py-0.5 rounded-sm bg-[#4A4A3C]/10 text-[#4A4A3C] border border-[#4A4A3C]/20 font-bold">
                          Tab: Daftar Salesman
                        </span>
                      </div>
                      <button
                        onClick={handleSyncSalesmenToSheets}
                        disabled={isSyncingSalesmen || salesmen.length === 0}
                        className="w-full py-2 bg-[#4A4A3C] text-white hover:bg-[#3A3A2F] disabled:bg-[#8C8C70]/30 disabled:cursor-not-allowed rounded-xl text-xs font-bold transition flex items-center justify-center gap-2 cursor-pointer shadow-3xs"
                      >
                        {isSyncingSalesmen ? (
                          <>
                            <span className="w-3 h-3 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                            Sinkronisasi Salesman...
                          </>
                        ) : (
                          <>
                            <Download className="w-3.5 h-3.5" />
                            Sinkronkan Data Salesman ({salesmen.length})
                          </>
                        )}
                      </button>
                    </div>

                    {/* 3. Database SKU Produk */}
                    <div className="p-3 bg-white border border-[#E5E5DF] rounded-2xl flex flex-col gap-2.5 shadow-2xs">
                      <div className="flex items-center justify-between">
                        <div>
                          <span className="text-xs font-bold text-[#4A4A3C] block">3. Database SKU Produk</span>
                          <span className="text-[10px] text-[#8C8C70]">Record lokal: <strong>{products.length} SKU</strong></span>
                        </div>
                        <span className="text-[10px] font-mono px-2 py-0.5 rounded-sm bg-[#3B3D2A]/10 text-[#3B3D2A] border border-[#3B3D2A]/20 font-bold">
                          Tab: Daftar Produk SKU
                        </span>
                      </div>
                      <button
                        onClick={handleSyncProductsToSheets}
                        disabled={isSyncingProducts || products.length === 0}
                        className="w-full py-2 bg-[#3B3D2A] text-white hover:bg-[#252719] disabled:bg-[#8C8C70]/30 disabled:cursor-not-allowed rounded-xl text-xs font-bold transition flex items-center justify-center gap-2 cursor-pointer shadow-3xs"
                      >
                        {isSyncingProducts ? (
                          <>
                            <span className="w-3 h-3 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                            Sinkronisasi SKU Produk...
                          </>
                        ) : (
                          <>
                            <Download className="w-3.5 h-3.5" />
                            Sinkronkan SKU Produk ({products.length})
                          </>
                        )}
                      </button>
                    </div>

                    {/* 4. Program Loyalti & Profiling Toko */}
                    <div className="p-3 bg-white border border-[#E5E5DF] rounded-2xl flex flex-col gap-2.5 shadow-2xs">
                      <div className="flex items-center justify-between">
                        <div>
                          <span className="text-xs font-bold text-[#4A4A3C] block">4. Program Loyalti Toko & Klaim</span>
                          <span className="text-[10px] text-[#8C8C70]">
                            Local: <strong>{customers.length} Toko</strong>, <strong>{loyaltyRedeemHistory.length} Klaim</strong>
                          </span>
                        </div>
                        <span className="text-[10px] font-mono px-2 py-0.5 rounded-sm bg-[#4A4238]/10 text-[#4A4238] border border-[#4A4238]/20 font-bold">
                          Multi-Tab Sync
                        </span>
                      </div>
                      <button
                        onClick={() => handleSyncLoyaltyToSheets()}
                        disabled={isSyncingLoyalty || customers.length === 0}
                        className="w-full py-2 bg-[#4A4238] text-white hover:bg-[#342D26] disabled:bg-[#8C8C70]/30 disabled:cursor-not-allowed rounded-xl text-xs font-bold transition flex items-center justify-center gap-2 cursor-pointer shadow-3xs"
                      >
                        {isSyncingLoyalty ? (
                          <>
                            <span className="w-3 h-3 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
                            Sinkronisasi Loyalti...
                          </>
                        ) : (
                          <>
                            <Download className="w-3.5 h-3.5" />
                            Sinkronkan Database Loyalti ({customers.length} Toko)
                          </>
                        )}
                      </button>
                    </div>
                  </div>
                </div>

              </div>

              {/* Right Column: Steps and Code block */}
              <div className="lg:col-span-2 flex flex-col gap-6">
                
                {/* Tutorial Panel */}
                <div className="bg-[#FAF9F6] rounded-3xl p-6 border border-[#E5E5DF] shadow-xs">
                  <h3 className="text-sm font-extrabold text-[#4A4A3C] uppercase tracking-wider mb-2 flex items-center gap-2">
                    <Sparkles className="w-4 h-4 text-[#5A5A40]" />
                    Langkah Integrasi spreadsheet (Paling Mudah & Instan)
                  </h3>
                  <p className="text-xs text-[#8C8C70] mb-5">
                    Ikuti panduan berikut untuk menghubungkan portal kpi dengan spreadsheet google secara gratis tanpa pusing setup GCP/Firestore yang terkunci:
                  </p>

                  <div className="space-y-5">
                    {/* Step 1 */}
                    <div className="flex gap-3">
                      <div className="w-6 h-6 rounded-full bg-[#5A5A40] text-[#FAF9F6] flex items-center justify-center text-[11px] font-black shrink-0 mt-0.5">
                        1
                      </div>
                      <div>
                        <p className="text-xs font-bold text-[#4A4A3C]">Buat Google Spreadsheet Baru</p>
                        <p className="text-[11px] text-[#8C8C70] mt-0.5">
                          Buka Google Drive atau Google Sheets, lalu buat lembar spreadsheet baru. Beri judul seperti <strong>"Laporan Audit KPI Sales"</strong>.
                        </p>
                      </div>
                    </div>

                    {/* Step 2 */}
                    <div className="flex gap-3">
                      <div className="w-6 h-6 rounded-full bg-[#5A5A40] text-[#FAF9F6] flex items-center justify-center text-[11px] font-black shrink-0 mt-0.5">
                        2
                      </div>
                      <div>
                        <p className="text-xs font-bold text-[#4A4A3C]">Buka Ekstensi Apps Script</p>
                        <p className="text-[11px] text-[#8C8C70] mt-0.5">
                          Di menu bar spreadsheet, klik <strong>Ekstensi (Extensions)</strong> &gt; pilih <strong>Apps Script</strong>.
                        </p>
                      </div>
                    </div>

                    {/* Step 3 */}
                    <div className="flex gap-3">
                      <div className="w-6 h-6 rounded-full bg-[#5A5A40] text-[#FAF9F6] flex items-center justify-center text-[11px] font-black shrink-0 mt-0.5">
                        3
                      </div>
                      <div className="w-full">
                        <p className="text-xs font-bold text-[#4A4A3C]">Salin dan Tempel Kode Apps Script</p>
                        <p className="text-[11px] text-[#8C8C70] mt-0.5 mb-2">
                          Hapus kode default di editor lalu paste kode optimal di bawah ini yang telah kami siapkan khusus untuk memformat kolom otomatis secara dinamis:
                        </p>
                        
                        {/* Monospace Code Board with copy trigger */}
                        <div className="relative bg-[#3A3A2C] rounded-2xl p-4 overflow-hidden border border-[#5A5A40]/30 shadow-inner">
                          <button
                            onClick={() => {
                              navigator.clipboard.writeText(APPS_SCRIPT_CODE_STENCIL);
                              showToast("Kode Apps Script disalin ke clipboard!", "success");
                            }}
                            className="absolute right-3 top-3 text-[10px] bg-[#FAF9F6]/10 text-[#FAF9F6] hover:bg-[#FAF9F6]/20 font-bold px-3 py-1.5 rounded-lg flex items-center gap-1.5 transition cursor-pointer"
                          >
                            <Copy className="w-3 h-3" />
                            Salin Kode
                          </button>
                          
                          <pre className="max-h-56 overflow-y-auto text-[10px] font-mono text-amber-200/90 leading-relaxed pr-8 pt-6 select-all scrollbar-thin">
                            {APPS_SCRIPT_CODE_STENCIL}
                          </pre>
                        </div>
                      </div>
                    </div>

                    {/* Step 4 */}
                    <div className="flex gap-3">
                      <div className="w-6 h-6 rounded-full bg-[#5A5A40] text-[#FAF9F6] flex items-center justify-center text-[11px] font-black shrink-0 mt-0.5">
                        4
                      </div>
                      <div>
                        <p className="text-xs font-bold text-[#4A4A3C]">Deploy ke Publik / Terapkan sebagai Web App</p>
                        <p className="text-[11px] text-[#8C8C70] mt-0.5 leading-relaxed">
                          Klik tombol <strong>Terapkan (Deploy)</strong> di kanan atas &gt; pilih <strong>Terapkan Baru (New Deployment)</strong>.<br />
                          1. Ubah ikon roda gigi Jenis ke: <strong>Aplikasi Web (Web App)</strong>.<br />
                          2. Setel jalankan sebagai: <strong>Saya (Me)</strong>.<br />
                          3. Setel Pengguna Akses ke: <strong>Siapa Saja (Anyone)</strong>, lalu klik <strong>Terapkan (Deploy)</strong>.<br />
                          <span className="text-rose-700 font-semibold italic text-[10px] block mt-1">
                            *Catatan Penting: Pastikan Anda memberikan izin akses Akun Google Anda saat pop-up autentikasi Google muncul, untuk mengizinkan input data dari aplikasi web ini.
                          </span>
                        </p>
                      </div>
                    </div>

                    {/* Step 5 */}
                    <div className="flex gap-3">
                      <div className="w-6 h-6 rounded-full bg-[#5A5A40] text-[#FAF9F6] flex items-center justify-center text-[11px] font-black shrink-0 mt-0.5">
                        5
                      </div>
                      <div>
                        <p className="text-xs font-bold text-[#4A4A3C]">Salin URL Aplikasi Web & Tempel</p>
                        <p className="text-[11px] text-[#8C8C70] mt-0.5">
                          Salin <strong>Web App URL</strong> yang diberikan, tempel di kolom "Koneksi Spreadsheet" di sebelah kiri, lalu mulailah mengirim data!
                        </p>
                      </div>
                    </div>

                  </div>
                </div>

              </div>
            </motion.div>
          )}

          {/* TAB 6: PROGRAM LOYALTI & CUSTOMER PROFILING (INTEGRASI SISTEM JIMMY & DKR SHOWN IN FLOWCHART) */}
          {activeTab === "loyalty" && (
            <motion.div
              key="loyalty-tab"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              className="space-y-6"
            >
              {customerMode ? (
                <CustomerLoyaltyPortal
                  customers={customers}
                  setCustomers={setCustomers}
                  sheetsScriptUrl={sheetsScriptUrl}
                  salesmen={salesmen}
                  handleSyncLoyaltyToSheets={handleSyncLoyaltyToSheets}
                  showToast={showToast}
                  setCustomerMode={setCustomerMode}
                  customerActiveSubTab={customerActiveSubTab}
                  setCustomerActiveSubTab={setCustomerActiveSubTab}
                  customerSearchQuery={customerSearchQuery}
                  setCustomerSearchQuery={setCustomerSearchQuery}
                  matchedCustomer={matchedCustomer}
                  setMatchedCustomer={setMatchedCustomer}
                  registerSuccessName={registerSuccessName}
                  setRegisterSuccessName={setRegisterSuccessName}
                  isCustomerSelfRegistering={isCustomerSelfRegistering}
                  setIsCustomerSelfRegistering={setIsCustomerSelfRegistering}
                  selfRegForm={selfRegForm}
                  setSelfRegForm={setSelfRegForm}
                  REWARDS_CATALOG={REWARDS_CATALOG}
                />
              ) : (
                <>
                  {/* Upper Loyalty Insights Bento Strip */}
              <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="bg-gradient-to-br from-[#FAF9F6] to-[#E5E5DF]/20 p-4 rounded-2xl border border-[#E5E5DF] shadow-xs">
                  <span className="text-[10px] font-bold text-[#8C8C70] block uppercase tracking-wider">Total Toko Diprofiling</span>
                  <span className="text-xl font-extrabold text-[#4A4A3C] block mt-1">{customers.length} Toko</span>
                  <span className="text-[10px] text-[#5A5A40] font-semibold">Tersinkron di Portal</span>
                </div>
                
                <div className="bg-gradient-to-br from-[#FAF9F6] to-amber-500/5 p-4 rounded-2xl border border-[#E5E5DF] shadow-xs">
                  <span className="text-[10px] font-bold text-[#8C8C70] block uppercase tracking-wider">Total Poin Terkumpul</span>
                  <div className="flex items-center gap-1.5 mt-1">
                    <Crown className="w-5 h-5 text-amber-500 shrink-0" />
                    <span className="text-xl font-extrabold text-[#4A4A3C]">{customers.reduce((sum, c) => sum + (c.points || 0), 0)} Poin</span>
                  </div>
                  <span className="text-[10px] text-[#8C8C70]">Siap ditukar merchant rewards</span>
                </div>

                <div className="bg-gradient-to-br from-[#FAF9F6] to-[#5A5A40]/5 p-4 rounded-2xl border border-[#E5E5DF] shadow-xs">
                  <span className="text-[10px] font-bold text-[#8C8C70] block uppercase tracking-wider">Toko Kelas Platinum / Gold</span>
                  <span className="text-xl font-extrabold text-[#5A5A40] block mt-1">
                    {customers.filter(c => c.tier === "Platinum" || c.tier === "Gold").length} Toko Utama
                  </span>
                  <span className="text-[10px] text-[#8C8C70]">Omzet &gt; Rp 8.000.000/bln</span>
                </div>

                <div className="bg-gradient-to-br from-[#FAF9F6] to-[#8C8C70]/10 p-4 rounded-2xl border border-[#E5E5DF] shadow-xs">
                  <span className="text-[10px] font-bold text-[#8C8C70] block uppercase tracking-wider">Total Klaim Hadiah</span>
                  <span className="text-xl font-extrabold text-[#8C8C70] block mt-1">{loyaltyRedeemHistory.length} Klaim</span>
                  <span className="text-[10px] text-emerald-700/80 font-semibold">Sukses didistribusikan</span>
                </div>
              </div>

              {/* Main Bento Layout Grid */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                
                {/* COLUMN 1: INTERACTIVE ACTION SYNC PANEL & REWARDS SIMULATOR */}
                <div className="lg:col-span-1 space-y-6">
                  
                  {/* Action 1: Register New Store Profiling */}
                  <div className="bg-[#FAF9F6] p-5 rounded-3xl border border-[#E5E5DF] shadow-xs">
                    <h3 className="text-xs font-black text-[#4A4A3C] uppercase tracking-wider mb-3 flex items-center gap-2">
                      <Sparkles className="w-4 h-4 text-amber-500" />
                      1. Customer Profiling Form
                    </h3>
                    <div className="max-h-[350px] overflow-y-auto pr-1.5 space-y-4">
                      <div>
                        <p className="text-[11px] text-[#8C8C70] leading-relaxed mb-3">
                          Sesuai diagram DKR Sales System, lakukan profiling dasar (Nama Toko, Estimasi Omzet, Status Sewa) untuk mendaftarkan tingkatan tier loyalitas toko baru.
                        </p>
                        <button
                          onClick={() => setIsCustomerModalOpen(true)}
                          className="w-full bg-[#5A5A40] hover:bg-[#4A4A3C] text-[#FAF9F6] hover:text-white font-bold text-xs uppercase tracking-wider py-3 px-4 rounded-xl transition cursor-pointer flex items-center justify-center gap-2"
                        >
                          <Plus className="w-4 h-4" />
                          Input Profil Toko Baru
                        </button>
                      </div>

                      {/* QR Code and Customer Link section */}
                      <div className="border-t border-dashed border-[#E5E5DF] pt-4 space-y-3">
                        <div className="flex items-center justify-between">
                          <span className="text-[10.5px] font-extrabold text-[#4A4A3C] uppercase flex items-center gap-1.5">
                            📠 Portal Mandiri Pelanggan
                          </span>
                          <span className="text-[8.5px] bg-[#5A5A40] text-amber-300 font-extrabold px-1.5 py-0.5 rounded uppercase font-mono tracking-wider">
                            QR CODE
                          </span>
                        </div>
                        
                        <p className="text-[10px] text-[#8C8C70] leading-relaxed">
                          Cetak flyer QR ini di outlet Semarang & Bobotsari. Pelanggan dapat scan mandiri untuk mendaftar profiling & memantau saldo poin toko!
                        </p>

                        {(() => {
                          const publicUrl = "https://docs.google.com/forms/d/e/1FAIpQLScX4GaMKzI0wjwDtzaXZ9r5IUiX5-uDSTCCFbitrH7Ckz-iQQ/viewform";
                          const qrApiUrl = `https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=${encodeURIComponent(publicUrl)}`;
                          
                          const handleDownloadQR = async () => {
                            try {
                              const response = await fetch(qrApiUrl);
                              const blob = await response.blob();
                              const url = window.URL.createObjectURL(blob);
                              const a = document.createElement("a");
                              a.href = url;
                              a.download = "QR_Code_Loyalty.png";
                              document.body.appendChild(a);
                              a.click();
                              document.body.removeChild(a);
                              window.URL.revokeObjectURL(url);
                              showToast("QR Code berhasil diunduh!", "success");
                            } catch (err) {
                              showToast("Gagal mengunduh QR Code.", "error");
                            }
                          };
                          
                          return (
                            <div className="bg-white p-2.5 rounded-2xl border border-[#E5E5DF] flex flex-col items-center justify-center text-center gap-2">
                              <img
                                src={qrApiUrl}
                                alt="Loyalty Portal Mandiri QR Code"
                                referrerPolicy="no-referrer"
                                className="w-28 h-28 border-2 border-[#5A5A40] rounded-xl shadow-xs"
                              />
                              
                              <div className="w-full space-y-1">
                                <button
                                  type="button"
                                  onClick={handleDownloadQR}
                                  className="w-full py-1.5 text-[9px] font-bold bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg transition mb-1"
                                >
                                  ⬇️ Unduh Gambar QR (PNG)
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    try {
                                      navigator.clipboard.writeText(publicUrl);
                                      showToast("Link Google Form berhasil disalin ke papan klip!", "success");
                                    } catch (err) {
                                      showToast("Gagal menyalin otomatis. Salin manual dari jendela alamat Anda.", "error");
                                    }
                                  }}
                                  className="w-full py-1.5 text-[9px] font-bold bg-[#E5E5DF]/50 hover:bg-[#E5E5DF] text-[#4A4A3C] rounded-lg transition"
                                >
                                  📋 Salin Link Google Form
                                </button>
                                <button
                                  type="button"
                                  onClick={() => window.open(publicUrl, "_blank")}
                                  className="w-full py-1.5 text-[9px] font-black bg-[#5A5A40] hover:bg-[#4A4A3C] text-[#FAF9F6] rounded-lg transition cursor-pointer"
                                >
                                  👀 Test / Buka Google Form
                                </button>
                              </div>
                            </div>
                          );
                        })()}
                      </div>
                      
                      {/* Google Form Integration Block */}
                      <div className="border-t border-dashed border-[#E5E5DF] pt-4 space-y-3">
                        <div className="flex items-center justify-between">
                          <span className="text-[10.5px] font-extrabold text-[#4A4A3C] uppercase flex items-center gap-1.5">
                            📊 Integrasi Google Form Publik
                          </span>
                          <span className="text-[8.5px] bg-red-700 text-white font-extrabold px-1.5 py-0.5 rounded uppercase font-mono tracking-wider">
                            BEBAS SANDBOX
                          </span>
                        </div>
                        
                        <p className="text-[10px] text-[#8C8C70] leading-relaxed">
                          Karena batasan preview sandbox, pelanggan eksternal tidak bisa membuka link portal mandiri di atas. 
                          <strong> Solusi Terbaik: Hubungkan dengan Google Form publik!</strong> Pelanggan mengisi form publik, lalu Anda sinkronkan di sini dengan mudah.
                        </p>
                        
                        <div className="space-y-2">
                          <button
                            type="button"
                            onClick={() => setIsGoogleFormGuideOpen(!isGoogleFormGuideOpen)}
                            className="w-full py-2 text-[9px] font-bold bg-amber-500/10 hover:bg-amber-500/20 text-amber-900 border border-amber-500/30 rounded-xl transition flex items-center justify-center gap-1.5"
                          >
                            📰 {isGoogleFormGuideOpen ? "Sembunyikan Panduan Setup" : "Buka Panduan & Auto-Script Google Form (3 Menit)"}
                          </button>

                          {isGoogleFormGuideOpen && (
                            <div className="bg-amber-500/5 p-3 rounded-2xl border border-amber-500/20 text-[10px] space-y-2 text-[#4A4A3C] leading-normal">
                              <p className="font-extrabold text-amber-900">
                                🚀 Cara Tercepat & Termudah:
                              </p>
                              <ol className="list-decimal pl-4.5 space-y-2">
                                <li>
                                  Buka Spreadsheet Anda, pilih menu <strong>Ekstensi (Extensions)</strong> &gt; <strong>Apps Script</strong>.
                                </li>
                                <li>
                                  Klik ikon <strong>(+) Tambah File</strong> &gt; pilih <strong>Skrip (Script)</strong>. Beri nama file baru: <code className="bg-white px-1 py-0.5 rounded border font-mono">FormGenerator.gs</code>.
                                </li>
                                <li>
                                  Salin seluruh kode script generator di bawah ini, lalu paste di file tersebut, kemudian klik ikon <strong>Simpan (💾)</strong>:
                                  
                                  <div className="relative mt-1.5 bg-[#3A3A2C] rounded-xl p-2.5 overflow-hidden">
                                    <button
                                      type="button"
                                      onClick={() => {
                                        const scriptText = `/**
 * CARA PAKAI:
 * 1. Di Google Sheets Anda, klik "Ekstensi" > "Apps Script".
 * 2. Klik ikon "+" lalu pilih "Skrip" (Script) untuk membuat file baru.
 * 3. Beri nama file baru tersebut: "FormGenerator.gs".
 * 4. Paste seluruh kode di bawah ini, lalu klik ikon simpan (💾).
 * 5. Pilih fungsi "createCustomerProfilingForm" di menu dropdown atas, lalu klik "Jalankan" (Run).
 * 6. Beri otorisasi jika diminta. Google Form akan otomatis dibuat & terhubung ke Spreadsheet Anda!
 */
function createCustomerProfilingForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.create("Customer Profiling & Loyalty Sign-up Form")
      .setDescription("Formulir pendaftaran Program Loyalitas & Profiling Dasar Toko DKR Retail.")
      .setAllowResponseEdits(true)
      .setConfirmationMessage("Terima kasih! Toko/Outlet Anda telah berhasil terdaftar ke sistem Loyalitas DKR Sales. Poin starter +50 Poin telah Anda dapatkan!");
  
  form.addTextItem().setTitle("Nama Toko / Outlet Anda").setRequired(true);
  form.addParagraphTextItem().setTitle("Alamat Lengkap Outlet").setRequired(true);
  
  var selectSales = form.addMultipleChoiceItem();
  selectSales.setTitle("Pilih Sales Wilayah Anda")
      .setChoiceValues(["RIZKY", "WILDAN", "RENY", "BUDI", "HENDRI", "FANDI", "ANTON"])
      .setRequired(true);
  
  var selectToko = form.addMultipleChoiceItem();
  selectToko.setTitle("Jenis & Sektor Toko")
      .setChoiceValues(["Sembako", "Kelontong Kecil", "Kelontong Grosir Utama", "Pengecer Keliling", "Lainnya"])
      .setRequired(true);
      
  form.addTextItem().setTitle("Estimasi Kas Belanja Outlet Per Bulan (Rupiah)")
      .setHelpText("Contoh: 5000000 (tulis angka saja tanpa Rp, titik, atau koma)")
      .setRequired(true);
      
  var selectSewa = form.addMultipleChoiceItem();
  selectSewa.setTitle("Status Sewa Bangunan")
      .setChoiceValues(["Milik Sendiri", "Sewa Kontrak Bulanan", "Milik Keluarga/Warisan"])
      .setRequired(true);
      
  form.addTextItem().setTitle("Lama Toko Berdiri (Tahun)")
      .setHelpText("Contoh: 3 (tulis angka umur berdiri toko saja)")
      .setRequired(true);
      
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  
  Logger.log("Form URL: " + form.getPublishedUrl());
  Browser.msgBox("SUKSES! Google Form telah dibuat di Drive Anda & dihubungkan ke spreadsheet ini sebagai tab respon baru! Tautkan link form tersebut ke poster/pelanggan Anda.");
}`;
                                        try {
                                          navigator.clipboard.writeText(scriptText);
                                          showToast("Script Generator Google Form berhasil disalin!", "success");
                                        } catch (_) {
                                          showToast("Gagal menyalin otomatis, silakan block manual teks code.", "error");
                                        }
                                      }}
                                      className="absolute right-2 top-2 text-[8px] bg-[#FAF9F6]/20 text-[#FAF9F6] hover:bg-[#FAF9F6]/30 font-bold px-2 py-1 rounded-md transition cursor-pointer"
                                    >
                                      Salin Script
                                    </button>
                                    <pre className="max-h-28 overflow-y-auto text-[8.5px] font-mono text-amber-200/90 leading-relaxed pt-3 pr-2 select-all scrollbar-thin">
{`function createCustomerProfilingForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.create("Customer Profiling & Loyalty Sign-up Form")
      .setDescription("Formulir pendaftaran Program Loyalitas & Profiling Dasar Toko DKR Retail.")
      .setAllowResponseEdits(true)
      .setConfirmationMessage("Terima kasih! Toko/Outlet Anda telah berhasil terdaftar ke sistem Loyalitas DKR Sales. Poin starter +50 Poin telah Anda dapatkan!");
  
  form.addTextItem().setTitle("Nama Toko / Outlet Anda").setRequired(true);
  form.addParagraphTextItem().setTitle("Alamat Lengkap Outlet").setRequired(true);
  
  var selectSales = form.addMultipleChoiceItem();
  selectSales.setTitle("Pilih Sales Wilayah Anda")
      .setChoiceValues(["RIZKY", "WILDAN", "RENY", "BUDI", "HENDRI", "FANDI", "ANTON"])
      .setRequired(true);
  
  var selectToko = form.addMultipleChoiceItem();
  selectToko.setTitle("Jenis & Sektor Toko")
      .setChoiceValues(["Sembako", "Kelontong Kecil", "Kelontong Grosir Utama", "Pengecer Keliling", "Lainnya"])
      .setRequired(true);
      
  form.addTextItem().setTitle("Estimasi Kas Belanja Outlet Per Bulan (Rupiah)").setRequired(true);
  
  var selectSewa = form.addMultipleChoiceItem();
  selectSewa.setTitle("Status Sewa Bangunan")
      .setChoiceValues(["Milik Sendiri", "Sewa Kontrak Bulanan", "Milik Keluarga/Warisan"])
      .setRequired(true);
      
  form.addTextItem().setTitle("Lama Toko Berdiri (Tahun)").setRequired(true);
  
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
}`}
                                    </pre>
                                  </div>
                                </li>
                                <li>
                                  Pilih fungsi <strong className="font-mono">createCustomerProfilingForm</strong> di dropdown atas editor Google Apps Script, lalu klik <strong>Jalankan (Run)</strong>.
                                </li>
                                <li>
                                  Google Form akan otomatis dibuat di Google Drive Anda, dan tab penampung respon bernama <strong className="font-mono">"Jawaban Formulir 1" (atau "Form Responses 1")</strong> akan muncul di spreadsheet.
                                </li>
                                <li>
                                  Hubungkan Google Form publik tersebut agar diisi oleh mitra ritel Anda secara bebas. Di sini, Anda tinggal mengklik tombol sinkronisasi di bawah!
                                </li>
                              </ol>
                            </div>
                          )}

                          <button
                            type="button"
                            onClick={() => window.open("https://docs.google.com/forms/d/e/1FAIpQLScX4GaMKzI0wjwDtzaXZ9r5IUiX5-uDSTCCFbitrH7Ckz-iQQ/viewform", "_blank")}
                            className="w-full mb-3 py-3 px-4 font-black rounded-xl transition cursor-pointer flex items-center justify-center gap-2 text-xs uppercase tracking-wider bg-indigo-700 text-white hover:bg-indigo-800 shadow-md hover:shadow-lg"
                          >
                            <ExternalLink className="w-4 h-4" />
                            Buka Google Form Profiling
                          </button>

                          <button
                            type="button"
                            disabled={isImportingGoogleForm}
                            onClick={handleImportFromGoogleForm}
                            className={`w-full py-3 px-4 font-black rounded-xl transition cursor-pointer flex items-center justify-center gap-2 text-xs uppercase tracking-wider ${
                              isImportingGoogleForm
                                ? "bg-amber-300 text-[#4A4A3C] animate-pulse cursor-not-allowed"
                                : "bg-gradient-to-r from-amber-500 to-[#5A5A40] text-white hover:opacity-95 shadow-md hover:shadow-lg"
                            }`}
                          >
                            <RefreshCw className={`w-4 h-4 ${isImportingGoogleForm ? "animate-spin" : ""}`} />
                            {isImportingGoogleForm ? "Menghubungkan & Sinkron..." : "🔄 Sinkronkan Respon Google Form"}
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Action 2: Tindakan Follow-Up Simulator (Section 4 in Flowchart) */}
                  <div className="bg-[#FAF9F6] p-5 rounded-3xl border border-[#E5E5DF] shadow-xs relative overflow-hidden">
                    <div className="absolute top-0 right-0 px-2.5 py-0.5 bg-amber-500 text-[8px] font-black text-white rounded-bl-xl uppercase tracking-widest">
                      Audit Action
                    </div>

                    <h3 className="text-xs font-black text-[#4A4A3C] uppercase tracking-wider mb-2 flex items-center gap-2">
                      <Clipboard className="w-4 h-4 text-[#5A5A40]" />
                      2. Tindakan / Follow-Up (Aris Desk)
                    </h3>
                    
                    <div className="max-h-[350px] overflow-y-auto pr-1.5 space-y-3.5">
                      <p className="text-[11px] text-[#8C8C70] leading-relaxed">
                        Setiap toko yang mendapat peringatan (misal lesu order atau pinalti) wajib diisi tindakan konkret oleh Kepala Sales.
                      </p>

                      <div className="space-y-3.5">
                        <div>
                          <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                            Pilih Toko Terdaftar:
                          </label>
                          <select
                            value={selectedCustomerIdForAction}
                            onChange={(e) => setSelectedCustomerIdForAction(e.target.value)}
                            className="w-full bg-white border border-[#E5E5DF] rounded-xl px-2.5 py-2 text-xs text-[#4A4A3C] font-semibold focus:outline-hidden focus:ring-1 focus:ring-[#5A5A40]"
                          >
                            {customers.map(c => (
                              <option key={c.id} value={c.id}>
                                {c.name} ({c.jenisToko} - {c.area})
                              </option>
                            ))}
                          </select>
                        </div>

                        <div className="grid grid-cols-2 gap-2">
                          <div>
                            <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                              Aksi Tindakan:
                            </label>
                            <select
                              value={followUpAction}
                              onChange={(e) => setFollowUpAction(e.target.value)}
                              className="w-full bg-white border border-[#E5E5DF] rounded-xl px-2 py-1.5 text-xs text-[#4A4A3C] focus:outline-hidden"
                            >
                              <option value="Kunjungan">🚶 Kunjungan Toko</option>
                              <option value="Kasih Promo">🎁 Kasih Promo</option>
                              <option value="Telepon">📞 Telepon Pengingat</option>
                              <option value="Penagihan">💰 Penagihan Nota</option>
                              <option value="Meeting">🤝 Meeting Sales</option>
                              <option value="Lainnya">📝 Lainnya</option>
                            </select>
                          </div>

                          <div>
                            <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                              Jadwal Tindakan:
                            </label>
                            <input
                              type="text"
                              readOnly
                              value={new Date().toLocaleDateString("id-ID")}
                              className="w-full bg-[#E5E5DF]/30 border border-[#E5E5DF] rounded-xl px-2 py-1.5 text-xs text-[#8C8C70] font-bold text-center font-mono"
                            />
                          </div>
                        </div>

                        <div>
                          <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                            Catatan & Detail Rencana Tindakan:
                          </label>
                          <textarea
                            rows={2}
                            value={followUpNotes}
                            onChange={(e) => setFollowUpNotes(e.target.value)}
                            placeholder="Contoh: Akan kunjungi toko untuk cek stok & tawarkan promo paket sembako..."
                            className="w-full bg-white border border-[#E5E5DF] rounded-xl p-2 text-xs text-[#4A4A3C] placeholder:text-[#8C8C70]/70 focus:outline-hidden focus:ring-1 focus:ring-[#5A5A40]"
                          />
                        </div>

                        <button
                          type="button"
                          onClick={handleAddFollowUpAction}
                          className="w-full bg-[#5A5A40] hover:bg-[#4A4A3C] text-[#FAF9F6] font-bold text-xs uppercase tracking-wider py-2.5 rounded-xl transition cursor-pointer"
                        >
                          Simpan Tindakan (Poin +15)
                        </button>
                      </div>
                    </div>
                  </div>

                </div>

                {/* COLUMN 2 & 3: DETAILS TABLE AND TIER LIST LOGS */}
                <div className="lg:col-span-2 space-y-6">
                  
                  {/* Customer database detailing list */}
                  <div className="bg-[#FAF9F6] rounded-3xl p-6 border border-[#E5E5DF] shadow-xs">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4.5">
                      <div>
                        <h3 className="text-sm font-extrabold text-[#4A4A3C] uppercase tracking-wider">
                          Daftar Profiling Retail DKR Sales
                        </h3>
                        <p className="text-[10px] text-[#8C8C70]">Klasifikasi loyalitas otomatis berdasarkan estimasi omzet bulanan</p>
                      </div>
                      <span className="text-[10px] px-3 py-1 bg-[#FAF9F6] border border-[#E5E5DF] rounded-lg font-mono text-[#8C8C70] font-bold shrink-0 self-start sm:self-auto">
                        Total {customers.length} Toko
                      </span>
                    </div>

                    <div className="overflow-y-auto max-h-[500px] pr-1.5">
                      <div className="overflow-x-auto">
                        <table className="w-full text-left border-collapse">
                          <thead>
                            <tr className="border-b border-[#E5E5DF] text-[10px] uppercase tracking-wider text-[#8C8C70] font-black bg-[#E5E5DF]/10">
                              <th className="py-3 px-3">NAMA TOKO</th>
                              <th className="py-3 px-3">Sales & Area</th>
                              <th className="py-3 px-3">Statistik Toko</th>
                              <th className="py-3 px-3">Poin Loyalitas</th>
                              <th className="py-3 px-3 text-right">Aksi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-[#E5E5DF]/50 text-xs">
                            {customers.map(c => {
                              // Define badge styling based on calculated Tiers
                              let tierBadgeColor = "bg-[#4A4A3C]/10 text-[#4A4A3C] border-gray-300";
                              let iconColor = "text-[#8C8C70]";
                              if (c.tier === "Platinum") {
                                tierBadgeColor = "bg-purple-100 text-purple-800 border-purple-300 font-extrabold";
                                iconColor = "text-purple-600";
                              } else if (c.tier === "Gold") {
                                tierBadgeColor = "bg-amber-100 text-amber-900 border-amber-300 font-bold";
                                iconColor = "text-amber-500";
                              } else if (c.tier === "Silver") {
                                tierBadgeColor = "bg-blue-100 text-blue-800 border-blue-200 font-semibold";
                                iconColor = "text-slate-500";
                              } else {
                                tierBadgeColor = "bg-orange-100 text-orange-900 border-orange-200";
                                iconColor = "text-orange-700";
                              }

                              return (
                                <React.Fragment key={c.id}>
                                  <tr className="hover:bg-[#E5E5DF]/10 transition-colors">
                                    <td className="py-3 px-3">
                                      <div className="flex flex-col gap-1">
                                        <span className="font-bold text-[#4A4A3C] text-sm uppercase">{c.name}</span>
                                        <span className="text-[10px] font-mono text-[#8C8C70] truncate max-w-[200px]" title={c.address}>
                                          📍 {c.address}
                                        </span>
                                        <div className="flex items-center gap-1.5 mt-0.5">
                                          <span className={`inline-flex items-center px-2 py-0.5 text-[8.5px] border rounded-full uppercase tracking-wider font-bold ${tierBadgeColor}`}>
                                            ★ Tier {c.tier}
                                          </span>
                                          <span className="text-[9px] text-[#8C8C70] bg-[#FAF9F6] border border-[#E5E5DF] px-1.5 rounded uppercase">
                                            {c.jenisToko}
                                          </span>
                                        </div>
                                      </div>
                                    </td>
                                    
                                    <td className="py-3 px-3">
                                      <div className="flex flex-col">
                                        <span className="font-bold text-[#4A4A3C] text-[11px]">{c.salesmanName}</span>
                                        <span className="text-[10px] text-[#8C8C70]">Area: {c.area}</span>
                                      </div>
                                    </td>

                                    <td className="py-3 px-3">
                                      <div className="flex flex-col gap-0.5 text-[10px] text-[#4A4A3C]">
                                        <div>Est. Omzet: <span className="font-bold font-mono">Rp {c.estimatedOmzet?.toLocaleString("id-ID")}</span></div>
                                        <div>Nota / Hari: <span className="font-bold font-mono">{c.notesPerDay || 0} nota</span></div>
                                        <div>Kepemilikan: <span className="italic text-[#8C8C70]">{c.ownership} ({c.storeAgeYears} th)</span></div>
                                      </div>
                                    </td>

                                    <td className="py-3 px-3">
                                      <div className="flex items-center gap-1.5">
                                        <div className="w-7 h-7 rounded-full bg-amber-500/10 flex items-center justify-center text-amber-600 font-bold shrink-0">
                                          🪙
                                        </div>
                                        <div className="flex flex-col">
                                          <span className="font-mono text-xs font-black text-[#4A4A3C]">{c.points || 0}</span>
                                          <span className="text-[9px] text-[#8C8C70] uppercase leading-none">Poin DKR</span>
                                        </div>
                                      </div>
                                    </td>

                                    <td className="py-3 px-3 text-right">
                                      <button
                                        onClick={() => handleDeleteCustomer(c.id, c.name)}
                                        className="p-1 text-[#8C8C70] hover:text-rose-600 rounded transition cursor-pointer"
                                        title="Hapus Profil Toko"
                                      >
                                        <Trash2 className="w-4 h-4" />
                                      </button>
                                    </td>
                                  </tr>

                                  {/* Nested Actions Log timeline for the Customer */}
                                  <tr>
                                    <td colSpan={5} className="py-2.5 px-6 bg-[#FAF9F6]/60 border-b border-[#E5E5DF]/45">
                                      <div className="flex flex-col gap-1">
                                        <span className="text-[9px] font-black text-[#8C8C70] uppercase tracking-wider block">
                                          Riwayat Tindakan & Follow-Up ({c.actionsLog?.length || 0}):
                                        </span>
                                        
                                        {c.actionsLog && c.actionsLog.length > 0 ? (
                                          <div className="space-y-1.5 mt-1 list-none pl-0">
                                            {c.actionsLog.map((log: any, idx: number) => (
                                              <div key={idx} className="flex flex-col sm:flex-row sm:items-center justify-between text-[11px] bg-white/50 p-2 rounded-lg border border-[#E5E5DF]/40 gap-1.5">
                                                <div className="flex items-center gap-2">
                                                  <span className="font-mono text-[9px] bg-[#E5E5DF] text-[#4A4A3C] px-1.5 py-0.5 rounded font-bold shrink-0">
                                                    {log.date}
                                                  </span>
                                                  <span className="font-bold text-[#5A5A40]">
                                                    {log.action}
                                                  </span>
                                                  <span className="text-[#8C8C70] italic truncate max-w-[300px]" title={log.notes}>
                                                    "-" {log.notes}
                                                  </span>
                                                </div>
                                                <span className="text-[9px] bg-emerald-500/15 text-emerald-800 font-bold px-1.5 py-0.5 rounded self-start sm:self-auto">
                                                  ✓ {log.status || "Selesai"}
                                                </span>
                                              </div>
                                            ))}
                                          </div>
                                        ) : (
                                          <span className="text-[10px] text-[#8C8C70] italic">Belum ada rincian tindakan follow-up audit. Pilih panel simulator sebelah kiri untuk menjadwalkan tindakan baru.</span>
                                        )}
                                      </div>
                                    </td>
                                  </tr>
                                </React.Fragment>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>

                </div>

              </div>
                </>
              )}
            </motion.div>
          )}

          {/* TAB 8: KLAIM HADIAH BRAND RE-ARCHITECTED TAB */}
          {activeTab === "claims" && (
            <motion.div
              key="claims-tab"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              className="space-y-6"
            >
              {/* Header Banner */}
              <div className="bg-gradient-to-r from-emerald-600 to-teal-700 rounded-3xl p-6 text-white shadow-lg relative overflow-hidden">
                <div className="absolute top-0 right-0 -mr-20 -mt-10 w-80 h-80 bg-white/10 rounded-full blur-2xl" />
                <div className="absolute bottom-0 right-2 w-32 h-32 bg-teal-400/20 rounded-full blur-xl" />
                
                <h2 className="text-xl md:text-2xl font-serif italic font-extrabold tracking-tight flex items-center gap-2">
                  <Gift className="w-6 h-6 shrink-0 text-white animate-bounce" />
                  PORTAL KLAIM HADIAH RETAIL DKR
                </h2>
                <p className="text-xs text-emerald-50/90 max-w-2xl mt-2 leading-relaxed">
                  Pusat penukaran & pengelolaan cinderamata loyalitas toko pelanggan terhubung langsung dengan basis data spreadsheet utama. Salesman DKR dapat mengkonversi poin toko terpilih secara instan dengan sinkronisasi Sheets otomatis harian.
                </p>

                {/* Pull Data Controls inside Header */}
                <div className="flex flex-wrap items-center gap-3 mt-6 pt-6 border-t border-white/20">
                  <button
                    onClick={() => handleSyncLoyaltyToSheets(undefined, undefined, false)}
                    disabled={isSyncingLoyalty}
                    className="bg-white hover:bg-emerald-50 text-emerald-900 font-extrabold text-xs uppercase tracking-wider px-5 py-3 rounded-xl transition shadow-sm hover:shadow-md flex items-center gap-2 cursor-pointer disabled:opacity-50"
                  >
                    <RefreshCw className={`w-4 h-4 ${isSyncingLoyalty ? "animate-spin" : ""}`} />
                    {isSyncingLoyalty ? "Menghubungi Google Sheets..." : "Sinkronisasikan Manual Ke Spreadsheet"}
                  </button>
                </div>
              </div>

              {/* Main Bento Layout Grid */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                
                {/* COLUMN 1: INTERACTIVE CLAIM REWARD MERCHANT CARD */}
                <div className="lg:col-span-1 space-y-6">
                  
                  {/* Action 3: Loyalty Gift Exchange Simulator */}
                  <div className="bg-[#FAF9F6] p-6 rounded-3xl border border-[#E5E5DF] shadow-xs relative overflow-hidden">
                    <div className="absolute top-0 right-0 px-2.5 py-0.5 bg-emerald-600 text-[8px] font-black text-white rounded-bl-xl uppercase tracking-widest">
                      Merchant Redeem
                    </div>

                    <h3 className="text-xs font-black text-[#4A4A3C] uppercase tracking-wider mb-2 flex items-center gap-2">
                      <Gift className="w-4 h-4 text-emerald-600" />
                      Klaim Reward Merchant
                    </h3>
                    
                    <div className="space-y-4">
                      <p className="text-[11px] text-[#8C8C70] leading-relaxed">
                        Mengkonversi poin loyalti yang diraih toko pelanggan setia dengan cinderamata atau subsidi tagihan penjualan DKR.
                      </p>

                      <div className="space-y-3.5">
                        <div>
                          <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                            Pilih Toko:
                          </label>
                          <select
                            value={selectedCustomerIdForRedeem}
                            onChange={(e) => setSelectedCustomerIdForRedeem(e.target.value)}
                            className="w-full bg-white border border-[#E5E5DF] rounded-xl px-2.5 py-2 text-xs text-[#4A4A3C] font-semibold focus:outline-hidden"
                          >
                            {customers.map(c => (
                              <option key={c.id} value={c.id}>
                                {c.name} ({c.points || 0} Poin Tersedia)
                              </option>
                            ))}
                          </select>
                        </div>

                        <div>
                          <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                            Katalog Hadiah (Sponsor):
                          </label>
                          <select
                            value={selectedRewardId}
                            onChange={(e) => setSelectedRewardId(e.target.value)}
                            className="w-full bg-white border border-[#E5E5DF] rounded-xl px-2.5 py-2 text-xs text-[#4A4A3C] focus:outline-hidden"
                          >
                            {[...rewardMerchants, ...katalogHadiah].map(r => (
                              <option key={r.id} value={r.id}>
                                {r.name} - Nilai {('pointsRequired' in r ? r.pointsRequired : (r as CatalogHadiah).pointsValue)} Poin
                              </option>
                            ))}
                          </select>
                        </div>

                        <button
                          type="button"
                          onClick={handleRedeemReward}
                          className="w-full bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs uppercase tracking-wider py-2.5 rounded-xl transition cursor-pointer"
                        >
                          Proses Pengambilan Hadiah
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Informational Rewards Display Panel */}
                  <div className="bg-[#FAF9F6] p-6 rounded-3xl border border-[#E5E5DF] shadow-xs">
                    <div className="flex items-center justify-between mb-3">
                      <h3 className="text-xs font-black text-[#4A4A3C] uppercase tracking-wider flex items-center gap-2">
                        <Crown className="w-4 h-4 text-amber-500" />
                        Katalog Resmi & Nilai Hadiah
                      </h3>
                      <div className="flex gap-2">
                        <button onClick={() => setRewardModal({ isOpen: true, type: 'merchant' })} className="text-[10px] bg-white border border-[#E5E5DF] px-2 py-1 rounded-md hover:bg-gray-50">+ Merchant</button>
                        <button onClick={() => setRewardModal({ isOpen: true, type: 'catalog' })} className="text-[10px] bg-white border border-[#E5E5DF] px-2 py-1 rounded-md hover:bg-gray-50">+ Katalog</button>
                      </div>
                    </div>
                    <div className="space-y-2.5">
                      {[...rewardMerchants, ...katalogHadiah].map(item => (
                        <div 
                          key={item.id} 
                          onClick={() => setRewardModal({ isOpen: true, type: 'pointsRequired' in item ? 'merchant' : 'catalog', item })}
                          className="p-3 bg-white border border-[#E5E5DF]/60 rounded-xl flex items-center justify-between gap-3 text-xs cursor-pointer hover:border-[#5A5A40] transition"
                        >
                          <div>
                            <div className="font-bold text-[#4A4A3C]">{item.name}</div>
                            <div className="text-[10px] text-[#8C8C70]">{'sponsor' in item ? `Sponsor: ${item.sponsor}` : item.description}</div>
                          </div>
                          <div className="text-emerald-700 font-extrabold pr-1 whitespace-nowrap">
                            {('pointsRequired' in item ? item.pointsRequired : (item as CatalogHadiah).pointsValue)} Poin
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                </div>

                {/* COLUMN 2 & 3: DETAILS TABLE AND TIER LIST LOGS */}
                <div className="lg:col-span-2 space-y-6">
                  
                  {/* Claims Rewards log list */}
                  <div className="bg-[#FAF9F6] rounded-3xl p-6 border border-[#E5E5DF] shadow-xs">
                    <h3 className="text-xs font-black text-[#4A4A3C] uppercase tracking-wider mb-3">
                      Riwayat Log Redeem Hadiah Portal
                    </h3>

                    <div className="max-h-[550px] overflow-y-auto pr-1.5 flex flex-col gap-3">
                      {loyaltyRedeemHistory.length > 0 ? (
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          {loyaltyRedeemHistory.map((h: any) => (
                            <div key={h.id} className="p-4 bg-white border border-[#E5E5DF] rounded-2xl flex flex-col justify-between gap-3 shadow-xs">
                              <div className="flex flex-col gap-1">
                                <div className="flex items-center justify-between">
                                  <span className="text-[9px] font-black uppercase text-[#8C8C70]/80 tracking-wider">Toko Penerima:</span>
                                  <span className="text-[9px] bg-emerald-500/10 text-emerald-800 font-extrabold px-1.5 py-0.5 rounded uppercase">
                                    Berhasil
                                  </span>
                                </div>
                                <span className="text-xs font-black text-[#4A4A3C] uppercase">{h.customerName}</span>
                                <div className="text-xs font-extrabold text-emerald-700 bg-emerald-500/5 px-2.5 py-1.5 rounded-xl border border-emerald-500/15 mt-1 flex items-center gap-1.5 font-sans">
                                  <span>🎁</span>
                                  <span>{h.rewardName}</span>
                                </div>
                              </div>
                              <div className="flex items-center justify-between border-t border-[#E5E5DF]/60 pt-2 text-[10px]">
                                <span className="font-mono bg-[#E5E5DF]/50 text-[#8C8C70] px-1.5 py-0.5 rounded">
                                  {h.date}
                                </span>
                                <span className="text-rose-800 font-extrabold font-mono text-[10px]">
                                  -{h.pointsSpent} Poin
                                </span>
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <div className="text-center p-8 bg-[#E5E5DF]/20 rounded-2xl border border-dashed border-[#E5E5DF] text-xs text-[#8C8C70] flex flex-col items-center justify-center gap-2">
                          <span className="text-lg">📭</span>
                          <p>Belum ada klaim hadiah yang terproses. Pilih panel di sebelah kiri untuk memproses klaim cinderamata pertama Anda.</p>
                        </div>
                      )}
                    </div>
                  </div>

                </div>

              </div>
            </motion.div>
          )}

          {/* TAB 7: KPI SALES CARD REKAP VIEWER MATCHING GRAPHICAL REQUIREMENTS */}
          {activeTab === "kpisales" && (
            <motion.div
              key="kpisales-tab"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              className="space-y-6"
            >
              {/* Header Banner */}
              <FarmerDashboard reports={reports} />
              <div className="bg-gradient-to-r from-rose-600 to-amber-600 rounded-3xl p-6 text-white shadow-lg relative overflow-hidden">
                <div className="absolute top-0 right-0 -mr-20 -mt-10 w-80 h-80 bg-white/10 rounded-full blur-2xl" />
                <div className="absolute bottom-0 right-2 w-32 h-32 bg-amber-400/20 rounded-full blur-xl" />
                
                <h2 className="text-xl md:text-2xl font-serif italic font-extrabold tracking-tight flex items-center gap-2">
                  <TrendingUp className="w-6 h-6 shrink-0 text-white" />
                  KPI SALES AUDIT PORTAL
                </h2>
                <p className="text-xs text-rose-50/90 max-w-2xl mt-2 leading-relaxed">
                  Halaman rekap audit KPI salesman. Anda dapat menarik data langsung dari tab spreadsheet Google Sheets atau menggunakan basis data backup luring. Kartu visual mendeteksi persentase dan merangkum metrics kepatuhan secara otomatis.
                </p>

                {/* Pull Data Controls inside Header */}
                <div className="flex flex-wrap items-center gap-3 mt-6 pt-6 border-t border-white/20">
                  <button
                    onClick={() => handleFetchReportsFromSheets(false)}
                    disabled={isFetchingReports}
                    className="bg-white hover:bg-rose-50 text-rose-900 font-extrabold text-xs uppercase tracking-wider px-5 py-3 rounded-xl transition shadow-sm hover:shadow-md flex items-center gap-2 cursor-pointer disabled:opacity-50"
                  >
                    <RefreshCw className={`w-4 h-4 ${isFetchingReports ? "animate-spin" : ""}`} />
                    {isFetchingReports ? "Menghubungi Google Sheets..." : "Tarik Data Real-time dari Google Sheets"}
                  </button>
                  
                  {lastFetchTime && (
                    <span className="text-[10px] font-mono font-bold bg-white/20 px-3 py-1 bg-rose-500/30 rounded-lg text-white">
                      Sinkron Terakhir: {lastFetchTime}
                    </span>
                  )}
                </div>
              </div>

              {/* Informational Integration Guide Card */}
              <div className="bg-[#FAF9F6] border border-[#E5E5DF] rounded-3xl p-6 shadow-xs relative">
                <div className="absolute top-4 right-4 bg-rose-500/10 text-rose-800 text-[10px] font-black uppercase px-2.5 py-1 rounded-md">
                  Alur Kerja Integrasi
                </div>
                <h3 className="text-sm font-bold text-[#4A4A3C] uppercase tracking-wider flex items-center gap-2 mb-3">
                  <Info className="w-4 h-4 text-rose-600" />
                  Bagaimana Cara Portal Menarik Data Laporan Sales?
                </h3>
                
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 text-[11px] leading-relaxed text-[#5A5A40] mt-4">
                  <div className="bg-[#FAF9F6] border border-[#E5E5DF]/60 p-4 rounded-2xl flex flex-col gap-1">
                    <span className="text-xs font-black text-rose-800 uppercase">1. Pengiriman Aksi getReports</span>
                    <p className="text-[#8C8C70] mt-1">
                      Menekan tombol "Tarik Data" memicu request POST bertransaksi ke URL Apps Script Anda dengan format instruksi <code className="bg-rose-50/80 px-1 py-0.5 rounded font-mono font-bold text-rose-700">action: "getReports"</code>.
                    </p>
                  </div>
                  <div className="bg-[#FAF9F6] border border-[#E5E5DF]/60 p-4 rounded-2xl flex flex-col gap-1">
                    <span className="text-xs font-black text-rose-800 uppercase">2. Pembacaan Baris Spreadsheet</span>
                    <p className="text-[#8C8C70] mt-1">
                      Apps Script diprogram membaca baris aktif pada tab lembar sebar bernilai <code className="bg-gray-100 px-1 py-0.5 rounded font-mono text-gray-800">"Laporan KPI Sales"</code>, mengonversi sel tanggal, TC, CP, EC, SKU, dan mengembalikan array data terstruktur.
                    </p>
                  </div>
                  <div className="bg-[#FAF9F6] border border-[#E5E5DF]/60 p-4 rounded-2xl flex flex-col gap-1">
                    <span className="text-xs font-black text-rose-800 uppercase">3. Parsing & Rangkuman Otomatis</span>
                    <p className="text-[#8C8C70] mt-1">
                      Data ditarik, dikomparasi secara real-time berdasarkan total hari lapor, lalu persentase Call Plan, Effective Call Rate, dan Target Sku dihitung presisi untuk mewarnai indikator KPI.
                    </p>
                  </div>
                </div>
              </div>

              {/* Kartu Rumus KPI */}
              <div className="bg-[#FAF9F6] border border-[#E5E5DF] rounded-3xl p-6 shadow-xs relative">
                <h3 className="text-sm font-bold text-[#4A4A3C] uppercase tracking-wider flex items-center gap-2 mb-3">
                  <TrendingUp className="w-4 h-4 text-emerald-600" />
                  Kartu Rumus KPI
                </h3>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-[11px] leading-relaxed text-[#5A5A40] mt-4">
                  <div className="bg-white border border-[#E5E5DF]/60 p-4 rounded-2xl">
                    <span className="font-extrabold text-emerald-800 uppercase block mb-1">Pencapaian Call Plan</span>
                    <p className="text-[#8C8C70] font-mono">CP / TC x 100%</p>
                  </div>
                  <div className="bg-white border border-[#E5E5DF]/60 p-4 rounded-2xl">
                    <span className="font-extrabold text-rose-800 uppercase block mb-1">Effective Call</span>
                    <p className="text-[#8C8C70] font-mono">EC / CP x 100%</p>
                  </div>
                  <div className="bg-white border border-[#E5E5DF]/60 p-4 rounded-2xl">
                    <span className="font-extrabold text-amber-800 uppercase block mb-1">Pencapaian SKU Fokus</span>
                    <p className="text-[#8C8C70] font-mono">SKU Terjual / Target SKU x 100%</p>
                  </div>
                </div>
              </div>

              {/* Salesman Focus Tabs */}
              <div className="flex flex-wrap bg-[#E5E5DF]/35 p-1 rounded-2xl gap-1 w-full relative mb-6">
                <button
                  type="button"
                  onClick={() => setKpiSalesmanTab("all")}
                  className={`flex-grow sm:flex-grow-0 px-6 py-2.5 rounded-xl text-xs uppercase font-extrabold cursor-pointer transition flex justify-center items-center gap-2 ${
                    kpiSalesmanTab === "all" 
                      ? "bg-[#5A5A40] text-white shadow-xs" 
                      : "text-[#5A5A40] hover:bg-[#E5E5DF]/50 font-bold"
                  }`}
                >
                  <Users className="w-4 h-4" />
                  Semua Salesman
                </button>
                <button
                  type="button"
                  onClick={() => setKpiSalesmanTab("Imam")}
                  className={`flex-grow sm:flex-grow-0 px-6 py-2.5 rounded-xl text-xs uppercase font-extrabold cursor-pointer transition flex justify-center items-center gap-2 ${
                    kpiSalesmanTab === "Imam" 
                      ? "bg-rose-600 text-white shadow-xs" 
                      : "text-rose-900 hover:bg-rose-500/10 font-bold"
                  }`}
                >
                  <User className="w-4 h-4" />
                  NOO Imam
                </button>
              </div>

              {kpiSalesmanTab === "all" && (
                <>
                  <div className="flex bg-[#E5E5DF]/35 p-1 rounded-2xl gap-1 w-full sm:w-auto relative mb-6">
                    <button
                      type="button"
                      onClick={() => setKpiTabMode("dashboard")}
                      className={`flex-1 sm:flex-none px-6 py-2.5 rounded-xl text-xs uppercase font-black cursor-pointer transition flex justify-center items-center gap-2 ${
                        kpiTabMode === "dashboard" 
                          ? "bg-[#5A5A40] text-white shadow-xs" 
                          : "text-[#5A5A40] hover:bg-[#E5E5DF]/50"
                      }`}
                    >
                      <TrendingUp className="w-4 h-4" />
                      Dashboard Visual
                    </button>
                    <button
                      type="button"
                      onClick={() => setKpiTabMode("goals")}
                      className={`flex-1 sm:flex-none px-6 py-2.5 rounded-xl text-xs uppercase font-black cursor-pointer transition flex justify-center items-center gap-2 ${
                        kpiTabMode === "goals" 
                          ? "bg-amber-600 text-white shadow-xs" 
                          : "text-[#5A5A40] hover:bg-[#E5E5DF]/50"
                      }`}
                    >
                      <Target className="w-4 h-4" />
                      Target Bulan
                    </button>
                  </div>

                  {kpiTabMode === "dashboard" && (
                    <>
              {/* Filtering & Source Status Controls */}
              <div className="bg-[#FAF9F6] border border-[#E5E5DF] rounded-3xl p-6 space-y-4">
                <div className="flex flex-wrap items-end justify-between gap-4">
                  {/* Col 1: Salesman and Timeframe Selection */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 w-full lg:w-auto grow">
                    {/* Dropdown Salesman */}
                    <div>
                      <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                        Filter Nama Salesman
                      </label>
                      <select
                        value={kpiSalesFilter}
                        onChange={(e) => setKpiSalesFilter(e.target.value)}
                        className="w-full bg-[#FAF9F6] border border-[#E5E5DF] text-[#4A4A3C] font-extrabold text-xs uppercase px-3 py-2.5 rounded-xl focus:outline-hidden cursor-pointer"
                      >
                        <option value="ALL">Semua Salesman</option>
                        {Array.from(new Set(
                          targetDatasetForKpi.map((r: any) => r.salesmanName.toUpperCase())
                        )).map(name => (
                          <option key={name} value={name}>{name}</option>
                        ))}
                      </select>
                    </div>

                    {/* Timeframe Selector Pills */}
                    <div>
                      <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                        Rentang Analisis KPI
                      </label>
                      <div className="flex border border-[#E5E5DF] rounded-xl overflow-hidden bg-white max-w-sm">
                        {(["all", "daily", "weekly", "monthly"] as const).map((mode) => {
                          const labels = { all: "Semua", daily: "Hari", weekly: "Minggu", monthly: "Bulan" };
                          const isActive = kpiTimeFrame === mode;
                          return (
                            <button
                              key={mode}
                              onClick={() => {
                                setKpiTimeFrame(mode);
                                // Reset sub-filters on change
                                setKpiSelectedDate("ALL");
                                setKpiSelectedWeek("ALL");
                                setKpiSelectedMonth("ALL");
                              }}
                              className={`flex-1 py-2 text-[11px] font-black uppercase text-center transition-all cursor-pointer ${
                                isActive
                                  ? "bg-[#5A5A40] text-[#FAF9F6] font-bold"
                                  : "text-[#8C8C70] hover:bg-[#E5E5DF]/20 hover:text-[#4A4A3C]"
                              }`}
                            >
                              {labels[mode]}
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Dynamic Sub-filter Selector based on timeframe choice */}
                    <div>
                      {kpiTimeFrame === "all" && (
                        <div>
                          <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                            Data Terakumulasi
                          </label>
                          <div className="text-xs font-bold text-[#5A5A40] bg-[#E5E5DF]/20 px-3 py-2.5 rounded-xl border border-[#E5E5DF]/50">
                            Menghitung Semua Hari Lapor
                          </div>
                        </div>
                      )}

                      {kpiTimeFrame === "daily" && (() => {
                        return (
                          <div>
                            <label className="block text-[10px] font-bold text-rose-800 uppercase tracking-wider mb-1">
                              Pilih Tanggal Spesifik
                            </label>
                            <div className="flex gap-1.5">
                              <input
                                type="date"
                                value={kpiSelectedDate === "ALL" ? "" : kpiSelectedDate}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  setKpiSelectedDate(val || "ALL");
                                }}
                                className="grow bg-[#FAF9F6] border border-rose-200 text-rose-900 font-black text-xs px-3 py-2 rounded-xl focus:outline-hidden cursor-pointer h-10"
                              />
                              {kpiSelectedDate !== "ALL" && (
                                <button
                                  type="button"
                                  onClick={() => setKpiSelectedDate("ALL")}
                                  className="px-2.5 bg-rose-100 hover:bg-rose-200 text-rose-800 font-bold text-[10px] uppercase rounded-xl cursor-pointer transition-all flex items-center justify-center border border-rose-200 h-10 shrink-0"
                                >
                                  Semua
                                </button>
                              )}
                            </div>
                          </div>
                        );
                      })()}

                      {kpiTimeFrame === "weekly" && (() => {
                        const targetDataset = targetDatasetForKpi;
                        const MONTH_NAMES_SHORT = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agt", "Sep", "Okt", "Nov", "Des"];
                        const getWeekKeyLocal = (dateStr: string) => {
                          const d = new Date(dateStr);
                          if (isNaN(d.getTime())) return { key: "2026-W01", label: "Minggu 1" };
                          const start = new Date(d.getFullYear(), 0, 1);
                          const diff = d.getTime() - start.getTime() + ((start.getTimezoneOffset() - d.getTimezoneOffset()) * 60000);
                          const oneDay = 1000 * 60 * 60 * 24;
                          const dayOfYear = Math.floor(diff / oneDay) + 1;
                          const weekNo = Math.ceil(dayOfYear / 7);
                          const mLabel = d.toLocaleString("id-ID", { month: "short" }) || MONTH_NAMES_SHORT[d.getMonth()];
                          return {
                            key: `${d.getFullYear()}-W${String(weekNo).padStart(2, '0')}`,
                            label: `Minggu Ke-${weekNo} (${mLabel} ${d.getFullYear()})`
                          };
                        };

                        // Find distinct weeks
                        const weekMap = new Map<string, string>();
                        targetDataset.forEach((r: any) => {
                          if (r.date) {
                            const info = getWeekKeyLocal(r.date);
                            weekMap.set(info.key, info.label);
                          }
                        });
                        const sortedWeeks = Array.from(weekMap.entries()).sort((a, b) => b[0].localeCompare(a[0]));

                        return (
                          <div>
                            <label className="block text-[10px] font-bold text-rose-800 uppercase tracking-wider mb-1">
                              Pilih Minggu Laporan
                            </label>
                            <select
                              value={kpiSelectedWeek}
                              onChange={(e) => setKpiSelectedWeek(e.target.value)}
                              className="w-full bg-[#FAF9F6] border border-rose-200 text-rose-900 font-extrabold text-xs uppercase px-3 py-2.5 rounded-xl focus:outline-hidden cursor-pointer"
                            >
                              <option value="ALL">Semua Minggu ({sortedWeeks.length} Periode)</option>
                              {sortedWeeks.map(([wkKey, wkLabel]) => (
                                <option key={wkKey} value={wkKey}>{wkLabel}</option>
                              ))}
                            </select>
                          </div>
                        );
                      })()}

                      {kpiTimeFrame === "monthly" && (() => {
                        const targetDataset = targetDatasetForKpi;
                        const MONTH_NAMES_LONG = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
                        const getMonthLabelLocal = (dateStr: string) => {
                          const d = new Date(dateStr);
                          if (isNaN(d.getTime())) return "Mei 2026";
                          return `${MONTH_NAMES_LONG[d.getMonth()]} ${d.getFullYear()}`;
                        };

                        const monthMap = new Map<string, string>();
                        targetDataset.forEach((r: any) => {
                          if (r.date) {
                            const key = r.date.substring(0, 7);
                            const label = getMonthLabelLocal(r.date);
                            monthMap.set(key, label);
                          }
                        });
                        const sortedMonths = Array.from(monthMap.entries()).sort((a, b) => b[0].localeCompare(a[0]));

                        return (
                          <div>
                            <label className="block text-[10px] font-bold text-rose-800 uppercase tracking-wider mb-1">
                              Pilih Bulan Laporan
                            </label>
                            <select
                              value={kpiSelectedMonth}
                              onChange={(e) => setKpiSelectedMonth(e.target.value)}
                              className="w-full bg-[#FAF9F6] border border-rose-200 text-rose-900 font-extrabold text-xs uppercase px-3 py-2.5 rounded-xl focus:outline-hidden cursor-pointer"
                            >
                              <option value="ALL">Semua Bulan ({sortedMonths.length} Bulan)</option>
                              {sortedMonths.map(([mKey, mLabel]) => (
                                <option key={mKey} value={mKey}>{mLabel}</option>
                              ))}
                            </select>
                          </div>
                        );
                      })()}
                    </div>
                  </div>

                  {/* Source status indicator */}
                  <div className="shrink-0">
                    <div className="bg-white border border-[#E5E5DF] p-3 rounded-2xl shadow-sm">
                      <span className="block text-[9px] font-extrabold text-[#8C8C70] uppercase mb-1.5 text-left">Sumber Data KPI</span>
                      <div className="flex bg-[#E5E5DF]/35 p-1 rounded-xl gap-1">
                        <button
                          type="button"
                          onClick={() => setKpiDataSource("sheets")}
                          className={`px-3 py-1.5 rounded-lg text-[10px] uppercase font-black cursor-pointer transition flex items-center gap-1.5 ${
                            kpiDataSource === "sheets" 
                              ? "bg-rose-600 text-white shadow-xs" 
                              : "text-[#5A5A40] hover:bg-[#E5E5DF]/50"
                          }`}
                        >
                          <span className={`w-1.5 h-1.5 rounded-full ${fetchedReports.length > 0 ? "bg-emerald-300 animate-pulse" : "bg-orange-400"}`} />
                          Google Sheets (Live)
                        </button>
                        <button
                          type="button"
                          onClick={() => setKpiDataSource("local")}
                          className={`px-3 py-1.5 rounded-lg text-[10px] uppercase font-black cursor-pointer transition flex items-center gap-1.5 ${
                            kpiDataSource === "local" 
                              ? "bg-[#5A5A40] text-white shadow-xs" 
                              : "text-[#5A5A40] hover:bg-[#E5E5DF]/50"
                          }`}
                        >
                          <span className="w-1.5 h-1.5 rounded-full bg-blue-400" />
                          Lokal/Luring
                        </button>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="border-t border-dashed border-[#E5E5DF] pt-3 flex flex-wrap items-center justify-between gap-2">
                  <div className="text-[10px] font-extrabold text-[#8C8C70] uppercase flex items-center gap-1.5">
                    <Info className="w-3.5 h-3.5 text-rose-500" />
                    <span>
                      Mode: {kpiTimeFrame === "all" && "Akumulasi Semua Data"}
                      {kpiTimeFrame === "daily" && "Breakdown Harian (Per Hari Kerja)"}
                      {kpiTimeFrame === "weekly" && "Kompilasi Performance Mingguan"}
                      {kpiTimeFrame === "monthly" && "Kompilasi Performance Bulanan"}
                    </span>
                  </div>
                  <div className="text-[9px] font-mono text-[#8C8C70]">
                    DKR Audit Engine v1.5.0
                  </div>
                </div>
              </div>

              {/* KPI Cards Grid */}
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                {(() => {
                  const targetDataset = targetDatasetForKpi;
                  
                  // Helpers declared locally inside renderer to be resilient:
                  const MONTH_NAMES_LONG = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
                  const MONTH_NAMES_SHORT = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agt", "Sep", "Okt", "Nov", "Des"];

                  const getMonthLabelLocal = (dateStr: string): string => {
                    const d = new Date(dateStr);
                    if (isNaN(d.getTime())) return "Mei 2026";
                    return `${MONTH_NAMES_LONG[d.getMonth()]} ${d.getFullYear()}`;
                  };

                  const getDayOfWeekLabelLocal = (dateStr: string): string => {
                    const d = new Date(dateStr);
                    if (isNaN(d.getTime())) return dateStr;
                    const days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
                    return `${days[d.getDay()]}, ${d.getDate()} ${MONTH_NAMES_LONG[d.getMonth()]} ${d.getFullYear()}`;
                  };

                  const getWeekKeyLocal = (dateStr: string) => {
                    const d = new Date(dateStr);
                    if (isNaN(d.getTime())) {
                      return { key: "2026-W01", label: "Minggu 1" };
                    }
                    const start = new Date(d.getFullYear(), 0, 1);
                    const diff = d.getTime() - start.getTime() + ((start.getTimezoneOffset() - d.getTimezoneOffset()) * 60000);
                    const oneDay = 1000 * 60 * 60 * 24;
                    const dayOfYear = Math.floor(diff / oneDay) + 1;
                    const weekNo = Math.ceil(dayOfYear / 7);
                    const mLabel = MONTH_NAMES_SHORT[d.getMonth()];
                    return {
                      key: `${d.getFullYear()}-W${String(weekNo).padStart(2, '0')}`,
                      label: `Minggu Ke-${weekNo} (${mLabel} ${d.getFullYear()})`
                    };
                  };

                  // We aggregate data into periods:
                  const groups: { [name: string]: {
                    salesmanName: string;
                    dates: Set<string>;
                    periodKey: string;
                    periodLabel: string;
                    tc: number;
                    cp: number;
                    ec: number;
                    skuTotal: number;
                    operationalCost: number;
                    billsReceived: number;
                  }} = {};

                  targetDataset.forEach((rep: any) => {
                    // 1. Filter by Salesman Name first
                    if (kpiSalesFilter !== "ALL" && rep.salesmanName.toUpperCase().trim() !== kpiSalesFilter) {
                      return;
                    }

                    // 2. Identify the active time frame period keys
                    let periodKey = "all";
                    let periodLabel = "Semua Hari Lapor";

                    if (kpiTimeFrame === "daily") {
                      periodKey = rep.date;
                      periodLabel = getDayOfWeekLabelLocal(rep.date);
                    } else if (kpiTimeFrame === "weekly") {
                      const wkInfo = getWeekKeyLocal(rep.date);
                      periodKey = wkInfo.key;
                      periodLabel = wkInfo.label;
                    } else if (kpiTimeFrame === "monthly") {
                      periodKey = rep.date.substring(0, 7);
                      periodLabel = getMonthLabelLocal(rep.date);
                    }

                    // 3. Filter by timeframe dynamic subfilters
                    if (kpiTimeFrame === "daily" && kpiSelectedDate !== "ALL" && periodKey !== kpiSelectedDate) {
                      return;
                    }
                    if (kpiTimeFrame === "weekly" && kpiSelectedWeek !== "ALL" && periodKey !== kpiSelectedWeek) {
                      return;
                    }
                    if (kpiTimeFrame === "monthly" && kpiSelectedMonth !== "ALL" && periodKey !== kpiSelectedMonth) {
                      return;
                    }

                    // Group key is the combo of salesman and period
                    const groupKey = `${rep.salesmanName.toUpperCase().trim()}___${periodKey}`;

                    if (!groups[groupKey]) {
                      groups[groupKey] = {
                        salesmanName: rep.salesmanName,
                        dates: new Set<string>(),
                        periodKey,
                        periodLabel,
                        tc: 0,
                        cp: 0,
                        ec: 0,
                        skuTotal: 0,
                        operationalCost: 0,
                        billsReceived: 0
                      };
                    }

                    groups[groupKey].dates.add(rep.date);
                    groups[groupKey].tc += Number(rep.tc || 0);
                    groups[groupKey].cp += Number(rep.cp || 0);
                    groups[groupKey].ec += Number(rep.ec || 0);
                    groups[groupKey].skuTotal += Number(rep.skuTotal || 0);
                    groups[groupKey].operationalCost += Number(rep.operationalCost || 0);
                    groups[groupKey].billsReceived += Number(rep.billsReceived || 0);
                  });

                  const groupedList = Object.values(groups);

                  if (kpiDataSource === "sheets" && fetchedReports.length === 0) {
                    return (
                      <div className="col-span-full text-center p-10 bg-rose-500/5 border border-dashed border-rose-200 rounded-3xl space-y-4">
                        <div className="text-rose-900 font-extrabold text-sm md:text-base">
                          Belum Ada Data Laporan Real-time di Google Sheets
                        </div>
                        <p className="text-xs text-[#8C8C70] max-w-lg mx-auto leading-relaxed">
                          Sistem mendeteksi lembar sebar Google Sheets Anda masih kosong atau data belum ditarik. Silakan masukkan URL Apps Script di tab <strong className="text-[#5A5A40]">Google Sheets Linker</strong> kemudian tekan tombol <strong className="text-rose-700">Tarik Data Real-time dari Google Sheets</strong> di bagian atas halaman ini untuk sinkronisasi.
                        </p>
                        <div className="flex flex-wrap justify-center gap-3">
                          <button
                            type="button"
                            onClick={() => handleFetchReportsFromSheets(false)}
                            disabled={isFetchingReports}
                            className="bg-rose-600 hover:bg-rose-700 text-white font-extrabold text-[11px] uppercase tracking-wider px-5 py-2.5 rounded-xl transition cursor-pointer disabled:opacity-50"
                          >
                            {isFetchingReports ? "Menghubungi Google Sheets..." : "Tarik Data Sekarang"}
                          </button>
                          <button
                            type="button"
                            onClick={() => setKpiDataSource("local")}
                            className="bg-[#5A5A40] hover:bg-[#4A4A3C] text-white font-extrabold text-[11px] uppercase tracking-wider px-5 py-2.5 rounded-xl transition cursor-pointer"
                          >
                            Gunakan Data Backup Lokal
                          </button>
                        </div>
                      </div>
                    );
                  }

                  if (groupedList.length === 0) {
                    return (
                      <div className="col-span-full text-center p-12 bg-[#FAF9F6] border border-[#E5E5DF] rounded-3xl text-sm text-[#8C8C70] font-bold">
                        Tidak ada data pencapaian KPI yang cocok dengan filter parameter terpilih.
                      </div>
                    );
                  }

                  return groupedList.map((g: any, idx: number) => {
                    const cntDays = g.dates.size || 1;
                    const tcSum = g.tc;
                    const cpSum = g.cp;
                    const ecSum = g.ec;
                    const skuSum = g.skuTotal;
                    
                    const cpPct = tcSum > 0 ? (cpSum / tcSum) * 100 : 0;
                    const ecPct = cpSum > 0 ? (ecSum / cpSum) * 100 : 0;
                    
                    const targetSku = Math.round(112.5 * cntDays);
                    const skuPct = targetSku > 0 ? (skuSum / targetSku) * 100 : 0;

                    // Suitability decision: CP % >= 80% and EC % >= 40%
                    const isLayak = cpPct >= 80 && ecPct >= 40;

                    const smMatch = salesmen.find(s => s.name.toUpperCase().trim() === g.salesmanName.toUpperCase().trim());
                    const salesmanArea = smMatch ? smMatch.area : (g.salesmanName.toUpperCase().trim() === "RINO" ? "CILONGOK" : "DKR SEKTOR");

                    const initialChar = g.salesmanName ? g.salesmanName.charAt(0).toUpperCase() : "S";

                    // Period Title based on active mode
                    let headerPeriodText = "";
                    let headerIcon = <Calendar className="w-3.5 h-3.5 text-rose-500" />;
                    
                    if (kpiTimeFrame === "all") {
                      headerPeriodText = `REKAP TOTAL • ${cntDays} HARI LAPOR`;
                      headerIcon = <Award className="w-3.5 h-3.5 text-rose-500" />;
                    } else if (kpiTimeFrame === "daily") {
                      headerPeriodText = g.periodLabel.toUpperCase();
                      headerIcon = <Clock className="w-3.5 h-3.5 text-rose-500" />;
                    } else if (kpiTimeFrame === "weekly") {
                      headerPeriodText = `${g.periodLabel.toUpperCase()} • ${cntDays} HARI`;
                      headerIcon = <CalendarRange className="w-3.5 h-3.5 text-orange-500" />;
                    } else if (kpiTimeFrame === "monthly") {
                      headerPeriodText = `${g.periodLabel.toUpperCase()} • ${cntDays} HARI`;
                      headerIcon = <CalendarDays className="w-3.5 h-3.5 text-indigo-500" />;
                    }

                    return (
                      <motion.div
                        key={g.salesmanName + "-" + g.periodKey + "-" + idx}
                        initial={{ opacity: 0, y: 10 }}
                        animate={{ opacity: 1, y: 0 }}
                        transition={{ delay: idx * 0.05 }}
                        className="w-full bg-[#FAF9F6] border-t-4 border-rose-600 rounded-3xl p-5 border border-[#E5E5DF] shadow-md hover:shadow-xl transition-all relative flex flex-col justify-between text-left"
                      >
                        <div className="flex items-center justify-between text-[#8C8C70] text-[10px] font-black uppercase tracking-wider mb-2">
                          <span className="flex items-center gap-1 text-[9px] truncate max-w-[70%]">
                            {headerIcon}
                            <span className="truncate">{headerPeriodText}</span>
                          </span>
                          <span className={`px-2.5 py-0.5 rounded-full font-black text-[9px] shrink-0 ${
                            isLayak 
                              ? "bg-emerald-100 text-emerald-800 border border-emerald-200" 
                              : "bg-rose-100 text-rose-800 border border-rose-200"
                          }`}>
                            {isLayak ? "LAYAK" : "TIDAK LAYAK"}
                          </span>
                        </div>

                        <div className="flex items-start justify-between mt-1 mb-4">
                          <div>
                            <h4 className="text-xl font-serif italic font-black tracking-tight text-[#4A4A3C]">
                              {g.salesmanName}
                            </h4>
                          </div>
                          <div className="bg-[#5A5A40]/10 text-[#5A5A40] text-[9.5px] font-extrabold px-2 py-0.5 rounded-lg flex items-center gap-1 uppercase shrink-0">
                            <MapPin className="w-3 h-3 text-[#5A5A40]" />
                            {salesmanArea}
                          </div>
                        </div>

                        <div className="grid grid-cols-4 gap-2 mb-4 text-center">
                          <div className="bg-rose-600 rounded-2xl p-1.5 text-white">
                            <span className="block text-[8px] font-bold opacity-80 uppercase tracking-widest leading-none">TC</span>
                            <span className="text-sm font-extrabold font-mono leading-tight">{tcSum}</span>
                          </div>
                          <div className="bg-[#E5E5DF]/40 text-[#4A4A3C] rounded-2xl p-1.5 border border-[#E5E5DF]">
                            <span className="block text-[8px] font-extrabold text-[#8C8C70] uppercase tracking-widest leading-none">CP</span>
                            <span className="text-sm font-extrabold font-mono leading-tight">{cpSum}</span>
                          </div>
                          <div className="bg-rose-600 rounded-2xl p-1.5 text-white">
                            <span className="block text-[8px] font-bold opacity-80 uppercase tracking-widest leading-none">EC</span>
                            <span className="text-sm font-extrabold font-mono leading-tight">{ecSum}</span>
                          </div>
                          <div className="bg-[#E5E5DF]/40 text-[#4A4A3C] rounded-2xl p-1.5 border border-[#E5E5DF]">
                            <span className="block text-[8px] font-extrabold text-[#8C8C70] uppercase tracking-widest leading-none">SKU</span>
                            <span className="text-sm font-extrabold font-mono leading-tight">{skuSum}</span>
                          </div>
                        </div>

                        <div className="space-y-2.5 mt-2 mb-4">
                          <div>
                            <div className="flex justify-between text-[9px] font-black uppercase text-[#8C8C70] mb-0.5">
                              <span>Pencapaian Call Plan (CP/TC)</span>
                              <span className="text-emerald-700 font-mono font-bold">{cpPct.toFixed(1)}%</span>
                            </div>
                            <div className="w-full h-2 bg-[#E5E5DF]/40 rounded-full overflow-hidden">
                              <div
                                style={{ width: `${Math.min(cpPct, 100)}%` }}
                                className="h-full bg-emerald-500 transition-all duration-550"
                              />
                            </div>
                          </div>

                          <div>
                            <div className="flex justify-between text-[9px] font-black uppercase text-[#8C8C70] mb-0.5">
                              <span>Tingkat Effective Call (EC/CP)</span>
                              <span className="text-rose-700 font-mono font-bold">{ecPct.toFixed(1)}%</span>
                            </div>
                            <div className="w-full h-2 bg-[#E5E5DF]/40 rounded-full overflow-hidden">
                              <div
                                style={{ width: `${Math.min(ecPct, 100)}%` }}
                                className="h-full bg-rose-500 transition-all duration-550"
                              />
                            </div>
                          </div>

                          <div>
                            <div className="flex justify-between text-[9px] font-black uppercase text-[#8C8C70] mb-0.5">
                              <span>Pencapaian SKU Fokus ({skuSum}/{targetSku} Sku)</span>
                              <span className="text-amber-700 font-mono font-bold">{skuPct.toFixed(1)}%</span>
                            </div>
                            <div className="w-full h-2 bg-[#E5E5DF]/40 rounded-full overflow-hidden">
                              <div
                                style={{ width: `${Math.min(skuPct, 100)}%` }}
                                className="h-full bg-amber-500 transition-all duration-550"
                              />
                            </div>
                          </div>
                        </div>

                        <div className="border border-dashed border-[#E5E5DF] rounded-2xl p-3 bg-[#E5E5DF]/10 mb-4 flex justify-between gap-2 text-left">
                          <div>
                            <span className="text-[8px] font-extrabold text-[#8C8C70] flex items-center gap-1 uppercase tracking-wide">
                              <Boxes className="w-3 h-3 text-[#8C8C70]" />
                              Fokus Produk
                            </span>
                            <span className="text-[10px] font-black text-[#5A5A40] block mt-0.5 leading-none">
                              CB-YPP, TJ-YPP-PU
                            </span>
                          </div>
                          <div className="text-right border-l border-[#E5E5DF] pl-3 shrink-0">
                            <span className="text-[8px] font-extrabold text-[#8C8C70] flex items-center justify-end gap-1 uppercase tracking-wide">
                              <DollarSign className="w-3 h-3 text-amber-600" />
                              Biaya Operasional
                            </span>
                            <span className="text-[10.5px] font-mono font-black text-amber-700 block mt-0.5 leading-none">
                              Rp {((typeof g.operationalCost === 'number' && !isNaN(g.operationalCost)) ? g.operationalCost : 0).toLocaleString("id-ID")}
                            </span>
                          </div>
                        </div>

                        <div className="border-t border-dotted border-[#E5E5DF] pt-3 mt-1 flex items-center justify-between text-left">
                          <div className="flex items-center gap-2.5">
                            <div className="w-7 h-7 rounded-full bg-[#5A5A40] text-white flex items-center justify-center font-serif text-[11px] font-extrabold italic">
                              {initialChar}
                            </div>
                            <div>
                              <span className="text-[8px] font-extrabold text-[#8C8C70] block uppercase tracking-wide leading-none">Total Tagihan Termasuk Transfer & Giro</span>
                              <span className="text-[11.5px] font-mono font-black text-indigo-800 mt-0.5 block leading-none">
                                Rp {(g.billsReceived + (g.billsTransfer || 0) + (g.billsGiro || 0)).toLocaleString("id-ID")}
                              </span>
                            </div>
                          </div>
                          <div>
                            <span className="bg-[#E5E5DF]/60 text-[#4A4A3C] text-[8.5px] font-black rounded-md px-1.5 py-0.5 font-mono">
                              {cntDays} HARI
                            </span>
                          </div>
                        </div>
                      </motion.div>
                    );
                  });
                })()}
              </div>
              </>
              )}

              {kpiTabMode === "goals" && (
                <div className="bg-[#FAF9F6] border border-[#E5E5DF] rounded-3xl p-6 shadow-xs animate-in fade-in slide-in-from-bottom-4 space-y-6">
                  <div className="flex flex-wrap items-center justify-between gap-4 border-b border-[#E5E5DF] pb-4">
                    <h3 className="text-base font-black text-[#5A5A40] uppercase tracking-wider flex items-center gap-2">
                      <Target className="w-5 h-5 text-amber-600" />
                      Manajemen Target KPI Bulanan
                    </h3>
                    <div className="w-full sm:w-64">
                      <label className="block text-[10px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">Pilih Bulan Target</label>
                      <input 
                        type="month" 
                        value={goalSelectedMonth}
                        onChange={(e) => setGoalSelectedMonth(e.target.value)}
                        className="w-full bg-white border border-[#E5E5DF] px-3 py-2 text-sm font-bold text-[#4A4A3C] rounded-xl focus:outline-hidden focus:border-amber-400"
                      />
                    </div>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="border-b-2 border-[#E5E5DF]">
                          <th className="py-3 px-2 text-[10px] font-black text-[#5A5A40] uppercase">Nama Salesman</th>
                          <th className="py-3 px-2 text-[10px] font-black text-[#5A5A40] uppercase">Target TC</th>
                          <th className="py-3 px-2 text-[10px] font-black text-[#5A5A40] uppercase">Target CP</th>
                          <th className="py-3 px-2 text-[10px] font-black text-[#5A5A40] uppercase">Target EC</th>
                          <th className="py-3 px-2 text-[10px] font-black text-[#5A5A40] uppercase">Target SKU</th>
                        </tr>
                      </thead>
                      <tbody>
                        {salesmen.filter(s => s.isActive).map(s => {
                          const goal = salesmanGoals.find(g => g.salesmanId === s.id && g.monthString === goalSelectedMonth) || {
                            tcTarget: 0,
                            cpTarget: 0,
                            ecTarget: 0,
                            skuTarget: 0
                          };

                          let actualTc = 0;
                          let actualCp = 0;
                          let actualEc = 0;
                          let actualSku = 0;

                          targetDatasetForKpi.forEach((r: any) => {
                            if (r.salesmanName.toUpperCase().trim() === s.name.toUpperCase().trim()) {
                              if (r.date && r.date.startsWith(goalSelectedMonth)) {
                                actualTc += Number(r.tc || 0);
                                actualCp += Number(r.cp || 0);
                                actualEc += Number(r.ec || 0);
                                actualSku += Number(r.skuTotal || 0);
                              }
                            }
                          });

                          const tcPct = goal.tcTarget > 0 ? (actualTc / goal.tcTarget) * 100 : 0;
                          const cpPct = goal.cpTarget > 0 ? (actualCp / goal.cpTarget) * 100 : 0;
                          const ecPct = goal.ecTarget > 0 ? (actualEc / goal.ecTarget) * 100 : 0;
                          const skuPct = goal.skuTarget > 0 ? (actualSku / goal.skuTarget) * 100 : 0;
                          
                          return (
                            <tr key={s.id} className="border-b border-[#E5E5DF]/50 hover:bg-[#E5E5DF]/20 transition-colors">
                              <td className="py-3 px-2 font-bold text-[#4A4A3C] text-sm uppercase align-top pt-4">{s.name}</td>
                              
                              <td className="py-3 px-2 align-top">
                                <div className="space-y-2">
                                  <input 
                                    type="number" min="0" 
                                    className="w-full bg-white border border-[#E5E5DF] rounded-lg px-2 py-1.5 text-xs font-mono font-bold text-center focus:outline-hidden focus:border-amber-400"
                                    value={goal.tcTarget}
                                    placeholder="Target"
                                    onChange={(e) => handleUpdateGoal(s.id, goalSelectedMonth, "tcTarget", parseInt(e.target.value) || 0)}
                                    onFocus={(e) => e.target.value === "0" && e.target.select()}
                                  />
                                  <div className="flex justify-between text-[9px] font-black uppercase text-[#8C8C70]">
                                    <span>Aktual: {actualTc}</span>
                                    <span className={tcPct >= 100 ? "text-emerald-600" : tcPct > 0 ? "text-amber-600" : ""}>{tcPct.toFixed(0)}%</span>
                                  </div>
                                  <div className="w-full h-1.5 bg-[#E5E5DF]/60 rounded-full overflow-hidden">
                                    <div style={{ width: `${Math.min(tcPct, 100)}%` }} className={`h-full ${tcPct >= 100 ? "bg-emerald-500" : "bg-amber-500"}`} />
                                  </div>
                                </div>
                              </td>

                              <td className="py-3 px-2 align-top">
                                <div className="space-y-2">
                                  <input 
                                    type="number" min="0" 
                                    className="w-full bg-white border border-[#E5E5DF] rounded-lg px-2 py-1.5 text-xs font-mono font-bold text-center focus:outline-hidden focus:border-amber-400"
                                    value={goal.cpTarget}
                                    placeholder="Target"
                                    onChange={(e) => handleUpdateGoal(s.id, goalSelectedMonth, "cpTarget", parseInt(e.target.value) || 0)}
                                    onFocus={(e) => e.target.value === "0" && e.target.select()}
                                  />
                                  <div className="flex justify-between text-[9px] font-black uppercase text-[#8C8C70]">
                                    <span>Aktual: {actualCp}</span>
                                    <span className={cpPct >= 100 ? "text-emerald-600" : cpPct > 0 ? "text-amber-600" : ""}>{cpPct.toFixed(0)}%</span>
                                  </div>
                                  <div className="w-full h-1.5 bg-[#E5E5DF]/60 rounded-full overflow-hidden">
                                    <div style={{ width: `${Math.min(cpPct, 100)}%` }} className={`h-full ${cpPct >= 100 ? "bg-emerald-500" : "bg-amber-500"}`} />
                                  </div>
                                </div>
                              </td>

                              <td className="py-3 px-2 align-top">
                                <div className="space-y-2">
                                  <input 
                                    type="number" min="0" 
                                    className="w-full bg-white border border-[#E5E5DF] rounded-lg px-2 py-1.5 text-xs font-mono font-bold text-center focus:outline-hidden focus:border-amber-400"
                                    value={goal.ecTarget}
                                    placeholder="Target"
                                    onChange={(e) => handleUpdateGoal(s.id, goalSelectedMonth, "ecTarget", parseInt(e.target.value) || 0)}
                                    onFocus={(e) => e.target.value === "0" && e.target.select()}
                                  />
                                  <div className="flex justify-between text-[9px] font-black uppercase text-[#8C8C70]">
                                    <span>Aktual: {actualEc}</span>
                                    <span className={ecPct >= 100 ? "text-emerald-600" : ecPct > 0 ? "text-amber-600" : ""}>{ecPct.toFixed(0)}%</span>
                                  </div>
                                  <div className="w-full h-1.5 bg-[#E5E5DF]/60 rounded-full overflow-hidden">
                                    <div style={{ width: `${Math.min(ecPct, 100)}%` }} className={`h-full ${ecPct >= 100 ? "bg-emerald-500" : "bg-amber-500"}`} />
                                  </div>
                                </div>
                              </td>

                              <td className="py-3 px-2 align-top">
                                <div className="space-y-2">
                                  <input 
                                    type="number" min="0" 
                                    className="w-full bg-white border border-[#E5E5DF] rounded-lg px-2 py-1.5 text-xs font-mono font-bold text-center focus:outline-hidden focus:border-amber-400"
                                    value={goal.skuTarget}
                                    placeholder="Target"
                                    onChange={(e) => handleUpdateGoal(s.id, goalSelectedMonth, "skuTarget", parseInt(e.target.value) || 0)}
                                    onFocus={(e) => e.target.value === "0" && e.target.select()}
                                  />
                                  <div className="flex justify-between text-[9px] font-black uppercase text-[#8C8C70]">
                                    <span>Aktual: {actualSku}</span>
                                    <span className={skuPct >= 100 ? "text-emerald-600" : skuPct > 0 ? "text-amber-600" : ""}>{skuPct.toFixed(0)}%</span>
                                  </div>
                                  <div className="w-full h-1.5 bg-[#E5E5DF]/60 rounded-full overflow-hidden">
                                    <div style={{ width: `${Math.min(skuPct, 100)}%` }} className={`h-full ${skuPct >= 100 ? "bg-emerald-500" : "bg-amber-500"}`} />
                                  </div>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                        {salesmen.filter(s => s.isActive).length === 0 && (
                          <tr><td colSpan={5} className="py-4 text-center text-xs text-[#8C8C70]">Belum ada salesman aktif didatabase.</td></tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

                </>
              )}

              {kpiSalesmanTab !== "all" && (() => {
                const activeSalesmanName = kpiSalesmanTab === "Aris" ? "Aris" : "Imam";
                const foundSalesman = salesmen.find(s => s.name.toLowerCase() === activeSalesmanName.toLowerCase()) || {
                  id: kpiSalesmanTab === "Aris" ? "s-7" : "s-8",
                  name: activeSalesmanName,
                  area: kpiSalesmanTab === "Aris" ? "Semarang" : "Demak",
                  phone: kpiSalesmanTab === "Aris" ? "084567890123" : "08567890124",
                  isActive: true
                };

                // filter reports by that salesman
                const salesmanReports = targetDatasetForKpi.filter(
                  r => r.salesmanName.toUpperCase().trim() === activeSalesmanName.toUpperCase()
                );

                const sTc = salesmanReports.reduce((sum, r) => sum + Number(r.tc || 0), 0);
                const sCp = salesmanReports.reduce((sum, r) => sum + Number(r.cp || 0), 0);
                const sEc = salesmanReports.reduce((sum, r) => sum + Number(r.ec || 0), 0);
                const sSku = salesmanReports.reduce((sum, r) => sum + Number(r.skuTotal || 0), 0);
                const sCost = salesmanReports.reduce((sum, r) => sum + Number(r.operationalCost || 0), 0);
                const sBill = salesmanReports.reduce((sum, r) => sum + Number(r.billsReceived || 0) + Number(r.billsTransfer || 0) + Number(r.billsGiro || 0), 0);

                const sCpPct = sTc > 0 ? (sCp / sTc) * 100 : 0;
                const sEcPct = sCp > 0 ? (sEc / sCp) * 105 : 0; // standard scaling for aesthetics
                const sCntDays = salesmanReports.length || 1;
                const sTargetSku = Math.round(112.5 * sCntDays);
                const sSkuPct = sTargetSku > 0 ? (sSku / sTargetSku) * 100 : 0;

                const isLayak = sCpPct >= 80 && sEcPct >= 40;

                // Tracking NOO
                const sNooLogs = nooRecords.filter(
                  n => n.salesmanName.toLowerCase().trim() === activeSalesmanName.toLowerCase()
                );
                const totalWarung = sNooLogs.reduce((sum, n) => sum + n.warung, 0);
                const totalStore = sNooLogs.reduce((sum, n) => sum + n.store, 0);
                const totalKiosk = sNooLogs.reduce((sum, n) => sum + n.kiosk, 0);
                const totalWholesaler = sNooLogs.reduce((sum, n) => sum + n.wholesaler, 0);
                const totalNooCount = totalWarung + totalStore + totalKiosk + totalWholesaler;

                const handleAddNoo = (e: React.FormEvent) => {
                  e.preventDefault();
                  const newRecord: NooRecord = {
                    id: `noo-${Date.now()}`,
                    salesmanId: foundSalesman.id,
                    salesmanName: foundSalesman.name,
                    date: newNooDate,
                    warung: newNooWarung,
                    store: newNooStore,
                    kiosk: newNooKiosk,
                    wholesaler: newNooWholesaler
                  };
                  const updatedRecords = [newRecord, ...nooRecords];
                  setNooRecords(updatedRecords);
                  showToast(`Berhasil menyimpan log NOO harian untuk ${foundSalesman.name}!`, "success");
                  // reset input counts
                  setNewNooWarung(0);
                  setNewNooStore(0);
                  setNewNooKiosk(0);
                  setNewNooWholesaler(0);

                  if (sheetsScriptUrl) {
                    handleSyncNooToSheets(updatedRecords);
                  }
                };

                const handleDeleteNoo = (id: string) => {
                  const updatedRecords = nooRecords.filter(n => n.id !== id);
                  setNooRecords(updatedRecords);
                  showToast("Record NOO berhasil dihapus", "info");

                  if (sheetsScriptUrl) {
                    handleSyncNooToSheets(updatedRecords, true);
                  }
                };

                return (
                  <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-300">
                    
                    {/* PROFILE HERO CARD */}
                    <div className="bg-[#FAF9F6] border border-[#E5E5DF] rounded-3xl p-6 shadow-sm flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
                      <div className="flex items-center gap-4">
                        <div className="w-14 h-14 rounded-full bg-rose-600 text-white flex items-center justify-center font-serif text-2xl font-black italic shadow-md shrink-0">
                          {activeSalesmanName.charAt(0).toUpperCase()}
                        </div>
                        <div>
                          <div className="flex items-center gap-2">
                            <h3 className="text-xl font-serif font-black text-[#4A4A3C] italic">{activeSalesmanName}</h3>
                            <span className="bg-emerald-100 text-emerald-800 text-[9px] font-black px-2 py-0.5 rounded-md uppercase tracking-wide">
                              AKTIF
                            </span>
                          </div>
                          <p className="text-xs text-[#8C8C70] font-bold flex items-center gap-2 mt-1">
                            <MapPin className="w-3.5 h-3.5 text-rose-500" /> Area: {foundSalesman.area || "Semarang"}
                            <span className="text-[#E5E5DF]">|</span> Telp: {foundSalesman.phone || "-"}
                          </p>
                        </div>
                      </div>

                      <div className="flex flex-wrap gap-2">
                        <span className={`px-4 py-2 rounded-2xl text-xs font-black uppercase text-center tracking-wider shrink-0 ${
                          isLayak 
                            ? "bg-emerald-100 text-emerald-800 border border-emerald-200" 
                            : "bg-rose-100 text-rose-800 border border-rose-200"
                        }`}>
                          KAPASITAS: {isLayak ? "LAYAK KPI" : "TIDAK LAYAK"}
                        </span>
                        <div className="bg-[#5A5A40]/10 text-[#5A5A40] text-xs font-black px-4 py-2 rounded-2xl flex items-center gap-2">
                          <Calendar className="w-4 h-4" />
                          {sCntDays} Hari Kerja Audit
                        </div>
                      </div>
                    </div>

                    {/* SPREADSHEET NOO SYNC BAR */}
                    {sheetsScriptUrl ? (
                      <div className="bg-[#FAF9F6] border border-[#E5E5DF] rounded-3xl p-5 flex flex-col sm:flex-row items-center justify-between gap-4 shadow-xs">
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-xl bg-emerald-50 text-emerald-700 flex items-center justify-center font-bold shrink-0">
                            <span className="text-lg">📊</span>
                          </div>
                          <div>
                            <h4 className="text-xs font-black text-emerald-800 uppercase tracking-wider">Koneksi Spreadsheet Aktif</h4>
                            <p className="text-[10px] text-[#8C8C70] font-bold mt-1">
                              Auto-sync aktif: Setiap penambahan atau penghapusan log NOO langsung terkirim ke Google Sheets!
                            </p>
                          </div>
                        </div>
                        <div className="flex gap-2 w-full sm:w-auto">
                          <button
                            type="button"
                            onClick={() => handleFetchNooFromSheets()}
                            disabled={isFetchingNoo}
                            className="flex-1 sm:flex-none bg-white hover:bg-gray-50 border border-[#E5E5DF] text-[#5A5A40] text-xs font-black px-4 py-2.5 rounded-2xl flex items-center justify-center gap-2 transition cursor-pointer disabled:opacity-50"
                          >
                            <RefreshCw className={`w-3.5 h-3.5 ${isFetchingNoo ? "animate-spin" : ""}`} />
                            Tarik Data Sheets
                          </button>
                          <button
                            type="button"
                            onClick={() => handleSyncNooToSheets()}
                            disabled={isSyncingNoo}
                            className="flex-1 sm:flex-none bg-rose-600 hover:bg-rose-700 text-white text-xs font-black px-4 py-2.5 rounded-2xl flex items-center justify-center gap-2 transition cursor-pointer disabled:opacity-50 shadow-sm"
                          >
                            <Send className={`w-3.5 h-3.5 ${isSyncingNoo ? "animate-spin" : ""}`} />
                            Kirim Manual
                          </button>
                        </div>
                      </div>
                    ) : (
                      <div className="bg-amber-50 border border-amber-200 rounded-3xl p-5 flex items-center gap-3">
                        <span className="text-xl">⚠️</span>
                        <div>
                          <h4 className="text-xs font-black text-amber-800 uppercase">Spreadsheet Belum Terhubung</h4>
                          <p className="text-[10px] text-amber-900/80">Silakan hubungkan URL Google Apps Script Anda di tab <strong className="text-rose-700">Google Sheets Linker</strong> untuk sinkronisasi otomatis harian NOO.</p>
                        </div>
                      </div>
                    )}

                    <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                      {/* COLUMN 1 & 2: KPI RESULTS & METRICS */}
                      <div className="lg:col-span-2 space-y-6">
                        
                        {/* THE METRICS SUMMARY GRID */}
                        <div className="bg-[#FAF9F6] border border-[#E5E5DF] rounded-3xl p-6 shadow-xs space-y-4">
                          <h4 className="text-xs font-black text-[#5A5A40] uppercase tracking-wider flex items-center gap-1.5 border-b border-[#E5E5DF] pb-3">
                            <TrendingUp className="w-4 h-4 text-rose-600" />
                            REKAP METRIC KPI - {activeSalesmanName.toUpperCase()}
                          </h4>
                          
                          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 text-center">
                            <div className="bg-white border border-[#E5E5DF]/50 rounded-2xl p-3">
                              <span className="block text-[9px] font-bold text-[#8C8C70] uppercase">Total Call (TC)</span>
                              <span className="text-xl font-bold font-mono text-[#4A4A3C] block mt-1">{sTc}</span>
                            </div>
                            <div className="bg-white border border-[#E5E5DF]/50 rounded-2xl p-3">
                              <span className="block text-[9px] font-bold text-[#8C8C70] uppercase">Call Plan (CP)</span>
                              <span className="text-xl font-bold font-mono text-[#4A4A3C] block mt-1">{sCp}</span>
                            </div>
                            <div className="bg-white border border-[#E5E5DF]/50 rounded-2xl p-3">
                              <span className="block text-[9px] font-bold text-[#8C8C70] uppercase">Effective (EC)</span>
                              <span className="text-xl font-bold font-mono text-[#4A4A3C] block mt-1">{sEc}</span>
                            </div>
                            <div className="bg-white border border-[#E5E5DF]/50 rounded-2xl p-3">
                              <span className="block text-[9px] font-bold text-[#8C8C70] uppercase">SKU Terjual</span>
                              <span className="text-xl font-bold font-mono text-[#4A4A3C] block mt-1">{sSku}</span>
                            </div>
                          </div>

                          <div className="space-y-4 pt-2">
                            {/* CP/TC PROGRESS */}
                            <div>
                              <div className="flex justify-between text-xs font-bold text-[#5A5A40] mb-1">
                                <span>PENCAPAIAN CALL PLAN (TC terhadap CP)</span>
                                <span className={`font-mono text-xs font-black ${sCpPct >= 80 ? "text-emerald-700" : "text-rose-700"}`}>
                                  {sCpPct.toFixed(1)}% {sCpPct >= 80 ? "✓ (Target >= 80%)" : "✗ (Kurang)"}
                                </span>
                              </div>
                              <div className="w-full h-3 bg-[#E5E5DF]/50 rounded-full overflow-hidden">
                                <div style={{ width: `${Math.min(sCpPct, 100)}%` }} className={`h-full transition-all duration-300 ${sCpPct >= 80 ? "bg-emerald-500" : "bg-rose-500"}`} />
                              </div>
                            </div>

                            {/* EC/CP PROGRESS */}
                            <div>
                              <div className="flex justify-between text-xs font-bold text-[#5A5A40] mb-1">
                                <span>EFFECTIVE CALL RATE (EC terhadap CP)</span>
                                <span className={`font-mono text-xs font-black ${sEcPct >= 40 ? "text-emerald-700" : "text-rose-700"}`}>
                                  {sEcPct.toFixed(1)}% {sEcPct >= 40 ? "✓ (Target >= 40%)" : "✗ (Kurang)"}
                                </span>
                              </div>
                              <div className="w-full h-3 bg-[#E5E5DF]/50 rounded-full overflow-hidden">
                                <div style={{ width: `${Math.min(sEcPct, 100)}%` }} className={`h-full transition-all duration-300 ${sEcPct >= 40 ? "bg-emerald-500" : "bg-rose-500"}`} />
                              </div>
                            </div>

                            {/* SKU PROGRESS */}
                            <div>
                              <div className="flex justify-between text-xs font-bold text-[#5A5A40] mb-1">
                                <span>SKU AKUMULASI PENCAPAIAN ({sSku} Sku/Target {sTargetSku} Sku)</span>
                                <span className="font-mono text-xs font-black text-amber-700">{sSkuPct.toFixed(1)}%</span>
                              </div>
                              <div className="w-full h-3 bg-[#E5E5DF]/50 rounded-full overflow-hidden">
                                <div style={{ width: `${Math.min(sSkuPct, 100)}%` }} className="h-full bg-amber-500 transition-all duration-300" />
                              </div>
                            </div>
                          </div>

                          <div className="grid grid-cols-2 gap-4 pt-4 border-t border-dashed border-[#E5E5DF] text-xs">
                            <div className="bg-[#E5E5DF]/20 p-3 rounded-2xl">
                              <span className="block text-[10px] text-[#8C8C70] font-bold uppercase">BIAYA OPERASIONAL TOTAL:</span>
                              <span className="text-[#4A4A3C] font-mono font-black text-sm block mt-0.5">Rp {sCost.toLocaleString("id-ID")}</span>
                            </div>
                            <div className="bg-[#E5E5DF]/20 p-3 rounded-2xl">
                              <span className="block text-[10px] text-[#8C8C70] font-bold uppercase">TOTAL PENAGIHAN:</span>
                              <span className="text-indigo-800 font-mono font-black text-sm block mt-0.5">Rp {sBill.toLocaleString("id-ID")}</span>
                            </div>
                          </div>
                        </div>

                        {/* NOO TRACKING SECTION */}
                        <div className="bg-[#FAF9F6] border border-[#E5E5DF] rounded-3xl p-6 shadow-sm space-y-4">
                          <div className="flex justify-between items-center border-b border-[#E5E5DF] pb-3">
                            <h4 className="text-xs font-black text-rose-700 uppercase tracking-wider flex items-center gap-1.5">
                              <Crown className="w-4 h-4 text-rose-600 animate-bounce" />
                              NEW OUTLET OPENING (NOO) HARIAN
                            </h4>
                            <span className="text-[10px] font-black bg-rose-100 text-rose-800 px-3 py-1 rounded-full uppercase">
                              Cumulative: {totalNooCount} Outlet
                            </span>
                          </div>

                          {/* Cumulative NOO Grid */}
                          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                            <div className="bg-amber-500/10 border border-amber-500/20 rounded-2xl p-3 text-center">
                              <span className="block text-[9px] font-bold text-amber-800 uppercase">Warung / Toko Kelontong</span>
                              <span className="text-2xl font-black font-mono text-amber-900 block mt-1">{totalWarung}</span>
                            </div>
                            <div className="bg-rose-500/10 border border-rose-500/20 rounded-2xl p-3 text-center">
                              <span className="block text-[9px] font-bold text-rose-800 uppercase">Store / Toko Modern</span>
                              <span className="text-2xl font-black font-mono text-rose-900 block mt-1">{totalStore}</span>
                            </div>
                            <div className="bg-indigo-500/10 border border-indigo-500/20 rounded-2xl p-3 text-center">
                              <span className="block text-[9px] font-bold text-indigo-800 uppercase">Kios / Outlet Atap</span>
                              <span className="text-2xl font-black font-mono text-[#312E81] block mt-1">{totalKiosk}</span>
                            </div>
                            <div className="bg-emerald-500/10 border border-emerald-500/20 rounded-2xl p-3 text-center">
                              <span className="block text-[9px] font-bold text-emerald-800 uppercase">Grosir / Wholesaler</span>
                              <span className="text-2xl font-black font-mono text-[#064E3B] block mt-1">{totalWholesaler}</span>
                            </div>
                          </div>

                          {/* Historical records table */}
                          <div className="space-y-4">
                            <span className="text-[10px] font-extrabold text-[#8C8C70] uppercase block">Riwayat Log NOO {foundSalesman.name}:</span>
                            {sNooLogs.length === 0 ? (
                              <div className="text-center py-6 border border-dashed border-[#E5E5DF] rounded-2xl text-xs text-[#8C8C70] font-bold">
                                Tidak ada log harian NOO. Silakan tambahkan pada form di samping kanan.
                              </div>
                            ) : (
                              <div className="overflow-x-auto border border-[#E5E5DF] rounded-2xl bg-white">
                                <table className="w-full text-left border-collapse text-xs">
                                  <thead>
                                    <tr className="bg-[#E5E5DF]/20 border-b border-[#E5E5DF] text-[9px] font-black uppercase text-[#8C8C70]">
                                      <th className="p-2.5">Tanggal</th>
                                      <th className="p-2.5 text-center">Warung</th>
                                      <th className="p-2.5 text-center">Toko/Store</th>
                                      <th className="p-2.5 text-center">Kios</th>
                                      <th className="p-2.5 text-center">Grosir</th>
                                      <th className="p-2.5 text-center">Total harian</th>
                                      <th className="p-2.5 text-center">Aksi</th>
                                    </tr>
                                  </thead>
                                  <tbody className="divide-y divide-[#E5E5DF]/30">
                                    {sNooLogs.map(log => {
                                      const rowSum = log.warung + log.store + log.kiosk + log.wholesaler;
                                      return (
                                        <tr key={log.id} className="hover:bg-[#E5E5DF]/10 transition-colors bg-white text-[#4A4A3C]">
                                          <td className="p-2.5 font-bold font-mono">{log.date}</td>
                                          <td className="p-2.5 text-center font-semibold text-amber-700">{log.warung}</td>
                                          <td className="p-2.5 text-center font-semibold text-rose-700">{log.store}</td>
                                          <td className="p-2.5 text-center font-semibold text-indigo-700">{log.kiosk}</td>
                                          <td className="p-2.5 text-center font-semibold text-emerald-700">{log.wholesaler}</td>
                                          <td className="p-2.5 text-center font-black bg-rose-50 text-rose-900">{rowSum}</td>
                                          <td className="p-2.5 text-center">
                                            <button
                                              type="button"
                                              onClick={() => handleDeleteNoo(log.id)}
                                              className="p-1 hover:bg-rose-100 rounded-lg text-rose-600 transition cursor-pointer"
                                              title="Hapus baris harian ini"
                                            >
                                              <Trash2 className="w-3.5 h-3.5" />
                                            </button>
                                          </td>
                                        </tr>
                                      );
                                    })}
                                  </tbody>
                                </table>
                              </div>
                            )}
                          </div>
                        </div>

                      </div>

                      {/* COLUMN 3: NOO INPUT CONTROL PANEL */}
                      <div>
                        <div className="bg-[#FAF9F6] border border-[#E5E5DF] rounded-3xl p-6 shadow-sm space-y-4 stick top-6">
                          <h4 className="text-sm font-bold text-[#4A4A3C] uppercase tracking-wider flex items-center gap-1.5 border-b border-[#E5E5DF] pb-3">
                            <Plus className="w-4 h-4 text-rose-600 animate-pulse" />
                            LOG NOO HARIAN (+ HARIAN)
                          </h4>
                          <p className="text-[10px] text-[#8C8C70] leading-relaxed uppercase font-bold">
                            Catat pembukaan outlet baru salesman <strong className="text-rose-800">{activeSalesmanName}</strong> secara harian ke dalam database.
                          </p>

                          <form onSubmit={handleAddNoo} className="space-y-4">
                            <div>
                              <label className="block text-[9px] font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                                Tanggal Log Outlet
                              </label>
                              <input
                                type="date"
                                required
                                value={newNooDate}
                                onChange={(e) => setNewNooDate(e.target.value)}
                                className="w-full bg-white border border-[#E5E5DF] text-[#4A4A3C] font-black text-xs px-3 py-2.5 rounded-xl focus:outline-hidden"
                              />
                            </div>

                            {/* WARUNG */}
                            <div className="bg-white border border-[#E5E5DF] rounded-2xl p-3 flex items-center justify-between">
                              <div>
                                <span className="text-[10px] font-bold text-amber-800 block uppercase leading-none font-sans">Warung Kelontong</span>
                                <span className="text-[8px] text-[#8C8C70] uppercase mt-0.5 block">Warung kecil & depot</span>
                              </div>
                              <div className="flex items-center gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => setNewNooWarung(p => Math.max(0, p - 1))}
                                  className="w-8 h-8 rounded-xl bg-[#E5E5DF]/40 hover:bg-[#E5E5DF]/70 text-[#4A4A3C] flex items-center justify-center font-bold text-lg select-none cursor-pointer"
                                >
                                  -
                                </button>
                                <span className="w-8 text-center font-bold font-mono text-sm">
                                  {newNooWarung}
                                </span>
                                <button
                                  type="button"
                                  onClick={() => setNewNooWarung(p => p + 1)}
                                  className="w-8 h-8 rounded-xl bg-amber-500 text-white hover:bg-amber-600 flex items-center justify-center font-bold text-lg select-none cursor-pointer"
                                >
                                  +
                                </button>
                              </div>
                            </div>

                            {/* TOKO MODERN */}
                            <div className="bg-white border border-[#E5E5DF] rounded-2xl p-3 flex items-center justify-between">
                              <div>
                                <span className="text-[10px] font-bold text-rose-800 block uppercase leading-none font-sans">Toko / Store</span>
                                <span className="text-[8px] text-[#8C8C70] uppercase mt-0.5 block">Toko sedang, minimarket</span>
                              </div>
                              <div className="flex items-center gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => setNewNooStore(p => Math.max(0, p - 1))}
                                  className="w-8 h-8 rounded-xl bg-[#E5E5DF]/40 hover:bg-[#E5E5DF]/70 text-[#4A4A3C] flex items-center justify-center font-bold text-lg select-none cursor-pointer"
                                >
                                  -
                                </button>
                                <span className="w-8 text-center font-bold font-mono text-sm">
                                  {newNooStore}
                                </span>
                                <button
                                  type="button"
                                  onClick={() => setNewNooStore(p => p + 1)}
                                  className="w-8 h-8 rounded-xl bg-rose-600 text-white hover:bg-rose-700 flex items-center justify-center font-bold text-lg select-none cursor-pointer"
                                >
                                  +
                                </button>
                              </div>
                            </div>

                            {/* KIOS */}
                            <div className="bg-white border border-[#E5E5DF] rounded-2xl p-3 flex items-center justify-between">
                              <div>
                                <span className="text-[10px] font-bold text-indigo-800 block uppercase leading-none font-sans">Kios Atap</span>
                                <span className="text-[8px] text-[#8C8C70] uppercase mt-0.5 block">Kios pasar & tenda</span>
                              </div>
                              <div className="flex items-center gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => setNewNooKiosk(p => Math.max(0, p - 1))}
                                  className="w-8 h-8 rounded-xl bg-[#E5E5DF]/40 hover:bg-[#E5E5DF]/70 text-[#4A4A3C] flex items-center justify-center font-bold text-lg select-none cursor-pointer"
                                >
                                  -
                                </button>
                                <span className="w-8 text-center font-bold font-mono text-sm">
                                  {newNooKiosk}
                                </span>
                                <button
                                  type="button"
                                  onClick={() => setNewNooKiosk(p => p + 1)}
                                  className="w-8 h-8 rounded-xl bg-indigo-600 text-white hover:bg-indigo-700 flex items-center justify-center font-bold text-lg select-none cursor-pointer"
                                >
                                  +
                                </button>
                              </div>
                            </div>

                            {/* GROSIR */}
                            <div className="bg-white border border-[#E5E5DF] rounded-2xl p-3 flex items-center justify-between">
                              <div>
                                <span className="text-[10px] font-bold text-emerald-800 block uppercase leading-none font-sans">Wholesaler / Grosir</span>
                                <span className="text-[8px] text-[#8C8C70] uppercase mt-0.5 block">Toko agen besar</span>
                              </div>
                              <div className="flex items-center gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => setNewNooWholesaler(p => Math.max(0, p - 1))}
                                  className="w-8 h-8 rounded-xl bg-[#E5E5DF]/40 hover:bg-[#E5E5DF]/70 text-[#4A4A3C] flex items-center justify-center font-bold text-lg select-none cursor-pointer"
                                >
                                  -
                                </button>
                                <span className="w-8 text-center font-bold font-mono text-sm">
                                  {newNooWholesaler}
                                </span>
                                <button
                                  type="button"
                                  onClick={() => setNewNooWholesaler(p => p + 1)}
                                  className="w-8 h-8 rounded-xl bg-emerald-600 text-white hover:bg-emerald-700 flex items-center justify-center font-bold text-lg select-none cursor-pointer"
                                >
                                  +
                                </button>
                              </div>
                            </div>

                            <button
                              type="submit"
                              className="w-full bg-[#5A5A40] hover:bg-[#4A4A3C] text-white font-black text-xs uppercase tracking-wider py-3.5 rounded-xl shadow-md hover:shadow-lg transition cursor-pointer flex justify-center items-center gap-1.5"
                            >
                              <Check className="w-4 h-4" /> Simpan Log NOO Harian
                            </button>
                          </form>
                        </div>

                        {/* LIST OF RECENT KPI REPORTS FILED */}
                        <div className="bg-[#FAF9F6] border border-[#E5E5DF] rounded-3xl p-6 shadow-sm mt-6 space-y-4">
                          <h4 className="text-xs font-black text-[#5A5A40] uppercase tracking-wider flex items-center gap-1.5 border-b border-[#E5E5DF] pb-3">
                            <FileText className="w-4 h-4 text-rose-500" />
                            LAPORAN HARIAN TERAKHIR
                          </h4>
                          {salesmanReports.length === 0 ? (
                            <p className="text-[11px] text-[#8C8C70] text-center font-semibold italic">Belum ada laporan audit KPI terdaftar.</p>
                          ) : (
                            <div className="space-y-3 max-h-72 overflow-y-auto pr-1">
                              {salesmanReports.slice(0, 5).map(rep => (
                                <div key={rep.id} className="bg-white p-3 rounded-2xl border border-[#E5E5DF]/60 text-[11px] space-y-2">
                                  <div className="flex justify-between font-bold">
                                    <span className="font-mono text-rose-800">{rep.date}</span>
                                    <span className="text-[#8C8C70]">{rep.cycle}</span>
                                  </div>
                                  <div className="grid grid-cols-2 gap-1 text-[10px] text-[#5A5A40]">
                                    <div>TC/CP/EC: <strong className="text-gray-800 font-mono">{rep.tc}/{rep.cp}/{rep.ec}</strong></div>
                                    <div>SKU Terjual: <strong className="text-gray-800 font-mono">{rep.skuTotal}</strong></div>
                                  </div>
                                  {rep.notes && (
                                    <p className="text-[10px] text-[#8C8C70] leading-snug border-t border-gray-100 pt-1.5 italic">
                                      "{rep.notes}"
                                    </p>
                                  )}
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      </div>

                    </div>
                  </div>
                );
              })()}

            </motion.div>
          )}

        </AnimatePresence>

        </div>

      {/* --- DATABASE DIALOGS / MODALS SIMULATION --- */}

      {/* 0. CUSTOMER PROFILING LOYALTY MODAL */}
      <AnimatePresence>
        {isCustomerModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-[#4A4A3C]/30 backdrop-blur-xs">
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-[#FAF9F6] rounded-3xl p-6 w-full max-w-lg border border-[#E5E5DF] shadow-2xl relative max-h-[90vh] overflow-y-auto"
            >
              <button
                onClick={() => setIsCustomerModalOpen(false)}
                className="absolute right-4 top-4 text-[#8C8C70] hover:text-[#4A4A3C] transition cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>

              <h3 className="text-base font-black text-[#5A5A40] uppercase tracking-wider mb-1 flex items-center gap-2">
                <Sparkles className="w-5 h-5 text-amber-500" />
                Customer Profiling & Loyalty Sign-up
              </h3>
              <p className="text-[11px] text-[#8C8C70] uppercase tracking-widest mb-4">
                Input Profil Toko & Klasifikasi Tingkat Poin
              </p>

              <hr className="border-[#E5E5DF] mb-4" />

              <form onSubmit={handleSaveCustomerProfile} className="space-y-4">
                
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-[10px] font-bold text-[#4A4A3C] uppercase tracking-wider mb-1">
                      Nama Toko / Outlet *
                    </label>
                    <input
                      type="text"
                      required
                      placeholder="Contoh: TOKO ANUGERAH UTAMA"
                      value={newCustomer.name}
                      onChange={(e) => setNewCustomer({ ...newCustomer, name: e.target.value })}
                      className="w-full bg-[#E5E5DF]/15 border border-[#E5E5DF] rounded-xl px-3 py-2 text-xs font-semibold focus:outline-hidden"
                    />
                  </div>

                  <div>
                    <label className="block text-[10px] font-bold text-[#4A4A3C] uppercase tracking-wider mb-1">
                      Jenis Toko / Outlet
                    </label>
                    <select
                      value={newCustomer.jenisToko}
                      onChange={(e) => setNewCustomer({ ...newCustomer, jenisToko: e.target.value })}
                      className="w-full bg-white border border-[#E5E5DF] rounded-xl px-2 py-2 text-xs text-[#4A4A3C] focus:outline-hidden"
                    >
                      <option value="Sembako">🌾 Toko Sembako</option>
                      <option value="Kelontong">🛒 Toko Kelontong</option>
                      <option value="Grosir">📦 Toko Grosir</option>
                      <option value="ATK / Lainnya">📎 ATK / Retail Lainnya</option>
                    </select>
                  </div>
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-[#4A4A3C] uppercase tracking-wider mb-1">
                    Alamat Lengkap Toko
                  </label>
                  <input
                    type="text"
                    placeholder="Contoh: Jl. Diponegoro No. 45, Semarang Tengah"
                    value={newCustomer.address}
                    onChange={(e) => setNewCustomer({ ...newCustomer, address: e.target.value })}
                    className="w-full bg-[#E5E5DF]/15 border border-[#E5E5DF] rounded-xl px-3 py-2 text-xs focus:outline-hidden"
                  />
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-[10px] font-bold text-[#4A4A3C] uppercase tracking-wider mb-1">
                      Sales Representative (Penanggung Jawab)
                    </label>
                    <select
                      value={newCustomer.salesmanName}
                      onChange={(e) => {
                        const sName = e.target.value;
                        const match = salesmen.find(s => s.name === sName);
                        setNewCustomer({
                          ...newCustomer,
                          salesmanName: sName,
                          area: match?.area || "Semarang"
                        });
                      }}
                      className="w-full bg-white border border-[#E5E5DF] rounded-xl px-2 py-2 text-xs text-[#4A4A3C] focus:outline-hidden"
                    >
                      {salesmen.map(s => (
                        <option key={s.id} value={s.name}>
                          {s.name} ({s.area})
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-[10px] font-bold text-[#4A4A3C] uppercase tracking-wider mb-1">
                      Area Distribusi
                    </label>
                    <input
                      type="text"
                      value={newCustomer.area}
                      onChange={(e) => setNewCustomer({ ...newCustomer, area: e.target.value })}
                      placeholder="Nama Area"
                      className="w-full bg-white border border-[#E5E5DF] rounded-xl px-3 py-2 text-xs font-bold text-[#4A4A3C] focus:outline-hidden"
                    />
                  </div>
                </div>

                <hr className="border-[#E5E5DF]/40 my-2" />

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 bg-[#E5E5DF]/20 p-3.5 rounded-2xl border border-[#E5E5DF]/40">
                  
                  <div>
                    <label className="block text-[9px] font-black text-[#5A5A40] uppercase tracking-wider mb-1">
                      Omzet Per Bulan *
                    </label>
                    <input
                      type="number"
                      required
                      min="0"
                      value={newCustomer.estimatedOmzet}
                      onChange={(e) => setNewCustomer({ ...newCustomer, estimatedOmzet: Number(e.target.value) })}
                      className="w-full bg-white border border-[#E5E5DF] rounded-xl px-2 py-1.5 text-xs font-mono font-bold text-[#4A4A3C]"
                    />
                    <span className="text-[8px] text-[#8C8C70] block mt-1 leading-none">Min. Rp 5jt (Silver), 8jt (Gold), 15jt (Plat)</span>
                  </div>

                  <div>
                    <label className="block text-[9px] font-black text-[#5A5A40] uppercase tracking-wider mb-1">
                      Nota / Hari (DKR)
                    </label>
                    <input
                      type="number"
                      required
                      min="0"
                      value={newCustomer.notesPerDay}
                      onChange={(e) => setNewCustomer({ ...newCustomer, notesPerDay: Number(e.target.value) })}
                      className="w-full bg-white border border-[#E5E5DF] rounded-xl px-2 py-1.5 text-xs text-center font-mono text-[#4A4A3C]"
                    />
                  </div>

                  <div>
                    <label className="block text-[9px] font-black text-[#5A5A40] uppercase tracking-wider mb-1">
                      Umur Gabung (Th)
                    </label>
                    <input
                      type="number"
                      required
                      min="0"
                      value={newCustomer.storeAgeYears}
                      onChange={(e) => setNewCustomer({ ...newCustomer, storeAgeYears: Number(e.target.value) })}
                      className="w-full bg-white border border-[#E5E5DF] rounded-xl px-2 py-1.5 text-xs text-center font-mono text-[#4A4A3C]"
                    />
                  </div>

                </div>

                <div className="flex items-center gap-3">
                  <span className="text-[10px] font-bold text-[#8C8C70] uppercase">Status Kepemilikan Bangunan:</span>
                  <label className="flex items-center gap-1.5 text-xs text-[#4A4A3C] font-semibold">
                    <input
                      type="radio"
                      name="ownership"
                      value="Milik Sendiri"
                      checked={newCustomer.ownership === "Milik Sendiri"}
                      onChange={() => setNewCustomer({ ...newCustomer, ownership: "Milik Sendiri" })}
                    />
                    Milik Sendiri
                  </label>
                  <label className="flex items-center gap-1.5 text-xs text-[#4A4A3C] font-semibold">
                    <input
                      type="radio"
                      name="ownership"
                      value="Sewa Kontrak"
                      checked={newCustomer.ownership === "Sewa Kontrak"}
                      onChange={() => setNewCustomer({ ...newCustomer, ownership: "Sewa Kontrak" })}
                    />
                    Sewa Kontrak
                  </label>
                </div>

                <div className="pt-2 flex justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setIsCustomerModalOpen(false)}
                    className="px-4 py-2 text-xs font-bold text-[#8C8C70] uppercase bg-[#E5E5DF]/40 hover:bg-[#E5E5DF] rounded-xl transition cursor-pointer"
                  >
                    Batal
                  </button>
                  <button
                    type="submit"
                    className="px-5 py-2 text-xs font-bold text-white uppercase bg-[#5A5A40] hover:bg-[#4A4A3C] rounded-xl shadow-xs transition cursor-pointer"
                  >
                    Simpan Profiling (+50 Poin)
                  </button>
                </div>

              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* 1. SALESMAN MODAL */}
      <AnimatePresence>
        {salesmanModal.isOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-[#4A4A3C]/30 backdrop-blur-xs">
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-[#FAF9F6] rounded-3xl p-6 w-full max-w-md border border-[#E5E5DF] shadow-2xl relative"
            >
              <button
                onClick={() => setSalesmanModal({ ...salesmanModal, isOpen: false })}
                className="absolute right-4 top-4 text-[#8C8C70] hover:text-[#4A4A3C] transition cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>

              <h3 className="text-base font-bold text-[#4A4A3C] font-serif italic uppercase mb-4 flex items-center gap-2">
                <Database className="w-5 h-5 text-[#5A5A40]" />
                {salesmanModal.id ? "Edit Data Salesman" : "Tambah Salesman Baru"}
              </h3>

              <form onSubmit={handleSaveSalesman} className="flex flex-col gap-4">
                <div>
                  <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                    Nama Salesman (Kapital) *
                  </label>
                  <input
                    type="text"
                    required
                    value={salesmanModal.name}
                    onChange={(e) => setSalesmanModal({ ...salesmanModal, name: e.target.value })}
                    placeholder="Contoh: RENY, BUDI, dsb."
                    className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl px-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] text-[#4A4A3C] uppercase font-bold"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                    Wilayah Kunjungan / Area
                  </label>
                  <input
                    type="text"
                    value={salesmanModal.area}
                    onChange={(e) => setSalesmanModal({ ...salesmanModal, area: e.target.value })}
                    placeholder="Contoh: Semarang Barat"
                    className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl px-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] text-[#4A4A3C]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                    Nomor HP Salesman
                  </label>
                  <input
                    type="text"
                    value={salesmanModal.phone}
                    onChange={(e) => setSalesmanModal({ ...salesmanModal, phone: e.target.value })}
                    placeholder="Contoh: 081234567xxx"
                    className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl px-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] text-[#4A4A3C] font-mono"
                  />
                </div>

                <div className="flex gap-2.5 pt-2">
                  <button
                    type="submit"
                    className="flex-1 bg-[#5A5A40] hover:bg-[#4A4A3C] text-[#FAF9F6] font-bold text-xs uppercase tracking-wider py-3 rounded-xl transition cursor-pointer"
                  >
                    Simpan ke DB
                  </button>
                  <button
                    type="button"
                    onClick={() => setSalesmanModal({ ...salesmanModal, isOpen: false })}
                    className="bg-[#E5E5DF]/50 hover:bg-[#E5E5DF] text-[#4A4A3C] font-bold text-xs uppercase px-4 py-3 rounded-xl transition duration-150 cursor-pointer"
                  >
                    Batal
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* 2. PRODUCT MODAL */}
      <AnimatePresence>
        {productModal.isOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-[#4A4A3C]/30 backdrop-blur-xs">
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-[#FAF9F6] rounded-3xl p-6 w-full max-w-md border border-[#E5E5DF] shadow-2xl relative"
            >
              <button
                onClick={() => setProductModal({ ...productModal, isOpen: false })}
                className="absolute right-4 top-4 text-[#8C8C70] hover:text-[#4A4A3C] transition cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>

              <h3 className="text-base font-bold text-[#4A4A3C] font-serif italic uppercase mb-4 flex items-center gap-2">
                <ShoppingBag className="w-5 h-5 text-[#5A5A40]" />
                {productModal.id ? "Edit Data Produk" : "Tambah Produk Baru"}
              </h3>

              <form onSubmit={handleSaveProduct} className="flex flex-col gap-4">
                <div>
                  <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                    Nama SKU Produk (Kapital) *
                  </label>
                  <input
                    type="text"
                    required
                    value={productModal.name}
                    onChange={(e) => setProductModal({ ...productModal, name: e.target.value })}
                    placeholder="Contoh: TISSUE YO PIPI POP UP"
                    className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl px-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] text-[#4A4A3C] uppercase font-bold"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                    Kode Singkat SKU
                  </label>
                  <input
                    type="text"
                    value={productModal.skuCode}
                    onChange={(e) => setProductModal({ ...productModal, skuCode: e.target.value })}
                    placeholder="Contoh: TJ-YPP-PU"
                    className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl px-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] text-[#4A4A3C]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-[#8C8C70] uppercase tracking-wider mb-1">
                    Kategori / Group Produk
                  </label>
                  <input
                    type="text"
                    value={productModal.category}
                    onChange={(e) => setProductModal({ ...productModal, category: e.target.value })}
                    placeholder="Contoh: Tissue, Baby Care, dsb."
                    className="w-full bg-[#FAF9F6] border border-[#E5E5DF] rounded-xl px-4 py-3 text-sm focus:outline-hidden focus:ring-2 focus:ring-[#5A5A40] text-[#4A4A3C]"
                  />
                </div>

                <div className="flex gap-2.5 pt-2">
                  <button
                    type="submit"
                    className="flex-1 bg-[#5A5A40] hover:bg-[#4A4A3C] text-[#FAF9F6] font-bold text-xs uppercase tracking-wider py-3 rounded-xl transition cursor-pointer"
                  >
                    Simpan ke DB
                  </button>
                  <button
                    type="button"
                    onClick={() => setProductModal({ ...productModal, isOpen: false })}
                    className="bg-[#E5E5DF]/50 hover:bg-[#E5E5DF] text-[#4A4A3C] font-bold text-xs uppercase px-4 py-3 rounded-xl transition duration-150 cursor-pointer"
                  >
                    Batal
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* 2.5. PRODUCT DELETE CONFIRMATION MODAL */}
      <AnimatePresence>
        {deleteConfirmProductId && (() => {
          const productToDelete = products.find(p => p.id === deleteConfirmProductId);
          if (!productToDelete) return null;
          return (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-[#4A4A3C]/45 backdrop-blur-xs">
              <motion.div
                initial={{ scale: 0.95, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                exit={{ scale: 0.95, opacity: 0 }}
                className="bg-[#FAF9F6] rounded-3xl p-6 w-full max-w-md border border-[#E5E5DF] shadow-2xl relative"
              >
                <button
                  onClick={() => setDeleteConfirmProductId(null)}
                  className="absolute right-4 top-4 text-[#8C8C70] hover:text-[#4A4A3C] transition cursor-pointer"
                  type="button"
                >
                  <X className="w-5 h-5" />
                </button>

                <h3 className="text-base font-bold text-rose-800 font-serif italic uppercase mb-3 flex items-center gap-2">
                  <AlertTriangle className="w-5 h-5 text-rose-600" />
                  Konfirmasi Hapus Produk
                </h3>

                <div className="space-y-3 bg-rose-50/70 border border-rose-100 p-4 rounded-2xl mb-5">
                  <p className="text-xs text-rose-900 leading-relaxed font-semibold">
                    Apakah Anda yakin ingin menghapus produk ini dari database?
                  </p>
                  
                  <div className="border-t border-dashed border-rose-200/50 pt-3 text-xs space-y-1.5 text-gray-700 font-bold">
                    <div className="flex justify-between">
                      <span className="text-gray-400 font-normal font-sans">Nama Produk:</span>
                      <span className="text-rose-950 uppercase font-black">{productToDelete.name}</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-gray-400 font-normal font-sans">Kode SKU:</span>
                      <span className="font-mono bg-[#E5E5DF]/50 px-1.5 py-0.5 rounded text-[10px]">{productToDelete.skuCode || "-"}</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-gray-400 font-normal font-sans">Kategori:</span>
                      <span className="text-gray-800">{productToDelete.category || "Umum"}</span>
                    </div>
                  </div>
                </div>

                <div className="flex gap-2.5">
                  <button
                    onClick={confirmDeleteProduct}
                    className="flex-1 bg-rose-600 hover:bg-rose-700 text-[#FAF9F6] font-bold text-xs uppercase tracking-wider py-3 rounded-xl transition cursor-pointer shadow-sm shadow-rose-600/10"
                    type="button"
                  >
                    Ya, Hapus Permanen
                  </button>
                  <button
                    type="button"
                    onClick={() => setDeleteConfirmProductId(null)}
                    className="bg-[#E5E5DF]/50 hover:bg-[#E5E5DF] text-[#4A4A3C] font-bold text-xs uppercase px-4 py-3 rounded-xl transition duration-150 cursor-pointer"
                  >
                    Batal
                  </button>
                </div>
              </motion.div>
            </div>
          );
        })()}
      </AnimatePresence>

      {/* 2.6. LOYALTY CUSTOMER DELETE CONFIRMATION MODAL */}
      <AnimatePresence>
        {deleteConfirmCustomerId && (() => {
          const customerToDelete = customers.find(c => c.id === deleteConfirmCustomerId);
          if (!customerToDelete) return null;
          return (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-[#4A4A3C]/45 backdrop-blur-xs">
              <motion.div
                initial={{ scale: 0.95, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                exit={{ scale: 0.95, opacity: 0 }}
                className="bg-[#FAF9F6] rounded-3xl p-6 w-full max-w-md border border-[#E5E5DF] shadow-2xl relative"
              >
                <button
                  onClick={() => setDeleteConfirmCustomerId(null)}
                  className="absolute right-4 top-4 text-[#8C8C70] hover:text-[#4A4A3C] transition cursor-pointer"
                  type="button"
                >
                  <X className="w-5 h-5" />
                </button>

                <h3 className="text-base font-bold text-rose-800 font-serif italic uppercase mb-3 flex items-center gap-2">
                  <AlertTriangle className="w-5 h-5 text-rose-600" />
                  Konfirmasi Hapus Profiling Loyalti
                </h3>

                <div className="space-y-3 bg-rose-50/70 border border-rose-100 p-4 rounded-2xl mb-5">
                  <p className="text-xs text-rose-900 leading-relaxed font-semibold">
                    Apakah Anda yakin ingin menghapus data profiling Toko {customerToDelete.name.toUpperCase()}? Tindakan ini akan menyinkronkan data terbaru ke Google Sheets secara real-time!
                  </p>
                  
                  <div className="border-t border-dashed border-rose-200/50 pt-3 text-xs space-y-1.5 text-gray-700 font-bold">
                    <div className="flex justify-between">
                      <span className="text-gray-400 font-normal font-sans">Nama Toko:</span>
                      <span className="text-rose-950 uppercase font-black">{customerToDelete.name}</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-gray-400 font-normal font-sans">Kelas Tier:</span>
                      <span className="text-emerald-700 font-extrabold uppercase">★ Tier {customerToDelete.tier}</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-gray-400 font-normal font-sans">Salesman:</span>
                      <span className="text-gray-800">{customerToDelete.salesmanName} ({customerToDelete.area})</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-gray-400 font-normal font-sans">Poin Saat Ini:</span>
                      <span className="text-amber-700 font-bold font-mono">{customerToDelete.points} Poin</span>
                    </div>
                  </div>
                </div>

                <div className="flex gap-2.5">
                  <button
                    onClick={confirmDeleteCustomer}
                    className="flex-1 bg-rose-600 hover:bg-rose-700 text-[#FAF9F6] font-bold text-xs uppercase tracking-wider py-3 rounded-xl transition cursor-pointer shadow-sm shadow-rose-600/10"
                    type="button"
                  >
                    Ya, Hapus Profiling
                  </button>
                  <button
                    type="button"
                    onClick={() => setDeleteConfirmCustomerId(null)}
                    className="bg-[#E5E5DF]/50 hover:bg-[#E5E5DF] text-[#4A4A3C] font-bold text-xs uppercase px-4 py-3 rounded-xl transition duration-150 cursor-pointer"
                  >
                    Batal
                  </button>
                </div>
              </motion.div>
            </div>
          );
        })()}
      </AnimatePresence>

      {/* 3. REPORT DELETE CONFIRMATION MODAL */}
      <AnimatePresence>
        {deleteConfirmId && (() => {
          const reportToDelete = reports.find(r => r.id === deleteConfirmId);
          if (!reportToDelete) return null;
          return (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-[#4A4A3C]/45 backdrop-blur-xs">
              <motion.div
                initial={{ scale: 0.95, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                exit={{ scale: 0.95, opacity: 0 }}
                className="bg-[#FAF9F6] rounded-3xl p-6 w-full max-w-md border border-[#E5E5DF] shadow-2xl relative"
              >
                <button
                  onClick={() => setDeleteConfirmId(null)}
                  className="absolute right-4 top-4 text-[#8C8C70] hover:text-[#4A4A3C] transition cursor-pointer"
                  type="button"
                >
                  <X className="w-5 h-5" />
                </button>

                <h3 className="text-base font-bold text-rose-800 font-serif italic uppercase mb-3 flex items-center gap-2">
                  <AlertTriangle className="w-5 h-5 text-rose-600" />
                  Konfirmasi Hapus Laporan
                </h3>

                <div className="space-y-3 bg-rose-50/70 border border-rose-100 p-4 rounded-2xl mb-5">
                  <p className="text-xs text-rose-900 leading-relaxed font-semibold">
                    Apakah Anda yakin ingin menghapus data laporan audit KPI berikut? Tindakan ini tidak dapat dibatalkan.
                  </p>
                  
                  <div className="border-t border-dashed border-rose-200/50 pt-3 text-xs space-y-1.5 text-gray-700 font-bold">
                    <div className="flex justify-between">
                      <span className="text-gray-400 font-normal font-sans">Salesman:</span>
                      <span className="text-rose-950 uppercase font-black">{reportToDelete.salesmanName}</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-gray-400 font-normal font-sans">Tanggal:</span>
                      <span>{reportToDelete.date}</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-gray-400 font-normal font-sans">Siklus:</span>
                      <span className="bg-rose-100 text-rose-800 px-1.5 py-0.5 rounded text-[10px] font-extrabold uppercase">{reportToDelete.cycle}</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-gray-400 font-normal font-sans">Metrik (TC/CP/EC/SKU):</span>
                      <span>{reportToDelete.tc} / {reportToDelete.cp} / {reportToDelete.ec} / {reportToDelete.skuTotal}</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-gray-400 font-normal font-sans">Tagihan Tunai:</span>
                      <span className="text-emerald-700 font-mono">Rp {reportToDelete.billsReceived.toLocaleString("id-ID")}</span>
                    </div>
                  </div>
                </div>

                <div className="flex gap-2.5">
                  <button
                    onClick={confirmDeleteReport}
                    className="flex-1 bg-rose-600 hover:bg-rose-700 text-[#FAF9F6] font-bold text-xs uppercase tracking-wider py-3 rounded-xl transition cursor-pointer shadow-sm shadow-rose-600/10"
                    type="button"
                  >
                    Ya, Hapus Permanen
                  </button>
                  <button
                    type="button"
                    onClick={() => setDeleteConfirmId(null)}
                    className="bg-[#E5E5DF]/50 hover:bg-[#E5E5DF] text-[#4A4A3C] font-bold text-xs uppercase px-4 py-3 rounded-xl transition duration-150 cursor-pointer"
                  >
                    Batal
                  </button>
                </div>
              </motion.div>
            </div>
          );
        })()}
      </AnimatePresence>

      {/* --- FOOTER DESCRIPTOR --- */}
      <footer className="bg-[#FAF9F6] border-t border-[#E5E5DF] mt-12 py-8 text-center text-xs text-[#8C8C70]">
        <div className="max-w-2xl mx-auto px-4 flex flex-col gap-2">
          <p className="font-bold text-[#5A5A40] uppercase tracking-widest">
            Audit KPI Sales • Auditor Portal
          </p>
          <p className="leading-relaxed">
            Menyelaraskan laporan harian sales force dengan fungsionalitas visual yang bersih, parsing otomatis, dan basis data sederhana yang responsif.
          </p>
          <p className="mt-2 font-mono text-[10px]">
            © {new Date().getFullYear()} AI Studio. All rights reserved.
          </p>
        </div>
      </footer>
      </div>

      {/* Modals */}
      <RewardModal 
        isOpen={rewardModal.isOpen}
        type={rewardModal.type}
        item={rewardModal.item}
        onClose={() => setRewardModal({ isOpen: false, type: 'merchant' })}
        onSave={handleSaveReward}
        onDelete={handleDeleteReward}
      />
    </div>
  );
}
