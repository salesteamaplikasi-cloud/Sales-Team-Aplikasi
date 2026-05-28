export interface Customer {
  id: string;
  nama: string;
  jalan: string;
  kota: string;
}

export interface VisitSchedule {
  salesman: string;
  kunjungan: string;
  customers: Customer[];
}

export const CUSTOMER_DATA: VisitSchedule[] = [
  {
    salesman: "ARIS",
    kunjungan: "SENIN GANJIL",
    customers: [
      { id: "CUST-00788", nama: "TK SONDANG", jalan: "SUMBANG", kota: "BANYUMAS" },
      { id: "CUST-00884", nama: "HR (CILONGOK)", jalan: "JL. CILONGOK - SUDIMARA KM 01, TERMINAL ANGKUDES", kota: "CILONGOK" },
      { id: "CUST-00888", nama: "HR (LOSARI)", jalan: "JL. LOSARI CIPENDOK KM 1", kota: "CIPENDOK" },
      { id: "CUST-00944", nama: "TK.FORMOSA MART", jalan: "KR.GAMBAS, PADAMARA", kota: "PURBALINGGA" }
    ]
  },
  {
    salesman: "ARIS",
    kunjungan: "SELASA GANJIL",
    customers: [
      { id: "CUST-00056", nama: "TOKO SATU HATI/KIYAN WANGI HABIBA", jalan: "PETANAHAN", kota: "KEBUMEN" },
      { id: "CUST-00155", nama: "KASYATI SNACK", jalan: "PSR GROPAK", kota: "KEBUMEN" },
      { id: "CUST-00154", nama: "KENCANAMULYA", jalan: "PETANAHAN", kota: "KEBUMEN" },
      { id: "CUST-00156", nama: "SANG BARU SWALAYAN", jalan: "GROPAK", kota: "KEBUMEN" }
    ]
  },
  {
    salesman: "IMAM",
    kunjungan: "SENIN GANJIL",
    customers: [
      { id: "CUST-00101", nama: "YANTO", jalan: "ARCAWINANGUN", kota: "BANYUMAS" },
      { id: "CUST-00106", nama: "BUDI DAYA", jalan: "MERSI", kota: "BANYUMAS" },
      { id: "CUST-00107", nama: "SINAR BUANA", jalan: "PURWOKERTO", kota: "BANYUMAS" }
    ]
  }
];
