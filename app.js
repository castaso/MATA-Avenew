/* ============================================================
   MATA Dashboard – App Logic & Data
   Jargas Gabungan – Dashboard Weekly Ke-11
   Data source: Dashboard Gabungan Jargas Week 11.pptx
   ============================================================ */

'use strict';

/* ── ① MASTER DATA ──────────────────────────────────────── */
const DATA = {
  project: {
    name: "Dashboard Gabungan Jargas",
    week: 11,
    period: "28 Januari – 04 Februari 2026",
    status_date: "04 Februari 2026",
    total_sr: 115264,
    total_contract: 961913187634,
    ministry: "Direktorat Perencanaan & Pembangunan Infrastruktur Migas, KESDM",
    total_safe_hours: 2530813,
    total_manpower: 2684
  },

  packages: [
    /* ── Paket 1 ─────────────────────────────────────────── */
    {
      id: 1,
      code: "P1",
      name: "Kota Jambi, Kab. TJB & Pelalawan",
      short: "Jambi Region",
      epc: "KSO NOOREL-KSM",
      pmc: "PT. Surveyor Indonesia",
      contract_value: 182717782371,
      total_sr: 22972,
      duration: "240 hari",
      start: "19 November 2025",
      color: "#1e6aff",
      grad: "linear-gradient(135deg,#1e6aff,#00c6ff)",
      safe_hours: 472684,
      progress_plan: null,
      progress_actual: null,
      progress_dev: null,
      locations: [
        {
          code: "JBI", name: "Kota Jambi", sr: 13235,
          contract: 105884462327, safe_hours: 276336,
          workforce: { "EPC Staff": 28, "PMC": 32, "Pekerja Konstruksi": 570, "CAS": 1 },
          progress_plan: null, progress_actual: null, progress_dev: null,
          permits: [
            { name: "Izin Lingkungan", status: "Proses" },
            { name: "Izin Penempatan RS", status: "Proses" },
            { name: "Izin Tie In", status: "Open" },
            { name: "Izin Penggalian", status: "Close" },
            { name: "Izin Pemanfaatan Jalan Nasional", status: "Proses" }
          ],
          konstruksi: null, weekly: null, sk_sectors: null
        },
        {
          code: "TJB", name: "Kab. Tanjung Jabung Barat", sr: 6661,
          contract: 46639522633, safe_hours: 111480,
          workforce: { "EPC Staff": 20, "PMC": 23, "Pekerja Konstruksi": 120, "CAS": 1 },
          progress_plan: null, progress_actual: null, progress_dev: null,
          permits: [
            { name: "Izin Lingkungan", status: "Open" },
            { name: "Izin Penempatan RS", status: "Open" },
            { name: "Izin Tie In", status: "Open" },
            { name: "Izin Pemanfaatan Jalan", status: "Close" }
          ],
          konstruksi: null, weekly: null, sk_sectors: null
        },
        {
          code: "PLW", name: "Kab. Pelalawan", sr: 3076,
          contract: 30193797412, safe_hours: 84868,
          workforce: { "EPC Staff": 21, "PMC": 13, "Pekerja Konstruksi": 54, "CAS": 1 },
          progress_plan: null, progress_actual: null, progress_dev: null,
          permits: [
            { name: "Izin Lingkungan", status: "Proses" },
            { name: "Izin Penempatan RS", status: "Proses" },
            { name: "Izin Tie In", status: "Proses" },
            { name: "Izin Penggalian", status: "Proses" },
            { name: "Izin Pemanfaatan Jalan Nasional", status: "Proses" },
            { name: "Izin Crossing Pipa TGI", status: "Proses" }
          ],
          konstruksi: null, weekly: null, sk_sectors: null
        }
      ]
    },

    /* ── Paket 2 ─────────────────────────────────────────── */
    {
      id: 2,
      code: "P2",
      name: "Kab. MUBA, OKU Timur & Indramayu",
      short: "Sumatra + Jabar",
      epc: "PT. Pratiwi Putri Sulung",
      pmc: "PT. Amythas",
      contract_value: 233685508475,
      total_sr: 22087,
      tkdn: 67.84,
      duration: "240 hari",
      start: "17 November 2025",
      color: "#00c6ff",
      grad: "linear-gradient(135deg,#0072ff,#00c6ff)",
      safe_hours: 553014,
      manpower: 828,
      progress_plan: null, progress_actual: null, progress_dev: null,
      locations: [
        {
          code: "MUBA", name: "Kab. Musi Banyuasin", sr: 5143,
          contract: 72957976853, safe_hours: 180976,
          progress_plan: 65.453, progress_actual: 58.104, progress_dev: -7.349,
          workforce: { "Galian/Rojok": 123, "Operator+Helper": 14, "Staff Kantor": 20, "PMC/CAS": 17, "SR": 30, "SK": 24 },
          permits: [
            { name: "Izin Prinsip", status: "OPEN" },
            { name: "Izin UKL-UPL", status: "OPEN" },
            { name: "Izin PUPR", status: "OPEN" },
            { name: "IZIN BBPJN", status: "OPEN" },
            { name: "Izin Tie-In/Crossing Pipa", status: "OPEN" },
            { name: "Sosialisasi Pemerintahan", status: "CLOSED" }
          ],
          weekly: {
            labels: ["W1","W2","W3","W4","W5","W6","W7","W8","W9","W10","W11","W12","W13","W14","W15","W16","W17","W18","W19","W20","W21","W22","W23","W24","W25","W26","W27","W28","W29","W30","W31","W32","W33","W34"],
            plan:   [0.0, 9.5, 28.0, 45.0, 49.5, 53.0, 56.5, 59.0, 61.126, 63.429, 65.453, 67.558, 69.513, 71.377, 73.100, 74.818, 76.761, 79.0, 81.5, 83.5, 85.0, 87.0, 88.5, 90.0, 92.0, 93.5, 94.5, 96.0, 98.0, 100.0, 100.0, 100.0, 100.0, 100.0],
            actual: [3.5, 32.0, 46.5, 50.0, 53.0, 54.0, 55.0, 55.5, 55.710, 57.311, 58.104, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null],
            current_week: 11
          },
          konstruksi: [
            { item: "PE 63 mm", pct: 63.035 },
            { item: "SR",       pct: 14.622 },
            { item: "SK",       pct: 93.136 }
          ],
          sk_sectors: [
            { name: "RS 1 – Desa Letang",           boq: 1551, terpasang: 1198 },
            { name: "RS 3 – Desa Babat Banyuasin",  boq: 1709, terpasang: 1709 },
            { name: "RS 4 – Desa Tanjung Kerang",   boq: 1883, terpasang: 1883 }
          ]
        },
        {
          code: "OKUT", name: "Kab. OKU Timur", sr: 4148,
          contract: 49751120209, safe_hours: 186836,
          progress_plan: 66.980, progress_actual: 80.401, progress_dev: 13.421,
          workforce: { "Galian PE63": 53, "Operator Pneumatic": 6, "Operator": 14, "SR": 30, "SK": 22, "MGRT": 10 },
          permits: [
            { name: "Izin UKL/UPL & Lingkungan", status: "On Progress" },
            { name: "Izin Prinsip", status: "Closed" },
            { name: "Izin PUPR Provinsi", status: "On Progress" },
            { name: "Izin BBWS Sumsel", status: "On Progress" },
            { name: "Izin Utilitas Eksisting", status: "On Progress" },
            { name: "Sosialisasi Pemerintahan", status: "Closed" }
          ],
          weekly: {
            labels: ["M-2","W9","W10","W11","W12","W13","W14","W15","W16","W17"],
            plan:   [59.533,62.098,64.640,66.980,69.056,71.144,73.036,74.937,76.870,78.231],
            actual: [69.841,70.771,74.828,80.401,null,null,null,null,null,null],
            current_week: 3
          },
          konstruksi: [
            { item: "PE 180 mm", pct: 100.0  },
            { item: "PE 63 mm",  pct: 99.06  },
            { item: "SR",        pct: 24.54  },
            { item: "MGRT",      pct: 23.02  },
            { item: "SK",        pct: 79.90  }
          ],
          sk_sectors: [
            { name: "RS 1 – Mojosari & Serbaguna",         boq: 223,  terpasang: 223  },
            { name: "RS 3 – Gedung Rejo & Harjowinangun",  boq: 302,  terpasang: 302  },
            { name: "RS 4 – Pujorahayu & Tegalrejo",       boq: 1848, terpasang: 1848 },
            { name: "RS 5 – Gumawang & Bedilan",           boq: 1775, terpasang: 941  }
          ]
        },
        {
          code: "IDM", name: "Kab. Indramayu", sr: 12796,
          contract: 109019979770, safe_hours: 185229,
          progress_plan: 70.064, progress_actual: 68.154, progress_dev: -1.910,
          workforce: { "Galian": 194, "Operator": 7, "SK": 49 },
          permits: [
            { name: "Izin UKL/UPL & Lingkungan",   status: "ON PROCESS" },
            { name: "Izin Prinsip",                 status: "ON PROCESS" },
            { name: "Izin PUPR",                    status: "ON PROCESS" },
            { name: "IZIN BBPJN",                   status: "ON PROCESS" },
            { name: "Izin BBWS",                    status: "ON PROCESS" },
            { name: "Izin Utilitas (Pertagas)",     status: "ON PROCESS" },
            { name: "Izin Utilitas (Pertamina EP)", status: "ON PROCESS" },
            { name: "Sosialisasi Pemerintahan",     status: "CLOSED" }
          ],
          weekly: {
            labels: ["W9","W10","W11","W12","W13","W14","W15","W16","W17"],
            plan:   [65.315,67.567,70.064,72.017,73.132,74.265,75.469,76.770,80.134],
            actual: [65.447,67.186,68.154,null,null,null,null,null,null],
            current_week: 2
          },
          konstruksi: [
            { item: "PE 180 mm", pct: 42.302 },
            { item: "PE 63 mm",  pct: 73.711 },
            { item: "SK",        pct: 87.817 }
          ],
          sk_sectors: [
            { name: "RS 1 – Pertamina SKG KHT",         boq: 1141, terpasang: 1141 },
            { name: "RS 2 – Pertamina SKG KHT",         boq: 1889, terpasang: 1740 },
            { name: "RS 3 – Desa Karanganyar",          boq: 1812, terpasang: 1450 },
            { name: "RS 5 – Desa Karanganyar",          boq: 1576, terpasang: 1266 },
            { name: "RS 6 – Desa Karanganyar",          boq: 1456, terpasang: 1409 },
            { name: "RS 7 – Kec. Kandanghaur",          boq: 1722, terpasang: 1381 },
            { name: "RS 8 – Kec. Kandanghaur",          boq: 1562, terpasang: 1338 },
            { name: "RS 9 – Desa Ilir Kandanghaur",     boq: 1638, terpasang: 1512 }
          ]
        }
      ]
    },

    /* ── Paket 3 ─────────────────────────────────────────── */
    {
      id: 3,
      code: "P3",
      name: "Kab. Batang, Kendal & Wajo",
      short: "Jateng + Sulsel",
      epc: "KSO NOOREL-KSM",
      pmc: "KSO Wahana Krida Konsulindo – Tracon",
      contract_value: 186834052464,
      total_sr: 22213,
      tkdn: 48,
      duration: "240 hari",
      start: "17/18 November 2025",
      color: "#38ef7d",
      grad: "linear-gradient(135deg,#11998e,#38ef7d)",
      safe_hours: 610436,
      progress_plan: 49.37,
      progress_actual: 77.85,
      progress_dev: 28.48,
      weekly_combined: {
        labels: ["W1-4","W5-9","W10","W11","W12","W13","W14","W15","W16","W17","W18","W19"],
        plan:   [43.43,47.06,48.43,49.37,50.16,50.94,51.73,52.72,53.51,53.97,53.97,53.97],
        actual: [61.46,76.23,76.97,77.85,null,null,null,null,null,null,null,null],
        current_week: 3
      },
      locations: [
        {
          code: "BTG", name: "Kab. Batang", sr: 5800,
          contract: 36009983023, safe_hours: 108079,
          progress_plan: 49.61, progress_actual: 83.79, progress_dev: 34.18,
          tkdn: 50,
          workforce: { "Galian": 16, "Operator+Helper": 10, "Rekondisi": 4, "SR": 0, "SK": 16, "Surveyor": 5 },
          permits: [
            { name: "Izin Dinas Lingkungan Hidup",   status: "Open" },
            { name: "Izin Penempatan RS",             status: "Closed" },
            { name: "Izin Andalalin",                 status: "Closed" },
            { name: "Izin Penggalian Pipa (PU Kab)", status: "Closed" },
            { name: "Izin Penggalian Pipa (BBPJN)",  status: "Open" },
            { name: "Izin Perlintasan Sungai (Pusdataru)", status: "Open" },
            { name: "Izin Perlintasan Sungai (BBWS)",     status: "Open" }
          ],
          weekly: {
            labels: ["W1-4","W5-9","W10","W11","W12","W13","W14","W15","W16","W17","W18","W19"],
            plan:   [42.36,46.95,48.45,49.61,50.56,51.50,52.44,53.39,54.33,54.84,54.84,54.84],
            actual: [68.77,83.27,83.33,83.79,null,null,null,null,null,null,null,null],
            current_week: 3
          },
          konstruksi: [
            { item: "PE 180 mm", pct: 38.20 },
            { item: "PE 63 mm",  pct: 86.00 },
            { item: "SK",        pct: 72.41 },
            { item: "SR",        pct: 0.00  }
          ],
          sk_sectors: [
            { name: "RS 1 – Kantor Kec. Grinsing",  boq: 1987, terpasang: 854  },
            { name: "RS 3 – Kantor Desa Grinsing",  boq: 1922, terpasang: 1919 },
            { name: "RS 4 – Balai Desa Lebo",       boq: 1891, terpasang: 1427 }
          ]
        },
        {
          code: "KDL", name: "Kab. Kendal", sr: 12086,
          contract: 83650481260, safe_hours: 351201,
          progress_plan: 46.44, progress_actual: 79.27, progress_dev: 32.83,
          tkdn: 46,
          workforce: { "Galian PE180": 40, "Galian PE63": 85, "Operator+Helper": 16, "SR": 27, "SK": 56, "Surveyor": 6 },
          permits: [
            { name: "Izin UKL/UPL & Lingkungan",         status: "Open 70%" },
            { name: "Izin Penempatan RS/MRS/PRS",         status: "Open" },
            { name: "Izin Penggalian (PUPR Kab)",         status: "Close" },
            { name: "Izin Penggalian (BBPJN)",            status: "Open 70%" },
            { name: "Izin Perlintasan Sungai (BBWS)",     status: "Open 80%" },
            { name: "Izin Perlintasan Sungai (PU SDA)",   status: "Open 80%" },
            { name: "Izin Perlintasan Kereta (KAI)",      status: "Open 20%" },
            { name: "Persetujuan Teknis Lalu Lintas",     status: "Close" }
          ],
          weekly: {
            labels: ["W1-4","W5-9","W10","W11","W12","W13","W14","W15","W16","W17","W18","W19"],
            plan:   [40.76,44.12,45.49,46.44,47.23,48.01,48.80,49.58,50.36,50.80,50.80,50.80],
            actual: [59.06,78.23,78.76,79.27,null,null,null,null,null,null,null,null],
            current_week: 3
          },
          konstruksi: [
            { item: "PE 180 mm", pct: 81.09 },
            { item: "PE 63 mm",  pct: 93.22 },
            { item: "SK",        pct: 76.61 },
            { item: "SR",        pct: 6.81  }
          ],
          sk_sectors: [
            { name: "RS 1 – Desa Sumber Agung",   boq: 965,  terpasang: 879  },
            { name: "RS 2 – Desa Bumiayu",        boq: 1658, terpasang: 1563 },
            { name: "RS 3 – Desa Penyangkringan", boq: 1852, terpasang: 1378 },
            { name: "RS 5 – Desa Karang Anom",    boq: 1257, terpasang: 667  },
            { name: "RS 6 – Desa Sambong Sari",   boq: 1293, terpasang: 1177 },
            { name: "RS 7 – Desa Tratemulyo",     boq: 1314, terpasang: 1061 },
            { name: "RS 8 – Desa Pucuk Sari",     boq: 645,  terpasang: 1149 },
            { name: "RS 9 – Tanah Desa Weleri",   boq: 3102, terpasang: 1373 }
          ]
        },
        {
          code: "WJO", name: "Kab. Wajo", sr: 4327,
          contract: 45107778655, safe_hours: 150156,
          progress_plan: 54.62, progress_actual: 70.47, progress_dev: 15.85,
          tkdn: 49,
          workforce: { "Galian PE180": 103, "Galian PE63": 104, "Operator": 5, "SR": 5, "SK": 17, "Spv/Admin": 9 },
          permits: [
            { name: "Izin Dinas Lingkungan Hidup", status: "Open" },
            { name: "Izin Penempatan RS",          status: "Open" },
            { name: "Izin Tie In (Pertagas)",      status: "Closed" },
            { name: "Izin Penggelaran Pipa (PU)",  status: "Closed" },
            { name: "Izin Penggelaran Pipa (BBPJN Sulsel)", status: "Open" },
            { name: "Rekomendasi Andalalin",       status: "Open" }
          ],
          weekly: {
            labels: ["W1-4","W5-9","W10","W11","W12","W13","W14","W15","W16","W17","W18","W19"],
            plan:   [49.24,52.61,53.86,54.62,55.27,55.93,56.59,58.03,58.69,59.16,59.16,59.16],
            actual: [60.49,66.90,68.33,70.47,null,null,null,null,null,null,null,null],
            current_week: 3
          },
          konstruksi: [
            { item: "PE 180 mm", pct: 50.89 },
            { item: "PE 63 mm",  pct: 35.25 },
            { item: "SK",        pct: 55.63 }
          ],
          sk_sectors: [
            { name: "RS 1",          boq: 818,  terpasang: 605 },
            { name: "RS 3",          boq: 1068, terpasang: 625 },
            { name: "RS 4",          boq: 993,  terpasang: 17  },
            { name: "RS 5",          boq: 1151, terpasang: 861 },
            { name: "Penetrasi RS 26", boq: 297, terpasang: 299 }
          ]
        }
      ]
    },

    /* ── Paket 4 ─────────────────────────────────────────── */
    {
      id: 4,
      code: "P4",
      name: "Kab. Demak, Kota Bontang & Tana Tidung",
      short: "Jateng + Kaltara",
      epc: "PT. Noorel Idea",
      pmc: "PT. Prosys Bangun Persada",
      contract_value: 196756145588,
      total_sr: 26137,
      tkdn: 46.45,
      duration: "240 hari",
      start: "20 November 2025",
      color: "#f97316",
      grad: "linear-gradient(135deg,#f97316,#ffd200)",
      safe_hours: 465791,
      progress_plan: null,
      progress_actual: null,
      progress_dev: null,
      locations: [
        {
          code: "DMK", name: "Kab. Demak", sr: 14859,
          contract: 104337884177, safe_hours: 388648,
          progress_plan: null, progress_actual: null, progress_dev: null,
          workforce: { "EPC Staff": 34, "PMC": 30, "CAS": 1, "Tim SK": 48, "PE 63": 426, "PE 180": 7, "Validasi": 7 },
          permits: [
            { name: "Izin Lingkungan",               status: "Proses" },
            { name: "Izin Penempatan RS/MRS/PRS",    status: "Close" },
            { name: "Izin Tie In",                   status: "Proses" },
            { name: "Izin Pemanfaatan Jalan Prov. Jateng", status: "Proses" },
            { name: "Izin BBWS",                     status: "Proses" },
            { name: "Izin PT KAI",                   status: "Proses" }
          ],
          konstruksi: null, weekly: null, sk_sectors: null
        },
        {
          code: "BTN", name: "Kota Bontang", sr: 10553,
          contract: 42760315304, safe_hours: 42259,
          progress_plan: null, progress_actual: null, progress_dev: null,
          workforce: { "EPC Staff": 20, "PMC": 24, "CAS": 1, "Tim SK": 40, "PE 63+125": 59, "Operator Joint": 2, "Validasi": 11 },
          permits: [
            { name: "Izin Lingkungan",            status: "Proses" },
            { name: "Izin Penempatan RS",          status: "Close" },
            { name: "Izin Tie In",                 status: "Proses" },
            { name: "Izin Pemanfaatan Jalan",      status: "Proses" },
            { name: "Izin BBWS",                   status: "Proses" }
          ],
          konstruksi: null, weekly: null, sk_sectors: null
        },
        {
          code: "TTD", name: "Kab. Tana Tidung", sr: 725,
          contract: 25240823572, safe_hours: 34884,
          progress_plan: null, progress_actual: null, progress_dev: null,
          workforce: { "EPC Staff": 17, "PMC": 15, "CAS": 1, "Tim SK": 14, "PE 63": 50, "PE 180": 11, "Helper": 13 },
          permits: [
            { name: "Izin Lingkungan",        status: "Proses" },
            { name: "Izin Penempatan RS",     status: "Close" },
            { name: "Izin Tie In",            status: "Proses" },
            { name: "Izin Pemanfaatan Jalan", status: "Open" },
            { name: "Izin BBWS",              status: "Proses" }
          ],
          konstruksi: null, weekly: null, sk_sectors: null
        }
      ]
    },

    /* ── Paket 5 ─────────────────────────────────────────── */
    {
      id: 5,
      code: "P5",
      name: "Kab. Gresik, Sidoarjo & Samarinda",
      short: "Jatim + Kaltim",
      epc: "PT. Panca Indah Jayamahe",
      pmc: "PT. Amythas",
      contract_value: 161919698736,
      total_sr: 21855,
      tkdn: 60.30,
      duration: "240 hari",
      start: "17 November 2025",
      color: "#a855f7",
      grad: "linear-gradient(135deg,#8e44ad,#a855f7)",
      safe_hours: 428888,
      manpower: 753,
      progress_plan: 59.155,
      progress_actual: 71.056,
      progress_dev: 11.901,
      locations: [
        {
          code: "GRK", name: "Kab. Gresik", sr: 7013,
          contract: 59322619533, safe_hours: 183964,
          progress_plan: 59.19, progress_actual: 72.40, progress_dev: 13.21,
          workforce: { "PIJ Staff": 19, "PMC": 23, "CAS": 1, "Manpower": 168 },
          permits: [
            { name: "PKKPR",       status: "Open" },
            { name: "UKL/UPL",     status: "Open" },
            { name: "Izin BBWS",   status: "Open" },
            { name: "Izin BBPJN",  status: "Open" },
            { name: "Izin PUPR",   status: "Open" },
            { name: "Izin MRS/RS", status: "Open" }
          ],
          konstruksi: null, weekly: null, sk_sectors: null
        },
        {
          code: "SDA", name: "Kab. Sidoarjo", sr: 7223,
          contract: 46288294091, safe_hours: 151254,
          progress_plan: 63.10, progress_actual: 80.16, progress_dev: 17.06,
          workforce: { "PIJ Staff": 20, "PMC": 23, "CAS": 1, "Manpower": 183 },
          permits: [
            { name: "PKKPR",            status: "Open" },
            { name: "Izin PU Binamarga",status: "Open" },
            { name: "UKL/UPL",          status: "Open" },
            { name: "Izin BBWS",        status: "Open" },
            { name: "Izin RS",          status: "Open" }
          ],
          weekly: {
            labels: ["W8","W9","W10","W11","W12","W13","W14"],
            plan:   [54.43,57.31,60.27,63.10,65.85,71.18,73.34],
            actual: [75.93,78.68,79.56,80.16,null,null,null],
            current_week: 3
          },
          konstruksi: [
            { item: "PE 63 mm",  pct: (71537/110398)*100 },
            { item: "PE 125 mm", pct: (2447/2729)*100   },
            { item: "PE 180 mm", pct: (1832/3404)*100   },
            { item: "SR",        pct: (51/7223)*100     },
            { item: "SK",        pct: (4173/7223)*100   }
          ],
          sk_sectors: null
        },
        {
          code: "SMD", name: "Kota Samarinda", sr: 7619,
          contract: 55575590395, safe_hours: 93670,
          progress_plan: 55.789, progress_actual: 61.924, progress_dev: 6.153,
          workforce: { "PIJ Staff": 22, "PMC": 24, "Tim PPK": 1, "Manpower": 268 },
          permits: [
            { name: "Izin Galian",        status: "Proses" },
            { name: "UKL/UPL",            status: "Proses" },
            { name: "Izin Penempatan RS",  status: "Close" }
          ],
          weekly: {
            labels: ["W7","W8","W9","W10","W11","W12","W13"],
            plan:   [46.23,48.09,50.13,52.94,55.79,58.51,61.39],
            actual: [57.72,58.66,59.42,60.83,61.92,null,null],
            current_week: 4
          },
          konstruksi: null, sk_sectors: null
        }
      ]
    }
  ]
};

/* ── ② CHART REGISTRY ────────────────────────────────────── */
const charts = {};

function destroyChart(id) {
  if (charts[id]) { charts[id].destroy(); delete charts[id]; }
}

/* ── ③ HELPERS ────────────────────────────────────────────── */
const fmt = {
  pct:  v => v == null ? '—' : v.toFixed(2) + '%',
  pctS: v => v == null ? '—' : (v > 0 ? '+' : '') + v.toFixed(2) + '%',
  num:  v => v == null ? '—' : v.toLocaleString('id-ID'),
  rp:   v => 'Rp ' + (v / 1e9).toFixed(2) + ' M',
  jam:  v => (v / 1000).toFixed(1) + 'K Jam',
};

function statusClass(s) {
  if (!s) return 'badge-proses';
  const u = s.toUpperCase();
  if (u === 'CLOSED' || u === 'CLOSE') return 'badge-close';
  if (u === 'OPEN' || u === 'ON PROCESS' || u === 'ON PROGRESS') return 'badge-open';
  return 'badge-proses';
}

function devClass(v) {
  if (v == null) return '';
  return v >= 0 ? 'positive' : 'negative';
}

function devProg(v) {
  if (v == null) return '';
  return v >= 0 ? 'pos' : 'neg';
}

function barActualClass(v) {
  if (v == null) return '';
  return v >= 0 ? 'ahead' : 'behind';
}

/* ── ④ CHART FACTORY ─────────────────────────────────────── */
function getChartDefaults() {
  const root = getComputedStyle(document.documentElement);
  const textColor = root.getPropertyValue('--text-primary').trim() || '#f0f4ff';
  const mutedColor = root.getPropertyValue('--text-muted').trim() || '#4a5878';
  const gridColor = root.getPropertyValue('--border').trim() || 'rgba(255,255,255,0.04)';

  return {
    plugins: { legend: { display: false }, tooltip: {
      backgroundColor: 'rgba(10,15,28,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      borderWidth: 1,
      titleColor: '#f0f4ff',
      bodyColor: '#8898b8',
      padding: 10,
      callbacks: {
        label: ctx => ' ' + ctx.dataset.label + ': ' + (ctx.parsed.y != null ? ctx.parsed.y.toFixed(2) + '%' : '—')
      }
    }},
    scales: {
      x: {
        grid: { color: gridColor },
        ticks: { color: mutedColor, font: { size: 10, family: 'Inter' } }
      },
      y: {
        grid: { color: gridColor },
        ticks: { color: mutedColor, font: { size: 10, family: 'Inter' }, callback: v => v + '%' },
        min: 0, max: 100
      }
    },
    animation: { duration: 700, easing: 'easeOutQuart' },
    responsive: true,
    maintainAspectRatio: false,
    interaction: { mode: 'index', intersect: false }
  };
}

function buildSCurveChart(canvasId, weekly) {
  destroyChart(canvasId);
  const el = document.getElementById(canvasId);
  if (!el || !weekly) return;

  const planLength = weekly.plan.length;
  const projectionData = Array(planLength).fill(null);
  
  // Find last actual non-null point
  let lastActualIdx = -1;
  for (let i = planLength - 1; i >= 0; i--) {
    if (weekly.actual[i] != null) {
      lastActualIdx = i;
      break;
    }
  }

  if (lastActualIdx >= 0 && lastActualIdx < planLength - 1) {
    projectionData[lastActualIdx] = weekly.actual[lastActualIdx];
    projectionData[planLength - 1] = weekly.plan[planLength - 1]; // Project to final plan target
  }

  charts[canvasId] = new Chart(el, {
    type: 'line',
    data: {
      labels: weekly.labels,
      datasets: [
        {
          label: 'Kum. Rencana',
          data: weekly.plan,
          borderColor: 'rgba(255,255,255,0.3)',
          backgroundColor: 'transparent',
          borderWidth: 2,
          borderDash: [5, 4],
          pointRadius: 3,
          pointBackgroundColor: 'rgba(255,255,255,0.3)',
          tension: 0.35,
          spanGaps: false
        },
        {
          label: 'Kum. Aktual',
          data: weekly.actual,
          borderColor: '#1e6aff',
          backgroundColor: 'rgba(30,106,255,0.08)',
          fill: true,
          borderWidth: 2.5,
          pointRadius: 4,
          pointBackgroundColor: '#1e6aff',
          pointHoverRadius: 6,
          tension: 0.35,
          spanGaps: false
        },
        {
          label: 'Proyeksi',
          data: projectionData,
          borderColor: '#ffc107',
          backgroundColor: 'transparent',
          borderWidth: 2,
          borderDash: [4, 4],
          pointRadius: 3,
          pointBackgroundColor: '#ffc107',
          tension: 0,
          spanGaps: true
        }
      ]
    },
    options: { ...getChartDefaults() }
  });
}

function buildBarChart(canvasId, konstruksi) {
  destroyChart(canvasId);
  const el = document.getElementById(canvasId);
  if (!el || !konstruksi || konstruksi.length === 0) return;
  charts[canvasId] = new Chart(el, {
    type: 'bar',
    data: {
      labels: konstruksi.map(k => k.item),
      datasets: [{
        label: '% Terpasang',
        data: konstruksi.map(k => k.pct),
        backgroundColor: konstruksi.map(k =>
          k.pct >= 80 ? 'rgba(56,239,125,0.7)' :
          k.pct >= 50 ? 'rgba(30,106,255,0.7)' :
          'rgba(255,210,0,0.7)'
        ),
        borderRadius: 6,
        borderSkipped: false
      }]
    },
    options: {
      ...getChartDefaults(),
      scales: {
        ...getChartDefaults().scales,
        y: { ...getChartDefaults().scales.y, max: 105 }
      },
      plugins: {
        ...getChartDefaults().plugins,
        legend: { display: false },
        tooltip: {
          ...getChartDefaults().plugins.tooltip,
          callbacks: {
            label: ctx => ' ' + ctx.parsed.y.toFixed(2) + '%'
          }
        }
      }
    }
  });
}

function buildDonutChart(canvasId, pkg) {
  destroyChart(canvasId);
  const el = document.getElementById(canvasId);
  if (!el || pkg.progress_actual == null) return;
  const actual = pkg.progress_actual;
  const remaining = 100 - actual;
  charts[canvasId] = new Chart(el, {
    type: 'doughnut',
    data: {
      labels: ['Realisasi', 'Sisa'],
      datasets: [{
        data: [actual, remaining],
        backgroundColor: [pkg.color || '#1e6aff', 'rgba(255,255,255,0.05)'],
        borderWidth: 0,
        hoverOffset: 6
      }]
    },
    options: {
      cutout: '72%',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: { label: ctx => ' ' + ctx.parsed.toFixed(2) + '%' }
        }
      },
      animation: { duration: 900, easing: 'easeOutQuart' }
    }
  });
}

function buildOverviewDonut(canvasId, locations) {
  destroyChart(canvasId);
  const el = document.getElementById(canvasId);
  if (!el) return;
  const locs = locations.filter(l => l.progress_actual != null);
  if (!locs.length) return;
  const COLORS = ['#1e6aff','#00c6ff','#38ef7d','#ffd200','#f97316','#a855f7','#14b8a6','#ff4e50'];
  charts[canvasId] = new Chart(el, {
    type: 'doughnut',
    data: {
      labels: locs.map(l => l.name),
      datasets: [{
        data: locs.map(l => l.progress_actual),
        backgroundColor: COLORS.slice(0, locs.length),
        borderWidth: 2,
        borderColor: '#0d1320',
        hoverOffset: 8
      }]
    },
    options: {
      cutout: '60%',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'right',
          labels: { color: '#8898b8', font: { size: 10 }, boxWidth: 10, padding: 10 }
        },
        tooltip: {
          callbacks: {
            label: ctx => ' ' + ctx.label + ': ' + ctx.parsed.toFixed(2) + '%'
          }
        }
      },
      animation: { duration: 900 }
    }
  });
}

/* ── ⑤ PAGE RENDERERS ────────────────────────────────────── */

/* ─── Overview Page ──────────────────────────────────────── */
function renderOverview() {
  const p = DATA.project;
  const pkgs = DATA.packages;

  // Compute aggregates for packages that have progress
  const withProgress = pkgs.filter(pkg => pkg.progress_actual != null);
  const allLocs = pkgs.flatMap(pkg => pkg.locations).filter(l => l.progress_actual != null);
  const behind = allLocs.filter(l => l.progress_dev < 0).length;
  const ahead  = allLocs.filter(l => l.progress_dev >= 0).length;

  const html = `
<div class="section-header">
  <div>
    <h2>📊 Ringkasan Gabungan – Dashboard Weekly Ke-11</h2>
    <p>${p.period} &nbsp;|&nbsp; ${p.ministry}</p>
  </div>
  <div class="topbar-chips">
    <span class="chip chip-blue">⏱ Minggu 11 / 34</span>
    <span class="chip chip-yellow">Status: ${p.status_date}</span>
  </div>
</div>

<div class="kpi-grid">
  <div class="kpi-card" data-variant="blue">
    <span class="kpi-icon">🏘️</span>
    <div class="kpi-value">${fmt.num(p.total_sr)}</div>
    <div class="kpi-label">Total SR Keseluruhan</div>
  </div>
  <div class="kpi-card" data-variant="green">
    <span class="kpi-icon">🏗️</span>
    <div class="kpi-value">5</div>
    <div class="kpi-label">Paket Pekerjaan Aktif</div>
  </div>
  <div class="kpi-card" data-variant="teal">
    <span class="kpi-icon">🛡️</span>
    <div class="kpi-value">${(p.total_safe_hours/1e6).toFixed(2)}M</div>
    <div class="kpi-label">Total Jam Kerja Aman</div>
  </div>
  <div class="kpi-card" data-variant="yellow">
    <span class="kpi-icon">👷</span>
    <div class="kpi-value">${ahead}</div>
    <div class="kpi-label">Lokasi Ahead of Schedule</div>
    <div class="kpi-delta positive">▲ dari ${allLocs.length} lokasi terdata</div>
  </div>
  <div class="kpi-card" data-variant="red">
    <span class="kpi-icon">⚠️</span>
    <div class="kpi-value">${behind}</div>
    <div class="kpi-label">Lokasi Behind Schedule</div>
    <div class="kpi-delta negative">▼ perlu perhatian</div>
  </div>
  <div class="kpi-card" data-variant="purple">
    <span class="kpi-icon">💰</span>
    <div class="kpi-value">Rp 962M</div>
    <div class="kpi-label">Total Nilai Kontrak</div>
  </div>
</div>

<h3 style="font-size:14px;font-weight:700;color:var(--text-primary);margin-bottom:14px">Ringkasan Progres per Paket</h3>
<div class="packages-overview">
${pkgs.map(pkg => `
  <div class="pkg-overview-card" style="--pkg-grad:${pkg.grad}" onclick="navigateTo('pkg${pkg.id}')">
    <div class="pkg-num">Paket ${pkg.id}</div>
    <h3 style="color:${pkg.color}">${pkg.short}</h3>
    <div class="pkg-sr">📍 ${fmt.num(pkg.total_sr)} SR &nbsp;·&nbsp; ${pkg.epc}</div>
    ${pkg.progress_actual != null ? `
    <div class="pkg-progress-group">
      <div class="pkg-progress-item">
        <span class="label text-muted">Rencana</span>
        <div class="bar"><div class="bar-fill plan" style="width:${pkg.progress_plan}%"></div></div>
        <span class="val text-muted">${fmt.pct(pkg.progress_plan)}</span>
      </div>
      <div class="pkg-progress-item">
        <span class="label" style="color:${pkg.color}">Aktual</span>
        <div class="bar"><div class="bar-fill actual ${pkg.progress_dev >= 0 ? 'ahead' : 'behind'}" style="width:${Math.min(pkg.progress_actual,100)}%"></div></div>
        <span class="val" style="color:${pkg.color}">${fmt.pct(pkg.progress_actual)}</span>
      </div>
      <div class="pkg-progress-item">
        <span class="label text-muted">Deviasi</span>
        <div class="bar" style="background:transparent"></div>
        <span class="val ${devClass(pkg.progress_dev)}">${fmt.pctS(pkg.progress_dev)}</span>
      </div>
    </div>` : `
    <div style="margin-top:12px;font-size:11px;color:var(--text-muted);font-style:italic">
      Progress data tersedia per lokasi → klik untuk detail
    </div>`}
    <div style="display:flex;justify-content:space-between;align-items:center;margin-top:14px;padding-top:10px;border-top:1px solid var(--border)">
      <span style="font-size:10.5px;color:var(--text-muted)">🛡️ ${fmt.jam(pkg.safe_hours)}</span>
      ${pkg.tkdn ? `<span style="font-size:10.5px;color:var(--text-muted)">TKDN ${pkg.tkdn}%</span>` : ''}
      <span style="font-size:10.5px;color:${pkg.color}">Detail →</span>
    </div>
  </div>`).join('')}
</div>

<div class="content-grid" style="margin-top:8px">
  <div class="card">
    <div class="card-header">
      <div><h3>Progress Kumulatif per Lokasi (W11)</h3><p>Hanya lokasi dengan data progress tersedia</p></div>
    </div>
    <div class="card-body">
      ${allLocs.map(loc => {
        const devCls = barActualClass(loc.progress_dev);
        return `
        <div class="prog-row">
          <div class="prog-meta">
            <span class="prog-label">${loc.name}</span>
            <div class="prog-values">
              <span class="prog-plan">${fmt.pct(loc.progress_plan)}</span>
              <span class="prog-actual" style="color:${devCls === 'ahead' ? 'var(--green)' : devCls === 'behind' ? 'var(--red)' : 'var(--blue)'}">${fmt.pct(loc.progress_actual)}</span>
              <span class="prog-dev ${devProg(loc.progress_dev)}">${fmt.pctS(loc.progress_dev)}</span>
            </div>
          </div>
          <div class="prog-bar-track">
            <div class="prog-bar-plan" style="width:${loc.progress_plan}%"></div>
            <div class="prog-bar-actual ${devCls}" style="width:${Math.min(loc.progress_actual,100)}%"></div>
          </div>
        </div>`;
      }).join('')}
    </div>
  </div>

  <div class="card">
    <div class="card-header"><div><h3>Safety – Total Jam Kerja Aman</h3><p>Per paket pekerjaan</p></div></div>
    <div class="card-body">
      ${pkgs.map(pkg => `
      <div class="prog-row">
        <div class="prog-meta">
          <span class="prog-label">${pkg.short}</span>
          <span style="font-size:12px;font-weight:700;color:var(--green);font-family:'JetBrains Mono',monospace">${fmt.num(pkg.safe_hours)} Jam</span>
        </div>
        <div class="prog-bar-track">
          <div class="prog-bar-actual ahead" style="width:${(pkg.safe_hours/610436)*100}%"></div>
        </div>
      </div>`).join('')}
      <div style="margin-top:16px;padding-top:12px;border-top:1px solid var(--border);display:flex;justify-content:space-between;align-items:center">
        <span style="font-size:12px;color:var(--text-secondary);font-weight:600">Grand Total</span>
        <span style="font-size:18px;font-weight:800;color:var(--green);font-family:'JetBrains Mono',monospace">${fmt.num(p.total_safe_hours)} Jam</span>
      </div>
    </div>
  </div>
</div>`;

  document.getElementById('page-overview').innerHTML = html;
}

/* ─── Package Page ───────────────────────────────────────── */
function renderPackagePage(pkgIdx) {
  const pkg = DATA.packages[pkgIdx];
  const container = document.getElementById('package-content');
  if (!container) return;

  // Aggregated workforce
  const allWorkers = {};
  pkg.locations.forEach(loc => {
    if (loc.workforce) {
      Object.entries(loc.workforce).forEach(([role, count]) => {
        allWorkers[role] = (allWorkers[role] || 0) + count;
      });
    }
  });
  const totalWorkers = Object.values(allWorkers).reduce((a, b) => a + b, 0);

  // Count permits
  let totalPermits = 0, closedPermits = 0;
  pkg.locations.forEach(loc => {
    if (loc.permits) {
      totalPermits += loc.permits.length;
      closedPermits += loc.permits.filter(p => {
        const u = (p.status || '').toUpperCase();
        return u === 'CLOSED' || u === 'CLOSE';
      }).length;
    }
  });

  const html = `
<div class="section-header">
  <div>
    <h2 style="color:${pkg.color}">📦 Paket ${pkg.id} – ${pkg.name}</h2>
    <p>${pkg.epc} &nbsp;|&nbsp; PMC: ${pkg.pmc}</p>
  </div>
  <div class="topbar-chips">
    <span class="chip chip-blue">📅 ${pkg.duration}</span>
    <span class="chip chip-yellow">🚀 Start: ${pkg.start}</span>
    ${pkg.tkdn ? `<span class="chip chip-green">🏭 TKDN ${pkg.tkdn}%</span>` : ''}
  </div>
</div>

<div class="kpi-grid">
  <div class="kpi-card" data-variant="blue">
    <span class="kpi-icon">🏘️</span>
    <div class="kpi-value">${fmt.num(pkg.total_sr)}</div>
    <div class="kpi-label">Total SR</div>
  </div>
  <div class="kpi-card" data-variant="teal">
    <span class="kpi-icon">💰</span>
    <div class="kpi-value">${fmt.rp(pkg.contract_value)}</div>
    <div class="kpi-label">Nilai Kontrak</div>
  </div>
  <div class="kpi-card" data-variant="green">
    <span class="kpi-icon">🛡️</span>
    <div class="kpi-value">${fmt.jam(pkg.safe_hours)}</div>
    <div class="kpi-label">Jam Kerja Aman</div>
  </div>
  <div class="kpi-card" data-variant="yellow">
    <span class="kpi-icon">👷</span>
    <div class="kpi-value">${fmt.num(totalWorkers)}</div>
    <div class="kpi-label">Total Manpower</div>
  </div>
  ${pkg.progress_actual != null ? `
  <div class="kpi-card" data-variant="${pkg.progress_dev >= 0 ? 'green' : 'red'}">
    <span class="kpi-icon">${pkg.progress_dev >= 0 ? '📈' : '📉'}</span>
    <div class="kpi-value" style="color:${pkg.progress_dev >= 0 ? 'var(--green)' : 'var(--red)'}">${fmt.pct(pkg.progress_actual)}</div>
    <div class="kpi-label">Progress Aktual</div>
    <div class="kpi-delta ${pkg.progress_dev >= 0 ? 'positive' : 'negative'}">${fmt.pctS(pkg.progress_dev)} vs rencana</div>
  </div>` : ''}
  <div class="kpi-card" data-variant="purple">
    <span class="kpi-icon">📋</span>
    <div class="kpi-value">${closedPermits}/${totalPermits}</div>
    <div class="kpi-label">Izin Selesai</div>
  </div>
</div>

<div class="content-grid" style="margin-bottom:20px">
  ${(pkg.weekly_combined || pkg.progress_actual != null) ? `
  <div class="card">
    <div class="card-header">
      <div><h3>S-Curve Progress – Paket ${pkg.id}</h3><p>Kumulatif Rencana vs Aktual</p></div>
      <div class="scurve-legend">
        <div class="legend-item"><div class="legend-dot" style="background:rgba(255,255,255,0.3)"></div>Rencana</div>
        <div class="legend-item"><div class="legend-dot" style="background:${pkg.color}"></div>Aktual</div>
      </div>
    </div>
    <div class="card-body"><div class="chart-wrap"><canvas id="scurve-pkg-${pkg.id}"></canvas></div></div>
  </div>` : ''}

  ${pkg.progress_actual != null ? `
  <div class="card">
    <div class="card-header"><div><h3>Donut Progress</h3><p>Realisasi Paket ${pkg.id}</p></div></div>
    <div class="card-body" style="display:flex;justify-content:center;align-items:center">
      <div style="width:200px;height:200px;position:relative">
        <canvas id="donut-pkg-${pkg.id}"></canvas>
        <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center">
          <div style="font-size:22px;font-weight:800;color:${pkg.color}">${fmt.pct(pkg.progress_actual)}</div>
          <div style="font-size:10px;color:var(--text-muted)">Aktual</div>
        </div>
      </div>
    </div>
  </div>` : ''}
</div>

<h3 style="font-size:14px;font-weight:700;color:var(--text-primary);margin-bottom:14px">Progress per Lokasi</h3>
<div class="content-grid" style="margin-bottom:20px">
  <div class="card col-span-2">
    <div class="card-body">
      ${pkg.locations.map((loc, i) => {
        const hasProgress = loc.progress_actual != null;
        const devCls = hasProgress ? barActualClass(loc.progress_dev) : '';
        return `
        <div class="prog-row" style="cursor:pointer" onclick="navigateTo('loc-${pkg.id}-${i}')">
          <div class="prog-meta">
            <span class="prog-label">${loc.name} <span style="font-size:10px;color:var(--text-muted)">(${loc.code})</span></span>
            <div class="prog-values">
              ${hasProgress ? `
              <span class="prog-plan">${fmt.pct(loc.progress_plan)}</span>
              <span class="prog-actual" style="color:${devCls === 'ahead' ? 'var(--green)' : devCls === 'behind' ? 'var(--red)' : 'var(--blue)'}">${fmt.pct(loc.progress_actual)}</span>
              <span class="prog-dev ${devProg(loc.progress_dev)}">${fmt.pctS(loc.progress_dev)}</span>
              ` : `<span style="font-size:11px;color:var(--text-muted);font-style:italic">Data tersedia di halaman lokasi →</span>`}
            </div>
          </div>
          ${hasProgress ? `
          <div class="prog-bar-track">
            <div class="prog-bar-plan" style="width:${loc.progress_plan}%"></div>
            <div class="prog-bar-actual ${devCls}" style="width:${Math.min(loc.progress_actual,100)}%"></div>
          </div>` : `
          <div class="prog-bar-track">
            <div class="prog-bar-plan" style="width:0%"></div>
          </div>`}
        </div>`;
      }).join('')}
    </div>
  </div>
</div>

<h3 style="font-size:14px;font-weight:700;color:var(--text-primary);margin-bottom:14px">Tenaga Kerja Gabungan</h3>
<div class="content-grid" style="margin-bottom:20px">
  <div class="card col-span-2">
    <div class="card-body">
      <div class="worker-chips">
        ${Object.entries(allWorkers).map(([role, count]) => `
          <div class="worker-chip">
            <div class="role">${role}</div>
            <div class="count">${count}</div>
          </div>`).join('')}
      </div>
    </div>
  </div>
</div>

<h3 style="font-size:14px;font-weight:700;color:var(--text-primary);margin-bottom:14px">Status Perizinan per Lokasi</h3>
<div class="content-grid" style="margin-bottom:20px">
  ${pkg.locations.map((loc, i) => `
  <div class="card">
    <div class="card-header"><div><h3>${loc.name}</h3><p>${loc.permits ? loc.permits.length + ' izin' : 'Belum ada data'}</p></div></div>
    <div class="card-body">
      <div class="permit-grid">
        ${(loc.permits || []).map(p => `
          <div class="permit-item">
            <div class="permit-name">${p.name}</div>
            <div class="permit-statuses"><span class="badge ${statusClass(p.status)}">${p.status}</span></div>
          </div>`).join('')}
      </div>
    </div>
  </div>`).join('')}
</div>
`;

  container.innerHTML = html;

  // Render charts
  setTimeout(() => {
    const weekly = pkg.weekly_combined;
    if (weekly) {
      buildSCurveChart(`scurve-pkg-${pkg.id}`, weekly);
    }
    if (pkg.progress_actual != null) {
      buildDonutChart(`donut-pkg-${pkg.id}`, pkg);
    }
  }, 80);
}

function renderLocationPanel(loc, pkgId, locIdx) {
  const permitHtml = loc.permits && loc.permits.length > 0 ? loc.permits.map(p => `
    <div class="permit-item">
      <div class="permit-name">${p.name}</div>
      <div class="permit-statuses">
        <span class="badge ${statusClass(p.status)}">${p.status}</span>
      </div>
    </div>`).join('') : '<div class="text-muted" style="font-size:13px;font-style:italic">Data perizinan belum tersedia</div>';

  const workforceHtml = loc.workforce && Object.keys(loc.workforce).length > 0 ? Object.entries(loc.workforce).map(([role, count]) => `
    <div class="worker-chip">
      <div class="role">${role}</div>
      <div class="count">${count}</div>
    </div>`).join('') : '<div class="text-muted" style="font-size:13px;font-style:italic">Data tenaga kerja belum tersedia</div>';

  const weeklyHtml = loc.weekly ? buildWeeklyTableHtml(loc.weekly) : '<div class="text-muted" style="font-size:13px;font-style:italic;text-align:center;padding:20px 0">Data tabel progress mingguan belum tersedia</div>';

  const skHtml = loc.sk_sectors && loc.sk_sectors.length > 0 ? `
  <div class="card">
    <div class="card-header"><div><h3>Capel & SK Terpasang per RS</h3></div></div>
    <div class="card-body" style="padding:0">
      <table class="data-table">
        <thead><tr><th>Sektor</th><th style="text-align:right">BOQ</th><th style="text-align:right">SK Terpasang</th><th style="text-align:right">%</th></tr></thead>
        <tbody>
          ${loc.sk_sectors.map(s => {
            const pct = (s.terpasang / s.boq * 100).toFixed(1);
            return `<tr>
              <td>${s.name}</td>
              <td style="text-align:right" class="mono">${fmt.num(s.boq)}</td>
              <td style="text-align:right" class="mono highlight">${fmt.num(s.terpasang)}</td>
              <td style="text-align:right" class="mono ${pct >= 80 ? 'positive' : pct >= 50 ? '' : 'negative'}">${pct}%</td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>
  </div>` : `
  <div class="card">
    <div class="card-header"><div><h3>Capel & SK Terpasang per RS</h3></div></div>
    <div class="card-body" style="display:flex;align-items:center;justify-content:center;min-height:100px;">
      <div class="text-muted" style="font-size:13px;font-style:italic">Data Capel & SK Terpasang belum tersedia</div>
    </div>
  </div>`;

return `
<div id="loc-panel-${pkgId}-${locIdx}">
  <div class="kpi-grid" style="grid-template-columns:repeat(auto-fill,minmax(140px,1fr));margin-bottom:20px">
    <div class="kpi-card" data-variant="blue">
      <span class="kpi-icon">🏘️</span>
      <div class="kpi-value">${fmt.num(loc.sr)}</div>
      <div class="kpi-label">Jumlah SR</div>
    </div>
    <div class="kpi-card" data-variant="teal">
      <span class="kpi-icon">💰</span>
      <div class="kpi-value">${fmt.rp(loc.contract)}</div>
      <div class="kpi-label">Nilai Kontrak</div>
    </div>
    <div class="kpi-card" data-variant="green">
      <span class="kpi-icon">🛡️</span>
      <div class="kpi-value">${fmt.jam(loc.safe_hours)}</div>
      <div class="kpi-label">Jam Kerja Aman</div>
    </div>
    ${loc.progress_actual != null ? `
    <div class="kpi-card" data-variant="${loc.progress_dev >= 0 ? 'green' : 'red'}">
      <span class="kpi-icon">${loc.progress_dev >= 0 ? '📈' : '📉'}</span>
      <div class="kpi-value" style="color:${loc.progress_dev >= 0 ? 'var(--green)' : 'var(--red)'}">${fmt.pct(loc.progress_actual)}</div>
      <div class="kpi-label">Progress Aktual</div>
      <div class="kpi-delta ${loc.progress_dev >= 0 ? 'positive' : 'negative'}">${fmt.pctS(loc.progress_dev)} vs rencana</div>
    </div>` : `
    <div class="kpi-card" data-variant="purple">
      <span class="kpi-icon">📈</span>
      <div class="kpi-value" style="font-size:18px;line-height:28px;color:var(--text-muted)">—</div>
      <div class="kpi-label">Progress Aktual</div>
      <div class="kpi-delta" style="color:var(--text-muted)">Data belum tersedia</div>
    </div>`}
  </div>

  <div class="content-grid" style="margin-bottom:20px">
    <div class="card">
      <div class="card-header">
        <div><h3>S-Curve Progress – ${loc.name}</h3><p>Kumulatif Rencana vs Aktual</p></div>
        <div class="scurve-legend">
          <div class="legend-item"><div class="legend-dot" style="background:rgba(255,255,255,0.3)"></div>Rencana</div>
          <div class="legend-item"><div class="legend-dot" style="background:#1e6aff"></div>Aktual</div>
          <div class="legend-item"><div class="legend-dot" style="background:#ffc107"></div>Proyeksi</div>
        </div>
      </div>
      <div class="card-body">
        ${loc.weekly ? `<div class="chart-wrap"><canvas id="scurve-${pkgId}-${locIdx}"></canvas></div>` : `<div class="chart-wrap" style="display:flex;align-items:center;justify-content:center;color:var(--text-muted);font-size:13px;font-style:italic">Data S-Curve belum tersedia untuk lokasi ini</div>`}
      </div>
    </div>
    <div class="card">
      <div class="card-header"><div><h3>Konstruksi Terpasang (%)</h3><p>Per komponen</p></div></div>
      <div class="card-body">
        ${loc.konstruksi && loc.konstruksi.length > 0 ? `<div class="chart-wrap short"><canvas id="bar-${pkgId}-${locIdx}"></canvas></div>` : `<div class="chart-wrap short" style="display:flex;align-items:center;justify-content:center;color:var(--text-muted);font-size:13px;font-style:italic">Data Konstruksi belum tersedia</div>`}
      </div>
    </div>
  </div>

  <div class="content-grid" style="margin-bottom:20px">
    <div class="card col-span-2">
      <div class="card-header"><div><h3>Tabel Progress Mingguan</h3></div></div>
      <div class="card-body ${loc.weekly ? 'weekly-table-wrap' : ''}">${weeklyHtml}</div>
    </div>
  </div>

  <div class="content-grid" style="margin-bottom:20px">
    <div class="card">
      <div class="card-header"><div><h3>Status Perizinan</h3></div></div>
      <div class="card-body"><div class="permit-grid">${permitHtml}</div></div>
    </div>
    <div class="card">
      <div class="card-header"><div><h3>Tenaga Kerja</h3></div></div>
      <div class="card-body">
        <div class="worker-chips">${workforceHtml}</div>
      </div>
    </div>
  </div>

  ${skHtml}
</div>`;
}

function buildWeeklyTableHtml(weekly) {
  const cw = weekly.current_week - 1; // index of current week
  const headerCells = weekly.labels.map((l, i) =>
    `<th class="${i === cw ? 'current-week' : ''}">${l}</th>`).join('');

  const planRow = weekly.plan.map((v, i) =>
    `<td class="${i === cw ? 'current-week' : ''}">${v != null ? v.toFixed(3) + '%' : ''}</td>`).join('');
  const actualRow = weekly.actual.map((v, i) =>
    `<td class="${i === cw ? 'current-week' : ''}">${v != null ? v.toFixed(3) + '%' : ''}</td>`).join('');
  const devRow = weekly.labels.map((_, i) => {
    const a = weekly.actual[i];
    const pl = weekly.plan[i];
    if (a == null || pl == null) return `<td class="${i === cw ? 'current-week' : ''}"></td>`;
    const d = a - pl;
    return `<td class="${i === cw ? 'current-week ' : ''}${d >= 0 ? 'positive' : 'negative'}">${(d > 0 ? '+' : '') + d.toFixed(3)}%</td>`;
  }).join('');

  return `
<table class="weekly-table">
  <thead><tr><th></th>${headerCells}</tr></thead>
  <tbody>
    <tr class="kum-plan-row"><td>Kum. Rencana</td>${planRow}</tr>
    <tr class="kum-actual-row"><td>Kum. Aktual</td>${actualRow}</tr>
    <tr class="dev-row"><td>Deviasi</td>${devRow}</tr>
  </tbody>
</table>`;
}

function renderLocationCharts(loc, pkgId, locIdx) {
  if (loc.weekly) buildSCurveChart(`scurve-${pkgId}-${locIdx}`, loc.weekly);
  if (loc.konstruksi && loc.konstruksi.length > 0) buildBarChart(`bar-${pkgId}-${locIdx}`, loc.konstruksi);
}

/* ── ⑥ NAVIGATION ────────────────────────────────────────── */
let currentPage = 'overview';
let renderedPages = new Set();

function navigateTo(pageKey) {
  // Hide current
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));

  // Highlight nav button
  const navBtn = document.querySelector(`[data-page="${pageKey}"]`);
  if (navBtn) navBtn.classList.add('active');

  if (pageKey === 'overview') {
    const target = document.getElementById('page-overview');
    if (target) target.classList.add('active');
    
    if (!renderedPages.has(pageKey)) {
      renderedPages.add(pageKey);
      renderOverview();
    }
    
    document.getElementById('topbar-title').textContent = 'Dashboard Gabungan – W11';
    document.getElementById('topbar-sub').textContent = 'Ringkasan seluruh paket pekerjaan Jargas';
  } else if (pageKey.startsWith('pkg')) {
    const target = document.getElementById('page-package');
    if (target) target.classList.add('active');
    
    const pkgIdx = parseInt(pageKey.replace('pkg', '')) - 1;
    const pkg = DATA.packages[pkgIdx];
    renderPackagePage(pkgIdx);
    
    document.getElementById('topbar-title').textContent = `Paket ${pkg.id} – ${pkg.short}`;
    document.getElementById('topbar-sub').textContent = `${pkg.name} | ${pkg.epc} | ${fmt.num(pkg.total_sr)} SR`;
  } else if (pageKey.startsWith('loc-')) {
    const target = document.getElementById('page-location');
    if (target) target.classList.add('active');
    
    const parts = pageKey.split('-');
    const pkgId = parseInt(parts[1]);
    const locIdx = parseInt(parts[2]);
    renderLocationPage(pkgId, locIdx);
    
    const pkg = DATA.packages[pkgId - 1];
    const loc = pkg.locations[locIdx];
    document.getElementById('topbar-title').textContent = `${loc.name}`;
    document.getElementById('topbar-sub').textContent = `Paket ${pkg.id} – ${pkg.name} | ${fmt.num(loc.sr)} SR`;
  }

  currentPage = pageKey;
}

function renderLocationPage(pkgId, locIdx) {
  const pkg = DATA.packages[pkgId - 1];
  const loc = pkg.locations[locIdx];
  const container = document.getElementById('location-content');
  if (!container) return;

  container.innerHTML = renderLocationPanel(loc, pkgId, locIdx);

  // Render charts for this location
  setTimeout(() => renderLocationCharts(loc, pkgId, locIdx), 80);
}

/* ── ⑦ INIT ──────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  // Setup theme toggle
  const themeToggle = document.getElementById('theme-toggle');
  const themeLabel = document.getElementById('theme-label');
  
  const currentTheme = localStorage.getItem('mata-theme') || 'dark';
  if (currentTheme === 'light') {
    document.documentElement.setAttribute('data-theme', 'light');
    if(themeLabel) themeLabel.textContent = 'Dark Mode';
  }
  
  if (themeToggle) {
    themeToggle.addEventListener('click', () => {
      let theme = document.documentElement.getAttribute('data-theme');
      if (theme === 'light') {
        document.documentElement.removeAttribute('data-theme');
        localStorage.setItem('mata-theme', 'dark');
        if(themeLabel) themeLabel.textContent = 'Light Mode';
      } else {
        document.documentElement.setAttribute('data-theme', 'light');
        localStorage.setItem('mata-theme', 'light');
        if(themeLabel) themeLabel.textContent = 'Dark Mode';
      }
      
      // Re-render charts on theme change to update colors
      renderedPages.clear(); 
      navigateTo(currentPage);
    });
  }

  // Generate Dynamic Sidebar
  const navContainer = document.getElementById('dynamic-nav');
  if (navContainer && DATA && DATA.packages) {
    let navHtml = '';
    DATA.packages.forEach(pkg => {
      navHtml += `
        <button class="nav-item nav-pkg-btn" data-page="pkg${pkg.id}">
          <span class="nav-icon" style="font-size:11px">📦</span>
          <span>Paket ${pkg.id} – ${pkg.short || pkg.name.split(' ')[0]}</span>
        </button>`;
      pkg.locations.forEach((loc, i) => {
        navHtml += `
          <button class="nav-item nested" data-page="loc-${pkg.id}-${i}">
            <span>${loc.name}</span>
          </button>
        `;
      });
    });
    navContainer.innerHTML = navHtml;
  }

  // Setup nav items
  document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', () => navigateTo(item.dataset.page));
  });

  // Load overview immediately
  navigateTo('overview');
});
