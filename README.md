# 🚀 Insight-Dev: Master Engineering Playbook & LinkedIn Storytelling

Repository ini berisi **24 studi kasus engineering nyata**, terbagi seimbang antara **Penanganan Error Spesifik (Behind The Bug)** dan **Inovasi Arsitektur, Improvement Sistem & Tips-Trik Teknis**, yang diekstraksi dari pengalaman nyata dalam membangun dan mengoptimalkan ekosistem:
* **🤖 AI Automation & Agentic Workflow (MCP & Multi-Agent)**
* **⚡ DevOps & Cloud Infrastructure (n8n, Docker, Cloudflare Zero Trust)**
* **🛠️ Senior Fullstack & Database Engineering (Next.js, React, MySQL/MariaDB)**
* **📊 Enterprise ERP & Data Analytics (Microsoft Dynamics 365 BC, Goodhart's Law Cap)**

---

## 📁 Struktur File di Repository

```text
├── Master_Engineering_Playbook_Case_Studies.xlsx    # [MASTER WORKBOOK] 24 Kasus Lengkap (5 Sheet Komprehensif)
├── AI_Automation_Error_Handling_Case_Studies.xlsx    # Workbook Khusus 12 Error Spesifik & Root Cause Fix
├── AI_Automation_Fullstack_DevOps_Case_Studies.xlsx  # Workbook Awal (Overview Kasus Umum)
├── generate_master_case_studies.py                  # Skrip generator openpyxl untuk Master Workbook
├── generate_case_studies.py                         # Skrip generator openpyxl untuk Error Workbook
└── README.md                                        # Dokumentasi lengkap & ringkasan 24 kasus
```

---

## 📑 5 Sheet di Dalam `Master_Engineering_Playbook_Case_Studies.xlsx`

1. **`Katalog Lengkap (24 Kasus)`**
   * Direktori terpadu 24 kasus engineering dengan label kategori: *Penanganan Error, Improvement Sistem, Strategi Arsitektur, dan Tips & Trik Teknis*.
2. **`Kasus Error Spesifik` (12 Kasus Error)**
   * Bedah tuntas error konkret: *Kode Error, Gejala & Pesan Error Asli, Cara Tracing Log/Debugging, Akar Masalah Teknis, dan Solusi Perbaikan Kode/Query*.
3. **`Improvement & Strategi` (12 Kasus Non-Error)**
   * Bedah inovasi non-error: *Tantangan Bisnis, Konsep Desain Arsitektur, Langkah Implementasi, dan Efisiensi Terukur*.
4. **`LinkedIn Content Playbook` (24 Postingan Lengkap)**
   * 24 konten siap posting ke LinkedIn berbasis **STAR Framework** (*Situation, Task, Action, Result*), dilengkapi dengan *3-line Hook, Visual Format, CTA, dan Hashtags*.
5. **`Arsitektur & Best Practices`**
   * Peta ekosistem enterprise, prinsip Zero Trust, pola Staging Buffer, Normalisasi SQL, dan Hybrid Caching.

---

## 🌟 Daftar 24 Kasus Engineering Lengkap

### Bagian A: 12 Kasus Penanganan Error Spesifik (Behind The Bug)

| No | Kode / Nama Error | Domain | Akar Masalah (Root Cause) | Solusi & Hasil Konkret |
|:---:|:---|:---|:---|:---|
| **01** | **HTTP 503 Gateway Timeout** | DevOps & Automation | Query n8n memuat 15.400 baris histori tak berindeks dalam satu payload. | Refactoring SQL view ke ~480 baris aktif; runtime drop 98% (16 min ➡️ 20 sec). |
| **02** | **HTTP 400 Grid Limits (10M Cells)** | DevOps & Automation | Append baris raw log approval terus-menerus ke Google Sheets setiap 5 menit. | Migrasi ke Single Row State Update + Pre-agregasi SQL; sheet ramping (<5.000 baris). |
| **03** | **Double-Offset Timezone (+7h)** | Fullstack & Data | `toISOString()` mengubah WIB ke UTC (+0) lalu disimpan ke MySQL DATETIME non-tz. | Standardisasi string `YYYY-MM-DD HH:mm:ss` tanpa double offset; 100% false late hilang. |
| **04** | **Async Sorting Reversal** | Fullstack & Database | Auto-increment ID merekam waktu webhook selesai, bukan waktu tanda tangan user. | Normalisasi view: `ORDER BY COALESCE(task_start_date, action_date), id` (5.367 baris rapi). |
| **05** | **ECONNREFUSED Port 3306** | DevOps & Security | Firewall cloud server memblokir akses inbound port 3306 publik demi keamanan. | Cloudflare Access TCP Tunnel via mTLS; 0 open port di firewall publik, latensi <40ms. |
| **06** | **False Late: Click vs Real Date** | Fullstack & System | `completed_at` diisi `NOW()` saat tombol diklik, bukan tanggal koding/deploy selesai. | Pemisahan kolom `completedAt` faktual dengan audit log activity; skor pulih ke 5.00/5.00. |
| **07** | **PO Stuck: Non-Physical Lines** | Fullstack & Logistics | API ERP menyatukan baris Item fisik dengan GL Account (ongkir) dan Jasa. | Domain filter pipeline: pisahkan item fisik (wajib scan) vs jasa (auto-exclude gate-in). |
| **08** | **Form 403: Periodicity Dirty State** | Fullstack & Web UI | Detector `checkDirtyState()` hanya mengecek nilai numerik target, bukan periodisitas. | Refactoring form state evaluator mencakup `Tipe_Periode_Indikator`; simpan instan. |
| **09** | **Bobot 160% & Data Drift** | Data Engineering & SQL | SUM bobot statis mencampur indikator Bulanan dengan Kuartalan di master view. | Window function: `Weight / SUM(Weight) OVER (PARTITION BY Pegawai, Periode)` (0 drift). |
| **10** | **HTTP 429 Too Many Requests** | DevOps & Automation | Looping n8n menembak Google Sheets API baris demi baris melebihi kuota 60 req/min. | Mengganti loop dengan Batch Payload Update ber-delay jeda 1.000ms (1 request per run). |
| **11** | **MCP Argument Rejection** | AI Automation & LLM | LLM mengasumsikan parameter generik CLI yang ditolak skema ketat tool MCP. | Strict JSON Schema dengan parameter hint eksplisit; eksekusi agent otonom 100% mulus. |
| **12** | **300% Cap Bias (Goodhart's Law)** | Data Science & Analytics | Rumus linier mentah membuat capaian 300% di task mudah menutupi task penting. | Formula Cap 120%: `min(1.20, Actual/Target) * Bobot` dengan konversi proporsional di akhir. |

---

### Bagian B: 12 Kasus Improvement, Strategi & Tips-Trik Teknis (Non-Error)

| No | Kategori | Topik / Judul Kasus | Konsep Desain & Pola Arsitektur | Dampak & Efisiensi |
|:---:|:---|:---|:---|:---|
| **13** | **Strategi** | 2-Phase Staging Buffer Multi-Agent | Agent log draft ke buffer (`PENDING`), AI mereview & batch promote ke tabel produksi. | Zero task spam, integritas data proyek 100%. |
| **14** | **Improvement**| Hybrid Caching ERP Dynamics 365 | Read-Optimized local cache (<150ms) + on-demand live ERP verification saat submit. | Respon barcode turun 4s ➡️ <150ms; antrean truk hilang. |
| **15** | **Strategi** | Dynamic Cross-Org RBAC SQL Views | Delegasi akses via tabel independen & dynamic security subquery di master SQL view. | Konfigurasi visual di admin console tanpa deploy frontend. |
| **16** | **Tips & Trik**| Coalesce Multi-Tier Sorting Async | Formula `ORDER BY COALESCE(start_date, action_date, created_at), sequence_id`. | Urutan kronologis bisnis selalu stabil dan konsisten. |
| **17** | **Strategi** | Zero Trust TCP Tunneling DB Cloud | Daemon `cloudflared access tcp` memetakan port 3306 lokal ke cloud via mTLS. | 0 open port di firewall publik, aman dari bot internet. |
| **18** | **Tips & Trik**| Idempotent Weight Normalization | Window function: `Weight / SUM(Weight) OVER (PARTITION BY Pegawai, Periode)`. | Total bobot selalu 100% otomatis, zero data drift. |
| **19** | **Improvement**| Compute-at-Source vs Low-Code Mem | Pindahkan komputasi dan deduplikasi berat ke SQL View di level database. | Hemat memori container 85%, runtime turun 98%. |
| **20** | **Tips & Trik**| Dual-Date Architecture ERP/PMS | Kolom `completedAt` (business date) dipisah dari `activity_log` (audit date). | Memungkinkan backdating faktual tanpa langgar compliance. |
| **21** | **Strategi** | Custom MCP Server untuk Autonomous AI | Membangun 25+ tool RPC terstruktur dengan JSON schema validation untuk LLM. | AI bekerja otonom sebagai project assistant enterprise. |
| **22** | **Tips & Trik**| Domain-Driven Ingestion Filtering | Pisahkan tipe Item fisik dari GL Account / Jasa di tahap data pipeline awal. | Layar scanner gudang 100% bersih hanya barang fisik. |
| **23** | **Strategi** | Non-Linear Cap 120% vs Goodhart | Fungsi saturasi `min(1.20, Actual/Target)` dengan konversi proporsional di akhir. | Mencegah gaming-the-system, motivasi tim proporsional. |
| **24** | **Tips & Trik**| Multi-Field Dirty State di React | Deep-comparison evaluator pada seluruh atribut konfigurasi objek state formulir. | Tombol simpan selalu responsif terhadap perubahan metadata. |

---

## 🛠️ Menjalankan Generator

```bash
pip install openpyxl
python generate_master_case_studies.py
```

---

*Author: Haidir Magribi ([@haydiru](https://github.com/haydiru))*
