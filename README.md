# 🚀 Insight-Dev: Engineering Problem-Solving & Error-Handling Playbook

Repository ini berisi kumpulan studi kasus teknikal, arsitektur sistem, dan playbook konten LinkedIn yang diekstraksi dari pengalaman nyata dalam memecahkan masalah kompleks di bidang:
* **🤖 AI Automation & Agentic Workflow**
* **⚡ DevOps & Cloud Infrastructure**
* **🛠️ Senior Fullstack & Database Engineering**
* **📊 Data Engineering & Analytics**

---

## 📁 Struktur Repository

```text
├── AI_Automation_Error_Handling_Case_Studies.xlsx    # [NEW] Master Excel penanganan error spesifik per kasus ("Behind The Bug")
├── AI_Automation_Fullstack_DevOps_Case_Studies.xlsx  # Master Excel 12 studi kasus komprehensif & content calendar umum
├── generate_case_studies.py                         # Skrip generator openpyxl untuk memperbarui file Excel
└── README.md                                        # Dokumentasi overview dan ringkasan kasus
```

---

## 📑 Struktur Workbook Baru: `AI_Automation_Error_Handling_Case_Studies.xlsx`

File ini dirancang khusus untuk membedah insiden teknis **per error secara terpisah dan mendalam**:

1. **`Penanganan Error Spesifik`**
   * Membedah 12 error teknis konkret: *Kode & Nama Error, Gejala & Pesan Error Nyata, Langkah Diagnosa/Debugging, Akar Masalah (Root Cause), Solusi Konkret & Perbaikan Kode/Query, Hasil Terukur, dan Key Engineering Takeaway*.
2. **`LinkedIn Story Per Error` ("Behind The Bug")**
   * Format storytelling LinkedIn siap posting per error: *Opening Hook 3 baris, The Error & Symptom, The 'Aha!' Moment, Solusi Teknis Ringkas, Call to Action, dan Hashtags*.
3. **`Arsitektur Pencegahan Error`**
   * Panduan arsitektur pencegahan (*Zero Trust, Delta Filtering at Source, Idempotent Normalization, Hybrid Caching, Strict JSON Schema*).

---

## 🔍 Ringkasan 12 Error Spesifik yang Dibedah

| No | Kode / Nama Error Spesifik | Domain | Akar Masalah (Root Cause) | Solusi & Hasil Konkret |
|:---:|:---|:---|:---|:---|
| **01** | **HTTP 503 Gateway Timeout** | DevOps & Automation | Query n8n memuat 15.400 baris histori tak berindeks dalam satu payload. | Refactoring SQL view ke ~480 baris aktif; runtime turun 98% (16 min ➡️ 20 sec). |
| **02** | **HTTP 400 Grid Limits (10M Cells)** | DevOps & Automation | Append baris raw log approval terus-menerus ke Google Sheets setiap 5 menit. | Migrasi ke Single Row State Update + Pre-agregasi SQL; sheet ramping (<5.000 baris). |
| **03** | **Double-Offset Timezone (+7h)** | Fullstack & Data | `toISOString()` mengubah WIB ke UTC (+0) lalu disimpan ke MySQL DATETIME non-tz. | Standardisasi string `YYYY-MM-DD HH:mm:ss` tanpa double offset; 100% false late hilang. |
| **04** | **Async Sorting Reversal** | Fullstack & Database | Auto-increment ID merekam waktu webhook selesai, bukan waktu tanda tangan user. | Normalisasi view: `ORDER BY COALESCE(task_start_date, action_date), id` (5.367 baris rapi). |
| **05** | **ECONNREFUSED Port 3306** | DevOps & Cloud Security | Firewall cloud server memblokir akses inbound port 3306 publik demi keamanan. | Cloudflare Access TCP Tunnel via mTLS; 0 open port di firewall publik, latensi <40ms. |
| **06** | **False Late: Click vs Real Date** | Fullstack & System Design | `completed_at` diisi `NOW()` saat tombol diklik, bukan tanggal koding/deploy selesai. | Pemisahan kolom `completedAt` faktual dengan audit log activity; skor pulih ke 5.00/5.00. |
| **07** | **PO Stuck: Non-Physical Lines** | Fullstack & Logistics | API ERP menyatukan baris Item fisik dengan GL Account (ongkir) dan Jasa. | Domain filter pipeline: pisahkan item fisik (wajib scan) vs jasa (auto-exclude gate-in). |
| **08** | **Form 403: Periodicity Dirty State** | Fullstack & Web UI | Detector `checkDirtyState()` hanya mengecek nilai numerik target, bukan periodisitas. | Refactoring form state evaluator mencakup `Tipe_Periode_Indikator`; simpan instan. |
| **09** | **Bobot 160% & Data Drift** | Data Engineering & SQL | SUM bobot statis mencampur indikator Bulanan dengan Kuartalan di master view. | Window function: `Weight / SUM(Weight) OVER (PARTITION BY Pegawai, Periode)` (0 drift). |
| **10** | **HTTP 429 Too Many Requests** | DevOps & Automation | Looping n8n menembak Google Sheets API baris demi baris melebihi kuota 60 req/min. | Mengganti loop dengan Batch Payload Update ber-delay jeda 1.000ms (1 request per run). |
| **11** | **MCP Argument Rejection** | AI Automation & LLM | LLM mengasumsikan parameter generik CLI yang ditolak skema ketat tool MCP. | Strict JSON Schema dengan parameter hint eksplisit; eksekusi agent otonom 100% mulus. |
| **12** | **300% Cap Bias (Goodhart's Law)** | Data Science & Analytics | Rumus linier mentah membuat capaian 300% di task mudah menutupi task penting. | Formula Cap 120%: `min(1.20, Actual/Target) * Bobot` dengan konversi proporsional di akhir. |

---

## 🛠️ Menjalankan Generator

```bash
pip install openpyxl
python generate_case_studies.py
```

---

*Author: Haidir Magribi ([@haydiru](https://github.com/haydiru))*
