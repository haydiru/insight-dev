# 🚀 Insight-Dev: Engineering Problem-Solving & LinkedIn Playbook

Repository ini berisi kumpulan studi kasus teknikal, arsitektur sistem, dan playbook konten LinkedIn yang diekstraksi dari pengalaman nyata dalam memecahkan masalah kompleks di bidang:
* **🤖 AI Automation & Agentic Workflow**
* **⚡ DevOps & Cloud Infrastructure**
* **🛠️ Senior Fullstack & Database Engineering**
* **📊 Data Engineering & Analytics**

---

## 📁 Struktur Repository

```text
├── AI_Automation_Fullstack_DevOps_Case_Studies.xlsx  # Master file Excel berisi 12 studi kasus & content calendar
├── generate_case_studies.py                         # Skrip generator openpyxl untuk memperbarui file Excel
└── README.md                                        # Dokumentasi overview dan ringkasan kasus
```

---

## 📑 Struktur Workbook Excel (`.xlsx`)

File **`AI_Automation_Fullstack_DevOps_Case_Studies.xlsx`** memiliki 3 sheet utama:

1. **`Studi Kasus & Problem Solving`**
   * Berisi 12 studi kasus teknikal nyata dengan analisis mendalam: *The Problem, Root Cause Analysis, How to Solve, Technical & Business Impact, Key Engineering Takeaways*, dan *Tech Stack*.
2. **`LinkedIn Content Plan`**
   * Rencana publikasi berbasis **STAR Framework** (*Situation, Task, Action, Result*), dilengkapi dengan *Catchy 3-line Hooks*, rekomendasi format visual (*Carousel, Storytelling, Architecture Diagram*), *Call to Action (CTA)*, dan *Hashtags*.
3. **`Architecture & Tech Stack`**
   * Peta komponen ekosistem enterprise (*Model Context Protocol, Cloudflare Zero Trust, n8n, MariaDB, Next.js, Microsoft Dynamics 365 BC*) serta pola desain arsitektur yang teruji di lingkungan produksi.

---

## 🌟 Top 12 Studi Kasus Engineering

| ID | Pilar / Domain | Judul Kasus / Problem Statement | Ringkasan Solusi Teknis |
|:---:|:---|:---|:---|
| **01** | AI Automation | Multi-Agent Staging Buffer vs Polusi Database Produksi | 2-Phase Staging Buffer Pattern (Draft -> Review & Deduplication -> Batch Promotion via MCP). |
| **02** | DevOps & Automation | Optimasi n8n: Dari 16 Menit ke 20 Detik (Fix HTTP 503 & 400) | Delta query SQL View, pemangkasan 15.400 baris ke 480 baris aktif, dan auto-retry guard. |
| **03** | Fullstack & Database | Menangani Urutan Audit Log Acak Akibat Multi-Approval Async | Multi-level coalesce timestamp `ORDER BY COALESCE(task_start_date, action_date), id`. |
| **04** | DevOps & Security | Akses Database Cloud Aman Tanpa Buka Port Publik (Cloudflare) | Cloudflare Access TCP Tunnel (Zero Trust), 0 open ports di firewall publik server. |
| **05** | Fullstack & Data | Timezone Hell UTC vs WIB: Misteri Turunnya Skor KPI Karyawan | Standarisasi MySQL DATETIME tanpa double offset timezone Asia/Jakarta di API contract layer. |
| **06** | Fullstack & System | The Backdating Dilemma: Audit Timestamp vs Business Date | Pemisahan kolom `completedAt` faktual dengan `activity_log` audit trail di skema PMS. |
| **07** | Fullstack & ERP | Hybrid Caching: Mencegah Delay Scan Kedatangan Gudang (BC 365) | Read-Optimized staging cache (<150ms) + on-demand live ERP verification saat submit final. |
| **08** | Data Engineering | Zero Data Drift: Multi-Periodisitas KPI (Bulan vs Kuartal) | Dynamic weight normalization engine di SQL View (Migration 018) teruji 4.437 baris deviasi 0. |
| **09** | Fullstack & Security | Dynamic Cross-Org RBAC: Akses Lintas Divisi Tanpa Hardcoding | Tabel delegasi independen dengan dynamic security subquery di core master view. |
| **10** | Fullstack & Logistics | Filter Cerdas PO Fisik vs Non-Fisik pada Modul Gate-In Gudang | Klasifikasi domain fisik (Item/Fixed Asset) vs non-fisik (GL Account/Ongkir/Jasa Pasang). |
| **11** | AI Automation & LLM | Membangun Custom MCP Server untuk Mengontrol ERP & PMS | Implementasi 25+ callable tool schemas JSON Schema untuk autonomous agent execution. |
| **12** | Data Analytics | Algoritma Cap 120% Non-Linier: Mencegah Manipulasi Skor Tim | Formula matematis `min(1.20, Actual/Target) * Weight` dengan konversi proporsional 1-5 di akhir. |

---

## 🛠️ Cara Menjalankan Ulang Generator

Pastikan Python dan modul `openpyxl` terpasang:

```bash
pip install openpyxl
python generate_case_studies.py
```

---

*Author: Haidir Magribi ([@haydiru](https://github.com/haydiru))*
