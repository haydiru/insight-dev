import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_master_excel():
    wb = openpyxl.Workbook()

    # Styling Palette
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid") # Deep Slate Navy
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    accent_fill_1 = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") # Very light slate
    accent_fill_2 = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid") # White

    # Category badges
    badge_error = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")       # Rose / Red tint (Error)
    badge_improve = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")     # Emerald / Green (Improvement)
    badge_strategy = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")    # Sky Blue (Strategi)
    badge_tips = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")        # Amber / Yellow (Tips & Trik)

    # Domain Badges
    badge_ai = PatternFill(start_color="EDE9FE", end_color="EDE9FE", fill_type="solid")          # Purple
    badge_devops = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")      # Yellow
    badge_fullstack = PatternFill(start_color="BAE6FD", end_color="BAE6FD", fill_type="solid")   # Light Sky
    badge_data = PatternFill(start_color="BBF7D0", end_color="BBF7D0", fill_type="solid")        # Light Emerald

    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    cell_font = Font(name="Calibri", size=10, color="0F172A")
    cell_bold = Font(name="Calibri", size=10, bold=True, color="0F172A")
    title_font = Font(name="Calibri", size=15, bold=True, color="0F172A")
    subtitle_font = Font(name="Calibri", size=10, italic=True, color="64748B")

    # =========================================================================
    # DATA DEFINITIONS: 24 KASUS LENGKAP (12 Error + 12 Non-Error)
    # =========================================================================

    # 12 KASUS ERROR SPESIFIK
    error_cases = [
        (
            1,
            "Penanganan Error",
            "DevOps & Automation",
            "HTTP 503 Service Unavailable / Gateway Timeout",
            "Workflow n8n sinkronisasi form FPPP (3pwAHMpguHsPHW8O) gantung selama 16 menit lalu crash dengan 503 Gateway Timeout.",
            "1. Cek log eksekusi n8n #953532.\n2. Cek slow query log MariaDB.\n3. Query get v_fppp_comments membaca 15.400 baris unindexed dalam 1 payload.",
            "n8n memuat 15.400 baris dalam satu payload memory node, memblokir event loop dan melewati batas timeout proxy Nginx/Cloudflare (300s).",
            "Refactoring SQL View v_fppp_comments di MariaDB: Filter status form aktif, memangkas dari 15.400 baris ke ~480 baris aktif, dan deploy versi aktif.",
            "• Runtime drop 98% (dari 16 menit jadi 20 detik).\n• Nol insiden timeout HTTP 503.\n• Penggunaan CPU container n8n turun drastis.",
            "Hindari full historical scan pada workflow automasi jika bisnis hanya membutuhkan delta status transaksi aktif.",
            "n8n, Docker, MariaDB, SQL Query Tuning, Cloudflare",
            "Saring di level SQL View sebelum data ditarik ke engine automasi."
        ),
        (
            2,
            "Penanganan Error",
            "DevOps & Automation",
            "HTTP 400 Grid Limits Exceeded (Google Sheets API)",
            "Error: 400 Bad Request: 'This action would increase the number of cells above the limit of 10,000,000 cells' pada node Google Sheets n8n.",
            "1. Cek parameter node append row di n8n.\n2. Hitung jumlah kolom x baris di spreadsheet.\n3. Ditemukan akumulasi baris tanpa pembersihan berbulan-bulan.",
            "Workflow lama melakukan append raw log per komentar approval tiap 5 menit, membuat sheet membengkak melebihi batas 10 juta sel.",
            "Ubah pola integrasi dari append raw log menjadi Single Row State Update (1 baris per form) dengan pre-agregasi komentar langsung di SQL view.",
            "• Ukuran spreadsheet stabil di bawah 5.000 baris.\n• Bebas dari error kuota 10 juta sel selamanya.\n• Spreadsheet dibuka super cepat tanpa lag.",
            "Google Sheets adalah alat visualisasi/kolaborasi, bukan database transaksional. Lakukan agregasi di database sebelum menulis ke Sheets.",
            "Google Sheets API, n8n, SQL Views, MariaDB",
            "Pola Single Row State Update + Pre-Aggregation SQL."
        ),
        (
            3,
            "Penanganan Error",
            "Fullstack & Data Integrity",
            "Double-Offset Timezone Shift (+7h / -7h Bug)",
            "Task yang selesai jam 16:30 WIB otomatis ditandai LATE dan tanggal jatuh tempo mundur 1 hari di dashboard pelaporan KPI dan Web-OTD.",
            "1. Inspect network payload: browser kirim '2026-10-15T09:30:00.000Z'.\n2. Cek kolom database: tersimpan '2026-10-15 09:30:00' non-timezone.\n3. Server anggap jam 09:30 pagi.",
            "Frontend Next.js mengeksekusi toISOString() yang memotong WIB ke UTC (+0). Backend menyimpannya ke MySQL DATETIME tanpa zona waktu, bergeser 7 jam.",
            "1. Hentikan toISOString() pada form submit.\n2. Kirim format string standar MySQL: YYYY-MM-DD HH:mm:ss berbasis Asia/Jakarta.\n3. Set session timezone MySQL ke '+07:00'.",
            "• 100% false-positive keterlambatan task hilang.\n• Skor on-time karyawan kembali akurat dan adil.\n• Timestamp antar modul sinkron 100%.",
            "Jangan pernah mempercayakan konversi UTC otomatis client JavaScript jika database target menggunakan tipe DATETIME non-timezone.",
            "Next.js, React, Node.js, MySQL DATETIME, Timezone Asia/Jakarta",
            "Explicit DateTime Contract tanpa double offset di API layer."
        ),
        (
            4,
            "Penanganan Error",
            "Fullstack & Database",
            "Asynchronous Out-of-Order Approval Sequence",
            "Riwayat komentar approval form FPPP di dashboard terbalik: Komentar Approver tahap 3 muncul di atas Approver tahap 1 (tercatat 5.367 baris urutan terbalik).",
            "1. Jalankan kueri: SELECT id, form_id, comment, created_at FROM comments.\n2. Bandingkan ID auto-increment dengan tanggal riil action_date di audit log DingTalk.",
            "Webhook diterima dan diproses async oleh multiple workers. Auto-increment ID merekam urutan webhook selesai di server, bukan waktu pejabat menandatangani dokumen.",
            "Merancang SQL View v_fppp_comments dengan multi-level coalesce sorting: ORDER BY COALESCE(task_start_date, action_date) ASC, id ASC per nomor form.",
            "• 5.367 baris log komentar kembali ke urutan kronologis bisnis yang 100% tepat.\n• Laporan audit trail persetujuan valid untuk audit kepatuhan.",
            "Di sistem event-driven asynchronous, nomor Auto-Increment ID bukanlah penanda waktu kejadian bisnis yang valid.",
            "MySQL / MariaDB, SQL Advanced Views, DingTalk Webhooks, REST API",
            "Multi-Tier Coalesce Sorting (Timestamp Bisnis > ID Fallback)."
        ),
        (
            5,
            "Penanganan Error",
            "DevOps & Cloud Security",
            "ECONNREFUSED & Port 3306 Blocked (Remote DB Access)",
            "Error: connect ECONNREFUSED 127.0.0.1:3306 saat AI Agent lokal dan script analitik mencoba menghubungkan ke MariaDB cloud di server produksi.",
            "1. Test koneksi port 3306 via telnet/nc: connection refused.\n2. Cek firewall UFW server cloud: port 3306 ditutup rapat demi keamanan dari serangan brute-force.",
            "Database produksi di cloud tidak boleh membuka port 3306 ke internet terbuka demi mencegah eksploitasi zero-day dan serangan hacker.",
            "Membangun konektivitas via Cloudflare Zero Trust Access TCP Tunnel: Daemon lokal cloudflared access tcp memetakan port 3306 lokal ke database remote via mTLS aman.",
            "• Port 3306 publik tetap tertutup total (0 open attack surface).\n• AI Agent dan local tooling dapat query langsung dengan latensi rendah (<40ms).",
            "Keamanan enterprise modern tidak boleh menghambat developer. Gunakan Zero Trust TCP Tunnel daripada membuka port database publik.",
            "Cloudflare Zero Trust, cloudflared CLI, Linux Firewall (UFW), MariaDB",
            "Zero Trust Network Access (ZTNA) Tunneling."
        ),
        (
            6,
            "Penanganan Error",
            "Fullstack & System Architecture",
            "False Late Flag: System Click Date vs Factual Completion Date",
            "Koding selesai & deploy tgl 19, namun developer baru klik centang 'DONE' di sistem project tracker tgl 26. Sistem mencatat late 7 hari dan skor anjlok ke 1.97/5.00.",
            "1. Audit tabel kpi_task_facts: completed_at bernilai 2026-10-26 sedangkan agreed_due_date 2026-10-19.\n2. Mengecek kode API mutasi status task.",
            "Skema database lama hanya memiliki satu kolom update: completed_at = NOW(). Tidak ada pemisahan antara waktu klik sistem dengan waktu penyelesaian fisik di lapangan.",
            "1. Tambahkan parameter completedAt faktual pada tool update_task (bisa backdate dengan validasi tidak boleh di masa depan).\n2. Kolom activity_log tetap mencatat audit klik asli.",
            "• Skor KPI karyawan pulih dari 1.97 menjadi 5.00/5.00 secara sah.\n• Sistem tetap memenuhi syarat audit kepatuhan karena tanggal klik tetap tercatat di log terpisah.",
            "Selalu pisahkan 'Audit Trail Timestamp' (kapan data diinput) dengan 'Business Effective Date' (kapan peristiwa bisnis terjadi).",
            "Strata PMS, Model Context Protocol, REST API Design, Database Architecture",
            "Dual-Date Architecture (Business Date vs Audit Date)."
        ),
        (
            7,
            "Penanganan Error",
            "Fullstack & Logistics Domain",
            "PO Stuck in Pending State due to Non-Physical GL Lines",
            "Status Purchase Order (PO) di sistem gudang Web-OTD tidak pernah bisa Released (menggantung selamanya), menyebabkan metrik pengiriman terdistorsi.",
            "1. Query tabel purchase_lines di database staging.\n2. Membandingkan baris PO dengan barang fisik yang tiba di truk logistik.\n3. Menemukan baris tipe GL Account (Ongkir) dan Jasa.",
            "API ERP mencampur item fisik dengan item non-fisik (ongkos angkut, asuransi, jasa pemasangan). Petugas gudang tidak bisa men-scan barcode barang yang tidak berwujud fisik.",
            "Membangun domain filtering pipeline: Filter item bertipe Item fisik & Fixed Asset untuk diverifikasi fisik, dan otomatis exclude baris jasa/GL account dari syarat gate-in.",
            "• Antarmuka staf gudang 100% bersih hanya berisi barang fisik.\n• Status PO Released menjadi akurat dan selesai tepat waktu.\n• Metrik OTD kredibel.",
            "Jangan menelan data ERP mentah-mentah. Pahami perbedaan antara 'dokumen akuntansi pembelian' dengan 'kejadian fisik di gudang'.",
            "Microsoft Dynamics 365 BC, Next.js, MariaDB Staging, Logistics Domain",
            "Domain-Driven Ingestion Filtering."
        ),
        (
            8,
            "Penanganan Error",
            "Fullstack & Web Engineering",
            "Target Save Authorization 403 on Periodicity-Only Change",
            "Admin mengubah periodisitas indikator KPI dari Bulanan ke Kuartalan, tetapi tombol simpan mati atau API mengembalikan 403 Forbidden 'No Changes Detected'.",
            "1. Inspect event handler onChange pada dropdown periodisitas.\n2. Periksa fungsi checkDirtyState() di state manajemen form React.",
            "Logika pendeteksi dirty state hanya mengecek perubahan pada nilai numerik Target (target !== initialTarget). Perubahan metadata Tipe_Periode_Indikator diabaikan.",
            "1. Memperbaiki fungsi dirty check di React untuk mengevaluasi Tipe_Periode_Indikator.\n2. Menambahkan otorisasi payload terpisah dan perbaikan nilai target KPI-921 di backend.",
            "• Admin dapat menyimpan pembaruan tipe periodisitas secara instan.\n• Form state sinkron sempurna antara UI dan database.",
            "Saat merancang formulir web kompleks, pastikan dirty state detector mengevaluasi seluruh atribut konfigurasi, bukan hanya input numerik utama.",
            "React, Next.js, State Management, API Authorization, Form Validation",
            "Comprehensive Multi-Field Dirty State Checking."
        ),
        (
            9,
            "Penanganan Error",
            "Data Engineering & SQL",
            "Weight Distortion & Data Drift pada Multi-Periodicity SQL View",
            "Saat indikator Kuartalan digabung dengan Bulanan dalam satu Master View KPI, total bobot per pegawai membengkak jadi 160% dan angka capaian historis bergeser (Data Drift).",
            "1. Kueri: SELECT SUM(Weight) FROM v_master_view GROUP BY Pegawai, Periode.\n2. Ditemukan nilai SUM > 1.0 (bobot pecah).\n3. Audit 4.437 baris historis menunjukkan pergeseran.",
            "Formula SQL lama menjumlahkan seluruh bobot indikator secara statis tanpa memperhitungkan apakah indikator tersebut aktif dievaluasi di bulan berjalan atau di akhir kuartal.",
            "Mengimplementasikan Migration 018: dynamic weight normalization engine di SQL View v_target_cross_periode menggunakan window function: Weight / SUM(Weight) OVER (PARTITION BY Pegawai, Periode).",
            "• Total bobot selalu otomatis ternormalisasi tepat 100%.\n• Audit 4.437 baris historis membuktikan 0 penyimpangan (Zero Data Drift).",
            "Idempotensi dan normalisasi bobot dinamis di SQL View menjamin konsistensi laporan historis meskipun skema bisnis berubah di tengah jalan.",
            "MySQL Advanced Views, SQL Window Functions, Database Migration, Data Quality Audit",
            "Idempotent Dynamic Weight Normalization via SQL Window Functions."
        ),
        (
            10,
            "Penanganan Error",
            "DevOps & Automation",
            "Google Sheets API 429 Too Many Requests (Rate Limit Quota)",
            "Error: 429 Too Many Requests: 'Read requests per minute per user exceeded' saat workflow n8n memperbarui baris data spreadsheet secara individual.",
            "1. Cek log node loop di n8n: n8n mengeksekusi request update per item secara berurutan dalam loop cepat.\n2. Bandingkan dengan kuota Google Sheets API (60 req/min).",
            "Pola loop per item menembak Google API ratusan kali dalam beberapa detik, langsung memicu rate limiting dari Google Cloud.",
            "Mengubah strategi n8n dari 'Single Item Loop' menjadi 'Batch Update': Seluruh data dikonsolidasikan dalam 1 array JSON lalu dikirim dalam 1 HTTP call dengan rate throttle delay 1.000ms.",
            "• Penggunaan API call turun dari 350 call menjadi 1 call per run.\n• Bebas dari error 429 rate limit.\n• Kecepatan sinkronisasi 10x lebih cepat.",
            "Selalu gunakan pola Batching saat berinteraksi dengan API pihak ketiga yang memiliki rate limit ketat.",
            "Google Sheets API, n8n, Cloud Integration, Rate Limiting & Throttling",
            "Batch Payload Aggregation Pattern."
        ),
        (
            11,
            "Penanganan Error",
            "AI Automation & LLM",
            "Tool Execution Argument Rejection in AI Agent MCP Runtime",
            "Error: 'Encountered error in tool execution: unrecognized argument Cwd / toolSummary rejected' saat AI Agent memanggil tools sistem manajemen proyek.",
            "1. Periksa schema tool JSON di direktori mcp/strata/.\n2. Bandingkan payload arguments yang digenerate oleh LLM dengan parameter yang diizinkan schema.",
            "LLM mengasumsikan parameter generik CLI (seperti Cwd, toolAction) yang tidak terdaftar pada skema ketat Model Context Protocol (additionalProperties: false).",
            "Memperbarui System Prompt dan Schema Definition MCP Server: Menyediakan deskripsi parameter eksplisit dan error hint yang jelas sehingga agent otomatis menyesuaikan argumen sesuai schema resmi.",
            "• Panggilan tool MCP oleh AI Agent berjalan 100% mulus tanpa argumen invalid.\n• Operasi automasi mandiri agent stabil tanpa intervensi manual.",
            "JSON Schema pada Model Context Protocol (MCP) harus eksplisit dan menyertakan instruksi error-handling agar LLM tidak berhalusinasi argumen.",
            "Model Context Protocol (MCP), LLM Agent Architecture, JSON Schema, REST RPC",
            "Schema-Driven Agentic Tool-Use Validation."
        ),
        (
            12,
            "Penanganan Error",
            "Data Science & Performance",
            "Metric Distortion from Uncapped Linear KPI Formula (300% Bias)",
            "Karyawan yang menyelesaikan 1 task remeh sebesar 300% mendapatkan skor total sempurna (5.00), meskipun task prioritas utama bisnis tidak selesai sama sekali.",
            "1. Analisis formula: Skor = SUM((Actual / Target) * Bobot).\n2. Cek simulasi: Over-achievement di indikator berbobot kecil menutupi kegagalan di indikator berbobot besar.",
            "Formula linier tanpa pembatas melanggar prinsip Goodhart's Law: karyawan terdorong mengejar target mudah secara berlebihan demi mendongkrak skor akhir.",
            "Menerapkan algoritma Cap 120%: Nilai capaian per indikator dibatasi maksimal 1.20 (120%). Konversi ke skala standar 1-5 dilakukan pada aggregasi akhir seluruh bobot.",
            "• Penilaian kinerja menjadi adil, seimbang, dan anti-gaming.\n• Mendorong penyelesaian seluruh pilar pekerjaan secara proporsional.",
            "Rumus evaluasi kinerja harus memiliki batas saturasi (cap) agar tidak mendistorsi fokus strategis organisasi.",
            "Business Analytics, Mathematical Modeling, KPI Formula, Data Visualization",
            "Non-Linear Saturation Cap Algorithm."
        )
    ]

    # 12 KASUS NON-ERROR (Improvement, Strategi & Tips/Trik)
    non_error_cases = [
        (
            13,
            "Strategi Arsitektur",
            "AI Automation & Systems",
            "Arsitektur 2-Phase Staging Buffer untuk Multi-Agentic AI System",
            "Developer menggunakan multi-agent AI (Claude Code, Antigravity, Codex) lintas berbagai repo. Jika tiap agent langsung menulis ke database produksi, terjadi task spam dan duplikasi.",
            "Kebutuhan integrasi AI Agent yang otonom namun terkontrol, tanpa merusak konsistensi skema project management utama perusahaan.",
            "Pola 2-Phase Staging Buffer: Agent mencatat log teknis mentah ke buffer pribadi (log_work_note) berstatus PENDING. Di akhir siklus/minggu, AI Agent mereview, mendeduplikasi data, lalu batch promote ke project resmi.",
            "• Zero duplicate tasks di sistem produksi.\n• Konteks pengerjaan tetap utuh lintas agent.\n• Tracking pekerjaan otomatis sinkron dengan metrik performa.",
            "Jangan biarkan AI Agent melakukan direct write ke tabel utama produksi. Selalu gunakan Staging Buffer Pattern untuk review dan agregasi cerdas.",
            "Model Context Protocol (MCP), Node.js, MariaDB, Multi-Agent Architecture",
            "Two-Phase Staging Buffer Pattern (Draft ➡️ Review ➡️ Batch Promote)."
        ),
        (
            14,
            "Improvement Sistem",
            "Fullstack & ERP Integration",
            "Arsitektur Hybrid Caching untuk Live ERP Integration (Dynamics 365 BC)",
            "Aplikasi Web-OTD di gudang mengalami delay 3-5 detik saat tim lapangan men-scan barcode Purchase Order untuk gate-in karena API ERP lambat saat peak hour.",
            "Tantangan latensi ERP enterprise yang tinggi dan rate limit API yang ketat tidak cocok untuk operasional frontline yang membutuhkan respon instan (<200ms).",
            "Arsitektur Hybrid Caching: Staging database lokal diupdate via sync worker untuk pencarian cepat (<100ms), lalu API Business Central live hanya dipanggil on-demand saat submit akhir sebagai validasi.",
            "• Waktu pencarian PO di gudang terpangkas dari 4 detik menjadi <150 milidetik.\n• Menghilangkan antrian fisik truk pengiriman di gerbang gudang.\n• Menghemat kuota API request ERP hingga 80%.",
            "Jangan pasang ERP kompleks sebagai backend query langsung aplikasi operasional frontline. Gunakan Read-Optimized Cache untuk kecepatan dan Live Verification untuk konsistensi.",
            "Microsoft Dynamics 365 BC API, Next.js, MariaDB Staging, Web Barcode Scanner",
            "Read-Optimized Staging Cache + On-Demand Live Verification."
        ),
        (
            15,
            "Strategi Arsitektur",
            "Fullstack & Security",
            "Dynamic Cross-Org RBAC via Database Security Views Tanpa Hardcoding",
            "Tim eksekutif dan auditor Procurement butuh memantau data performa divisi lain, tetapi sistem keamanan awal mengisolasi akses secara kaku berdasarkan departemen login.",
            "Kebutuhan delegasi akses dinamis yang fleksibel, aman, dan dapat dikonfigurasi kapan saja tanpa perlu deploy kode frontend baru.",
            "Membangun skema tabel kpi_access_delegations dan mengintegrasikan subquery Delegated_Viewers langsung di v_master_view_v2. Diuji dengan automated test suite 13 skenario akses.",
            "• Delegasi akses dapat diatur dinamis via Admin Console tanpa deploy ulang.\n• Keamanan data terjamin 100% lolos uji regresi hak akses.\n• Menghilangkan duplikasi query di berbagai frontend.",
            "Pindahkan logika otorisasi multi-tenant/cross-department ke dalam Dynamic Security Views, bukan disebar di puluhan komponen frontend.",
            "Role-Based Access Control (RBAC), MySQL Security Views, Next.js Auth, QA Automation",
            "Database-Level Dynamic Security Subquery."
        ),
        (
            16,
            "Tips & Trik Teknis",
            "Fullstack & Database",
            "Coalesce Multi-Tier Sorting untuk Menjinakkan Data Event-Driven Async",
            "Data transaksi atau log yang datang dari webhook multi-channel sering kali memiliki atribut waktu yang berbeda-beda (ada yang start_date, action_date, atau created_at kosong).",
            "Kebutuhan menampilkan audit trail yang 100% akurat secara bisnis meskipun payload pihak ketiga memiliki format atribut tanggal yang tidak konsisten.",
            "Menerapkan teknik multi-tier fallback: ORDER BY COALESCE(task_start_date, action_date, created_at) ASC, sequence_id ASC di level SQL View.",
            "• Menghasilkan urutan kronologis bisnis yang 100% stabil dan konsisten.\n• Mencegah bug tampilan riwayat yang melompat-lompat di frontend.",
            "Selalu siapkan hirarki fallback timestamp saat mengonsumsi data dari integrasi pihak ketiga.",
            "SQL Views, Data Modeling, Event-Driven Architecture, PostgreSQL / MySQL",
            "Multi-Tier Coalesce Timestamp Hierarchy."
        ),
        (
            17,
            "Strategi Arsitektur",
            "DevOps & Security",
            "Zero Trust TCP Tunneling untuk AI Agent & Local Tooling",
            "Local dev machine dan AI agent butuh terhubung langsung ke MariaDB produksi di cloud server tanpa membuka port publik dan tanpa VPN korporat yang rumit.",
            "Keamanan zero-trust menuntut tidak adanya port publik yang terekspos ke internet terbuka (Zero Attack Surface).",
            "Mengimplementasikan Cloudflare Zero Trust Access TCP Tunnel: Daemon background cloudflared access tcp berjalan di laptop developer, memetakan localhost:3306 ke cloud via mTLS.",
            "• Port 3306 publik tetap tertutup total (0 open attack surface).\n• AI Agent dan local tooling dapat query langsung dengan latensi rendah (<40ms).\n• Bebas dari scanning bot hacker.",
            "Gunakan Zero Trust TCP Tunneling untuk menghubungkan local tooling dan worker ke private database cloud tanpa repot setup VPN korporat.",
            "Cloudflare Zero Trust, cloudflared CLI, Linux Firewall (UFW), MariaDB",
            "Zero Trust Network Access (ZTNA) Tunneling."
        ),
        (
            18,
            "Tips & Trik Teknis",
            "Data Engineering & SQL",
            "Idempotent Dynamic Weight Normalization dengan SQL Window Function",
            "Perusahaan memiliki struktur indikator yang dinamis: beberapa indikator bernilai bulanan, beberapa kuartalan. Frontend membutuhkan total bobot selalu 100% pada filter apa pun.",
            "Menghindari pembuatan puluhan view berbeda untuk setiap kombinasi periode evaluasi.",
            "Menggunakan SQL Window Function: Weight / SUM(Weight) OVER (PARTITION BY Pegawai, Periode) AS Normalized_Weight langsung di dalam reporting master view.",
            "• Total bobot selalu otomatis terkunci di 100% pada filter bulan apa pun.\n• Menjamin zero data drift pada 4.437 baris audit historis.",
            "Gunakan SQL Window Function untuk kalkulasi proporsional dinamis di level database, menjaga frontend tetap bersih dan bebas kalkulasi rumit.",
            "SQL Window Functions, MySQL / MariaDB, Analytical Views, Business Intelligence",
            "Partitioned Dynamic Normalization Formula."
        ),
        (
            19,
            "Improvement Sistem",
            "DevOps & Automation",
            "Query Delta Filtering at Source vs Low-Code In-Memory Aggregation",
            "Platform low-code automasi (seperti n8n/Zapier) sering lambat dan boros memori jika disuruh memfilter ribuan baris data di dalam memorinya sendiri.",
            "Mengoptimalkan alur kerja automasi agar hemat resource server dan tidak memicu timeout gateway.",
            "Pindahkan logika filtering, deduplikasi, dan agregasi ke SQL View di level database (Delta Filtering at Source). Node automasi hanya menerima payload bersih yang siap kirim.",
            "• Penggunaan memori container turun drastis hingga 85%.\n• Waktu eksekusi automasi terpangkas hingga 98%.\n• Menghilangkan risiko out-of-memory crash.",
            "Biarkan database melakukan tugas komputasi data (karena didesain untuk itu), dan biarkan engine automasi fokus pada orchestrating HTTP & webhooks.",
            "n8n Automation, Database Optimization, Docker Resource Tuning, System Architecture",
            "Compute-at-Source (Pushdown Processing) Pattern."
        ),
        (
            20,
            "Tips & Trik Teknis",
            "Fullstack & System Architecture",
            "Pemisahan Audit Trail Timestamp vs Business Effective Date",
            "Pekerjaan fisik atau coding di lapangan sering kali diselesaikan lebih dulu, namun pencatatan administrasi di sistem baru dilakukan beberapa hari kemudian.",
            "Sistem evaluasi performa membutuhkan data penyelesaian faktual tanpa merusak integritas audit log sistem.",
            "Merancang dua kolom waktu terpisah pada entitas: completedAt (tanggal faktual penyelesaian bisnis) dan activity_log.created_at (waktu saat tombol di klik di sistem).",
            "• Sistem mampu menampung update pekerjaan backdate secara sah.\n• Skor performa KPI 100% akurat sesuai tanggal deploy sesungguhnya.\n• Kepatuhan audit tetap terjaga 100%.",
            "Selalu pisahkan 'Audit Trail Timestamp' dengan 'Business Effective Date' dalam arsitektur ERP & Project Management.",
            "Database Schema Design, Audit Compliance, ERP & PMS Architecture",
            "Dual-Date Architecture (Effective Date vs Audit Date)."
        ),
        (
            21,
            "Strategi Arsitektur",
            "AI Automation & LLM",
            "Membangun Custom Model Context Protocol (MCP) Server untuk Internal Tools",
            "AI Agent (Claude Code / Antigravity) tidak memiliki akses natif untuk mengecek backlog proyek, membuat tiket, atau memutasi database sistem enterprise internal.",
            "Menghubungkan AI Agent langsung ke infrastruktur operasional bisnis secara aman, terstandarisasi, dan terstruktur.",
            "Mengimplementasikan Model Context Protocol (MCP) Server khusus untuk software internal. Menyediakan 25+ tool RPC terstruktur dengan validasi JSON Schema ketat.",
            "• AI Agent kini bertindak sebagai autonomous project assistant.\n• Seluruh commit kode otomatis dicatat dan diselaraskan ke KPI.\n• Peningkatan efisiensi operasional developer hingga lebih dari 10x lipat.",
            "Masa depan AI bukan sekadar tanya-jawab di chat, melainkan 'Agentic Tool-Use' melalui Model Context Protocol (MCP) yang langsung terkoneksi ke infrastruktur bisnis.",
            "Model Context Protocol (MCP), LLM Agent Architecture, JSON Schema, REST RPC",
            "Agentic Tool-Use Architecture."
        ),
        (
            22,
            "Tips & Trik Teknis",
            "Fullstack & Logistics Domain",
            "Domain-Driven Filtering pada Data Pipeline Pengadaan & Gudang",
            "Sistem gudang sering menampilkan data non-fisik (ongkir, asuransi, jasa) yang membingungkan operator scanner karena API ERP menyatukan semua jenis purchase line.",
            "Memastikan hanya entitas fisik riil yang diproses oleh modul operasional logistik.",
            "Membangun klasifikasi domain pada tahap ingest data: Filter baris bertipe Item fisik & Fixed Asset untuk scan barcode gate-in, dan otomatis exclude/settle baris non-fisik.",
            "• Antarmuka staf gudang 100% bersih hanya berisi barang nyata.\n• Status PO Released/Received menjadi akurat secara real-time.\n• Laporan On-Time Delivery (OTD) perusahaan menjadi kredibel.",
            "Pahami model domain bisnis fisik sebelum mengonsumsi data ERP. Apa yang dianggap 'item pembelian' oleh Finance belum tentu 'barang fisik' bagi tim Logistik.",
            "Domain-Driven Design (DDD), Logistics Tech, Data Pipeline, ERP Integration",
            "Domain-Driven Ingestion Filtering."
        ),
        (
            23,
            "Strategi Arsitektur",
            "Data Science & Analytics",
            "Formula Non-Linier Cap 120% & Proportional Aggregation",
            "Formula linier standar (Actual / Target * Weight) memiliki kelemahan: over-performance pada satu indikator mudah menutupi kegagalan fatal pada indikator kritis.",
            "Mencegah manipulasi metrik (gaming the system) dan menjaga keselarasan pencapaian di seluruh pilar strategis perusahaan.",
            "Menerapkan algoritma Cap 120%: Nilai capaian per indikator dibatasi maksimal 1.20 (120%). Konversi ke skala standar 1-5 dilakukan pada aggregasi akhir seluruh bobot.",
            "• Menghilangkan manipulasi metrik (anti-gaming).\n• Mendorong penyelesaian seluruh pilar pekerjaan secara proporsional.\n• Penilaian kinerja menjadi adil, seimbang, dan diterima seluruh tim.",
            "Rumus evaluasi kinerja harus memiliki batas saturasi (cap) agar tidak mendistorsi fokus strategis organisasi (Goodhart's Law).",
            "Business Analytics, Mathematical Modeling, KPI Formula, Performance Metrics",
            "Non-Linear Saturation Cap Algorithm."
        ),
        (
            24,
            "Tips & Trik Teknis",
            "Fullstack & Web UI",
            "Robust Dirty-State Detection untuk Form Kompleks di React/Next.js",
            "Pada formulir admin dengan banyak atribut (numerik, dropdown, switch, tags), tombol simpan sering tidak aktif saat user hanya mengubah atribut sekunder/metadata.",
            "Menghindari bug UX di mana user merasa sudah mengubah konfigurasi tetapi sistem menolak menyimpan karena menganggap 'tidak ada perubahan'.",
            "Membangun dirty-state detector komprehensif yang melakukan deep-comparison pada seluruh atribut objek state formulir, bukan hanya atribut numerik utama.",
            "• Tombol simpan selalu responsif terhadap setiap perubahan atribut form.\n• Mengeliminasi komplain bug dari user admin HR.\n• Validasi form berjalan konsisten.",
            "Saat merancang formulir web dengan banyak tab/dropdown, pastikan dirty-state evaluation mencakup seluruh konfigurasi state objek.",
            "React, Next.js, Form State Management, UX Engineering, JavaScript",
            "Comprehensive Multi-Field Dirty State Evaluation."
        )
    ]

    all_cases = error_cases + non_error_cases

    # Badge map helper
    type_badge_map = {
        "Penanganan Error": badge_error,
        "Improvement Sistem": badge_improve,
        "Strategi Arsitektur": badge_strategy,
        "Tips & Trik Teknis": badge_tips
    }

    # =========================================================================
    # SHEET 1: Master Katalog 24 Kasus (All Cases)
    # =========================================================================
    ws1 = wb.active
    ws1.title = "Katalog Lengkap (24 Kasus)"
    ws1.views.sheetView[0].showGridLines = True

    ws1['A1'] = "MASTER KATALOG 24 KASUS: PENANGANAN ERROR, IMPROVEMENT, STRATEGI & TIPS-TRIK"
    ws1['A1'].font = title_font
    ws1['A2'] = "Koleksi lengkap insight engineering riil dari proyek nyata: 12 Kasus Error-Handling spesifik + 12 Kasus Arsitektur, Improvement & Tips-Trik untuk personal branding LinkedIn."
    ws1['A2'].font = subtitle_font

    headers_1 = [
        ("No", 6),
        ("Kategori Kasus", 20),
        ("Pilar Peran", 22),
        ("Topik / Judul Kasus", 30),
        ("Tantangan / Gejala Masalah", 38),
        ("Akar Masalah / Konteks Strategis", 38),
        ("Solusi Konkret & Implementasi", 45),
        ("Hasil Terukur & Dampak", 35),
        ("Key Engineering Takeaways", 38),
        ("Tech Stack Terkait", 25)
    ]

    for col_idx, (h_name, width) in enumerate(headers_1, 1):
        cell = ws1.cell(row=4, column=col_idx, value=h_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
        ws1.column_dimensions[get_column_letter(col_idx)].width = width

    ws1.row_dimensions[4].height = 28

    current_row = 5
    for case in all_cases:
        # Unpack first 10 fields for sheet 1
        no, ktype, role, title, chal, root, sol, res, take, tech = case[:10]

        ws1.row_dimensions[current_row].height = 95
        row_fill = accent_fill_1 if current_row % 2 == 1 else accent_fill_2

        row_vals = [no, ktype, role, title, chal, root, sol, res, take, tech]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws1.cell(row=current_row, column=col_idx, value=val)
            cell.font = cell_font
            cell.border = border_thin
            cell.fill = row_fill

            if col_idx == 1:
                cell.alignment = Alignment(horizontal="center", vertical="top")
                cell.font = cell_bold
            elif col_idx == 2:
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                cell.fill = type_badge_map.get(val, row_fill)
                cell.font = cell_bold
            elif col_idx in [3, 4]:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.font = cell_bold
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        current_row += 1

    ws1.freeze_panes = "A5"

    # =========================================================================
    # SHEET 2: Kasus Penanganan Error Spesifik (12 Kasus Error)
    # =========================================================================
    ws2 = wb.create_sheet(title="Kasus Error Spesifik")
    ws2.views.sheetView[0].showGridLines = True

    ws2['A1'] = "DEEP DIVE: 12 PENANGANAN ERROR SPESIFIK & ROOT CAUSE DEBUGGING"
    ws2['A1'].font = title_font
    ws2['A2'] = "Bedah tuntas insiden teknis error konkret: Pesan error riil, cara tracing log, diagnosa akar masalah, perbaikan kode, dan metrik hasil."
    ws2['A2'].font = subtitle_font

    headers_2 = [
        ("No", 6),
        ("Kode & Nama Error Spesifik", 26),
        ("Pilar Peran", 20),
        ("Gejala & Pesan Error Nyata", 35),
        ("Langkah Diagnosa / Cara Debugging", 38),
        ("Akar Masalah Teknis (Root Cause)", 38),
        ("Solusi Konkret & Perbaikan Kode / Query", 45),
        ("Hasil Setelah Perbaikan (Impact)", 32),
        ("Key Engineering Takeaway", 35)
    ]

    for col_idx, (h_name, width) in enumerate(headers_2, 1):
        cell = ws2.cell(row=4, column=col_idx, value=h_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
        ws2.column_dimensions[get_column_letter(col_idx)].width = width

    ws2.row_dimensions[4].height = 28

    current_row = 5
    for case in error_cases:
        no, ktype, role, title, chal, root_diag, root_cause, sol, res, take = case[0], case[1], case[2], case[3], case[4], case[5], case[6], case[7], case[8], case[9]
        row_vals = [no, title, role, chal, root_diag, root_cause, sol, res, take]

        ws2.row_dimensions[current_row].height = 95
        row_fill = accent_fill_1 if current_row % 2 == 1 else accent_fill_2

        for col_idx, val in enumerate(row_vals, 1):
            cell = ws2.cell(row=current_row, column=col_idx, value=val)
            cell.font = cell_font
            cell.border = border_thin
            cell.fill = row_fill

            if col_idx == 1:
                cell.alignment = Alignment(horizontal="center", vertical="top")
                cell.font = cell_bold
            elif col_idx == 2:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.font = cell_bold
                cell.fill = badge_error
            elif col_idx == 3:
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                cell.font = cell_bold
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        current_row += 1

    ws2.freeze_panes = "A5"

    # =========================================================================
    # SHEET 3: Improvement, Strategi & Tips/Trik (12 Kasus Non-Error)
    # =========================================================================
    ws3 = wb.create_sheet(title="Improvement & Strategi")
    ws3.views.sheetView[0].showGridLines = True

    ws3['A1'] = "DEEP DIVE: 12 IMPROVEMENT SISTEM, STRATEGI ARSITEKTUR & TIPS-TRIK TEKNIS"
    ws3['A1'].font = title_font
    ws3['A2'] = "Bedah inovasi arsitektural: Caching cerdas, otomasi multi-agent, keamanan zero trust, normalisasi SQL, dan tips desain sistem tanpa nunggu sistem error."
    ws3['A2'].font = subtitle_font

    headers_3 = [
        ("No", 6),
        ("Kategori Inovasi", 20),
        ("Topik / Judul Inovasi", 30),
        ("Tantangan & Kebutuhan Bisnis", 38),
        ("Konsep Desain & Pola Arsitektur", 40),
        ("Langkah Implementasi Teknis", 45),
        ("Efisiensi & Dampak Terukur", 35),
        ("Pelajaran untuk Tech Lead / Senior Engineer", 38),
        ("Tech Stack", 25)
    ]

    for col_idx, (h_name, width) in enumerate(headers_3, 1):
        cell = ws3.cell(row=4, column=col_idx, value=h_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
        ws3.column_dimensions[get_column_letter(col_idx)].width = width

    ws3.row_dimensions[4].height = 28

    current_row = 5
    for case in non_error_cases:
        no, ktype, role, title, chal, root, sol, res, take, tech, pattern = case
        row_vals = [no, ktype, title, chal, pattern, sol, res, take, tech]

        ws3.row_dimensions[current_row].height = 95
        row_fill = accent_fill_1 if current_row % 2 == 1 else accent_fill_2

        for col_idx, val in enumerate(row_vals, 1):
            cell = ws3.cell(row=current_row, column=col_idx, value=val)
            cell.font = cell_font
            cell.border = border_thin
            cell.fill = row_fill

            if col_idx == 1:
                cell.alignment = Alignment(horizontal="center", vertical="top")
                cell.font = cell_bold
            elif col_idx == 2:
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                cell.fill = type_badge_map.get(val, row_fill)
                cell.font = cell_bold
            elif col_idx == 3:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.font = cell_bold
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        current_row += 1

    ws3.freeze_panes = "A5"

    # =========================================================================
    # SHEET 4: LinkedIn Content Plan (24 Postingan Lengkap)
    # =========================================================================
    ws4 = wb.create_sheet(title="LinkedIn Content Playbook")
    ws4.views.sheetView[0].showGridLines = True

    ws4['A1'] = "LINKEDIN CONTENT PLAYBOOK (24 POSTINGAN: ERROR + IMPROVEMENT + STRATEGI + TIPS)"
    ws4['A1'].font = title_font
    ws4['A2'] = "Blueprint konten LinkedIn siap pakai berbasis STAR Framework, dibagi menjadi seri 'Behind The Bug', 'System Architecture', 'Optimization', dan 'Quick Tech Tips'."
    ws4['A2'].font = subtitle_font

    headers_4 = [
        ("Post ID", 8),
        ("Kategori Konten", 20),
        ("Topik Kasus", 28),
        ("Catchy Hook (3 Baris Pembuka)", 42),
        ("Format Visual", 18),
        ("Struktur STAR: Situation & Task", 35),
        ("Struktur STAR: Action & Engineering", 45),
        ("Struktur STAR: Result & Business Impact", 32),
        ("Call to Action (CTA)", 32),
        ("Hashtags Rekomendasi", 30)
    ]

    for col_idx, (h_name, width) in enumerate(headers_4, 1):
        cell = ws4.cell(row=4, column=col_idx, value=h_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
        ws4.column_dimensions[get_column_letter(col_idx)].width = width

    ws4.row_dimensions[4].height = 28

    linkedin_posts_24 = [
        # --- 12 ERROR POSTS ---
        (
            "POST-01", "Behind The Bug", "HTTP 503 Gateway Timeout di n8n",
            "🚨 'Workflow n8n Anda sering HTTP 503 saat jam sibuk?'\n\nGejalanya: workflow gantung 16 menit, lalu mati begitu saja. Developer sering menyalahkan spesifikasi server RAM/CPU.\n\nPadahal masalah aslinya ada di query data. Ini cara saya memperbaikinya: 👇",
            "Text Breakdown",
            "Workflow n8n sinkronisasi form FPPP crash dengan error 503 Gateway Timeout karena menarik 15.400 baris histori tak berindeks.",
            "Refactoring SQL View di MariaDB: Ganti pembacaan histori mati dengan filter form aktif (~480 baris) dan deploy versi aktif.",
            "Runtime terpangkas 98% (16 menit ➡️ 20 detik). Nol insiden error 503 sejak optimasi.",
            "Pernah ngalamin n8n gantung karena dataset membengkak? Apa langkah mitigasi pertama kalian?",
            "#DevOps #n8n #Automation #DatabaseOptimization #CloudArchitecture"
        ),
        (
            "POST-02", "Behind The Bug", "HTTP 400 Grid Limits Exceeded (Google Sheets)",
            "⚠️ 'This action would increase the number of cells above 10,000,000.'\n\nPernah lihat error ini di automasi Google Sheets? Artinya spreadsheet Anda sudah sekarat.\n\nIni alasan kenapa Google Sheets BUKAN database transaksional: 🧵",
            "Carousel (PDF)",
            "Node Google Sheets di workflow automasi menolak menulis data baru karena akumulasi baris mencapai batas 10 juta sel.",
            "Ubah pola integrasi dari append raw log menjadi Single Row State Update dengan pre-agregasi komentar langsung di database SQL.",
            "Ukuran spreadsheet stabil di bawah 5.000 baris, bebas crash kuota sel selamanya, dan file dibuka super cepat tanpa lag.",
            "Siapa di sini yang perusahaannya masih pakai Google Sheets rasa database produksi? Share suka dukanya!",
            "#GoogleSheetsAPI #DataPipelines #DevOps #Automation #TechLead"
        ),
        (
            "POST-03", "Behind The Bug", "Timezone Shift +7 Jam (False Late Flag)",
            "🕒 'Kerjaan selesai jam 4 sore, tapi dicatat terlambat oleh sistem.'\n\nSetelah diinvestigasi, akar masalahnya bukan karena developer telat kirim, tapi ada selisih 7 jam di kode JavaScript.\n\nSelamat datang di Timezone Hell: UTC vs Asia/Jakarta. 🔍",
            "Case Study Snippet",
            "Karyawan menyelesaikan tiket tepat waktu di sore hari, tetapi laporan HR mencatat tiket selesai di hari berikutnya dan ditandai LATE.",
            "Browser client mengeksekusi toISOString() yang memotong waktu lokal ke UTC (+0). Solusinya: standarisasi string YYYY-MM-DD HH:mm:ss murni tanpa double-offset.",
            "100% false-positive keterlambatan hilang, skor performa karyawan kembali adil dan objektif.",
            "Siapa yang pernah kena 'jebakan batman' timezone JavaScript vs Database? Angkat tangan! 🙋‍♂️",
            "#JavaScript #Fullstack #Nextjs #MySQL #CleanCode"
        ),
        (
            "POST-04", "Behind The Bug", "Urutan Approval Terbalik (Async Sorting Reversal)",
            "🤯 'Bagaimana bisa manajer menyetujui dokumen SEBELUM staff mengajukannya?'\n\nItu yang terlihat di audit log kami: 5.367 komentar approval posisinya terbalik total.\n\nPenyebabnya? Kami terlalu percaya pada Auto-Increment ID di sistem async: 👇",
            "Storytelling",
            "Audit trail persetujuan form acak-acakan karena webhook diterima secara asynchronous oleh multiple workers.",
            "Auto-increment ID mencatat waktu selesainya webhook di server, bukan waktu tanda tangan user. Solusinya: buat SQL view dengan fallback ORDER BY COALESCE(task_start_date, action_date), id.",
            "Seluruh 5.367 baris log persetujuan kembali berurutan sempurna sesuai timeline bisnis asli.",
            "Pelajaran penting: Jangan pernah pakai auto-increment ID untuk urutan kronologis bisnis pada sistem async terdistribusi!",
            "#SoftwareEngineering #Database #SQL #SystemDesign #EventDriven"
        ),
        (
            "POST-05", "Behind The Bug", "ECONNREFUSED Port 3306 (Secure Cloud Tunneling)",
            "🛑 'Error: connect ECONNREFUSED 127.0.0.1:3306'\n\nBanyak developer tergoda membuka port 3306 database cloud ke internet publik cuma biar bisa diakses dari laptop atau AI agent lokal.\n\nJangan lakukan itu! Ini cara elegan mengatasinya dengan Cloudflare Tunnel: 🛡️",
            "Architecture Diagram",
            "Local tooling dan AI Agent butuh direct query ke MariaDB cloud, tetapi port 3306 ditutup rapat oleh firewall server demi keamanan.",
            "Mengonfigurasi Cloudflare Zero Trust TCP Tunnel. Daemon lokal cloudflared access tcp menghubungkan port 3306 lokal ke database remote via mTLS aman.",
            "Database terlindungi total dengan 0 open port publik di internet, sementara developer & AI agent tetap bisa query dengan latensi <40ms.",
            "Gimana cara tim kalian mengamankan akses database staging/cloud saat ini? Masih pakai VPN lama atau Zero Trust?",
            "#CyberSecurity #DevOps #Cloudflare #ZeroTrust #CloudSecurity"
        ),
        (
            "POST-06", "Behind The Bug", "Skor Anjlok Akibat Klik Tanggal Selesai (Backdating Error)",
            "📉 'Skor KPI anjlok dari 5.00 jadi 1.97 hanya gara-gara lupa klik tombol di hari Jumat.'\n\nKode sudah naik ke production tepat waktu, tapi tombol 'Done' baru diklik hari Senin. Sistem menganggap developer terlambat 3 hari.\n\nIni arsitektur database untuk memecahkannya: 💡",
            "Carousel (PDF)",
            "Sistem tracking performa mengunci waktu penyelesaian menggunakan NOW() saat tombol diklik, menghukum karyawan yang telat update administrasi.",
            "Pisahkan skema: Kolom completedAt mencatat 'Business Effective Date' (kapan kerjaan nyata selesai), sedangkan audit trail log tetap mencatat 'System Timestamp' (kapan tombol diklik).",
            "Skor KPI mencerminkan kenyataan lapangan 5.00/5.00 tanpa mengorbankan integritas audit compliance HR.",
            "Apakah sistem evaluasi kerja di kantor kalian memfasilitasi backdate penyelesaian kerja yang faktual?",
            "#SoftwareArchitecture #DatabaseDesign #ERP #PeopleAnalytics #TechLeadership"
        ),
        (
            "POST-07", "Behind The Bug", "PO Menggantung Gara-Gara Baris Ongkir (ERP Domain Bug)",
            "🚚 'Kenapa status PO barang di gudang tidak pernah bisa 100% Selesai?'\n\nTernyata petugas gudang bingung mencari barcode untuk di-scan dari item bernama 'Biaya Ekspedisi' dan 'Jasa Pasang'.\n\nPentingnya Domain Classification pada data ERP: 🔍",
            "Text Breakdown",
            "Status Purchase Order menggantung di gudang karena sistem mewajibkan verifikasi fisik untuk seluruh baris data ERP.",
            "ERP menggabungkan Item fisik, GL Account, dan Item Charge dalam array yang sama. Solusinya: bangun data pipeline classifier yang memfilter hanya item fisik untuk scan barcode.",
            "Status PO Released menjadi akurat 100%, antarmuka petugas gudang bersih tanpa item hantu, dan tracking OTD menjadi valid.",
            "Pernah nemu data 'aneh' gara-gara perbedaan definisi orang IT dan orang Lapangan? Cerita di bawah ya!",
            "#Fullstack #LogisticsTech #ERP #DomainDrivenDesign #SupplyChain"
        ),
        (
            "POST-08", "Behind The Bug", "Form 403 Forbidden 'No Changes Detected' (Dirty State Bug)",
            "🐛 'Tombol Save mati padahal pilihan dropdown sudah diganti.'\n\nBug frontend klasik: Admin mengubah periodisitas indikator KPI dari Bulanan ke Kuartalan, tapi form menolak menyimpan karena menganggap 'tidak ada perubahan'.\n\nIni bedah logika dirty state form: 💻",
            "Code Snippet",
            "Form management menolak request submit karena dirty state detector hanya membandingkan nilai numerik target, mengabaikan perubahan dropdown tipe periode.",
            "Refactoring fungsi checkDirtyState() di state React untuk mengevaluasi seluruh atribut konfigurasi, dan memperbarui handler otorisasi backend.",
            "Admin bisa update konfigurasi periodisitas secara instan tanpa glitch form.",
            "Tips Frontend: Selalu gunakan robust deep-comparison atau form library schema untuk dirty checking!",
            "#React #Frontend #Fullstack #JavaScript #WebDev"
        ),
        (
            "POST-09", "Behind The Bug", "Bobot Membengkak 160% (Data Drift Multi-Periodicity)",
            "📊 'Saat total bobot indikator KPI tiba-tiba tembus 160%.'\n\nHR meminta indikator Kuartalan digabung dengan Bulanan dalam satu layar. Akibatnya, formula agregasi SQL pecah dan nilai historis 4.000 karyawan bergeser!\n\nIni cara saya mengatasinya dengan Window Function Normalization: 🧮",
            "Deep Dive Case Study",
            "Penggabungan indikator beda siklus waktu membuat total pembobotan lebih dari 100% dan mendistorsi historical audit data.",
            "Menerapkan Migration 018: dynamic weight normalization engine di SQL View menggunakan formula: Weight / SUM(Weight) OVER (PARTITION BY Pegawai, Periode).",
            "Total bobot selalu otomatis terkunci di 100% pada bulan apa pun, dan audit 4.437 baris membuktikan 0 data drift.",
            "Bagaimana cara database architect kalian menangani perubahan skema metrik di tengah tahun tanpa merusak data masa lalu?",
            "#SQL #DataEngineering #Database #Analytics #DataQuality"
        ),
        (
            "POST-10", "Behind The Bug", "Google Sheets HTTP 429 Too Many Requests (Rate Limit)",
            "🚦 'Error 429: Rate Limit Exceeded pada automasi n8n.'\n\nLooping 300 data baris per baris ke Google Sheets API langsung memicu rate limit 60 request/menit dari Google Cloud.\n\nSolusinya: Mengganti Loop dengan Batch Payload. 📦",
            "Case Study Snippet",
            "Workflow n8n mati di tengah jalan saat memproses ratusan pembaruan baris data ke spreadsheet klien.",
            "Ubah strategi dari looping per baris menjadi Batch Update: Gabungkan 300 baris ke dalam 1 array JSON payload dan kirim dalam 1 API call dengan delay jeda 1.000ms.",
            "Penggunaan kuota request API berkurang hingga 99%, proses eksekusi 10x lebih cepat dan bebas dari error 429.",
            "Ingat: Looping API request satu per satu di automasi adalah 'antipattern'. Selalu prioritaskan batching!",
            "#DevOps #n8n #Automation #GoogleCloud #APIDesign"
        ),
        (
            "POST-11", "Behind The Bug", "MCP Tool Argument Rejection in AI Agent Runtime",
            "🤖 'Error: Tool argument Cwd / toolSummary rejected by schema.'\n\nSaat menyuruh AI Agent (Claude/Antigravity) mengeksekusi tugas di server, agent malah gagal karena mengirim argumen yang tidak dikenal oleh protokol MCP.\n\nCara mendesain Schema MCP yang tahan halusinasi: 🛠️",
            "Technical Guide",
            "AI Agent berasumsi bebas terhadap nama parameter CLI generik, melanggar validasi ketat JSON Schema di Model Context Protocol.",
            "Memperbarui definisi skema tool MCP dengan additionalProperties: false dan memberikan contoh deskriptif pada parameter prompt.",
            "AI Agent mampu mengeksekusi operasi manajemen proyek internal secara mandiri 100% tanpa salah operan argumen.",
            "Siapa yang sudah mulai develop custom MCP Server untuk AI coding agent kalian? Bagikan tipsnya!",
            "#AIAgents #ModelContextProtocol #LLM #Automation #DevTools"
        ),
        (
            "POST-12", "Behind The Bug", "Gaming The System: 300% Cap Bias pada Evaluasi Tim",
            "🎯 'Satu orang dapat nilai 5.00 sempurna hanya karena menyelesaikan 1 tugas gampang sebanyak 300%.'\n\nSementara tugas strategis lainnya 0%. Inilah bahaya Goodhart's Law jika rumus KPI menggunakan formula linier mentah.\n\nIni matematika Algoritma Cap 120%: 📈",
            "Carousel (PDF)",
            "Formula linier (Actual / Target * Bobot) memungkinkan over-performance pada task gampang menutupi kegagalan fatal pada task krusial.",
            "Menerapkan fungsi saturasi matematis: Score = min(1.20, Actual / Target) * Weight, dan konversi ke skala proporsional 1-5 dilakukan pada agregasi akhir.",
            "Menghilangkan manipulasi metrik (anti-gaming), mendorong tim menyelesaikan seluruh pilar tanggung jawab secara seimbang.",
            "Pernahkah kalian melihat metrik perusahaan yang justru membuat tim fokus ke hal yang salah? Mari kita bedah!",
            "#PeopleAnalytics #DataScience #GoodhartsLaw #KPI #Management"
        ),

        # --- 12 NON-ERROR POSTS (Improvement, Strategi & Tips/Trik) ---
        (
            "POST-13", "Strategi Arsitektur", "2-Phase Staging Buffer untuk Multi-Agentic AI",
            "🤖 'Jangan biarkan AI Agent Anda menulis langsung ke database produksi.'\n\nSaat kita memakai banyak AI Agent (Claude Code, Antigravity, Codex) di berbagai repo, database tiket/proyek bisa penuh spam dan duplikasi.\n\nIni arsitektur 2-Phase Staging Buffer yang saya bangun: 👇",
            "Carousel (PDF)",
            "Menggunakan multi-agent AI untuk tracking pekerjaan harian tanpa merusak kerapian database produksi.",
            "Membuat private staging buffer (status PENDING). Agent hanya boleh mencatat draft log. Di akhir hari/minggu, AI Agent mereview, membersihkan duplikat, lalu batch promote ke project resmi.",
            "Zero task spam, histori pekerjaan 100% rapi dan siap diaudit oleh manajemen tanpa campur tangan manual.",
            "Bagaimana cara kalian mengontrol aksi AI Agent agar tidak merusak data produksi? Diskusi yuk!",
            "#AIEngineering #MultiAgent #ModelContextProtocol #Automation #SystemDesign"
        ),
        (
            "POST-14", "Improvement Sistem", "Hybrid Caching untuk Integrasi ERP Dynamics 365",
            "⚡ 'Memangkas response time scan gudang dari 4 detik ke <150ms.'\n\nAPI ERP kompleks seperti Microsoft Dynamics 365 BC tidak didesain untuk live query di setiap ketukan keyboard barcode scanner lapangan.\n\nIni strategi Hybrid Caching yang menyelamatkan operasional gudang kami: 🚀",
            "Architecture Diagram",
            "Menghilangkan antrean truk logistik di pintu gerbang akibat lambatnya pencarian PO di scanner gudang.",
            "Memisahkan Read Path dan Write Path: Sync worker mengisi cache lokal MariaDB untuk autocomplete instan (<150ms), dan memanggil live ERP API hanya saat final submit gate-in.",
            "Waktu respon pencarian barcode turun dari 4 detik ke <150ms, antrean gerbang gudang hilang total, dan hemat kuota API ERP 80%.",
            "Pernahkah caching menyelamatkan operasional fisik bisnis kalian? Bagikan pengalamannya!",
            "#ERP #MicrosoftDynamics #Fullstack #PerformanceOptimization #SupplyChain"
        ),
        (
            "POST-15", "Strategi Arsitektur", "Dynamic Cross-Org RBAC Tanpa Hardcoded Code",
            "🛡️ 'Buka akses cross-department tanpa bikin query database jadi spaghetti.'\n\nManajemen butuh tim Procurement memantau metrik divisi lain. Cara salah: nambah if-else di kode frontend. Cara benar: Dynamic SQL View RBAC.\n\nIni arsitektur keamanannya: 🔐",
            "Code Snippet",
            "Menyediakan delegasi hak akses lintas organisasi yang dinamis, secure, dan dapat diubah kapan saja tanpa deploy frontend baru.",
            "Membuat tabel delegasi kpi_access_delegations dan menyuntikkan subquery dinamis Delegated_Viewers ke dalam master SQL view, divalidasi 13 automated test.",
            "Admin HR bisa mengatur izin lintas departemen secara visual via admin console tanpa deploy ulang. Nol celah kebocoran data.",
            "Lebih suka handle RBAC di level Application Layer (Middleware) atau Database Layer (Views)? Kenapa?",
            "#CyberSecurity #RBAC #SoftwareEngineering #Backend #WebSecurity"
        ),
        (
            "POST-16", "Tips & Trik Teknis", "Coalesce Multi-Tier Sorting untuk Data Async",
            "💡 'Tips Database: Menjinakkan urutan data dari webhook multi-channel yang tidak konsisten.'\n\nSaat menerima data dari berbagai API eksternal, kadang tanggal persetujuan ada di 'start_date', kadang di 'action_date', atau bahkan kosong.\n\nIni trik SQL Coalesce Hierarchy yang wajib dicatat: 📝",
            "Code Snippet",
            "Menampilkan audit log yang urut secara kronologis bisnis dari berbagai webhook eksternal dengan format tanggal berbeda.",
            "Terapkan multi-tier sorting: ORDER BY COALESCE(task_start_date, action_date, created_at) ASC, sequence_id ASC di level view pelaporan.",
            "Urutan timeline keputusan bisnis tampil 100% konsisten tanpa glitch urutan melompat di frontend.",
            "Apa trik SQL andalan kalian saat membersihkan data tanggal yang formatnya berantakan?",
            "#SQL #DatabaseTips #Backend #DataEngineering #WebDev"
        ),
        (
            "POST-17", "Strategi Arsitektur", "Zero Trust TCP Tunneling untuk Private Cloud Database",
            "🔒 'Membuka akses remote ke database private cloud tanpa VPN dan tanpa port publik terbuka.'\n\nBanyak tim dev pusing saat AI agent atau worker lokal butuh akses ke DB cloud: pilihannya antara buka port publik (bahaya) atau setup VPN (ribet).\n\nIni solusi modern menggunakan Cloudflare Zero Trust: 🌐",
            "Architecture Diagram",
            "Menyediakan koneksi database privat berkecepatan tinggi untuk AI Agent lokal dan developer tanpa mengorbankan keamanan cloud firewall.",
            "Menjalankan daemon cloudflared access tcp di mesin lokal yang menghubungkan localhost:3306 ke database remote melalui terowongan terenkripsi mTLS.",
            "Port 3306 publik di internet tetap tertutup total (0 open port), koneksi stabil dengan latensi <40ms.",
            "Masih pakai VPN korporat atau sudah beralih ke Zero Trust Tunneling untuk akses remote server?",
            "#DevOps #ZeroTrust #Cloudflare #CloudArchitecture #CyberSecurity"
        ),
        (
            "POST-18", "Tips & Trik Teknis", "Idempotent Dynamic Weight Normalization (SQL Window Function)",
            "🧮 'Trik SQL: Normalisasi bobot dinamis 100% tanpa mengubah data historis.'\n\nSaat manajemen mengubah skema indikator (dari bulanan ke kuartalan), formula SUM biasa sering bikin total bobot pecah atau data lama bergeser.\n\nIni rumus SQL Window Function yang idempotent dan anti-drift: 👇",
            "Code Snippet",
            "Menghitung total bobot persentase dinamis agar selalu tepat 100% di setiap kombinasi filter periode evaluasi.",
            "Gunakan formula: Weight / SUM(Weight) OVER (PARTITION BY Pegawai, Periode) AS Normalized_Weight langsung di dalam SQL View.",
            "Total bobot selalu 100% secara otomatis pada bulan apa pun, dan audit 4.437 baris membuktikan zero data drift.",
            "Seberapa sering kalian menggunakan Window Functions untuk memecahkan logika bisnis rumit di database?",
            "#SQL #DataEngineering #Analytics #DatabaseDesign #BusinessIntelligence"
        ),
        (
            "POST-19", "Improvement Sistem", "Compute-at-Source (SQL View) vs Low-Code Aggregation",
            "⚡ 'Kenapa automasi low-code (n8n/Zapier) kalian lambat dan sering kehabisan memori?'\n\nKesalahan umum: Menarik ribuan baris data mentah lalu melakukan filter & sorting di dalam node automasi.\n\nPrinsip Compute-at-Source yang harus dipahami: 🧵",
            "Text Breakdown",
            "Mengoptimalkan workflow automasi integrasi agar hemat memori container dan tidak memicu timeout.",
            "Pindahkan komputasi dan deduplikasi berat ke SQL View di level database. Node automasi hanya menarik payload bersih (~480 baris) yang siap kirim.",
            "Memori container n8n hemat 85%, waktu eksekusi terpangkas 98% (dari 16 menit jadi 20 detik).",
            "Kapan kalian memutuskan logika data harus ditaruh di Database vs ditaruh di Backend / Automasi?",
            "#DevOps #n8n #Automation #SoftwareEngineering #Performance"
        ),
        (
            "POST-20", "Tips & Trik Teknis", "Dual-Date Architecture: Business Date vs Audit Date",
            "💡 'Tips Desain Skema ERP: Mengapa Anda Butuh 2 Kolom Tanggal Selesai.'\n\nBanyak developer hanya membuat satu kolom: completed_at = NOW(). Begitu user telat input administrasi, data analitik langsung kacau.\n\nIni pola Dual-Date Architecture yang elegan: 📝",
            "Carousel (PDF)",
            "Mencatat penyelesaian tugas yang faktual tanpa menghilangkan jejak rekam waktu input sistem.",
            "Sediakan dua kolom: completedAt (business effective date yang bisa di-backdate secara sah) dan activity_log (audit trail yang merekam waktu klik riil).",
            "Skor performa mencerminkan realitas lapangan 5.00/5.00 dan sistem tetap 100% lolos audit regulasi.",
            "Bagaimana arsitektur pencatatan tanggal selesai tugas di software internal kantor kalian?",
            "#SystemDesign #DatabaseDesign #SoftwareEngineering #ERP #CleanArchitecture"
        ),
        (
            "POST-21", "Strategi Arsitektur", "Membangun Custom MCP Server untuk Autonomous AI Worker",
            "🤖 'Dari chatbot tanya-jawab jadi autonomous AI Worker dengan Model Context Protocol (MCP).'\n\nBagaimana cara saya membuat AI Agent bisa otomatis memeriksa backlog proyek, membuat task, dan memperbarui timeline di database internal?\n\nIni arsitektur integrasi MCP yang saya bangun: 🌐",
            "Step-by-Step Guide",
            "Menghubungkan LLM (Claude/Antigravity) ke sistem manajemen proyek internal perusahaan agar AI bisa bekerja mandiri dan terstruktur.",
            "Membuat custom MCP Server dengan 25+ tool schemas (JSON Schema validation). Agent bisa inspect rencana kerja (get_plan), mencatat pekerjaan, dan mempromosikannya ke database produksi.",
            "Produktivitas tim meningkat drastis. Developer cukup review hasil kerja AI, seluruh pencatatan proyek dan kepatuhan SLA selesai otomatis.",
            "Sudah mulai eksplor Model Context Protocol (MCP) untuk workflow harian kalian? Fitur apa yang paling ngebantu?",
            "#AIAgents #ModelContextProtocol #LLM #Automation #FutureOfWork"
        ),
        (
            "POST-22", "Tips & Trik Teknis", "Domain-Driven Ingestion Filtering untuk Data ERP",
            "🚚 'Tips Integrasi ERP: Jangan pernah telan data pengadaan mentah-mentah ke modul gudang.'\n\nBagi orang Finance, 'Biaya Ekspedisi' adalah baris pembelian (PO Line). Tapi bagi orang Gudang, itu bukan barang fisik yang bisa di-scan barcode.\n\nPentingnya Domain Classification di tahap Ingestion: 📦",
            "Storytelling",
            "Memastikan modul gate-in penerimaan barang logistik hanya menampilkan entitas fisik nyata.",
            "Bangun filter klasifikasi di data pipeline: Pisahkan tipe Item fisik (wajib scan) dari GL Account/Jasa (auto-settle/exclude dari gate-in).",
            "Status PO Released menjadi akurat 100%, antarmuka petugas gudang bersih tanpa item hantu, dan tracking OTD menjadi valid.",
            "Pernah nemu data 'aneh' gara-gara perbedaan definisi orang IT, Finance, dan Lapangan?",
            "#Fullstack #LogisticsTech #ERP #DomainDrivenDesign #SupplyChain"
        ),
        (
            "POST-23", "Strategi Arsitektur", "Algoritma Non-Linier Cap 120% vs Goodhart's Law",
            "📈 'Kenapa formula KPI linier sederhana justru merusak motivasi tim Anda (Goodhart\'s Law).'\n\nJika skor dihitung linier tanpa batas, karyawan akan terdorong mengejar 1 tugas gampang sebanyak 300% demi menutupi kegagalan di tugas prioritas.\n\nIni matematika di balik Algoritma Cap 120%: 🎯",
            "Carousel (PDF)",
            "Merancang formula evaluasi kinerja yang adil, seimbang, dan anti-gaming.",
            "Terapkan fungsi saturasi matematis: Score = min(1.20, Actual / Target) * Weight. Konversi ke skala standar 1-5 dilakukan pada agregasi akhir seluruh bobot.",
            "Mendorong tim menyelesaikan seluruh pilar tanggung jawab secara seimbang tanpa ada indikator strategis yang dianaktirikan.",
            "Pernahkah kalian melihat metrik perusahaan yang justru membuat tim fokus ke hal yang salah? Mari kita bedah!",
            "#PeopleAnalytics #DataScience #GoodhartsLaw #KPI #Management"
        ),
        (
            "POST-24", "Tips & Trik Teknis", "Multi-Field Dirty-State Detection di Form React/Next.js",
            "💻 'Tips Frontend: Mencegah bug tombol Save mati saat konfigurasi form diubah.'\n\nPada form admin kompleks, sering kali user mengganti dropdown metadata (misal: tipe periodisitas), tapi tombol Simpan tetap disabled karena sistem mengira 'tidak ada perubahan'.\n\nIni trik deep dirty checking di React: 🛠️",
            "Code Snippet",
            "Memastikan tombol simpan formulir selalu mendeteksi perubahan atribut apa pun secara akurat.",
            "Gunakan evaluator deep-comparison pada seluruh state form konfigurasi (bukan hanya nilai input numerik utama), atau gunakan schema validation library seperti Zod.",
            "Admin bisa update konfigurasi periodisitas secara instan tanpa glitch form atau tombol mati.",
            "Bagaimana cara kalian menghandle dirty state di formulir kompleks React saat ini?",
            "#React #Frontend #Fullstack #JavaScript #WebDev"
        )
    ]

    current_row = 5
    for post in linkedin_posts_24:
        pid, ktype, title, hook, fmt, star_st, star_act, star_res, cta, tags = post

        ws4.row_dimensions[current_row].height = 95
        row_fill = accent_fill_1 if current_row % 2 == 1 else accent_fill_2

        row_vals = [pid, ktype, title, hook, fmt, star_st, star_act, star_res, cta, tags]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws4.cell(row=current_row, column=col_idx, value=val)
            cell.font = cell_font
            cell.border = border_thin
            cell.fill = row_fill

            if col_idx == 1:
                cell.alignment = Alignment(horizontal="center", vertical="top")
                cell.font = cell_bold
            elif col_idx == 2:
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                cell.font = cell_bold
                if "Bug" in val:
                    cell.fill = badge_error
                elif "Strategi" in val:
                    cell.fill = badge_strategy
                elif "Improvement" in val:
                    cell.fill = badge_improve
                else:
                    cell.fill = badge_tips
            elif col_idx == 3:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.font = cell_bold
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        current_row += 1

    ws4.freeze_panes = "A5"

    # =========================================================================
    # SHEET 5: Arsitektur Ekosistem & Best Practice
    # =========================================================================
    ws5 = wb.create_sheet(title="Arsitektur & Best Practices")
    ws5.views.sheetView[0].showGridLines = True

    ws5['A1'] = "PLAYBOOK ARSITEKTUR EKOSISTEM & BEST PRACTICE ENTERPRISE"
    ws5['A1'].font = title_font
    ws5['A2'] = "Peta komponen teknologi, fungsi arsitektural, dan pola desain yang diterapkan dalam memecahkan masalah riil di ekosistem perusahaan."
    ws5['A2'].font = subtitle_font

    headers_5 = [
        ("Layer / Komponen", 22),
        ("Teknologi Spesifik", 25),
        ("Peran Arsitektural", 35),
        ("Tantangan Produksi yang Dihadapi", 38),
        ("Pola Solusi & Best Practice Teruji", 40)
    ]

    for col_idx, (h_name, width) in enumerate(headers_5, 1):
        cell = ws5.cell(row=4, column=col_idx, value=h_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
        ws5.column_dimensions[get_column_letter(col_idx)].width = width

    ws5.row_dimensions[4].height = 28

    tech_patterns = [
        (
            "AI Agent Orchestration",
            "Model Context Protocol (MCP), Claude Code, Antigravity, Node.js",
            "Menghubungkan autonomous AI Agent ke sistem operasional internal (PMS, Database, n8n, Cloudflare).",
            "AI Agent berisiko membuat task spam, duplicate data, dan menjalankan aksi destructive di database produksi.",
            "Two-Phase Staging Buffer Pattern (Draft -> Review -> Promote) dengan validasi JSON Schema ketat."
        ),
        (
            "DevOps & Security",
            "Cloudflare Zero Trust, cloudflared, UFW Firewall, TLS/mTLS",
            "Akses remote aman untuk developer dan background worker ke database cloud tanpa public exposure.",
            "Port 3306 rentan diserang hacker jika dibuka ke internet; VPN korporat terlalu lambat dan rumit.",
            "Zero Trust TCP Tunneling: Port lokal dimapping ke private cloud endpoint terenkripsi end-to-end."
        ),
        (
            "Workflow Automation",
            "n8n Workflow Engine, Docker, Google Sheets API, DingTalk Webhooks",
            "Integrasi data antar aplikasi (DingTalk Form Approval -> n8n -> Database -> Google Sheets).",
            "HTTP 503 Gateway Timeouts dan HTTP 400 Grid Limits akibat pengiriman dataset raksasa tanpa batching.",
            "Compute-at-Source (SQL View Delta Filtering) + Throttled Batch Processing + Throttling Retry Guards."
        ),
        (
            "Database & Data Pipeline",
            "MySQL / MariaDB, SQL Advanced Views, Stored Procedures, Migrations",
            "Single Source of Truth untuk reporting metrik performa, audit trail approval, dan staging data.",
            "Data drift pada periodisitas berbeda (bulan vs kuartal), race condition multi-approval, dan timezone skew.",
            "Idempotent SQL Views, Multi-Tier Coalesce Sorting, Dynamic Weight Normalization, dan Strict Timezone Cutoffs."
        ),
        (
            "Enterprise ERP Integration",
            "Microsoft Dynamics 365 Business Central API, REST, JSON",
            "Sinkronisasi data pengadaan (Purchase Orders & Lines) dengan sistem tracking penerimaan logistik gudang.",
            "High latency saat jam sibuk dan rate limit ERP membuat aplikasi frontline lambat dan antrean menumpuk.",
            "Hybrid Read-Optimized Caching: Caching lokal untuk fast query + On-Demand Live Verification saat commit."
        ),
        (
            "Fullstack Web Application",
            "Next.js, React, Tailwind CSS, TypeScript, Yida Dashboard UI",
            "Antarmuka pengguna untuk monitoring KPI, mobile barcode scanner logistik, dan administrasi delegasi.",
            "Pergeseran tanggal akibat toISOString() UTC conversion dan state management form complex.",
            "Explicit DateTime String Normalization (Asia/Jakarta) + Dirty-State Detection + Responsive Visual Badges."
        )
    ]

    current_row = 5
    for item in tech_patterns:
        ws5.row_dimensions[current_row].height = 70
        row_fill = accent_fill_1 if current_row % 2 == 1 else accent_fill_2

        for col_idx, val in enumerate(item, 1):
            cell = ws5.cell(row=current_row, column=col_idx, value=val)
            cell.font = cell_font
            cell.border = border_thin
            cell.fill = row_fill

            if col_idx in [1, 2]:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.font = cell_bold
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        current_row += 1

    ws5.freeze_panes = "A5"

    output_file = r"C:\Website\Insight-dev\Master_Engineering_Playbook_Case_Studies.xlsx"
    wb.save(output_file)
    print(f"Master Excel workbook successfully created at: {output_file}")

if __name__ == "__main__":
    build_master_excel()
