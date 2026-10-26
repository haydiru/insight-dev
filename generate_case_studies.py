import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_excel():
    wb = openpyxl.Workbook()

    # Styling Palette
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid") # Deep Slate Navy
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    accent_fill_1 = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") # Very light slate
    accent_fill_2 = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid") # White

    # Error severity badges / Domain badges
    badge_devops = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")   # Amber
    badge_fullstack = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")# Sky Blue
    badge_ai = PatternFill(start_color="EDE9FE", end_color="EDE9FE", fill_type="solid")       # Purple
    badge_data = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")     # Emerald

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
    # SHEET 1: Katalog Kasus Error Spesifik & Root Cause Fix
    # =========================================================================
    ws1 = wb.active
    ws1.title = "Penanganan Error Spesifik"
    ws1.views.sheetView[0].showGridLines = True

    ws1['A1'] = "KATALOG PENANGANAN ERROR SPESIFIK & ROOT CAUSE FIX (ENGINEERING CASE STUDIES)"
    ws1['A1'].font = title_font
    ws1['A2'] = "Daftar insiden teknis riil yang dipisahkan secara spesifik per error: pesan error, cara diagnosa, akar masalah, perbaikan kode/query, dan dampak terukur."
    ws1['A2'].font = subtitle_font

    headers_1 = [
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

    for col_idx, (h_name, width) in enumerate(headers_1, 1):
        cell = ws1.cell(row=4, column=col_idx, value=h_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
        ws1.column_dimensions[get_column_letter(col_idx)].width = width

    ws1.row_dimensions[4].height = 28

    errors_data = [
        (
            1,
            "HTTP 503 Service Unavailable / Gateway Timeout",
            "DevOps & Automation",
            "Error: 503 Service Unavailable pada node HTTP n8n saat eksekusi workflow 'GET FORM FPPP ORI' (3pwAHMpguHsPHW8O). Workflow gantung selama 16 menit lalu crash.",
            "1. Membuka log execution n8n (#953532).\n2. Memeriksa query runtime di MariaDB slow query log.\n3. Menemukan query get v_fppp_comments memproses 15.400 baris secara unindexed dan unbatched.",
            "n8n memuat seluruh histori baris form ke dalam memori container dalam satu event loop cycle, menyebabkan proxy gateway (Nginx/Cloudflare) memutus koneksi karena melebihi batas timeout 300 detik.",
            "Refactoring SQL View v_fppp_comments di MariaDB: Menambahkan filter status form aktif, memangkas data mentah dari 15.400 baris menjadi ~480 baris aktif, dan mempublikasikan versi workflow teroptimasi.",
            "• Durasi runtime terpangkas dari 16 menit menjadi 20 detik (turun 98%).\n• Nol insiden timeout HTTP 503.",
            "Jangan pernah biarkan automasi ETL memuat full historical scan jika bisnis hanya butuh delta status transaksi aktif."
        ),
        (
            2,
            "HTTP 400 Grid Limits Exceeded (Google Sheets API)",
            "DevOps & Automation",
            "Error: 400 Bad Request: 'This action would increase the number of cells in the workbook above the limit of 10,000,000 cells' pada node Google Sheets n8n.",
            "1. Cek parameter append row pada node Google Sheets.\n2. Hitung jumlah kolom x jumlah baris di spreadsheet target.\n3. Ditemukan akumulasi baris append tanpa deduplikasi selama berbulan-bulan.",
            "Pola integrasi lama melakukan raw log append per approval comment secara terus-menerus setiap 5 menit, sehingga kapasitas sel spreadsheet 10 juta sel terlampaui.",
            "1. Mengubah pola integrasi dari append raw log menjadi Single Row State Update (1 form = 1 baris di sheet).\n2. Melakukan pre-aggregation kolom komentar di SQL view sebelum dikirim ke Google Sheets.",
            "• Sheet operasional tetap ramping (<5.000 baris).\n• Menghilangkan error kuota 10 juta sel selamanya.\n• Sheet dapat dibuka kembali tanpa lag.",
            "Google Sheets adalah alat visualisasi, bukan database transactional. Lakukan agregasi di database sebelum menulis ke Sheets."
        ),
        (
            3,
            "Double-Offset Timezone Shift (+7h / -7h Bug)",
            "Fullstack & Data Integrity",
            "Task yang diselesaikan jam 16:30 WIB otomatis ditandai 'LATE' (terlambat) dan tanggal jatuh tempo mundur satu hari di dashboard KPI dan Web-OTD.",
            "1. Inspect network payload di browser devtools: terkirim '2026-10-15T09:30:00.000Z'.\n2. Cek kolom database MySQL: tersimpan '2026-10-15 09:30:00' tanpa zona waktu.\n3. Server menafsirkannya sebagai jam 09:30 pagi.",
            "Frontend JavaScript melakukan toISOString() yang mengonversi waktu lokal WIB ke UTC (+0). Backend kemudian menyimpan string tersebut ke kolom MySQL DATETIME tanpa zona waktu, sehingga terjadi pergeseran jam.",
            "1. Menghentikan penggunaan toISOString() pada form submission.\n2. Mengirim format string standar MySQL: YYYY-MM-DD HH:mm:ss berbasis Asia/Jakarta.\n3. Menyelaraskan server session timezone MySQL ke '+07:00'.",
            "• 100% false-positive status terlambat teratasi.\n• Skor kepatuhan on-time karyawan kembali akurat dan adil.",
            "Jangan andalkan konversi UTC otomatis client JavaScript jika database target menggunakan tipe DATETIME non-timezone."
        ),
        (
            4,
            "Asynchronous Out-of-Order Approval Sequence",
            "Fullstack & Database",
            "Riwayat komentar approval form FPPP di dashboard terbalik: Komentar Approver tahap 3 muncul di atas, sedangkan Approver tahap 1 muncul di bawah (tercatat 5.367 baris urutan terbalik).",
            "1. Melakukan query: SELECT id, form_id, comment, created_at FROM comments.\n2. Membandingkan kolom id auto-increment dengan tanggal riil action_date di form audit log DingTalk.",
            "Webhook DingTalk diterima dan diproses secara asynchronous oleh multi-worker. Auto-increment ID merekam urutan selesainya webhook di server, BUKAN urutan waktu pejabat melakukan tanda tangan approval.",
            "Merancang SQL View baru v_fppp_comments dengan multi-level coalesce sorting: ORDER BY COALESCE(task_start_date, action_date) ASC, id ASC dikelompokkan per nomor form.",
            "• 5.367 baris urutan approval kembali ke urutan kronologis bisnis yang 100% tepat.\n• Laporan audit trail persetujuan valid untuk kepatuhan regulasi.",
            "Di sistem event-driven asynchronous, nomor Auto-Increment ID bukanlah penanda waktu kejadian bisnis yang valid."
        ),
        (
            5,
            "ECONNREFUSED & Port 3306 Blocked (Remote DB Access)",
            "DevOps & Cloud Security",
            "Error: connect ECONNREFUSED 127.0.0.1:3306 saat AI Agent lokal dan script analitik mencoba menghubungkan ke MariaDB cloud di mariadb.automasiworkflow.com.",
            "1. Test port publik via telnet/nc ke port 3306: koneksi rejected.\n2. Cek status firewall server cloud (UFW): port 3306 sengaja ditutup dari internet publik untuk proteksi keamanan.",
            "Database produksi di cloud tidak boleh membuka port 3306 ke internet terbuka demi mencegah serangan brute-force dan eksploitasi zero-day.",
            "Membangun konektivitas via Cloudflare Zero Trust Access TCP Tunnel. Menjalankan daemon background cloudflared access tcp yang memetakan localhost:3306 ke endpoint cloud via mTLS terenkripsi.",
            "• Port 3306 publik tetap tertutup total (0 open attack surface).\n• AI Agent lokal dapat menjalankan kueri langsung dengan latensi rendah (<40ms).",
            "Keamanan enterprise tidak boleh menghambat automasi developer. Gunakan Zero Trust TCP Tunnel daripada membuka port database publik."
        ),
        (
            6,
            "False Late Flag: System Click Date vs Factual Completion Date",
            "Fullstack & System Architecture",
            "Pekerjaan koding selesai dan deploy tgl 19, tetapi developer baru sempat centang tombol 'DONE' di sistem project tracker tgl 26. Sistem mencatat late 7 hari dan skor anjlok ke 1.97/5.00.",
            "1. Audit tabel kpi_task_facts: kolom completed_at bernilai '2026-10-26 08:46:00' sedangkan agreed_due_date '2026-10-19'.\n2. Mengecek kode API mutasi status task.",
            "Skema database lama hanya memiliki satu kolom update: completed_at = NOW(). Tidak ada pemisahan antara 'Waktu Sistem Mencatat' dengan 'Waktu Bisnis Selesai'.",
            "1. Menambahkan parameter completedAt eksplisit pada tool update_task (bisa backdate dengan validasi tidak boleh di masa depan).\n2. Memisahkan kolom faktual completedAt dengan activity_log audit trail yang tetap mencatat real-time klik.",
            "• Skor KPI karyawan pulih dari 1.97 menjadi 5.00/5.00 secara sah.\n• Sistem tetap memenuhi syarat audit karena tanggal modifikasi sistem tetap tercatat di log terpisah.",
            "Selalu pisahkan 'Audit Trail Timestamp' (kapan data diinput) dengan 'Business Effective Date' (kapan peristiwa bisnis terjadi)."
        ),
        (
            7,
            "PO Stuck in Pending State due to Non-Physical GL Lines",
            "Fullstack & Logistics Domain",
            "Status Purchase Order (PO) di sistem gudang Web-OTD tidak pernah bisa menjadi 'Received' / 'Released' (menggantung selamanya), menyebabkan metrik ketepatan waktu pengiriman terdistorsi.",
            "1. Query tabel purchase_lines untuk PO terkait di database staging.\n2. Membandingkan baris PO dengan fisik barang yang tiba di truk.\n3. Menemukan baris tipe GL Account (Ongkir) dan Fixed Asset non-fisik.",
            "Payload API Microsoft Dynamics 365 BC mencampur item fisik dengan item non-fisik (ongkos angkut, asuransi, jasa pemasangan). Petugas gudang tidak bisa men-scan barang yang tidak berwujud fisik.",
            "Membangun domain filtering pipeline: Mengelompokkan item bertipe 'Item' fisik yang wajib scan barcode, dan otomatis meng-exclude baris jasa/GL account dari kriteria penyelesaian gate-in gudang.",
            "• Antarmuka staf gudang 100% bersih hanya berisi barang fisik.\n• Status PO Released menjadi akurat dan selesai tepat waktu.",
            "Jangan menelan data ERP mentah-mentah. Pahami perbedaan antara 'dokumen akuntansi' dengan 'kejadian fisik di gudang'."
        ),
        (
            8,
            "Target Save Authorization 403 on Periodicity-Only Change",
            "Fullstack & Web Engineering",
            "Error: Saat admin mengubah periodisitas indikator KPI dari Bulanan ke Kuartalan, tombol simpan tidak aktif atau API mengembalikan 403 Forbidden 'No Changes Detected'.",
            "1. Inspect event handler onChange pada dropdown periodisitas.\n2. Periksa fungsi checkDirtyState() di form state React.",
            "Logika pendeteksi dirty state hanya mengecek perubahan pada nilai numerik Target (target !== initialTarget). Perubahan pada kolom metadata Tipe_Periode_Indikator tidak memicu flag form dirty.",
            "1. Memperbaiki fungsi dirty check di React untuk mengevaluasi Tipe_Periode_Indikator.\n2. Menambahkan otorisasi payload terpisah dan koreksi target id KPI-921 di backend.",
            "• Admin dapat menyimpan pembaruan tipe periodisitas secara instan.\n• Form state sinkron sempurna antara UI dan database.",
            "Saat merancang formulir web, pastikan dirty state detector mengevaluasi seluruh atribut konfigurasi, bukan hanya input numerik utama."
        ),
        (
            9,
            "Weight Distortion & Data Drift pada Multi-Periodicity SQL View",
            "Data Engineering & SQL",
            "Saat indikator Kuartalan digabungkan dengan indikator Bulanan dalam satu Master View KPI, total bobot per pegawai membengkak menjadi 160% dan angka capaian historis bergeser (Data Drift).",
            "1. Jalankan kueri: SELECT SUM(Weight) FROM v_master_view GROUP BY Pegawai, Periode.\n2. Ditemukan nilai SUM > 1.0 (bobot pecah).\n3. Audit 4.437 baris historis menunjukkan penyimpangan data lama.",
            "Formula SQL lama menjumlahkan seluruh bobot indikator secara statis tanpa memperhitungkan apakah indikator tersebut aktif dievaluasi di bulan berjalan atau di akhir kuartal.",
            "Mengimplementasikan Migration 018: dynamic weight normalization engine di SQL View v_target_cross_periode menggunakan window function: Weight / SUM(Weight) OVER (PARTITION BY Pegawai, Periode).",
            "• Total bobot selalu otomatis ternormalisasi tepat 100%.\n• Audit 4.437 baris historis membuktikan 0 penyimpangan (Zero Data Drift).",
            "Idempotensi dan normalisasi bobot dinamis di SQL View menjamin konsistensi laporan historis meskipun skema bisnis berubah di tengah jalan."
        ),
        (
            10,
            "Google Sheets API 429 Too Many Requests (Rate Limit Quota)",
            "DevOps & Automation",
            "Error: 429 Too Many Requests: 'Read requests per minute per user exceeded' saat workflow n8n memperbarui baris data spreadsheet secara individual.",
            "1. Cek log node loop di n8n: n8n mengeksekusi request PUT/POST Google Sheets per item secara berurutan dalam loop cepat.\n2. Bandingkan dengan kuota Google Sheets API (60 requests/minute).",
            "Pola loop per item menembak Google API ratusan kali dalam beberapa detik, langsung memicu mekanisme rate limiting dari Google Cloud.",
            "Mengubah strategi n8n dari 'Single Item Loop' menjadi 'Batch Update': Seluruh data dikonsolidasikan dalam satu array JSON lalu dikirim hanya dalam 1 HTTP call dengan rate throttle delay 1.000ms.",
            "• Penggunaan API call turun dari 350 call menjadi 1 call per run.\n• Bebas dari error 429 rate limit.",
            "Selalu gunakan pola Batching saat berinteraksi dengan API pihak ketiga yang memiliki rate limit ketat."
        ),
        (
            11,
            "Tool Execution Argument Rejection in AI Agent MCP Runtime",
            "AI Automation & LLM",
            "Error: 'Encountered error in tool execution: unrecognized argument Cwd / toolSummary rejected' saat AI Agent memanggil tools sistem manajemen proyek.",
            "1. Periksa schema tool JSON di direktori mcp/strata/.\n2. Bandingkan payload arguments yang digenerate oleh LLM dengan parameter yang diizinkan schema.",
            "LLM mengasumsikan parameter generik CLI (seperti Cwd, toolAction) yang tidak terdaftar pada skema ketat Model Context Protocol (additionalProperties: false).",
            "Memperbarui System Prompt dan Schema Definition MCP Server: Menyediakan deskripsi parameter eksplisit dan error hint yang jelas sehingga agent otomatis menyesuaikan argumen sesuai schema resmi.",
            "• Panggilan tool MCP oleh AI Agent berjalan 100% mulus tanpa argumen invalid.\n• Operasi automasi mandiri agent stabil tanpa intervensi manual.",
            "JSON Schema pada Model Context Protocol (MCP) harus eksplisit dan menyertakan instruksi error-handling agar LLM tidak berhalusinasi argumen."
        ),
        (
            12,
            "Metric Distortion from Uncapped Linear KPI Formula (300% Bias)",
            "Data Science & Performance",
            "Karyawan yang menyelesaikan 1 task remeh sebesar 300% mendapatkan skor total sempurna (5.00), meskipun task prioritas utama bisnis tidak selesai sama sekali.",
            "1. Analisis formula: Skor = SUM((Actual / Target) * Bobot).\n2. Cek simulasi: Over-achievement di indikator berbobot kecil menutupi kegagalan di indikator berbobot besar.",
            "Formula linier tanpa pembatas melanggar prinsip Goodhart's Law: karyawan terdorong mengejar target mudah secara berlebihan demi mendongkrak skor akhir.",
            "Menerapkan algoritma Cap 120%: Nilai capaian per indikator dibatasi maksimal 1.20 (120%). Konversi ke skala standar 1-5 dilakukan pada aggregasi akhir seluruh bobot.",
            "• Penilaian kinerja menjadi adil, seimbang, dan anti-gaming.\n• Mendorong penyelesaian seluruh pilar pekerjaan secara proporsional.",
            "Rumus evaluasi kinerja harus memiliki batas saturasi (cap) agar tidak mendistorsi fokus strategis organisasi."
        )
    ]

    fill_map = {
        "DevOps & Automation": badge_devops,
        "DevOps & Cloud Security": badge_devops,
        "Fullstack & Data Integrity": badge_fullstack,
        "Fullstack & Database": badge_fullstack,
        "Fullstack & System Architecture": badge_fullstack,
        "Fullstack & Logistics Domain": badge_fullstack,
        "Fullstack & Web Engineering": badge_fullstack,
        "Data Engineering & SQL": badge_data,
        "AI Automation & LLM": badge_ai,
        "Data Science & Performance": badge_data
    }

    current_row = 5
    for row_data in errors_data:
        ws1.row_dimensions[current_row].height = 95
        row_fill = accent_fill_1 if current_row % 2 == 1 else accent_fill_2

        for col_idx, val in enumerate(row_data, 1):
            cell = ws1.cell(row=current_row, column=col_idx, value=val)
            cell.font = cell_font
            cell.border = border_thin
            cell.fill = row_fill

            if col_idx == 1:
                cell.alignment = Alignment(horizontal="center", vertical="top")
                cell.font = cell_bold
            elif col_idx == 2:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.font = cell_bold
            elif col_idx == 3:
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                cell.fill = fill_map.get(val, row_fill)
                cell.font = cell_bold
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        current_row += 1

    ws1.freeze_panes = "A5"

    # =========================================================================
    # SHEET 2: LinkedIn Story Per Error ("Behind The Bug")
    # =========================================================================
    ws2 = wb.create_sheet(title="LinkedIn Story Per Error")
    ws2.views.sheetView[0].showGridLines = True

    ws2['A1'] = "LINKEDIN POST PLAYBOOK: 'BEHIND THE BUG' (STORYTELLING PER ERROR)"
    ws2['A1'].font = title_font
    ws2['A2'] = "Template konten LinkedIn yang siap diposting, berfokus pada 1 error spesifik per postingan untuk membangun kredibilitas teknis yang tajam."
    ws2['A2'].font = subtitle_font

    headers_2 = [
        ("Post ID", 8),
        ("Error Spesifik yang Dibahas", 28),
        ("Catchy Hook (3 Baris Pembuka)", 40),
        ("Format Visual", 18),
        ("Struktur Cerita: The Error & The Symptom", 38),
        ("Struktur Cerita: How We Debugged & Solved It", 45),
        ("Struktur Cerita: The Numbers (Hasil Terukur)", 32),
        ("Call to Action (CTA) Diskusi", 32),
        ("Hashtags Relevan", 30)
    ]

    for col_idx, (h_name, width) in enumerate(headers_2, 1):
        cell = ws2.cell(row=4, column=col_idx, value=h_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
        ws2.column_dimensions[get_column_letter(col_idx)].width = width

    ws2.row_dimensions[4].height = 28

    linkedin_stories = [
        (
            "ERR-01",
            "HTTP 503 Gateway Timeout di n8n",
            "🚨 'Workflow n8n Anda sering HTTP 503 saat jam sibuk?'\n\nGejalanya: workflow gantung 16 menit, lalu mati begitu saja. Developer sering menyalahkan spesifikasi server RAM/CPU.\n\nPadahal masalah aslinya ada di query data. Ini cara saya memperbaikinya: 👇",
            "Text Breakdown",
            "Workflow n8n sinkronisasi form FPPP crash dengan error 503 Gateway Timeout. Eksekusi memakan waktu 16+ menit karena menarik 15.400 baris histori data yang tidak perlu.",
            "Alih-alih menaikkan kapasitas RAM server, saya membedah query SQL View-nya: memotong pembacaan histori mati dan menyisakan hanya ~480 baris transaksi aktif.",
            "Waktu eksekusi drop 98% dari 16 menit menjadi 20 detik! Nol error 503 sejak optimasi diterapkan.",
            "Pernah ngalamin n8n gantung karena dataset membengkak? Apa langkah mitigasi pertama kalian?",
            "#DevOps #n8n #Automation #DatabaseOptimization #CloudArchitecture"
        ),
        (
            "ERR-02",
            "HTTP 400 Grid Limits Exceeded (Google Sheets)",
            "⚠️ 'This action would increase the number of cells above 10,000,000.'\n\nPernah lihat error ini di automasi Google Sheets? Artinya spreadsheet Anda sudah sekarat.\n\nIni alasan kenapa Google Sheets BUKAN database transaksional: 🧵",
            "Carousel (PDF)",
            "Node Google Sheets di workflow automasi menolak menulis data baru karena akumulasi baris mencapai batas 10 juta sel.",
            "Mengubah arsitektur integrasi: Hentikan raw append log per komentar approval. Ganti dengan state update 1 baris per form dengan pre-agregasi komentar langsung di database SQL.",
            "Ukuran spreadsheet stabil di bawah 5.000 baris, bebas crash kuota sel selamanya, dan file dibuka super cepat tanpa lag.",
            "Siapa di sini yang perusahaannya masih pakai Google Sheets rasa database produksi? Share cerita suka dukanya!",
            "#GoogleSheetsAPI #DataPipelines #DevOps #Automation #TechLead"
        ),
        (
            "ERR-03",
            "Timezone Shift +7 Jam (False Late Flag)",
            "🕒 'Kerjaan selesai jam 4 sore, tapi dicatat terlambat oleh sistem.'\n\nSetelah diinvestigasi, akar masalahnya bukan karena developer telat kirim, tapi ada selisih 7 jam di kode JavaScript.\n\nSelamat datang di Timezone Hell: UTC vs Asia/Jakarta. 🔍",
            "Case Study Snippet",
            "Karyawan menyelesaikan tiket tepat waktu di sore hari, tetapi laporan HR mencatat tiket selesai di hari berikutnya dan ditandai LATE.",
            "Browser client mengeksekusi toISOString() yang memotong waktu lokal ke UTC (+0). Lalu backend MySQL menyimpannya sebagai DATETIME lokal. Solusinya: standarisasi string YYYY-MM-DD HH:mm:ss murni tanpa double-offset.",
            "100% false-positive keterlambatan hilang, skor performa karyawan kembali adil dan objektif.",
            "Siapa yang pernah kena 'jebakan batman' timezone JavaScript vs Database? Angkat tangan! 🙋‍♂️",
            "#JavaScript #Fullstack #Nextjs #MySQL #CleanCode"
        ),
        (
            "ERR-04",
            "Urutan Approval Terbalik (Async Sorting Reversal)",
            "🤯 'Bagaimana bisa manajer menyetujui dokumen SEBELUM staff mengajukannya?'\n\nItu yang terlihat di audit log kami: 5.367 komentar approval posisinya terbalik total.\n\nPenyebabnya? Kami terlalu percaya pada Auto-Increment ID di sistem async: 👇",
            "Storytelling",
            "Audit trail persetujuan form acak-acakan karena webhook diterima secara asynchronous oleh multiple workers.",
            "Auto-increment ID mencatat waktu selesainya webhook di server, bukan waktu pejabat menandatangani dokumen. Solusinya: buat SQL view dengan fallback ORDER BY COALESCE(task_start_date, action_date), id.",
            "Seluruh 5.367 baris log persetujuan kembali berurutan sempurna sesuai timeline bisnis asli.",
            "Pelajaran penting: Jangan pernah pakai auto-increment ID untuk urutan kronologis bisnis pada sistem async terdistribusi!",
            "#SoftwareEngineering #Database #SQL #SystemDesign #EventDriven"
        ),
        (
            "ERR-05",
            "ECONNREFUSED Port 3306 (Secure Cloud Tunneling)",
            "🛑 'Error: connect ECONNREFUSED 127.0.0.1:3306'\n\nBanyak developer tergoda membuka port 3306 database cloud ke internet publik cuma biar bisa diakses dari laptop atau AI agent lokal.\n\nJangan lakukan itu! Ini cara elegan mengatasinya dengan Cloudflare Tunnel: 🛡️",
            "Architecture Diagram",
            "Local tooling dan AI Agent butuh direct query ke MariaDB cloud, tetapi port 3306 ditutup rapat oleh firewall server demi keamanan.",
            "Mengonfigurasi Cloudflare Zero Trust TCP Tunnel. Daemon lokal cloudflared access tcp menghubungkan port 3306 lokal ke database remote via mTLS aman.",
            "Database terlindungi total dengan 0 open port publik di internet, sementara developer & AI agent tetap bisa query dengan latensi <40ms.",
            "Gimana cara tim kalian mengamankan akses database staging/cloud saat ini? Masih pakai VPN lama atau Zero Trust?",
            "#CyberSecurity #DevOps #Cloudflare #ZeroTrust #CloudSecurity"
        ),
        (
            "ERR-06",
            "Skor Anjlok Akibat Klik Tanggal Selesai (Backdating Error)",
            "📉 'Skor KPI anjlok dari 5.00 jadi 1.97 hanya gara-gara lupa klik tombol di hari Jumat.'\n\nKode sudah naik ke production tepat waktu, tapi tombol 'Done' baru diklik hari Senin. Sistem menganggap developer terlambat 3 hari.\n\nIni arsitektur database untuk memecahkannya: 💡",
            "Carousel (PDF)",
            "Sistem tracking performa mengunci waktu penyelesaian menggunakan NOW() saat tombol diklik, menghukum karyawan yang telat update administrasi.",
            "Pisahkan skema: Kolom completedAt mencatat 'Business Effective Date' (kapan kerjaan nyata selesai), sedangkan audit trail log tetap mencatat 'System Timestamp' (kapan tombol diklik).",
            "Skor KPI mencerminkan kenyataan lapangan 5.00/5.00 tanpa mengorbankan integritas audit compliance HR.",
            "Apakah sistem evaluasi kerja di kantor kalian memfasilitasi backdate penyelesaian kerja yang faktual?",
            "#SoftwareArchitecture #DatabaseDesign #ERP #PeopleAnalytics #TechLeadership"
        ),
        (
            "ERR-07",
            "PO Menggantung Gara-Gara Baris Ongkir (ERP Domain Bug)",
            "🚚 'Kenapa status PO barang di gudang tidak pernah bisa 100% Selesai?'\n\nTernyata petugas gudang bingung mencari barcode untuk di-scan dari item bernama 'Biaya Ekspedisi' dan 'Jasa Pasang'.\n\nPentingnya Domain Classification pada data ERP: 🔍",
            "Text Breakdown",
            "Status Purchase Order menggantung di gudang karena sistem mewajibkan verifikasi fisik untuk seluruh baris data ERP.",
            "ERP menggabungkan Item fisik, GL Account, dan Item Charge dalam array yang sama. Solusinya: bangun data pipeline classifier yang memfilter hanya item fisik untuk scan barcode.",
            "Status PO Released menjadi akurat 100%, antarmuka petugas gudang bersih tanpa item hantu, dan tracking OTD menjadi valid.",
            "Pernah nemu data 'aneh' gara-gara definisi orang IT dan orang Lapangan beda jauh? Cerita di bawah ya!",
            "#Fullstack #LogisticsTech #ERP #DomainDrivenDesign #SupplyChain"
        ),
        (
            "ERR-08",
            "Form 403 Forbidden 'No Changes Detected' (Dirty State Bug)",
            "🐛 'Tombol Save mati padahal pilihan dropdown sudah diganti.'\n\nBug frontend klasik: Admin mengubah periodisitas indikator KPI dari Bulanan ke Kuartalan, tapi form menolak menyimpan karena menganggap 'tidak ada perubahan'.\n\nIni bedah logika dirty state form: 💻",
            "Code Snippet",
            "Form management menolak request submit karena dirty state detector hanya membandingkan nilai numerik target, mengabaikan perubahan dropdown tipe periode.",
            "Refactoring fungsi checkDirtyState() di state React untuk mengevaluasi seluruh atribut konfigurasi, dan memperbarui handler otorisasi backend.",
            "Admin bisa update konfigurasi periodisitas secara instan tanpa glitch form.",
            "Tips Frontend: Selalu gunakan robust deep-comparison atau form library schema (seperti Zod/React-Hook-Form) untuk dirty checking!",
            "#React #Frontend #Fullstack #JavaScript #WebDev"
        ),
        (
            "ERR-09",
            "Bobot Membengkak 160% (Data Drift Multi-Periodicity)",
            "📊 'Saat total bobot indikator KPI tiba-tiba tembus 160%.'\n\nHR meminta indikator Kuartalan digabung dengan Bulanan dalam satu layar. Akibatnya, formula agregasi SQL pecah dan nilai historis 4.000 karyawan bergeser!\n\nIni cara saya mengatasinya dengan Window Function Normalization: 🧮",
            "Deep Dive Case Study",
            "Penggabungan indikator beda siklus waktu membuat total pembobotan lebih dari 100% dan mendistorsi historical audit data.",
            "Menerapkan Migration 018: dynamic weight normalization engine di SQL View menggunakan formula: Weight / SUM(Weight) OVER (PARTITION BY Pegawai, Periode).",
            "Total bobot selalu otomatis terkunci di 100% pada bulan apa pun, dan audit 4.437 baris membuktikan 0 data drift.",
            "Bagaimana cara database architect kalian menangani perubahan skema metrik di tengah tahun tanpa merusak data masa lalu?",
            "#SQL #DataEngineering #Database #Analytics #DataQuality"
        ),
        (
            "ERR-10",
            "Google Sheets HTTP 429 Too Many Requests (Rate Limit)",
            "🚦 'Error 429: Rate Limit Exceeded pada automasi n8n.'\n\nLooping 300 data baris per baris ke Google Sheets API langsung memicu rate limit 60 request/menit dari Google Cloud.\n\nSolusinya: Mengganti Loop dengan Batch Payload. 📦",
            "Case Study Snippet",
            "Workflow n8n mati di tengah jalan saat memproses ratusan pembaruan baris data ke spreadsheet klien.",
            "Ubah strategi dari looping per baris menjadi Batch Update: Gabungkan 300 baris ke dalam 1 array JSON payload dan kirim dalam 1 API call dengan delay jeda 1.000ms.",
            "Penggunaan kuota request API berkurang hingga 99%, proses eksekusi 10x lebih cepat dan bebas dari error 429.",
            "Ingat: Looping API request satu per satu di automasi adalah 'antipattern'. Selalu prioritaskan batching!",
            "#DevOps #n8n #Automation #GoogleCloud #APIDesign"
        ),
        (
            "ERR-11",
            "MCP Tool Argument Rejection in AI Agent Runtime",
            "🤖 'Error: Tool argument Cwd / toolSummary rejected by schema.'\n\nSaat menyuruh AI Agent (Claude/Antigravity) mengeksekusi tugas di server, agent malah gagal karena mengirim argumen yang tidak dikenal oleh protokol MCP.\n\nCara mendesain Schema MCP yang tahan halusinasi: 🛠️",
            "Technical Guide",
            "AI Agent berasumsi bebas terhadap nama parameter CLI generik, melanggar validasi ketat JSON Schema di Model Context Protocol.",
            "Memperbarui definisi skema tool MCP dengan additionalProperties: false dan memberikan contoh deskriptif pada parameter prompt.",
            "AI Agent mampu mengeksekusi operasi manajemen proyek internal secara mandiri 100% tanpa salah operan argumen.",
            "Siapa yang sudah mulai develop custom MCP Server untuk AI coding agent kalian? Bagikan tipsnya!",
            "#AIAgents #ModelContextProtocol #LLM #Automation #DevTools"
        ),
        (
            "ERR-12",
            "Gaming The System: 300% Cap Bias pada Evaluasi Tim",
            "🎯 'Satu orang dapat nilai 5.00 sempurna hanya karena menyelesaikan 1 tugas gampang sebanyak 300%.'\n\nSementara tugas strategis lainnya 0%. Inilah bahaya Goodhart's Law jika rumus KPI menggunakan formula linier mentah.\n\nIni matematika Algoritma Cap 120%: 📈",
            "Carousel (PDF)",
            "Formula linier (Actual / Target * Bobot) memungkinkan over-performance pada task gampang menutupi kegagalan fatal pada task krusial.",
            "Menerapkan fungsi saturasi matematis: Score = min(1.20, Actual / Target) * Weight, dan konversi ke skala proporsional 1-5 dilakukan pada agregasi akhir.",
            "Menghilangkan manipulasi metrik (anti-gaming), mendorong tim menyelesaikan seluruh pilar tanggung jawab secara seimbang.",
            "Pernahkah kalian melihat metrik perusahaan yang justru membuat tim fokus ke hal yang salah? Mari kita bedah!",
            "#PeopleAnalytics #DataScience #GoodhartsLaw #KPI #Management"
        )
    ]

    current_row = 5
    for post in linkedin_stories:
        ws2.row_dimensions[current_row].height = 95
        row_fill = accent_fill_1 if current_row % 2 == 1 else accent_fill_2

        for col_idx, val in enumerate(post, 1):
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
            elif col_idx == 3:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.font = cell_bold
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        current_row += 1

    ws2.freeze_panes = "A5"

    # =========================================================================
    # SHEET 3: Playbook Arsitektur & Best Practice Pencegahan
    # =========================================================================
    ws3 = wb.create_sheet(title="Arsitektur Pencegahan Error")
    ws3.views.sheetView[0].showGridLines = True

    ws3['A1'] = "PLAYBOOK ARSITEKTUR & BEST PRACTICE PENCEGAHAN ERROR (ENTERPRISE GRADE)"
    ws3['A1'].font = title_font
    ws3['A2'] = "Prinsip arsitektural dan panduan teknis teruji untuk mencegah error berulang di lingkungan produksi."
    ws3['A2'].font = subtitle_font

    headers_3 = [
        ("Layer Sistem", 22),
        ("Kategori Error Umum", 25),
        ("Akar Masalah Arsitektural", 35),
        ("Prinsip Desain Pencegahan (Best Practice)", 45),
        ("Contoh Implementasi Riil di Project", 38)
    ]

    for col_idx, (h_name, width) in enumerate(headers_3, 1):
        cell = ws3.cell(row=4, column=col_idx, value=h_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
        ws3.column_dimensions[get_column_letter(col_idx)].width = width

    ws3.row_dimensions[4].height = 28

    prevention_playbook = [
        (
            "Workflow Automation Engine",
            "Timeout (503) & Rate Limit (429)",
            "Memproses raw query data berukuran besar langsung di node automasi tanpa batching.",
            "Delta Filtering at Source: Saring dan agregasikan data di level SQL View sebelum masuk ke engine automasi. Terapkan batch payload ber-delay.",
            "n8n Workflow FPPP: Memangkas 15.400 baris ke 480 baris aktif, eksekusi drop dari 16 menit ke 20 detik."
        ),
        (
            "Cloud & DevOps Security",
            "Port Exposure & Connection Refused",
            "Membuka port database (3306) ke internet publik demi akses remote local tooling.",
            "Zero Trust Network Access (ZTNA): Jangan pernah buka port publik. Gunakan secure mTLS TCP Tunneling untuk remote development dan worker.",
            "Cloudflare Access TCP Tunnel: Menghubungkan local AI agent ke remote MariaDB tanpa port terbuka di firewall."
        ),
        (
            "Database & Data Integrity",
            "Data Drift & Race Conditions",
            "Mengandalkan Auto-Increment ID untuk urutan kronologis dan formula agregasi statis tanpa window function.",
            "Event-Driven Coalesce Sorting & Idempotent Dynamic Normalization: Selalu gunakan timestamp bisnis nyata dan window partition.",
            "Migration 018 KPI: Multi-periodisitas normalisasi bobot 100% dinamis teruji 4.437 baris deviasi 0."
        ),
        (
            "Fullstack Web Application",
            "Timezone Shifts & False LATE Flags",
            "Konversi toISOString() otomatis di client browser yang mengubah jam WIB ke UTC tanpa kontrol server.",
            "Explicit DateTime Contract: Kirim format string YYYY-MM-DD HH:mm:ss murni tanpa offset dari client, dan validasi ketat di API gateway.",
            "Web-OTD & Next.js: Eliminasi false overdue flag dengan standarisasi format MySQL DATETIME Asia/Jakarta."
        ),
        (
            "Enterprise Integration",
            "ERP Latency & Document Locking",
            "Direct live API call pada setiap user interaction di antarmuka lapangan.",
            "Hybrid Read-Optimized Caching: Sediakan local read cache untuk operasional cepat (<150ms), dan panggil live API hanya saat final commit.",
            "Dynamics 365 BC + Web-OTD: Autocomplete scan barcode instan di gudang tanpa membebani server ERP."
        ),
        (
            "AI Agent & LLM Tooling",
            "Tool Schema Rejection & Hallucination",
            "Definisi schema tool yang longgar atau tidak menyediakan error feedback bagi agent.",
            "Strict JSON Schema with Descriptive Error Hinting: Berikan batasan parameter ketat (additionalProperties: false) dan contoh konkret.",
            "Strata MCP Server: 25+ tool teruji untuk automasi mutasi proyek mandiri oleh AI Agent."
        )
    ]

    current_row = 5
    for item in prevention_playbook:
        ws3.row_dimensions[current_row].height = 75
        row_fill = accent_fill_1 if current_row % 2 == 1 else accent_fill_2

        for col_idx, val in enumerate(item, 1):
            cell = ws3.cell(row=current_row, column=col_idx, value=val)
            cell.font = cell_font
            cell.border = border_thin
            cell.fill = row_fill

            if col_idx in [1, 2]:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.font = cell_bold
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        current_row += 1

    ws3.freeze_panes = "A5"

    output_file = r"C:\Website\Insight-dev\AI_Automation_Error_Handling_Case_Studies.xlsx"
    wb.save(output_file)
    print(f"Excel workbook successfully updated at: {output_file}")

if __name__ == "__main__":
    build_excel()
