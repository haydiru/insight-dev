import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_excel():
    wb = openpyxl.Workbook()

    # Define color palette & styles
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid") # Dark Slate Navy
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    sub_header_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    sub_header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    accent_fill_1 = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid") # Light grayish slate
    accent_fill_2 = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid") # Pure white

    badge_ai = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid")     # Indigo tint
    badge_devops = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid") # Amber tint
    badge_fullstack = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid") # Sky tint
    badge_data = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")    # Emerald tint

    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    cell_font = Font(name="Calibri", size=10, color="0F172A")
    cell_bold = Font(name="Calibri", size=10, bold=True, color="0F172A")
    title_font = Font(name="Calibri", size=16, bold=True, color="1E293B")
    subtitle_font = Font(name="Calibri", size=11, italic=True, color="64748B")

    # =========================================================================
    # SHEET 1: Studi Kasus & Technical Problem
    # =========================================================================
    ws1 = wb.active
    ws1.title = "Studi Kasus & Problem Solving"
    ws1.views.sheetView[0].showGridLines = True

    # Title block
    ws1['A1'] = "KUMPULAN STUDI KASUS ENGINEERING: AI AUTOMATION, FULLSTACK & DEVOPS"
    ws1['A1'].font = title_font
    ws1['A2'] = "Ekstraksi insight problem-solving, arsitektur sistem, dan optimasi performa dari percakapan engineering nyata untuk konten LinkedIn & personal branding."
    ws1['A2'].font = subtitle_font

    headers_1 = [
        ("No", 6),
        ("Domain / Pilar Peran", 22),
        ("Topik Kasus", 28),
        ("The Problem / Gejala Masalah", 38),
        ("Root Cause Analysis (Akar Masalah)", 38),
        ("How to Solve / Solusi Teknis", 45),
        ("Technical & Business Impact", 35),
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

    cases_data = [
        (
            1,
            "AI Automation & Systems",
            "Multi-Agent Staging Buffer vs Polusi Database Produksi",
            "Developer menggunakan multi-agent AI (Claude Code, Antigravity, Codex) lintas repo berbeda. Jika tiap agent langsung menulis task ke database produksi, terjadi task spam, deskripsi acak-acakan, dan data duplikat.",
            "Tidak ada layer decoupling/buffer antara eksekusi runtime AI Agent yang sangat granular dan dinamis dengan skema production project management yang terstruktur.",
            "Merancang arsitektur 2-Phase: AI Agent menulis ke Private Staging Buffer (log_work_note) berstatus PENDING. Di akhir hari/minggu, AI mereview, membersihkan duplikasi, lalu melakukan batch promote (promote_work_notes) ke task resmi.",
            "• Zero duplicate tasks di sistem produksi.\n• Konteks pengerjaan tetap utuh lintas agent.\n• Data tracking rapi dan siap audit tanpa campur tangan manual.",
            "Jangan biarkan AI Agent melakukan direct write ke tabel utama produksi. Selalu gunakan Staging Buffer Pattern untuk review dan agregasi cerdas.",
            "Model Context Protocol (MCP), Node.js, MariaDB, Claude Code, Antigravity"
        ),
        (
            2,
            "DevOps & Automation",
            "Optimasi n8n: Dari 16 Menit ke 20 Detik (Fix HTTP 503 & 400)",
            "Workflow n8n sinkronisasi form FPPP ke Google Sheets sering timeout (HTTP 503 Gateway Timeout) dan crash akibat melebihi kuota sel (HTTP 400 Grid Limits) dengan durasi eksekusi 16+ menit.",
            "Query SQL menarik seluruh dataset historis (>15.400 baris) tanpa delta filtering, lalu seluruh baris dikirim sekaligus ke node Google Sheets yang memiliki limitasi API rate & grid cell.",
            "1. Refactoring SQL view untuk memfilter hanya form aktif (~480 baris).\n2. Memisahkan payload berat dengan agregasi komentar tervalidasi.\n3. Menambahkan throttling & retry guard pada n8n node.",
            "• Waktu eksekusi drop 98% (dari 16 menit menjadi 20 detik).\n• Error rate 0% (tidak pernah HTTP 503/400 lagi).\n• Beban memori container n8n stabil di bawah ambang batas.",
            "Automasi low-code bukan alasan mengabaikan efisiensi database. Filter di level sumber data (SQL View), jangan andalkan node filter di dalam engine automasi.",
            "n8n Automation, Docker, MariaDB, Google Sheets API, SQL Query Tuning"
        ),
        (
            3,
            "Fullstack & Database",
            "Menangani Urutan Audit Log Acak Akibat Multi-Approval Async",
            "Urutan riwayat komentar persetujuan form FPPP di dashboard manajemen terbalik dan tidak urut (terdeteksi 5.367 baris urutan terbalik), membingungkan pimpinan saat audit keputusan.",
            "Approval form memiliki banyak approver paralel. Pengurutan awal hanya mengandalkan auto-increment ID yang dibuat berdasarkan kapan webhook selesai diterima, bukan kapan approval benar-benar terjadi.",
            "Membuat SQL View ternormalisasi v_fppp_comments dengan algoritma multi-level coalesce: ORDER BY COALESCE(task_start_date, action_date), id serta mengelompokkan berdasarkan nomor approval stage.",
            "• 5.367 baris log komentar kembali ke urutan kronologis asli.\n• Timeline persetujuan 100% presisi untuk kebutuhan audit kepatuhan ISO & internal control.",
            "Auto-increment ID bukan penanda waktu yang valid pada sistem terdistribusi. Gunakan kombinasi COALESCE timestamp bisnis dan sequence fallback.",
            "MySQL / MariaDB, SQL Views, DingTalk Open API, Event-driven Webhooks"
        ),
        (
            4,
            "DevOps & Security",
            "Akses Database Cloud Aman Tanpa Buka Port Publik (Cloudflare Tunnel)",
            "Local dev machine dan AI agent butuh terhubung langsung ke MariaDB produksi di cloud server, namun membuka port 3306 ke publik sangat berisiko terkena brute-force dan eksploitasi zero-day.",
            "Kebutuhan konektivitas remote fleksibel tanpa repot setup VPN korporat yang berat atau IP Whitelisting yang sering berubah akibat koneksi dinamis ISP.",
            "Mengimplementasikan Cloudflare Zero Trust Access TCP Tunnel. Daemon lokal (cloudflared access tcp) berjalan di port 3306 lokal dan me-route traffic secara terenkripsi mTLS ke edge server Cloudflare.",
            "• Port 3306 publik tetap tertutup total (0 open port di firewall server).\n• Koneksi stabil dengan latensi rendah (<40ms).\n• Akses otomatis untuk background task AI Agent.",
            "Keamanan enterprise modern tidak mengorbankan kecepatan developer. Gunakan Zero Trust TCP Tunnels daripada mengekspos port database ke internet.",
            "Cloudflare Zero Trust, cloudflared CLI, Linux Firewall (UFW), MariaDB"
        ),
        (
            5,
            "Fullstack & Data Integrity",
            "Timezone Hell UTC vs WIB: Misteri Turunnya Skor KPI Karyawan",
            "Pekerjaan yang diselesaikan sebelum jam 17:00 WIB tercatat terlambat (late) 1 hari di sistem pelaporan, menyebabkan skor performa karyawan anjlok tanpa kesalahan dari karyawan.",
            "Frontend Next.js mengirimkan tanggal dengan new Date().toISOString() yang otomatis mengonversi jam lokal WIB (+7) ke UTC, menyebabkan tanggal melompat mundur atau maju saat diparsing oleh server.",
            "1. Menghilangkan double-offset serialization di Next.js API route.\n2. Standardisasi format MySQL DATETIME murni (YYYY-MM-DD HH:mm:ss).\n3. Menetapkan strict cutoff transaksi berbasis timezone Asia/Jakarta di database level.",
            "• Mengeliminasi 100% false-positive keterlambatan task.\n• Skor KPI pegawai pulih secara adil sesuai fakta lapangan.\n• Data timestamp antar modul logistik dan HR 100% sinkron.",
            "Jangan pernah mempercayakan konversi timezone otomatis ke JavaScript Client. Selalu lakukan standardisasi format dan timezone parsing eksplisit di API gateway.",
            "Next.js, React, Node.js, MySQL DATETIME, Timezone Handling (Asia/Jakarta)"
        ),
        (
            6,
            "Fullstack & System Design",
            "The Backdating Dilemma: Audit Timestamp vs Business Completion Date",
            "Developer menyelesaikan koding tepat waktu pada tanggal 19, namun baru sempat mengklik tombol 'Complete' di sistem manajemen proyek pada tanggal 26. Sistem mencatat tanggal 26 sehingga dianggap terlambat 7 hari.",
            "Skema database hanya menyediakan satu kolom updated_at/completed_at yang otomatis diisi oleh NOW() pada saat mutasi data dilakukan (menggabungkan audit waktu klik dengan waktu kerja).",
            "Merekayasa ulang skema Strata & API MCP: memisahkan completedAt (tanggal faktual pekerjaan selesai) dengan activity_log (pencatatan kapan mutasi diinput). Ditambah fitur agreedDueDate & override reason.",
            "• Sistem mampu menampung update pekerjaan backdate secara sah.\n• Skor performa KPI 100% akurat sesuai tanggal deploy sesungguhnya.\n• Kepatuhan audit tetap terjaga karena riwayat klik tetap tersimpan.",
            "Pisahkan selalu 'Waktu Sistem Mencatat' (Audit Trail) dengan 'Waktu Kejadian Bisnis' (Business Effective Date) dalam arsitektur ERP & Project Management.",
            "Strata PMS, Model Context Protocol, REST API Design, Database Normalization"
        ),
        (
            7,
            "Fullstack & ERP Integration",
            "Hybrid Caching: Mencegah Overdue Scan Kedatangan Gudang (Dynamics 365)",
            "Aplikasi Web-OTD di gudang mengalami delay 3-5 detik saat tim lapangan men-scan barcode Purchase Order untuk gate-in, karena live call ke API Microsoft Dynamics 365 BC lambat saat traffic tinggi.",
            "Memanggil ERP API secara langsung pada setiap keystroke user di mobile/web scanner tanpa caching layer dan tanpa debounce request.",
            "Mengimplementasikan Arsitektur Hybrid Caching: Staging database lokal diupdate via sync worker untuk pencarian cepat (<100ms), lalu API Business Central live hanya dipanggil on-demand saat submit akhir sebagai validasi.",
            "• Waktu pencarian PO di gudang terpangkas dari 4 detik menjadi <150 milidetik.\n• Tidak terjadi antrian truk pengiriman di gerbang gudang.\n• Menghemat kuota API request ERP hingga 80%.",
            "Jangan pasang ERP kompleks sebagai backend query langsung aplikasi operasional frontline. Gunakan Read-Optimized Cache untuk kecepatan dan Live Verification untuk konsistensi.",
            "Microsoft Dynamics 365 Business Central API, Next.js, MariaDB Staging, Web Scanner"
        ),
        (
            8,
            "Data Engineering & SQL",
            "Zero Data Drift: Multi-Periodisitas KPI (Bulan vs Kuartal vs Tahun)",
            "Beberapa divisi dievaluasi bulanan, divisi lain kuartalan (Procurement) dan tahunan. Saat digabung dalam satu Master View pelaporan, kalkulasi bobot pecah (total bobot != 100%) dan data historis bergeser.",
            "Kueri agregasi SQL mengasumsikan seluruh baris indikator memiliki periodisitas bulanan seragam, sehingga indikator kuartalan membagi bobot secara tidak proporsional di bulan berjalan.",
            "Merancang Migration 018: menambahkan kolom Tipe_Periode_Indikator dan dynamic weight normalization engine di SQL view (v_target_cross_periode), menjamin total bobot selalu 100% dengan audit deviasi 0.",
            "• 4.437 baris historical data diaudit dengan 0 selisih (zero drift).\n• Fleksibilitas HR dalam mengukur indikator kuartal/semester/tahunan tanpa kustomisasi kode per divisi.",
            "Skema reporting enterprise harus tahan terhadap perbedaan siklus bisnis. Selesaikan normalisasi bobot di level SQL View idempotent agar frontend tetap simpel.",
            "MySQL Advanced Views, SQL Window Functions, Database Migration, Data Quality Audit"
        ),
        (
            9,
            "Fullstack & Security",
            "Dynamic Cross-Org RBAC: Membuka Akses Lintas Divisi Tanpa Hardcoding",
            "Tim eksekutif dan auditor Procurement butuh memantau data performa divisi lain, tetapi sistem keamanan awal mengisolasi akses secara kaku berdasarkan departemen login pengguna.",
            "Otorisasi di-hardcode dalam klausa WHERE pada kode frontend dashboard, sehingga penambahan hak akses khusus memerlukan deploy kode baru dan berisiko membocorkan data departemen lain.",
            "Membangun skema tabel kpi_access_delegations dan mengintegrasikan subquery Delegated_Viewers langsung di v_master_view_v2. Diuji dengan automated test suite 13 skenario otorisasi RBAC.",
            "• Delegasi akses dapat diatur dinamis via Admin Console tanpa deploy ulang.\n• Keamanan data terjamin 100% lolos uji regresi hak akses.\n• Menghilangkan duplikasi query di berbagai frontend.",
            "Pindahkan logika otorisasi multi-tenant/cross-department ke dalam Dynamic Security Views, bukan disebar di puluhan komponen frontend.",
            "Role-Based Access Control (RBAC), MySQL Security Views, Next.js Auth, QA Automation"
        ),
        (
            10,
            "Fullstack & Logistics",
            "Filter Cerdas PO Fisik vs Non-Fisik: Mencegah PO Menggantung di Gudang",
            "Status PO di aplikasi gudang Web-OTD tidak pernah mencapai 100% selesai karena tercampur baris non-fisik seperti 'Biaya Ekspedisi', 'Biaya Asuransi', dan 'Jasa Pemasangan'.",
            "Payload ERP Purchase Lines mencampur Item fisik, GL Account, dan Item Charge dalam array data yang sama tanpa flag klasifikasi penerimaan fisik gudang.",
            "Membangun filter pipeline cerdas: Mengelompokkan item bertipe Item fisik & Fixed Asset untuk diverifikasi fisik oleh petugas, dan otomatis mengecualikan item non-fisik/jasa dari syarat scan gate-in.",
            "• Antarmuka staf gudang bersih 100% hanya menampilkan barang nyata.\n• Status PO Released/Received menjadi akurat secara real-time.\n• Laporan On-Time Delivery (OTD) perusahaan menjadi kredibel.",
            "Pahami model domain bisnis fisik sebelum mengonsumsi data ERP. Apa yang dianggap 'item pembelian' oleh bagian Finance belum tentu 'barang fisik' bagi tim Logistik.",
            "Supply Chain & Logistics Workflow, Data Parsing, Next.js, Business Logic Design"
        ),
        (
            11,
            "AI Automation & LLM",
            "Membangun Custom MCP Server untuk Mengontrol ERP & PMS via LLM",
            "AI Coding Agent (seperti Claude / Antigravity) tidak bisa memantau project internal, membuat task, atau memperbarui timeline karena sistem PM berada di server private tanpa antarmuka AI standar.",
            "LLM out-of-the-box hanya bisa membaca file lokal dan tidak memiliki protokol standar untuk berinteraksi dengan sistem manajemen proyek internal perusahaan secara aman dan terstruktur.",
            "Mengimplementasikan Model Context Protocol (MCP) Server khusus untuk Strata PMS. Menyediakan 25+ callable tools dengan skema JSON terstruktur (create_tasks, update_task, log_work_note, dll).",
            "• AI Agent kini bertindak sebagai autonomous project assistant.\n• Seluruh commit kode otomatis dicatat dan diselaraskan ke KPI.\n• Peningkatan efisiensi operasional developer hingga lebih dari 10x lipat.",
            "Masa depan AI bukan sekadar tanya-jawab di chat, melainkan 'Agentic Tool-Use' melalui Model Context Protocol (MCP) yang langsung terkoneksi ke infrastruktur bisnis.",
            "Model Context Protocol (MCP), LLM Agent Architecture, JSON Schema, REST RPC"
        ),
        (
            12,
            "Data Science & Analytics",
            "Algoritma Cap 120% Non-Linier: Mencegah Manipulasi Skor Metrik Tim",
            "Seorang individu yang menyelesaikan satu tugas gampang sebesar 300% mendapatkan skor akhir sempurna (5.00), padahal target inti yang bernilai strategis gagal total.",
            "Formula linier standar (Actual / Target * Weight) memiliki kelemahan mendasar: over-performance pada satu indikator bisa menutupi under-performance fatal pada indikator lainnya.",
            "Menerapkan algoritma Cap 120% matematis: Skor indikator dibatasi maksimal 120% (1.20) dari target, lalu skor dikonversi ke skala proporsional 1-5 pada level agregasi akhir seluruh bobot.",
            "• Mencegah manipulasi metrik (gaming the system).\n• Mendorong pencapaian menyeluruh di setiap KPI strategis perusahaan.\n• Hasil penilaian menjadi adil, seimbang, dan diterima oleh seluruh stakeholder.",
            "Metrik yang salah akan menciptakan perilaku yang salah (Goodhart's Law). Selalu terapkan batas cap dan pembobotan proporsional pada sistem evaluasi kinerja.",
            "Business Analytics, Mathematical Modeling, KPI Formula, Data Visualization"
        )
    ]

    fill_map = {
        "AI Automation & Systems": badge_ai,
        "DevOps & Automation": badge_devops,
        "DevOps & Security": badge_devops,
        "Fullstack & Database": badge_fullstack,
        "Fullstack & Data Integrity": badge_fullstack,
        "Fullstack & System Design": badge_fullstack,
        "Fullstack & ERP Integration": badge_fullstack,
        "Fullstack & Security": badge_fullstack,
        "Fullstack & Logistics": badge_fullstack,
        "Data Engineering & SQL": badge_data,
        "AI Automation & LLM": badge_ai,
        "Data Science & Analytics": badge_data
    }

    current_row = 5
    for case in cases_data:
        ws1.row_dimensions[current_row].height = 95
        row_fill = accent_fill_1 if current_row % 2 == 1 else accent_fill_2

        for col_idx, val in enumerate(case, 1):
            cell = ws1.cell(row=current_row, column=col_idx, value=val)
            cell.font = cell_font
            cell.border = border_thin
            cell.fill = row_fill

            # Alignment logic
            if col_idx == 1:
                cell.alignment = Alignment(horizontal="center", vertical="top")
                cell.font = cell_bold
            elif col_idx == 2:
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                cell.fill = fill_map.get(val, row_fill)
                cell.font = cell_bold
            elif col_idx == 3:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.font = cell_bold
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        current_row += 1

    # Freeze panes on row 5
    ws1.freeze_panes = "A5"

    # =========================================================================
    # SHEET 2: LinkedIn Content Calendar & Hooks
    # =========================================================================
    ws2 = wb.create_sheet(title="LinkedIn Content Plan")
    ws2.views.sheetView[0].showGridLines = True

    ws2['A1'] = "LINKEDIN CONTENT CALENDAR & STORYTELLING PLAYBOOK"
    ws2['A1'].font = title_font
    ws2['A2'] = "Blueprint konten berbasis STAR Framework (Situation, Task, Action, Result) untuk personal branding sebagai AI Automation, Fullstack & DevOps Engineer."
    ws2['A2'].font = subtitle_font

    headers_2 = [
        ("Post ID", 8),
        ("Target Branding Persona", 22),
        ("Catchy LinkedIn Hook (3 Baris Pembuka)", 40),
        ("Format Konten Rekomendasi", 20),
        ("Framework STAR: Situation & Task", 35),
        ("Framework STAR: Action & Engineering", 45),
        ("Framework STAR: Result & Business Impact", 35),
        ("Call to Action (CTA) Penutup", 32),
        ("Rekomendasi Hashtags", 30)
    ]

    for col_idx, (h_name, width) in enumerate(headers_2, 1):
        cell = ws2.cell(row=4, column=col_idx, value=h_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
        ws2.column_dimensions[get_column_letter(col_idx)].width = width

    ws2.row_dimensions[4].height = 28

    linkedin_posts = [
        (
            "POST-01",
            "AI Automation Specialist",
            "🚨 'Jangan biarkan AI Agent mengotori database produksi Anda.'\n\nBanyak developer mulai pakai AI agent (Claude Code, Antigravity) untuk coding. Masalahnya muncul saat AI disuruh catat progress: database penuh spam task kecil & duplikasi.\n\nBegini cara saya mengatasinya dengan 2-Phase Staging Buffer: 👇",
            "Carousel (PDF Slide)",
            "SITUATION & TASK:\nMenggunakan 3 AI agent berbeda untuk multi-project. Target: Seluruh progress harian tercatat di sistem ERP/PMS tanpa human error dan tanpa bikin database berantakan.",
            "ACTION (ENGINEERING):\nMembuat private buffer table (PENDING status). Agent hanya boleh log catatan teknis mentah. Di akhir siklus, AI Agent melakukan konsolidasi semantik, mendeteksi duplikat, lalu melakukan batch promote ke official project task.",
            "RESULT:\n100% task rapi, zero redundancy, dan tracking pekerjaan otomatis sinkron dengan metrik performa tanpa buang waktu rekap manual.",
            "Kalian yang pakai AI coding agent, gimana cara tracking pekerjaan hariannya? Diskusi di komentar yuk! 👇",
            "#AIEngineering #Automation #ModelContextProtocol #Productivity #SoftwareArchitecture"
        ),
        (
            "POST-02",
            "DevOps & Automation",
            "⚡ 'Dari 16 menit jadi 20 detik.'\n\nWorkflow automasi n8n perusahaan kami tiba-tiba sering crash dengan error HTTP 503 & HTTP 400 Grid Limits. Setelah dicek, penyebabnya klasik tapi fatal.\n\nBukan salah servernya, tapi arsitektur datanya. Ini bedah kasusnya: 🧵",
            "Storytelling Breakdown",
            "SITUATION & TASK:\nWorkflow sinkronisasi form approval FPPP ke Google Sheets memakan waktu 16+ menit dan sering mati di tengah jalan saat beban jam kerja puncak.",
            "ACTION (ENGINEERING):\n1. Bongkar query SQL: ganti SELECT * historis 15.400 baris dengan View Delta aktif (~480 baris).\n2. Pindahkan logika heavy-aggregation dari node n8n ke SQL database view.\n3. Implementasi auto-retry guard dengan exponential backoff.",
            "RESULT:\n• Runtime terpangkas 98% (16 min ➡️ 20 sec).\n• Error rate 0%.\n• Konsumsi CPU container n8n turun drastis.",
            "Pernah ngalamin workflow low-code crash karena data membengkak? Apa solusi pertama yang kalian lakuin?",
            "#DevOps #n8n #Automation #DatabaseTuning #SystemArchitecture"
        ),
        (
            "POST-03",
            "Fullstack & Database",
            "🤯 'Waktu kami cek log persetujuan, urutannya terbalik semua.'\n\nAda 5.367 baris log komentar approval yang posisinya acak. Kenapa? Karena kami terlalu percaya sama Auto-Increment ID di sistem terdistribusi.\n\nPelajaran mahal tentang asynchronous event: 👇",
            "Case Study + Snippet",
            "SITUATION & TASK:\nForm approval enterprise melibatkan multi-approver paralel. Audit trail di dashboard harus menampilkan kronologi persetujuan yang 100% valid secara hukum dan bisnis.",
            "ACTION (ENGINEERING):\nAuto-increment ID mencatat kapan webhook diterima, bukan kapan approval diklik user. Solusinya: bangun View dengan multi-tier fallback: ORDER BY COALESCE(task_start_date, action_date), id.",
            "RESULT:\n5.367 baris approval tersusun sempurna sesuai kronologi nyata, sistem lolos uji audit kepatuhan tanpa kendala.",
            "Tips: Jangan pernah percaya auto-increment ID untuk urutan logikal bisnis di sistem async. Ada yang pernah kena jebakan ini?",
            "#SoftwareEngineering #Database #SQL #SystemDesign #BackendDeveloper"
        ),
        (
            "POST-04",
            "DevOps & Cloud Security",
            "🔒 'Stop buka port 3306 database cloud Anda ke internet publik!'\n\nBanyak yang masih pakai IP Whitelist atau bahkan 0.0.0.0/0 cuma biar local dev & AI agent bisa query ke database cloud.\n\nIni arsitektur aman menggunakan Cloudflare Zero Trust Tunnel: 🛠️",
            "Architecture Diagram",
            "SITUATION & TASK:\nAI Agent lokal butuh direct database access ke MariaDB server di cloud secara otomatis, tanpa membuka port publik dan tanpa VPN korporat yang rumit.",
            "ACTION (ENGINEERING):\n1. Setup Cloudflare Access TCP Tunnel di server.\n2. Jalankan background daemon cloudflared di local machine.\n3. Tunneling mTLS terenkripsi menghubungkan localhost:3306 ke remote cloud secara aman.",
            "RESULT:\n0 open ports di firewall publik. Bebas dari serangan brute-force scanning bot internet, latensi stabil di bawah 40ms.",
            "Gimana cara kalian connect ke database staging/prod dari local machine sekarang? Masih pakai SSH tunnel atau Zero Trust?",
            "#DevOps #CyberSecurity #Cloudflare #ZeroTrust #CloudArchitecture"
        ),
        (
            "POST-05",
            "Fullstack Senior Engineer",
            "🕒 'Selisih 7 jam yang bikin skor performa karyawan anjlok.'\n\nPernah dengar istilah 'Timezone Hell'? Ini kasus nyatanya: Karyawan submit kerjaan jam 4 sore WIB, tapi sistem malah mencatat itu jam 9 malam kemarin (UTC), lalu ditandai LATE.\n\nIni cara kami membasmi bug timezone selamanya: 👇",
            "Carousel (PDF Slide)",
            "SITUATION & TASK:\nMenyelaraskan pelaporan performa antara frontend Next.js, API backend, dan MySQL database agar tidak ada false overdue flag.",
            "ACTION (ENGINEERING):\nHilangkan toISOString() mentah di client. Terapkan explicit MySQL DATETIME format (YYYY-MM-DD HH:mm:ss) tanpa double offset timezone Asia/Jakarta di API contract layer.",
            "RESULT:\nSkor performa kembali akurat 100%, eliminasi komplain karyawan, dan standardisasi tanggal di seluruh modul perusahaan.",
            "Timezone bug adalah musuh bebuyutan fullstack developer. Share pengalaman timezone paling aneh yang pernah kalian temui!",
            "#Fullstack #WebDevelopment #JavaScript #Nextjs #CleanCode"
        ),
        (
            "POST-06",
            "Fullstack & System Architect",
            "📝 'Kerjaan selesai tepat waktu di lapangan, tapi dicap terlambat oleh sistem.'\n\nMasalah klasik di ERP/PM: Developer koding beres tgl 19, deploy tgl 19. Tapi baru sempat centang tombol 'Done' tgl 26. Sistem otomatis mencatat tgl 26.\n\nSolusinya: Memisahkan Audit Date vs Business Date. 💡",
            "Storytelling Breakdown",
            "SITUATION & TASK:\nSistem evaluasi performa membutuhkan data penyelesaian tugas yang faktual tanpa merusak integritas audit jejak rekam sistem.",
            "ACTION (ENGINEERING):\nRefactoring skema database: Kolom 'completedAt' mencatat kapan kerjaan fisik selesai (business date), sedangkan tabel activity log tetap merekam kapan tombol diklik (audit date).",
            "RESULT:\nSkor KPI naik dari 1.97 menjadi 5.00/5.00 secara sah, transparan, dan dapat diverifikasi oleh HR serta auditor independen.",
            "Apakah sistem manajemen proyek di kantor kalian memfasilitasi backdate penyelesaian faktual? Mari bahas arsitekturnya!",
            "#SystemDesign #DatabaseDesign #ERP #Productivity #EngineeringLeadership"
        ),
        (
            "POST-07",
            "Fullstack & Enterprise ERP",
            "📦 'Truk logistik antre di gerbang gudang karena API ERP lemot.'\n\nSaat tim gudang scan barcode penerimaan barang, aplikasi Web-OTD delay 4 detik karena menembak live API Microsoft Dynamics 365 BC secara bersamaan.\n\nSolusinya? Hybrid Read-Optimized Caching. 🚀",
            "Architecture Breakdown",
            "SITUATION & TASK:\nMempercepat proses gate-in barang gudang agar scan barcode instan (<200ms) tanpa mengorbankan validasi ketat dari ERP pusat.",
            "ACTION (ENGINEERING):\nPisahkan Read & Write path: Background worker mensinkronkan master PO ke local MariaDB cache untuk instant search di frontline. Live API ERP hanya dipanggil sekali saat submit final.",
            "RESULT:\nResponse time pencarian PO turun drastis dari 4 detik menjadi <150 milidetik. Antrian gerbang gudang hilang total.",
            "Kapan terakhir kali arsitektur caching menyelamatkan operasi fisik bisnis kalian? Cerita di bawah ya! 👇",
            "#ERP #MicrosoftDynamics #Fullstack #PerformanceOptimization #SupplyChain"
        ),
        (
            "POST-08",
            "Data Engineer & Architect",
            "📊 'Zero Data Drift: Mengelola Indikator KPI Bulanan vs Kuartalan dalam Satu Database View.'\n\nHR minta satu dashboard terpadu: Divisi IT dinilai tiap bulan, tapi Procurement tiap kuartal. Begitu digabung, bobot indikator berantakan dan data masa lalu berubah.\n\nIni solusinya dengan SQL View Normalization: 🔍",
            "Deep Dive Case Study",
            "SITUATION & TASK:\nMenyatukan metrik multi-periodisitas (bulanan, kuartalan, tahunan) ke dalam single master view tanpa mengubah angka historis yang sudah disahkan BOD.",
            "ACTION (ENGINEERING):\nMendesain migration schema 'Tipe_Periode_Indikator' dengan kalkulasi bobot proporsional dinamis di level view SQL (v_target_cross_periode). Dilengkapi script audit perbandingan baris.",
            "RESULT:\nAudit 4.437 baris historis membuktikan 0 deviasi (Zero Data Drift). HR bisa mengukur target fleksibel tanpa developer coding ulang view.",
            "Bagaimana cara tim data kalian menjaga konsistensi data historis saat ada perubahan skema metrik di tengah tahun?",
            "#DataEngineering #SQL #Analytics #DataQuality #BusinessIntelligence"
        ),
        (
            "POST-09",
            "Fullstack & Cyber Security",
            "🛡️ 'Buka akses cross-department tanpa bikin query database jadi spaghetti.'\n\nManajemen butuh tim Procurement memantau metrik divisi lain. Cara salah: nambah if-else di kode frontend. Cara benar: Dynamic SQL View RBAC.\n\nIni arsitektur keamanannya: 🔐",
            "Code & Architecture Snippet",
            "SITUATION & TASK:\nMenyediakan delegasi hak akses lintas organisasi yang dinamis, secure, dan dapat diubah kapan saja tanpa perlu deploy kode frontend baru.",
            "ACTION (ENGINEERING):\nMembuat tabel kpi_access_delegations dan memasukkan subquery dinamis Delegated_Viewers ke dalam master view. Dilengkapi automated test suite 13 skenario akses.",
            "RESULT:\nAdmin HR bisa mengatur izin lintas departemen secara visual via dashboard admin. Nol celah kebocoran data divisi non-terkait.",
            "Lebih suka handle RBAC di level Application Layer (Middleware) atau Database Layer (Row-Level Security/Views)? Kenapa?",
            "#CyberSecurity #RBAC #SoftwareEngineering #Backend #WebSecurity"
        ),
        (
            "POST-10",
            "Fullstack & Logistics",
            "🚚 'Jangan pernah telan raw data ERP mentah-mentah.'\n\nDi sistem gudang kami, status Purchase Order tidak pernah bisa 100% selesai. Usut punya usut, ada baris 'Biaya Ongkir' dan 'Jasa Pasang' yang dianggap barang fisik oleh sistem.\n\nPentingnya domain filtering di data pipeline: 👇",
            "Storytelling Breakdown",
            "SITUATION & TASK:\nMemastikan modul penerimaan gudang (Gate-In) hanya memproses barang fisik riil yang memang masuk ke area rak penyimpanan.",
            "ACTION (ENGINEERING):\nMembedah schema Purchase Lines ERP. Memisahkan tipe 'Item' & 'Fixed Asset' fisik dari 'GL Account' / 'Item Charge'. Menambahkan exception handler untuk material tempered glass protection.",
            "RESULT:\nDaftar scan gudang 100% bersih, status PO Released akurat, metrik ketepatan waktu pengiriman (OTD) perusahaan menjadi valid.",
            "Pernah nemu bug aneh gara-gara perbedaan definisi 'istilah bisnis' antara orang IT, Finance, dan Lapangan?",
            "#Fullstack #LogisticsTech #DataPipeline #DomainDrivenDesign #SoftwareEngineering"
        ),
        (
            "POST-11",
            "AI Automation & LLM Engineer",
            "🤖 'Dari chatbot tanya-jawab jadi AI Worker sungguhan.'\n\nBagaimana cara saya membuat AI Agent bisa otomatis memeriksa backlog proyek, membuat task, dan memperbarui timeline di database internal?\n\nKuncinya ada di Model Context Protocol (MCP): 🌐",
            "Step-by-Step Guide",
            "SITUATION & TASK:\nMenghubungkan LLM (Claude/Antigravity) ke sistem manajemen proyek internal perusahaan agar AI bisa bekerja secara mandiri dan terstruktur.",
            "ACTION (ENGINEERING):\nMembuat custom MCP Server dengan 25+ tool schemas (JSON Schema validation). Agent bisa inspect rencana kerja (get_plan), mencatat pekerjaan, dan mempromosikannya ke database produksi.",
            "RESULT:\nProduktivitas tim meningkat drastis. Developer cukup review hasil kerja AI, seluruh pencatatan proyek dan kepatuhan SLA selesai otomatis.",
            "Sudah mulai eksplor Model Context Protocol (MCP) untuk workflow harian kalian? Fitur apa yang paling ngebantu?",
            "#AIAgents #ModelContextProtocol #LLM #Automation #FutureOfWork"
        ),
        (
            "POST-12",
            "Data Science & Performance",
            "🎯 'Kenapa rumus KPI linier justru merusak motivasi tim Anda (Goodhart\'s Law).' \n\nSatu orang bisa dapat nilai 5.00 sempurna hanya dengan ngerjain tugas mudah sebanyak 300%, sementara tugas prioritas diabaikan. \n\nIni matematika di balik Algoritma Cap 120%: 📈",
            "Carousel (PDF Slide)",
            "SITUATION & TASK:\nMerancang sistem evaluasi performa karyawan berbasis data yang adil, seimbang, dan tidak bisa diakali (anti-gaming).",
            "ACTION (ENGINEERING):\nMengganti rumus linier murni dengan Cap 120%: Score = min(1.20, Actual / Target) * Weight. Konversi nilai ke skala 1-5 dilakukan pada aggregasi akhir seluruh bobot indikator.",
            "RESULT:\nMendorong karyawan mencapai target di seluruh pilar indikator secara proporsional. Penilaian performa diterima transparan oleh manajemen dan karyawan.",
            "Pernah melihat sistem KPI yang malah bikin orang fokus ke hal yang salah? Bagaimana formula yang ideal menurut kalian?",
            "#DataAnalytics #PeopleAnalytics #GoodhartsLaw #KPI #ManagementEngineering"
        )
    ]

    current_row = 5
    for post in linkedin_posts:
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
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                cell.font = cell_bold
            elif col_idx == 3:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.font = cell_bold
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        current_row += 1

    ws2.freeze_panes = "A5"

    # =========================================================================
    # SHEET 3: Tech Stack & Architecture Map
    # =========================================================================
    ws3 = wb.create_sheet(title="Architecture & Tech Stack")
    ws3.views.sheetView[0].showGridLines = True

    ws3['A1'] = "ECOSYSTEM ARCHITECTURE & ENGINEERING PATTERNS"
    ws3['A1'].font = title_font
    ws3['A2'] = "Peta komponen teknologi, fungsi arsitektural, dan pola desain yang diterapkan dalam memecahkan masalah riil di ekosistem perusahaan."
    ws3['A2'].font = subtitle_font

    headers_3 = [
        ("Layer / Komponen", 22),
        ("Teknologi Spesifik", 25),
        ("Peran Arsitektural", 35),
        ("Tantangan Produksi yang Dihadapi", 38),
        ("Pola Solusi / Design Pattern Teruji", 40)
    ]

    for col_idx, (h_name, width) in enumerate(headers_3, 1):
        cell = ws3.cell(row=4, column=col_idx, value=h_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
        ws3.column_dimensions[get_column_letter(col_idx)].width = width

    ws3.row_dimensions[4].height = 28

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
            "Delta Filtering at Source (SQL View Level) + Throttled Batch Processing + Throttling Retry Guards."
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
        ws3.row_dimensions[current_row].height = 70
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

    output_file = r"c:\Website\Employee Project & Stage Tracker\AI_Automation_Fullstack_DevOps_Case_Studies.xlsx"
    wb.save(output_file)
    print(f"Excel workbook successfully created at: {output_file}")

if __name__ == "__main__":
    build_excel()
