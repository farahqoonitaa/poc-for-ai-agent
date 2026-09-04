#!/usr/bin/env python3
"""
Generate an executive presentation deck (.pptx) for Konsolidasi Analyzer & Review Dashboard
Solution Architecture & Stakeholder Alignment.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck(output_path="Konsolidasi_Analyzer_Solution_Architecture.pptx"):
    prs = Presentation()
    # 16:9 widescreen
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_slide_layout = prs.slide_layouts[6]

    # Color Palette
    COLOR_DEEP_FOREST = RGBColor(15, 42, 29)    # #0F2A1D
    COLOR_FOREST      = RGBColor(27, 67, 50)    # #1B4332
    COLOR_SAGE        = RGBColor(82, 121, 111)  # #52796F
    COLOR_GOLD        = RGBColor(184, 144, 46)  # #B8902E
    COLOR_BG          = RGBColor(246, 243, 236) # #F6F3EC
    COLOR_WHITE       = RGBColor(255, 255, 255)
    COLOR_TEXT_DARK   = RGBColor(34, 30, 24)    # #221E18
    COLOR_TEXT_MUTED  = RGBColor(92, 86, 74)    # #5C564A
    COLOR_CARD_BG     = RGBColor(255, 255, 255)
    COLOR_CARD_BORDER = RGBColor(226, 219, 208)
    COLOR_DANGER      = RGBColor(166, 64, 42)
    COLOR_OK          = RGBColor(45, 106, 79)

    def set_bg(slide, color=COLOR_BG):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()
        return bg

    def add_header(slide, title_text, category_text="KONSOLIDASI ANALYZER — SOLUTION ARCHITECTURE"):
        # Category pill
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(10), Inches(0.4))
        tf = cat_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = category_text.upper()
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.7), Inches(0.7))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = COLOR_DEEP_FOREST

    def add_card(slide, left, top, width, height, title, body_bullets, accent_color=None):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_BG
        card.line.color.rgb = accent_color if accent_color else COLOR_CARD_BORDER
        card.line.width = Pt(1.5 if accent_color else 1.0)

        # Content text
        pad = Inches(0.22)
        tb = slide.shapes.add_textbox(left + pad, top + pad, width - (pad * 2), height - (pad * 2))
        tf = tb.text_frame
        tf.word_wrap = True

        p_title = tf.paragraphs[0]
        p_title.text = title
        p_title.font.size = Pt(14)
        p_title.font.bold = True
        p_title.font.color.rgb = accent_color if accent_color else COLOR_DEEP_FOREST
        p_title.space_after = Pt(8)

        for bullet in body_bullets:
            p = tf.add_paragraph()
            p.text = "• " + bullet
            p.font.size = Pt(11.5)
            p.font.color.rgb = COLOR_TEXT_MUTED
            p.space_after = Pt(4)

    # -------------------------------------------------------------
    # SLIDE 1: TITLE SLIDE (Dark Forest Executive)
    # -------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_slide_layout)
    set_bg(s1, COLOR_DEEP_FOREST)

    # Gold accent line
    line = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(1.8), Inches(2.0), Inches(0.06))
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_GOLD
    line.line.fill.background()

    tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(10.8), Inches(2.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p = tf1.paragraphs[0]
    p.text = "Konsolidasi Analyzer & Review Workspace"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    p2 = tf1.add_paragraph()
    p2.text = "AI Forensic Copilot & Human-in-the-Loop Audit Intelligence"
    p2.font.size = Pt(20)
    p2.font.color.rgb = COLOR_GOLD
    p2.space_before = Pt(10)

    p3 = tf1.add_paragraph()
    p3.text = "Solution Architecture & Stakeholder Alignment for Multi-Entity Holding Financial Consolidation"
    p3.font.size = Pt(13)
    p3.font.color.rgb = RGBColor(210, 205, 195)
    p3.space_before = Pt(16)

    # Meta box
    meta_box = s1.shapes.add_textbox(Inches(1.2), Inches(5.6), Inches(10.8), Inches(1.0))
    tf_meta = meta_box.text_frame
    p_meta = tf_meta.paragraphs[0]
    p_meta.text = "Entitas Sasaran: Yayasan (Holding Induk) & 5 Anak Perusahaan (PT Anak I – V)\nFokus Regulasi: PSAK 65 (Konsolidasian/NCI), PSAK 7 (Transaksi Berelasi), POJK Lembaga Keuangan"
    p_meta.font.size = Pt(11)
    p_meta.font.color.rgb = RGBColor(170, 165, 155)

    # -------------------------------------------------------------
    # SLIDE 2: THE EXECUTIVE CHALLENGE
    # -------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_slide_layout)
    set_bg(s2)
    add_header(s2, "Tantangan Konsolidasi Finansial Multientitas", "Latar Belakang & Masalah")

    c_w = Inches(3.64)
    c_h = Inches(4.8)
    top_y = Inches(1.7)

    add_card(s2, Inches(0.8), top_y, c_w, c_h, "1. Kebocoran Alokasi NCI (PSAK 65)", [
        "Kepemilikan minoritas (NCI 22% - 40%) pada 4 anak perusahaan rawan ketidaksesuaian matematis.",
        "Kasus PT Anak II: Laba NCI dilaporkan 24.4 M vs porsi proporsional 18.2 M (deviasi +34.1%).",
        "Dampak: Understatement laba yang dapat diatribusikan ke induk yayasan atau dividen terselubung.",
        "Spreadsheet manual tidak memiliki validasi otomatis atas hak efektif pemegang saham non-pengendali."
    ], COLOR_DANGER)

    add_card(s2, Inches(4.84), top_y, c_w, c_h, "2. Lonjakan Piutang Berelasi (PSAK 7)", [
        "Piutang pihak berelasi (related-party) konsolidasi melonjak tajam tanpa pertumbuhan omzet sepadan.",
        "Kasus PT Anak II: Piutang relasi melesat +158.8% YoY sementara pendapatan hanya tumbuh +4.9%.",
        "Risiko likuiditas terserap ke entitas afiliasi lain secara non-arms-length.",
        "Potensi bad debts tersembunyi dan komplikasi eliminasi intercompany pada akhir periode."
    ], COLOR_GOLD)

    add_card(s2, Inches(8.88), top_y, c_w, c_h, "3. Ketiadaan Audit Trail Terintegrasi", [
        "Review konsolidasi terfragmentasi antara email, catatan kertas kerja terpisah, dan spreadsheet.",
        "Komite Audit dan CFO tidak memiliki visibilitas atas tindak lanjut klarifikasi anomali.",
        "Tidak ada tracking formal atas temuan yang memerlukan jurnal penyesuaian audit vs justifikasi operasional.",
        "Keterlambatan mitigasi risiko hingga temuan muncul pada audit eksternal formal."
    ], COLOR_FOREST)

    # -------------------------------------------------------------
    # SLIDE 3: THE DUAL-ENGINE SOLUTION VISION
    # -------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_slide_layout)
    set_bg(s3)
    add_header(s3, "Solusi: Arsitektur Dual-Engine + Review Workspace", "Visi Arsitektur")

    add_card(s3, Inches(0.8), top_y, Inches(5.6), Inches(4.8), "Engine 1: Deterministic Accounting Engine", [
        "Perhitungan Konsolidasi Presisi 100%: Menjumlahkan revenue, laba, aset, liabilitas, ekuitas 6 entitas.",
        "Eliminasi Intercompany & Alokasi NCI: Formula baku PSAK 65 menghitung porsi induk vs hak minoritas.",
        "Forensic Rule Engine: Ambang deviasi matematis otomatis mendeteksi 5 pola anomali material.",
        "Zero Black-Box: Setiap angka memiliki sumber, as-of date, dan formula matematis yang dapat diverifikasi.",
        "Real-time Recalculation: Setiap perubahan angka di tabel langsung mengupdate rasio solvabilitas & profitabilitas."
    ], COLOR_FOREST)

    add_card(s3, Inches(6.8), top_y, Inches(5.6), Inches(4.8), "Engine 2: Cognitive AI Agent Copilot & Review", [
        "Conversational Copilot: Menjawab pertanyaan pimpinan secara kontekstual dalam Bahasa Indonesia analis.",
        "Forensic Reasoning & Rationale: Menjelaskan mengapa anomali terjadi dan implikasi risiko auditnya.",
        "Interactive Review Workspace: Setiap anomali memiliki status (Klarifikasi / Temuan / Disetujui) & catatan reviewer.",
        "Smart Action Buttons: Chat langsung mengarahkan dan menghighlight kartu review yang relevan di dashboard.",
        "OJK Sectoral Benchmarking: Membandingkan DER konsolidasi (0.96x) & ROA (5.4%) dengan standar industri."
    ], COLOR_GOLD)

    # -------------------------------------------------------------
    # SLIDE 4: END-TO-END SYSTEM ARCHITECTURE
    # -------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_slide_layout)
    set_bg(s4)
    add_header(s4, "Arsitektur Solusi 5-Tier End-to-End", "Technical Blueprint")

    tiers = [
        ("Tier 1: Ingestion Layer", ["Konektor ERP (SAP/Oracle)", "Excel / XBRL Parser", "Validasi Neraca CY & PY"]),
        ("Tier 2: Consolidation", ["Agregator Multientitas", "Matriks Eliminasi Relasi", "Kalkulasi NCI PSAK 65"]),
        ("Tier 3: Forensic Rules", ["Deteksi Mismatch NCI", "Lonjakan Piutang Relasi", "Spike Leverage & Margin"]),
        ("Tier 4: Agent Copilot", ["Multi-Agent Coordinator", "Forensic Explainer Agent", "Benchmarking Sektor OJK"]),
        ("Tier 5: Review UI", ["Chat Copilot Split View", "Review Action Cards", "Sign-Off & PDF Export"])
    ]

    t_w = Inches(2.26)
    t_h = Inches(4.8)
    for idx, (title, items) in enumerate(tiers):
        x = Inches(0.8 + idx * 2.42)
        add_card(s4, x, top_y, t_w, t_h, title, items, COLOR_DEEP_FOREST if idx % 2 == 0 else COLOR_FOREST)

    # -------------------------------------------------------------
    # SLIDE 5: FORENSIC ANOMALY DETECTION ENGINE
    # -------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_slide_layout)
    set_bg(s5)
    add_header(s5, "Mesin Deteksi Anomali Finansial & Ambang Batas", "Forensic Rules")

    rules_data = [
        ("1. NCI Profit Allocation Disparity", "Kritis (High)", "|Reported NCI - Computed NCI| > 8%", "Mencegah distorsi laba induk dan memastikan hak dividen pemegang saham minoritas terlindungi sesuai PSAK 65."),
        ("2. Related-Party Debt Spike", "Kritis (High)", "ΔPiutang Relasi - ΔRevenue > 30%", "Mendeteksi aliran likuiditas keluar ke entitas berelasi dan mencegah akumulasi kredit macet internal (PSAK 7)."),
        ("3. Leverage / DER Surge", "Sedang (Med)", "DER CY / DER PY - 1 > 18%", "Mengontrol risiko solvabilitas anak perusahaan sebelum melanggar debt covenant perbankan."),
        ("4. Profit Margin Distortion", "Rendah / Sedang", "|Margin CY - Margin PY| ≥ 3.5pp", "Mengidentifikasi keuntungan non-operasional satu kali (one-off) atau inefisiensi beban pokok operasional.")
    ]

    card_h = Inches(1.1)
    for idx, (r_name, sev, formula, desc) in enumerate(rules_data):
        y = Inches(1.7 + idx * 1.25)
        card = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), y, Inches(11.7), card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_BG
        card.line.color.rgb = COLOR_DANGER if "Kritis" in sev else COLOR_GOLD
        card.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(Inches(1.0), y + Inches(0.12), Inches(11.3), card_h - Inches(0.24))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = f"{r_name}  [{sev}]"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_DEEP_FOREST

        p2 = tf.add_paragraph()
        p2.text = f"Formula: {formula}  •  Tujuan Audit: {desc}"
        p2.font.size = Pt(10.8)
        p2.font.color.rgb = COLOR_TEXT_MUTED

    # -------------------------------------------------------------
    # SLIDE 6: REVIEW WORKSPACE & HITL WORKFLOW
    # -------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_slide_layout)
    set_bg(s6)
    add_header(s6, "Review Workspace: Hasil yang Perlu Di-Review", "Human-in-the-Loop")

    flow_cards = [
        ("Step 1: Automated Flagging", [
            "Rule Engine mengeksekusi pemeriksaan pada seluruh entitas.",
            "4 temuan anomali otomatis diklasifikasikan berdasarkan tingkat keparahan.",
            "Badge Kritis / Sedang / Rendah dimunculkan pada workspace."
        ]),
        ("Step 2: AI Rationale & Impact", [
            "Agent menghasilkan rekomendasi tindakan dan analisa dampak risiko.",
            "Referensi standar akuntansi PSAK 65 dan PSAK 7 dicantumkan.",
            "Chat copilot merangkum temuan untuk pimpinan secara instan."
        ]),
        ("Step 3: Reviewer Action", [
            "Auditor menetapkan status: Perlu Klarifikasi / Temuan / Disetujui.",
            "Input catatan auditor internal dan permintaan data pendukung.",
            "Fitur satu klik 'Kirim Notif Entitas' untuk surat konfirmasi."
        ]),
        ("Step 4: Formal Sign-off", [
            "Tracking progres: X dari Y temuan telah diselesaikan.",
            "Tombol pengesahan resmi penguncian kertas kerja telaah.",
            "Ekspor Lembar Review tercetak rapi siap rapat Komite Audit."
        ])
    ]

    for idx, (title, items) in enumerate(flow_cards):
        x = Inches(0.8 + idx * 2.96)
        add_card(s6, x, top_y, Inches(2.8), Inches(4.8), title, items, COLOR_FOREST)

    # -------------------------------------------------------------
    # SLIDE 7: SECTORAL & OJK BENCHMARKING
    # -------------------------------------------------------------
    s7 = prs.slides.add_slide(blank_slide_layout)
    set_bg(s7)
    add_header(s7, "Kesehatan Finansial Konsolidasi vs Benchmark OJK", "Analisis Sektoral")

    bench_cards = [
        ("Solvabilitas (DER)", "Konsolidasi: 0.96x", "Batas POJK: Max 5.0x", "STATUS: SANGAT AMAN", [
            "Leverage sangat konservatif, memberikan ruang ekspansi pembiayaan luas.",
            "Jauh di bawah batas risiko solvabilitas lembaga keuangan non-bank."
        ], COLOR_OK),
        ("Rentabilitas (ROA)", "Konsolidasi: 5.4%", "Benchmark Industri: 3.5% - 5.2%", "STATUS: OUTPERFORMING", [
            "Kemampuan menghasilkan laba dari aset berada di kuartil teratas industri.",
            "Laba konsolidasi tumbuh +14.2% YoY (325 M vs 284.7 M)."
        ], COLOR_OK),
        ("Transaksi Afiliasi (RP)", "Konsolidasi: 241 M", "Rasio terhadap Aset: 4.3%", "STATUS: MONITORING KHUSUS", [
            "Secara agregat terkendali, namun konsentrasi di PT Anak II (88 M) perlu audit khusus.",
            "Diperlukan pembatasan plafon kredit intercompany antar anak perusahaan."
        ], COLOR_GOLD)
    ]

    b_w = Inches(3.64)
    for idx, (title, val_entity, val_ind, status, details, color) in enumerate(bench_cards):
        x = Inches(0.8 + idx * 4.03)
        c = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, top_y, b_w, Inches(4.8))
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_CARD_BG
        c.line.color.rgb = color
        c.line.width = Pt(1.5)

        tb = s7.shapes.add_textbox(x + Inches(0.2), top_y + Inches(0.2), b_w - Inches(0.4), Inches(4.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = COLOR_DEEP_FOREST
        p.space_after = Pt(6)

        p_val = tf.add_paragraph()
        p_val.text = f"{val_entity}\n{val_ind}"
        p_val.font.size = Pt(11)
        p_val.font.bold = True
        p_val.font.color.rgb = COLOR_TEXT_DARK
        p_val.space_after = Pt(6)

        p_st = tf.add_paragraph()
        p_st.text = status
        p_st.font.size = Pt(10.5)
        p_st.font.bold = True
        p_st.font.color.rgb = color
        p_st.space_after = Pt(10)

        for d in details:
            p_d = tf.add_paragraph()
            p_d.text = "• " + d
            p_d.font.size = Pt(11)
            p_d.font.color.rgb = COLOR_TEXT_MUTED
            p_d.space_after = Pt(4)

    # -------------------------------------------------------------
    # SLIDE 8: STAKEHOLDER EXPECTATION ALIGNMENT MATRIX
    # -------------------------------------------------------------
    s8 = prs.slides.add_slide(blank_slide_layout)
    set_bg(s8)
    add_header(s8, "Matriks Keselarasan Ekspektasi Pemangku Kepentingan", "Stakeholder Alignment")

    matrix_items = [
        ("Komite Audit / Pengawas", "Mencegah temuan material & kebocoran NCI sebelum audit eksternal.", "Rule Engine otomatis memvalidasi alokasi laba PSAK 65 & menyajikan ranking anomali kritis."),
        ("Direktur Keuangan / CFO", "Mempercepat closing laporan konsolidasi tanpa risiko kesalahan manusia.", "Kompilasi rasio instan, simulasi angka interaktif, dan ringkasan eksekutif komite."),
        ("Tim Internal Audit", "Memiliki kertas kerja telaah terstruktur dengan audit trail yang sah.", "Review Workspace dengan status klarifikasi, catatan auditor, dan log pengesahan tertanggal."),
        ("CFO Anak Perusahaan", "Transparansi perhitungan eliminasi & rekonsiliasi yang adil.", "Rincian deviasi disajikan angka demi angka sehingga mudah diverifikasi dengan GL entitas."),
        ("IT & Governance", "Keamanan data laporan keuangan internal dan integritas sistem.", "Dapat berjalan offline tanpa kirim data ke LLM publik; append-only point-in-time snapshot.")
    ]

    for idx, (stakeholder, expectation, delivery) in enumerate(matrix_items):
        y = Inches(1.65 + idx * 1.0)
        c = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), y, Inches(11.7), Inches(0.9))
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_CARD_BG
        c.line.color.rgb = COLOR_CARD_BORDER

        tb = s8.shapes.add_textbox(Inches(1.0), y + Inches(0.08), Inches(11.3), Inches(0.75))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = stakeholder.upper()
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p2 = tf.add_paragraph()
        p2.text = f"Ekspektasi: \"{expectation}\"  ➔  Solusi: {delivery}"
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = COLOR_TEXT_DARK

    # -------------------------------------------------------------
    # SLIDE 9: PHASED IMPLEMENTATION ROADMAP
    # -------------------------------------------------------------
    s9 = prs.slides.add_slide(blank_slide_layout)
    set_bg(s9)
    add_header(s9, "Rencana Implementasi Bertahap Menuju Produksi", "Roadmap & Rollout")

    phases = [
        ("Fase 1: Working PoC (Selesai)", "Bulan 1", [
            "Rule engine konsolidasi 6 entitas tervalidasi.",
            "Conversational Copilot dengan prompt pills.",
            "Review Workspace 'Hasil yang Perlu Di-Review'.",
            "Ekspor laporan dan tracking status review."
        ], COLOR_OK),
        ("Fase 2: Integrasi & Pilot", "Bulan 2 - 3", [
            "Konektor otomatis ke GL/ERP anak perusahaan.",
            "Penyimpanan database snapshot point-in-time.",
            "Workflow notifikasi email/chat ke tim akuntansi anak.",
            "Pilot pengujian paralel dengan audit tahun sebelumnya."
        ], COLOR_GOLD),
        ("Fase 3: Enterprise Rollout", "Bulan 4+", [
            "Multi-tenant RBAC & Single Sign-On (SSO).",
            "Elimination engine otomatis tingkat transaksi (full GL).",
            "Portal telaah formal untuk auditor eksternal.",
            "Monitoring performa portofolio investasi holding terpadu."
        ], COLOR_FOREST)
    ]

    p_w = Inches(3.64)
    for idx, (p_title, p_time, bullets, col) in enumerate(phases):
        x = Inches(0.8 + idx * 4.03)
        add_card(s9, x, top_y, p_w, Inches(4.8), f"{p_title}\n({p_time})", bullets, col)

    # -------------------------------------------------------------
    # SLIDE 10: EXECUTIVE VALUE PROPOSITION & NEXT STEPS
    # -------------------------------------------------------------
    s10 = prs.slides.add_slide(blank_slide_layout)
    set_bg(s10, COLOR_DEEP_FOREST)

    tb10 = s10.shapes.add_textbox(Inches(1.2), Inches(1.0), Inches(10.8), Inches(1.5))
    tf10 = tb10.text_frame
    tf10.word_wrap = True

    p = tf10.paragraphs[0]
    p.text = "Kesimpulan Eksekutif & Langkah Tindak Lanjut"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    p2 = tf10.add_paragraph()
    p2.text = "Transformasi Pengawasan Finansial Konsolidasi: Proaktif, Akuntabel, dan Terdokumentasi Sempurna."
    p2.font.size = Pt(14)
    p2.font.color.rgb = COLOR_GOLD
    p2.space_before = Pt(8)

    # 3 Summary Cards
    s_w = Inches(3.4)
    s_h = Inches(3.6)
    s_top = Inches(2.8)

    summary_items = [
        ("Efisiensi Waktu Closing 60%", [
            "Deteksi instan atas anomali matematis eliminasi & alokasi NCI.",
            "Mengurangi iterasi koreksi lembar kerja pada akhir periode."
        ]),
        ("Mitigasi Risiko Audit 100%", [
            "Seluruh transaksi berelasi (PSAK 7) terpantau lonjakannya.",
            "Hak pemegang saham minoritas (PSAK 65) terlindungi secara transparan."
        ]),
        ("Langkah Berikutnya", [
            "1. Presentasi hasil telaah kepada Komite Audit Yayasan.",
            "2. Pengujian data historis audit 2025 untuk validasi sensitivitas.",
            "3. Integrasi feed ERP langsung untuk otomasi periode berikutnya."
        ])
    ]

    for idx, (title, bullets) in enumerate(summary_items):
        x = Inches(1.2 + idx * 3.8)
        card = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, s_top, s_w, s_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_FOREST
        card.line.color.rgb = COLOR_GOLD
        card.line.width = Pt(1.5)

        tb = s10.shapes.add_textbox(x + Inches(0.2), s_top + Inches(0.2), s_w - Inches(0.4), s_h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.size = Pt(14)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_WHITE
        p_t.space_after = Pt(10)

        for b in bullets:
            p_b = tf.add_paragraph()
            p_b.text = "• " + b
            p_b.font.size = Pt(11)
            p_b.font.color.rgb = RGBColor(215, 210, 200)
            p_b.space_after = Pt(6)

    prs.save(output_path)
    print(f"Presentation successfully created at: {output_path}")

if __name__ == "__main__":
    create_deck()
