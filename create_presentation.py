#!/usr/bin/env python3
"""
Generate an executive presentation deck (.pptx) in English with a clean White/Light theme.
Project: Financial Consolidation Analyzer & Review Workspace (AI Forensic Copilot)
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

    # Clean White Executive Color Palette
    COLOR_BG          = RGBColor(255, 255, 255) # Pure White #FFFFFF
    COLOR_CARD_BG     = RGBColor(248, 250, 252) # Soft Slate-50 #F8FAFC
    COLOR_CARD_BORDER = RGBColor(226, 232, 240) # Slate-200 #E2E8F0
    COLOR_PRIMARY     = RGBColor(15, 23, 42)    # Slate-900 #0F172A (Deep Navy/Black)
    COLOR_FOREST      = RGBColor(27, 67, 50)    # Deep Forest Green #1B4332
    COLOR_EMERALD     = RGBColor(16, 120, 85)   # Rich Emerald Green
    COLOR_GOLD        = RGBColor(180, 130, 20)  # Sophisticated Warm Gold #B48214
    COLOR_BLUE        = RGBColor(2, 132, 199)   # Vibrant Blue
    COLOR_TEXT_DARK   = RGBColor(30, 41, 59)    # Slate-800 #1E293B
    COLOR_TEXT_MUTED  = RGBColor(100, 116, 139)# Slate-500 #64748B
    COLOR_DANGER      = RGBColor(225, 29, 72)   # Rose-600 #E11D48
    COLOR_WARN        = RGBColor(217, 119, 6)   # Amber-600 #D97706
    COLOR_OK          = RGBColor(13, 148, 136)  # Teal-600 #0D9488
    COLOR_WHITE       = RGBColor(255, 255, 255)

    def set_white_bg(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_BG
        bg.line.fill.background()
        return bg

    def add_header(slide, title_text, category_text="FINANCIAL CONSOLIDATION ANALYZER &bull; SOLUTION ARCHITECTURE"):
        # Top Accent Line
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.04))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = COLOR_FOREST
        top_bar.line.fill.background()

        # Category pill
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(10), Inches(0.35))
        tf = cat_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = category_text.upper()
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.85), Inches(11.733), Inches(0.7))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = COLOR_PRIMARY

    def add_card(slide, left, top, width, height, title, body_bullets, top_border_color=None, bg_color=COLOR_CARD_BG):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = COLOR_CARD_BORDER
        card.line.width = Pt(1.0)

        # Optional top colored accent strip
        if top_border_color:
            strip = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, Inches(0.08))
            strip.fill.solid()
            strip.fill.fore_color.rgb = top_border_color
            strip.line.fill.background()

        pad = Inches(0.24)
        tb = slide.shapes.add_textbox(left + pad, top + pad, width - (pad * 2), height - (pad * 2))
        tf = tb.text_frame
        tf.word_wrap = True

        p_title = tf.paragraphs[0]
        p_title.text = title
        p_title.font.size = Pt(13.5)
        p_title.font.bold = True
        p_title.font.color.rgb = top_border_color if top_border_color else COLOR_PRIMARY
        p_title.space_after = Pt(8)

        for bullet in body_bullets:
            p = tf.add_paragraph()
            p.text = "• " + bullet
            p.font.size = Pt(11)
            p.font.color.rgb = COLOR_TEXT_DARK
            p.space_after = Pt(4)

    top_y = Inches(1.75)

    # -------------------------------------------------------------
    # SLIDE 1: TITLE SLIDE (Clean Minimalist Executive White)
    # -------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_slide_layout)
    set_white_bg(s1)

    # Left decorative vertical accent
    v_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(1.8), Inches(0.08), Inches(3.8))
    v_bar.fill.solid()
    v_bar.fill.fore_color.rgb = COLOR_FOREST
    v_bar.line.fill.background()

    tb1 = s1.shapes.add_textbox(Inches(1.5), Inches(1.7), Inches(10.5), Inches(3.2))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p_tag = tf1.paragraphs[0]
    p_tag.text = "ENTERPRISE SOLUTION ARCHITECTURE & STAKEHOLDER ALIGNMENT"
    p_tag.font.size = Pt(11)
    p_tag.font.bold = True
    p_tag.font.color.rgb = COLOR_GOLD

    p_main = tf1.add_paragraph()
    p_main.text = "Financial Consolidation Analyzer & Review Workspace"
    p_main.font.size = Pt(32)
    p_main.font.bold = True
    p_main.font.color.rgb = COLOR_PRIMARY
    p_main.space_before = Pt(8)

    p_sub = tf1.add_paragraph()
    p_sub.text = "AI Forensic Copilot & Human-in-the-Loop Audit Intelligence"
    p_sub.font.size = Pt(18)
    p_sub.font.bold = True
    p_sub.font.color.rgb = COLOR_FOREST
    p_sub.space_before = Pt(8)

    p_desc = tf1.add_paragraph()
    p_desc.text = "An automated system for group financial consolidation review, mathematical anomaly detection (IFRS 10 / PSAK 65 NCI & IAS 24 / PSAK 7 Related-Parties), and internal audit sign-off workflows."
    p_desc.font.size = Pt(12)
    p_desc.font.color.rgb = COLOR_TEXT_MUTED
    p_desc.space_before = Pt(14)

    # Bottom Metadata Bar
    bot_card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(5.6), Inches(10.9), Inches(1.1))
    bot_card.fill.solid()
    bot_card.fill.fore_color.rgb = COLOR_CARD_BG
    bot_card.line.color.rgb = COLOR_CARD_BORDER

    tb_meta = s1.shapes.add_textbox(Inches(1.5), Inches(5.7), Inches(10.3), Inches(0.9))
    tf_meta = tb_meta.text_frame
    tf_meta.word_wrap = True
    p_m1 = tf_meta.paragraphs[0]
    p_m1.text = "Target Structure: Holding Foundation (Induk) & 5 Operating Subsidiaries (Subsidiaries I – V)"
    p_m1.font.size = Pt(11.5)
    p_m1.font.bold = True
    p_m1.font.color.rgb = COLOR_PRIMARY

    p_m2 = tf_meta.add_paragraph()
    p_m2.text = "Standards: IFRS 10 / PSAK 65 (NCI Attribution) &bull; IAS 24 / PSAK 7 (Related-Party Transactions) &bull; POJK Prudential Ratios"
    p_m2.font.size = Pt(11)
    p_m2.font.color.rgb = COLOR_TEXT_MUTED
    p_m2.space_before = Pt(3)

    # -------------------------------------------------------------
    # SLIDE 2: THE MULTI-ENTITY CONSOLIDATION CHALLENGE
    # -------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_slide_layout)
    set_white_bg(s2)
    add_header(s2, "The Multi-Entity Consolidation Challenge", "Problem Statement & Context")

    c_w = Inches(3.64)
    c_h = Inches(4.8)

    add_card(s2, Inches(0.8), top_y, c_w, c_h, "1. NCI Allocation Leakage (IFRS 10)", [
        "Minority interests (NCI 22% - 40%) across 4 operating subsidiaries are prone to manual calculation errors.",
        "Case in Point: Subsidiary II reported NCI profit share of $24.4M vs. statutory share of $18.2M (+34.1% gap).",
        "Material Impact: Distorts parent net income downwards or masks disguised profit extractions.",
        "Traditional spreadsheets lack automated mathematical reconciliation of effective minority stakes."
    ], COLOR_DANGER)

    add_card(s2, Inches(4.84), top_y, c_w, c_h, "2. Related-Party Debt Spikes (IAS 24)", [
        "Intercompany receivables frequently surge without commensurate commercial revenue growth.",
        "Case in Point: Subsidiary II related-party receivables spiked +158.8% YoY while revenues grew just +4.9%.",
        "Severe Risk: Holding liquidity becomes trapped in affiliated entities under non-arms-length terms.",
        "Increases intercompany elimination friction and triggers uncollectible debt exposure."
    ], COLOR_WARN)

    add_card(s2, Inches(8.88), top_y, c_w, c_h, "3. Fragmented Audit Trails", [
        "Consolidation workpapers and variance inquiries are scattered across email threads and disconnected sheets.",
        "The Audit Committee and Group CFO lack real-time visibility into anomaly resolution statuses.",
        "No centralized tracking distinguishes accepted operational variances from required audit adjustments.",
        "Critical errors are often discovered late during external auditor year-end field examinations."
    ], COLOR_BLUE)

    # -------------------------------------------------------------
    # SLIDE 3: THE DUAL-ENGINE SOLUTION VISION
    # -------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_slide_layout)
    set_white_bg(s3)
    add_header(s3, "Dual-Engine Architecture: Precision Math + Cognitive AI", "Core Architecture")

    add_card(s3, Inches(0.8), top_y, Inches(5.6), Inches(4.8), "Engine 1: Deterministic Accounting Engine", [
        "100% Mathematical Precision: Rigorous aggregation of Revenues, Net Income, Assets, Liabilities, and Equity.",
        "Automated Intercompany Eliminations: Standardized matrices eliminate cross-holdings and reciprocal debt.",
        "Statutory NCI Attribution: Precise IFRS 10 formulas calculate parent earnings vs. minority equity claims.",
        "Zero Black-Box Calculations: Every number is derived from transparent, reproducible accounting formulas.",
        "Instant Dynamic Recalculation: Input modifications dynamically update solvency and profitability indicators."
    ], COLOR_FOREST)

    add_card(s3, Inches(6.8), top_y, Inches(5.6), Inches(4.8), "Engine 2: Cognitive AI Forensic Copilot", [
        "Conversational Copilot: Answers executive inquiries in natural, professional corporate finance language.",
        "Forensic Root-Cause Reasoning: Translates raw numerical deltas into concrete audit and compliance risk rationales.",
        "Actionable Review Workspace: Empowers auditors to mark items as Clarify, Audit Finding, or Approved.",
        "Contextual Smart Navigation: Conversational insights feature direct jump buttons to highlighted review cards.",
        "Prudential Benchmarking: Contextualizes group leverage (DER 0.96x) and returns (ROA 5.4%) against sector standards."
    ], COLOR_GOLD)

    # -------------------------------------------------------------
    # SLIDE 4: 5-TIER TECHNICAL BLUEPRINT
    # -------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_slide_layout)
    set_white_bg(s4)
    add_header(s4, "End-to-End 5-Tier Technical Blueprint", "System Architecture")

    tiers = [
        ("Tier 1: Ingestion Layer", ["ERP / SAP / Oracle GL", "Excel & XBRL Ingestor", "CY & PY Data Sanity Checks"], COLOR_PRIMARY),
        ("Tier 2: Consolidation", ["Multi-Entity Aggregator", "Elimination Engine", "IFRS 10 NCI Attribution"], COLOR_FOREST),
        ("Tier 3: Forensic Rules", ["NCI Mismatch Vector", "Related-Party Spike Detector", "Leverage & Margin Scanners"], COLOR_DANGER),
        ("Tier 4: AI Copilot", ["Multi-Agent Orchestrator", "Forensic Explainer Agent", "Regulatory Benchmark Agent"], COLOR_GOLD),
        ("Tier 5: Review UI", ["Interactive Split View", "Review Decision Cards", "Formal Sign-off & PDF Export"], COLOR_BLUE)
    ]

    t_w = Inches(2.26)
    t_h = Inches(4.8)
    for idx, (title, items, col) in enumerate(tiers):
        x = Inches(0.8 + idx * 2.42)
        add_card(s4, x, top_y, t_w, t_h, title, items, col)

    # -------------------------------------------------------------
    # SLIDE 5: FORENSIC ANOMALY DETECTION ENGINE
    # -------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_slide_layout)
    set_white_bg(s5)
    add_header(s5, "Forensic Anomaly Detection Engine: Vectors & Thresholds", "Quantitative Rules")

    rules_data = [
        ("1. NCI Profit Allocation Disparity", "CRITICAL", "|Reported NCI - Computed NCI| > 8%", "Prevents parent income distortion and safeguards minority shareholder dividend equity under IFRS 10 / PSAK 65.", COLOR_DANGER),
        ("2. Related-Party Debt Acceleration", "CRITICAL", "ΔRP Receivables - ΔRevenue > 30%", "Detects liquidity outflows into affiliated entities and halts hidden internal non-performing debt under IAS 24 / PSAK 7.", COLOR_DANGER),
        ("3. Sudden Leverage / DER Escalation", "MEDIUM", "DER CY / DER PY - 1 > 18%", "Monitors subsidiary debt growth before debt-to-equity ratios breach bank debt covenants or regulatory limits.", COLOR_WARN),
        ("4. Profit Margin Distortion", "MODERATE", "|Margin CY - Margin PY| ≥ 3.5 pp", "Verifies whether margin expansion/compression stems from genuine operations or one-off artificial transactions.", COLOR_OK)
    ]

    card_h = Inches(1.1)
    for idx, (r_name, sev, formula, desc, col) in enumerate(rules_data):
        y = Inches(1.75 + idx * 1.25)
        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), y, Inches(11.733), card_h)
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_CARD_BG
        c.line.color.rgb = COLOR_CARD_BORDER

        # Left indicator bar
        bar = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), y, Inches(0.08), card_h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = col
        bar.line.fill.background()

        tb = s5.shapes.add_textbox(Inches(1.05), y + Inches(0.12), Inches(11.3), card_h - Inches(0.24))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = f"{r_name}  [{sev}]"
        p.font.size = Pt(13.5)
        p.font.bold = True
        p.font.color.rgb = col

        p2 = tf.add_paragraph()
        p2.text = f"Formula: {formula}   •   Audit Objective: {desc}"
        p2.font.size = Pt(11)
        p2.font.color.rgb = COLOR_TEXT_DARK
        p2.space_before = Pt(3)

    # -------------------------------------------------------------
    # SLIDE 6: REVIEW WORKSPACE & HITL WORKFLOW
    # -------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_slide_layout)
    set_white_bg(s6)
    add_header(s6, "Actionable Review Workspace: Human-in-the-Loop Workflow", "Audit Operations")

    workflow_cards = [
        ("Step 1: Automated Flagging", [
            "Rule Engine scans all 6 entities instantly.",
            "Anomalies classified by severity: Critical, Medium, Moderate.",
            "Dynamic badges populate the review dashboard."
        ], COLOR_DANGER),
        ("Step 2: AI Reasoning & Impact", [
            "AI Agent derives analytical root causes and accounting impact.",
            "Relevant IFRS 10 / IAS 24 standards cited.",
            "Conversational Copilot delivers instant executive briefs."
        ], COLOR_GOLD),
        ("Step 3: Auditor Decisions", [
            "Auditor selects status: Clarify / Audit Finding / Approved.",
            "Input field records auditor remarks & requested documentation.",
            "One-click 'Request Clarification' logs notice to subsidiary."
        ], COLOR_FOREST),
        ("Step 4: Formal Sign-off", [
            "Real-time tracking: X of Y anomalies reviewed.",
            "Formal sign-off locks workpaper with timestamp.",
            "One-click clean print to PDF ready for the Audit Committee."
        ], COLOR_BLUE)
    ]

    for idx, (title, items, col) in enumerate(workflow_cards):
        x = Inches(0.8 + idx * 2.96)
        add_card(s6, x, top_y, Inches(2.8), Inches(4.8), title, items, col)

    # -------------------------------------------------------------
    # SLIDE 7: REGULATORY & SECTORAL BENCHMARKING
    # -------------------------------------------------------------
    s7 = prs.slides.add_slide(blank_slide_layout)
    set_white_bg(s7)
    add_header(s7, "Consolidated Financial Health vs. Sectoral Benchmarks", "Prudential Ratios")

    bench_cards = [
        ("Solvency (DER)", "Consolidated: 0.96x", "Regulatory Ceiling: Max 5.0x", "STATUS: HIGHLY CONSERVATIVE", [
            "Capital structure provides ample liquidity buffer for operational growth.",
            "Well below regulatory leverage thresholds for financial institutions."
        ], COLOR_OK),
        ("Profitability (ROA)", "Consolidated: 5.4%", "Industry Average: 3.5% - 5.2%", "STATUS: OUTPERFORMING", [
            "Group asset return sits in the upper quartile of microfinance peers.",
            "Consolidated Net Income expanded +14.2% YoY ($325M vs $284.7M)."
        ], COLOR_OK),
        ("Affiliated Debt (RP)", "Total RP Receivables: $241M", "Share of Total Assets: 4.3%", "STATUS: CONCENTRATION WATCH", [
            "Aggregate exposure is proportional to balance sheet scale.",
            "However, concentration at Subsidiary II ($88M) warrants tight monitoring."
        ], COLOR_WARN)
    ]

    for idx, (title, val_entity, val_ind, status, details, col) in enumerate(bench_cards):
        x = Inches(0.8 + idx * 4.03)
        c = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, top_y, Inches(3.64), Inches(4.8))
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_CARD_BG
        c.line.color.rgb = COLOR_CARD_BORDER

        # Top strip
        strip = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, top_y, Inches(3.64), Inches(0.08))
        strip.fill.solid()
        strip.fill.fore_color.rgb = col
        strip.line.fill.background()

        tb = s7.shapes.add_textbox(x + Inches(0.24), top_y + Inches(0.24), Inches(3.16), Inches(4.3))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = COLOR_PRIMARY
        p.space_after = Pt(6)

        p_v = tf.add_paragraph()
        p_v.text = f"{val_entity}\n{val_ind}"
        p_v.font.size = Pt(11)
        p_v.font.bold = True
        p_v.font.color.rgb = COLOR_TEXT_DARK
        p_v.space_after = Pt(6)

        p_s = tf.add_paragraph()
        p_s.text = status
        p_s.font.size = Pt(10)
        p_s.font.bold = True
        p_s.font.color.rgb = col
        p_s.space_after = Pt(10)

        for d in details:
            p_d = tf.add_paragraph()
            p_d.text = "• " + d
            p_d.font.size = Pt(11)
            p_d.font.color.rgb = COLOR_TEXT_DARK
            p_d.space_after = Pt(4)

    # -------------------------------------------------------------
    # SLIDE 8: STAKEHOLDER EXPECTATION ALIGNMENT MATRIX
    # -------------------------------------------------------------
    s8 = prs.slides.add_slide(blank_slide_layout)
    set_white_bg(s8)
    add_header(s8, "Stakeholder Expectation Alignment Matrix", "Value Delivery")

    matrix_items = [
        ("Audit Committee / Board of Supervisors", "Prevent material misstatements & NCI leakage prior to external audits.", "Automated IFRS 10 rule engine verifies profit distribution & flags critical anomalies instantly."),
        ("Group Chief Financial Officer (CFO)", "Accelerate group consolidation closing cycles with zero arithmetic risk.", "Dynamic KPI summaries, instant sensitivity simulations, and executive board-ready memos."),
        ("Internal Audit Leadership", "Maintain structured electronic workpapers with traceable sign-offs.", "Actionable Review Workspace with item status toggles, auditor notes, and timestamped audit trails."),
        ("Subsidiary CFOs (Operating Units)", "Ensure transparent, fair elimination and reconciliation calculations.", "Line-by-line breakdown of all variances enables rapid cross-checking against subsidiary general ledgers."),
        ("IT & Data Governance", "Strict data privacy with no exposure of unreleased financials to public LLMs.", "Runs entirely client-side/on-premise; append-only point-in-time snapshot persistence.")
    ]

    for idx, (stakeholder, expectation, delivery) in enumerate(matrix_items):
        y = Inches(1.7 + idx * 1.0)
        c = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), y, Inches(11.733), Inches(0.9))
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_CARD_BG
        c.line.color.rgb = COLOR_CARD_BORDER

        # Vertical indicator
        bar = s8.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), y, Inches(0.06), Inches(0.9))
        bar.fill.solid()
        bar.fill.fore_color.rgb = COLOR_FOREST
        bar.line.fill.background()

        tb = s8.shapes.add_textbox(Inches(1.0), y + Inches(0.08), Inches(11.3), Inches(0.75))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = stakeholder.upper()
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p2 = tf.add_paragraph()
        p2.text = f"Expectation: \"{expectation}\"  ➔  Delivery: {delivery}"
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = COLOR_TEXT_DARK

    # -------------------------------------------------------------
    # SLIDE 9: PHASED ENTERPRISE IMPLEMENTATION ROADMAP
    # -------------------------------------------------------------
    s9 = prs.slides.add_slide(blank_slide_layout)
    set_white_bg(s9)
    add_header(s9, "Phased Enterprise Implementation Roadmap", "Deployment Plan")

    phases = [
        ("Phase 1: Working PoC (Delivered)", "Month 1", [
            "Consolidation rule engine validated for 6 entities.",
            "Conversational AI Copilot with quick-action chips.",
            "Actionable 'Review Workspace' for anomaly tracking.",
            "Self-contained local execution & GitHub repository."
        ], COLOR_OK),
        ("Phase 2: Integration & Pilot", "Months 2 - 3", [
            "Direct ERP/GL database connectors (SAP/Oracle).",
            "Immutable point-in-time snapshot database store.",
            "Automated email/chat notification dispatch to subsidiaries.",
            "Parallel dry-run benchmarked against prior-year audited data."
        ], COLOR_GOLD),
        ("Phase 3: Full Enterprise Rollout", "Month 4+", [
            "Multi-tenant Role-Based Access Control (RBAC) & SSO.",
            "Transaction-level automated intercompany elimination.",
            "Dedicated secure portal for external audit review teams.",
            "Continuous group-wide financial health telemetry."
        ], COLOR_FOREST)
    ]

    for idx, (p_title, p_time, bullets, col) in enumerate(phases):
        x = Inches(0.8 + idx * 4.03)
        add_card(s9, x, top_y, Inches(3.64), Inches(4.8), f"{p_title}\n({p_time})", bullets, col)

    # -------------------------------------------------------------
    # SLIDE 10: EXECUTIVE VALUE PROPOSITION & NEXT STEPS
    # -------------------------------------------------------------
    s10 = prs.slides.add_slide(blank_slide_layout)
    set_white_bg(s10)
    add_header(s10, "Executive Value Proposition & Recommended Next Steps", "Strategic Impact")

    roi_cards = [
        ("60% Faster Closing Cycle", [
            "Automated anomaly screening eliminates tedious manual spreadsheet reconciliations.",
            "Immediate flagging of NCI disparities removes late-stage closing bottlenecks."
        ], COLOR_FOREST),
        ("100% Audit Readiness", [
            "Continuous adherence to IFRS 10 (NCI) and IAS 24 (Related Parties).",
            "Complete digital audit trail protects holding executives and board members."
        ], COLOR_GOLD),
        ("Recommended Next Steps", [
            "1. Present PoC findings to the Audit Committee and Group CFO.",
            "2. Execute pilot validation using 2025 audited financial statements.",
            "3. Establish direct subsidiary GL connectors for automated data feeds."
        ], COLOR_BLUE)
    ]

    for idx, (title, bullets, col) in enumerate(roi_cards):
        x = Inches(0.8 + idx * 4.03)
        add_card(s10, x, top_y, Inches(3.64), Inches(4.8), title, bullets, col)

    prs.save(output_path)
    print(f"White executive presentation created at: {output_path}")

if __name__ == "__main__":
    create_deck()
