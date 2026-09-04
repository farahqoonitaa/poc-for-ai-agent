# Konsolidasi Analyzer & Review Dashboard — AI Forensic Copilot

A modern, conversational Financial Forensic AI Agent and Review Workspace for consolidating financial statements across holding entities and subsidiaries (Yayasan and 5 subsidiaries).

Built with a forensic accounting rule engine, interactive AI chat interface, real-time KPI dashboards, sectoral benchmarking (OJK / Multifinance), and an auditor review tracking system.

![Konsolidasi Analyzer Preview](preview.png) *(Optional screenshot)*

---

## Key Features

### 1. 💬 AI Forensic Copilot Chat
- **Interactive Conversational Flow**: Query financial metrics, simulate scenarios, or request executive summaries.
- **Analytical Reasoning**: Displays step-by-step progress (`Menganalisis parameter...`, `Mengevaluasi kertas kerja 6 entitas...`, `Menghitung porsi NCI & anomali...`).
- **One-Click Prompt Chips**:
  - `📊 Jalankan Analisis & Tampilkan Dashboard`
  - `⚠️ Cek Alokasi NCI`
  - `🏢 Cek Transaksi Afiliasi`
  - `📈 Benchmark Industri OJK`
  - `✍️ Ringkasan Komite Audit`
- **Smart Jump Buttons**: Chat recommendations contain direct action buttons that scroll to and highlight relevant review items on the dashboard.

### 2. 📋 Review Workspace ("Hasil yang Perlu Di-Review")
- **Live Review Status Tracker**: Tracks `Belum Di-review`, `Perlu Klarifikasi`, `Temuan Audit`, and `Disetujui`.
- **Actionable Anomaly Cards**:
  - Anomaly severity badges (`Kritis`, `Sedang`, `Rendah`).
  - Mathematical disparity quantification (e.g. NCI calculated 18.2 M vs reported 24.4 M, +34.1% discrepancy).
  - Regulatory context (PSAK 65 & PSAK 7).
  - Auditor action dropdown (`Belum Di-review`, `Perlu Klarifikasi Manajemen`, `Jadikan Temuan Audit`, `Disetujui / Justified`).
  - Notes field for reviewer remarks and confirmation requests.
  - Automatic audit trail logging into chat history.

### 3. 📊 Executive Consolidation Dashboard & Sektoral Benchmark
- **Consolidated Financial Metrics**:
  - Revenue CY vs PY (with YoY growth %)
  - Net Income CY vs PY (with YoY growth %)
  - NCI Profit Share vs Parent Net Income
  - Consolidated Assets, ROA, and ROE
  - Debt-to-Equity Ratio (DER)
- **Entity Profit Contribution Visual Bar**: Visual proportional breakdown across all 6 entities.
- **Sektoral / OJK Benchmarking**: Comparison against Indonesian microfinance & multifinance standards.

### 4. 🖨️ Sign-Off & Formal Reporting
- Formal review sign-off toggle with timestamp and completion summary.
- Dedicated `@media print` stylesheet for generating clean PDF audit review sheets.

### 5. 📝 Interactive Data Simulation
- Live editable table for the 6 entities (`Yayasan`, `PT Anak I – V`).
- Any figure edited recalculates all metrics, flags, and chat context in real-time.

---

## Getting Started

### Running Locally

You can serve the project using any static file server:

```bash
# Using Python 3
python3 -m http.server 3000

# Or using Node.js / npx
npx serve .
```

Open [http://localhost:3000](http://localhost:3000) in your browser.

No build step or external dependencies required — runs completely client-side!
