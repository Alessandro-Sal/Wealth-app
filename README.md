# 📈 Personal Finance & Portfolio Tracker (Google Apps Script)

An advanced, mobile-first Single Page Application (SPA) built on **Google Apps Script** to track Net Worth, Expenses, and Investments. It features real-time market data, extensive fiscal reporting (specifically tailored for Italian "Zainetto Fiscale"), and AI-powered insights using **Google Gemini**.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-Google%20Apps%20Script-green)
![Status](https://img.shields.io/badge/status-Production-orange)

## ✨ Key Features

### 💰 Portfolio Management
- **Real-Time Tracking:** Live updates for Stocks, ETFs, and Crypto.
- **FIFO Engine:** Accurate calculations for realized/unrealized gains using First-In-First-Out logic.
- **Derivatives Support:** Native handling for Options and Futures (Short selling, Multipliers, Cash flow analysis).
- **Tax Optimization (Italian Law):**
  - **"Zainetto Fiscale":** Tracks capital losses expiry (4 years) and optimizes tax credits.
  - **Capital Gains Tax:** Auto-calculation of 26% tax on profits.
- **Drill-down Analytics:** Breakdown by Sector, Industry, and Country.

### 🤖 AI Integration (Gemini Pro)
- **Smart Input:** Voice-to-Text and Image-to-Text (OCR for receipts) to automatically log expenses.
- **Market Insights:** Generative AI analysis of macro-economic trends and portfolio sentiment.
- **Risk Analysis:** Automated stress testing and portfolio concentration report.
- **Chat Assistant:** Conversational mode to query your financial data.

### 📊 Dashboard & Analytics
- **Interactive Charts:** Powered by **Chart.js** (Sankey Flow, Asset Allocation, Monthly Trends).
- **Financial Runway:** Calculates survival months based on liquid cash and average expenses.
- **FIRE Progress:** Tracks progress towards Financial Independence (4% rule).
- **Travel Log:** Interactive Map (**Leaflet.js**) and cost breakdown for trips.

### 🔒 Security & UX
- **Privacy Mode:** Blurs sensitive values with a single tap or via PIN abort.
- **PIN Lock:** Secure access to the application.
- **iOS-Inspired UI:** Smooth animations, haptic feedback, dark mode, and "Pull-to-Refresh".
- **Automated Backups:** Nightly snapshots of the spreadsheet to Google Drive and monthly data freezing.

## 🛠️ Tech Stack

- **Backend:** Google Apps Script (Server-side JavaScript).
- **Frontend:** HTML5, CSS3 (iOS 17 Design System), Vanilla JavaScript.
- **Database:** Google Sheets.
- **Libraries:**
  - [Chart.js](https://www.chartjs.org/) (Data Visualization)
  - [Leaflet.js](https://leafletjs.com/) (Maps)
  - [SweetAlert2](https://sweetalert2.github.io/) (Modals & Alerts)
  - [Google Gemini API](https://deepmind.google/technologies/gemini/) (AI)

## 🚀 Installation & Setup

1.  **Create a Google Sheet:**
    * Create a new sheet.
    * Set up the required tabs: `Expenses Tracker`, `NW Analitico`, `Portfolio`, `Watchlist`, `Config`.
2.  **Open Apps Script:**
    * Go to `Extensions` > `Apps Script`.
3.  **Copy Code:**
    * Create the files listed in **Project Structure** above.
    * Paste the corresponding code into each file.
4.  **Configuration:**
    * **API Keys:** Set your `GEMINI_API_KEY` in `Script Properties` (Project Settings). **Do not hardcode it!**
    * **Triggers:** Set up time-driven triggers for:
      * `createNightlyBackup` (e.g., Daily at 3 AM).
      * `freezeMonthEndOnly` (e.g., Daily at 11:30 PM).
5.  **Deploy:**
    * Click `Deploy` > `New Deployment`.
    * Select type: `Web App`.
    * Execute as: `Me`.
    * Who has access: `Only myself` (Recommended).

## ⚙️ Configuration Snippets

### Privacy & Cut-off Dates
In `Sheets_Trading.gs`, adjust the stock retention date. Positions closed *before* this date will be hidden to keep the UI clean:
const CUTOFF_DATE = new Date("2026-01-30");

### 🤖 AI Model Configuration
To ensure the AI features work correctly, verify the model name in your AI script:
const MODEL_NAME = "gemini-1.5-flash"; // Or "gemini-pro"

## 🤝 Contributing

Contributions are welcome! Please read the [contribution guidelines](CONTRIBUTING.md) first.

# 📄 License

Distributed under the MIT License. See `LICENSE.md` for more information.

> **⚠️ Disclaimer:** This tool is for **informational purposes only**. Always verify tax calculations with a professional accountant. The authors are not responsible for financial losses or fiscal errors.