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

## 📂 Project Structure

Ensure your Apps Script project files match this structure:

### Backend (.gs)
| File | Description |
| :--- | :--- |
| `Code.gs` | Main server-side logic (doGet, API router). |
| `Script_Portfolio.gs` | Core Engine: FIFO calculations, Fiscal logic, Derivatives, and Crypto. |
| `Script_Backup.gs` | Automation: Nightly backups to Drive and monthly data freeze. |
| `Script_DebugAI.gs` | Utility: API connection testing and Gemini model listing. |

### Frontend (.html)
| File | Description |
| :--- | :--- |
| `index.html` | Main application entry point. |
| **Styles (CSS)** | |
| `css_Main.html` | Core styles, color variables, Dark Mode, and Privacy Mode. |
| `css_Components.html` | UI Elements: Buttons, Inputs, PIN Pad, Splash Screen. |
| `css_Market.html` | Financial Widgets: Ticker scroll, Watchlist grid, Budget bars. |
| `css_Modals.html` | Styles for Bottom Sheets, AI Overlays, and Calculators. |
| **Logic (JS)** | |
| `Html_Script_Init.html` | App startup logic, Caching (Fast/Heavy load), and Skeleton UI. |
| `Html_Script_Navigation.html` | Tab routing, PIN management, Pull-to-refresh, and UX handlers. |
| `Html_Script_Chart.html` | Renderers for Chart.js (Financial charts and Sankey). |
| `Html_Script_Market.html` | Watchlist logic, Live Portfolio updates, and TradingView widgets. |
| `Html_Script_Form.html` | Transaction forms, validation, and financial calculators. |
| `Html_Script_AI.html` | Gemini Chat logic, Voice/Image input, and AI security. |
| `Html_Script_Module.html` | Extra modules: Travel and Maps (Leaflet). |

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
In `Script_Portfolio.gs`, adjust the stock retention date. Positions closed *before* this date will be hidden to keep the UI clean:
```javascript
// Stocks closed before this date will be hidden from the view
const CUTOFF_DATE = new Date("2026-01-30");