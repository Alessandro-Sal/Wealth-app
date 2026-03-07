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

## 🏗️ Architecture & File Structure

The project follows a strict naming convention to separate concerns between Frontend (UI), Backend (Logic), and Database (Google Sheets):

* `Html_*.html` & `css_*.html`: View layer and client-side scripts.
* `App_*.js`: Server-side controllers and business logic.
* `Sheets_*.js`: Data access layer handling reads/writes to Google Sheets.


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

Distributed under the GPL-3.0 license. See `LICENSE.md` for more information.

> **⚠️ Disclaimer:** This tool is for **informational purposes only**. Always verify tax calculations with a professional accountant. The authors are not responsible for financial losses or fiscal errors.

```mermaid
flowchart TD
    %% --- Styles ---
    classDef frontend fill:#e1f5fe,stroke:#0288d1,stroke-width:2px,color:#01579b
    classDef backend fill:#f3e5f5,stroke:#7b1fa2,stroke-width:2px,color:#4a148c
    classDef database fill:#e8f5e9,stroke:#388e3c,stroke-width:2px,color:#1b5e20
    classDef trigger fill:#fff3e0,stroke:#f57c00,stroke-width:2px,color:#e65100

    %% --- 1. FRONTEND LAYER ---
    subgraph Frontend ["🖥️ FRONTEND LAYER (Browser UI)"]
        direction TB
        UI["Html_Index / Html_Body"]:::frontend
        CSS["Styles (css_Main, css_Modals, css_Market)"]:::frontend
        JS["Client JS (Init, Navigation, Forms, AI, Charts)"]:::frontend
        
        UI --- CSS
        UI --- JS
    end

    %% --- 2. BACKEND LAYER ---
    subgraph Backend ["⚙️ BACKEND CONTROLLERS (Apps Script)"]
        direction TB
        Router["App.js (Main Router & doGet)"]:::backend
        Core["Core & Config (App_Utils, Security, Cache, AI_Config)"]:::backend
        
        ReadOps["🔍 Data Fetching Controllers
        (DashboardData, LivePortfolio, GetDataYEARS)"]:::backend
        
        WriteOps["📝 Data Mutation Controllers
        (InvestmentsToDB, TransactionsToDB, DeleteRow)"]:::backend
        
        Modules["🧩 Feature Modules
        (App_AI, TaxOptimization, BudgetGuardian, News)"]:::backend
    end

    %% --- AUTOMATIONS ---
    Automations["⏰ Time-Driven Triggers
    (MondayBriefing, SniperAlert, MarketMail)"]:::trigger

    %% --- 3. DATABASE LAYER ---
    subgraph Database ["📊 DATABASE LAYER (Google Sheets)"]
        direction TB
        Connector["Sheets_Utils.js (DB Connector)"]:::database
        
        Sheets["🗂️ Sheet Managers
        (Trading, RealAssets, Taxes & Gains)"]:::database
        
        Maintenance["🧹 Maintenance
        (Backup, Archive, Automation)"]:::database
        
        Connector --> Sheets
        Connector --> Maintenance
    end

    %% --- RELATIONS & DATA FLOW ---
    
    %% Client to Server calls
    JS == "google.script.run" ==> Router
    JS -. "Async fetch calls" .-> ReadOps
    JS -. "Submit forms" .-> WriteOps
    JS -. "Module requests" .-> Modules

    %% Internal Backend logic
    Router -.-> Core
    ReadOps -. "uses" .-> Core
    WriteOps -. "uses" .-> Core
    Modules -. "uses" .-> Core

    %% Server to Database calls
    ReadOps == "Reads data" ==> Connector
    WriteOps == "Writes data" ==> Connector
    Modules == "Queries" ==> Connector
    Automations == "Scheduled execution" ==> Connector