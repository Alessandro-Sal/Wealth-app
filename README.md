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
graph TD
    %% ─────────────────────────────────────────
    %% ENTRY POINT
    %% ─────────────────────────────────────────
    USER["👤 Utente (Mobile / Browser)"]

    subgraph FRONTEND["🖥️ Frontend — HTML / CSS / Vanilla JS"]
        direction TB
        HI["Html_Index.html\n(SPA Shell / Router)"]
        HB["Html_Body.html\n(Layout & Componenti UI)"]

        subgraph JS_MODULES["Script Modules"]
            direction LR
            JS_INIT["Html_Script_Init\n(Bootstrap / Sessione)"]
            JS_NAV["Html_Script_Navigation\n(Routing SPA)"]
            JS_FORM["Html_Script_Form\n(Form Handler)"]
            JS_MOD["Html_Script_Module\n(Modali & Pannelli)"]
            JS_CHART["Html_Script_Chart\n(Chart.js Render)"]
            JS_MARKET["Html_Script_Market\n(Live Ticker)"]
            JS_NEWS["Html_Script_News\n(Feed Notizie)"]
            JS_AI["Html_Script_AI\n(Chat AI / OCR / Voice)"]
        end

        subgraph CSS_LAYERS["Stili CSS"]
            direction LR
            CSS_MAIN["css_Main\n(Layout iOS-style)"]
            CSS_COMP["css_Components\n(Card / Badge)"]
            CSS_MOD["css_Modals\n(Overlay / Dialog)"]
            CSS_MKT["css_Market\n(Ticker / Grafici)"]
            CSS_NEWS2["css_News\n(Feed Cards)"]
        end
    end

    subgraph BACKEND["⚙️ Backend — Google Apps Script"]
        direction TB

        APP_MAIN["App.js\n(doGet — Web App Entry)"]

        subgraph CORE["Core & Infrastruttura"]
            APP_UTILS["App_Utils.js\n(Helper comuni)"]
            APP_CACHE["App_Cache.js\n(CacheService wrapper)"]
            APP_SEC["App_Security.js\n(PIN / Privacy Mode)"]
            APP_OPT["App_Optimization.js\n(Batch & Performance)"]
        end

        subgraph DATA_LAYER["Data Layer — Lettura/Scrittura Sheets"]
            APP_DASH["App_DashboardData.js\n(Aggregazione NW)"]
            APP_TRANS_DB["App_TransactionsToDB.js\n(Salva Transazioni)"]
            APP_INV_DB["App_InvestmentsToDB.js\n(Salva Investimenti)"]
            APP_DEL["App_DeleteRowToSpreadsheets.js\n(Elimina righe)"]
            APP_YEARS["App_GetDataYEARS.js\n(Dati storici annui)"]
            APP_BANK["App_BankBalances.js\n(Saldi Conti)"]
            APP_FIXED["App_Fixed_expenses.js\n(Spese fisse & Sub.)"]
            APP_TRIPS["App_Trips.js\n(Travel Log / Mappa)"]
            APP_PENSION["App_Pension.js\n(Calcolo Pensione)"]
        end

        subgraph PORTFOLIO_ENGINE["📦 Portfolio Engine"]
            APP_LIVE["App_GetLivePortfolio.js\n(Prezzi real-time)"]
            APP_LASTINV["App_LastInvestments.js\n(Ultime operazioni)"]
            APP_WATCH["App_Watchlist.js\n(Target Price Alert)"]
            APP_SNIPER["App_SniperAlert.js\n(Alert Acquisto/Sell)"]
        end

        subgraph ANALYTICS["📊 Analytics & Reporting"]
            APP_CHART2["App_AnalyticsChart.js\n(Sankey / Allocation)"]
            APP_SAVINGS["App_MonthlySavings.js\n(Risparmio mensile)"]
            APP_CHART_SAV["App_MonthlyChartSavings.js\n(Chart Risparmi)"]
            APP_RUNWAY["App_FinancialRunaway.js\n(Runway / FIRE %)"]
            APP_BUDGET["App_BudgetStatus.js\n(Stato Budget)"]
            APP_GUARD["App_BudgetGuardian.js\n(Soglie Avviso)"]
            APP_REPORT["App_Reporting.js\n(Export Report)"]
            APP_LASTTX["App_LastTransactions.js\n(Ricerca TX)"]
        end

        subgraph TAX_MODULE["🧾 Fiscalità (Italian Law)"]
            APP_TAX["App_TaxOptimization.js\n(Zainetto Fiscale)"]
            SHEETS_TAX["Sheets_TaxOptimization.js"]
            SHEETS_TAXES["Sheets_Taxes_RealGains.js\n(Plusvalenze FIFO)"]
            SHEETS_TRADING["Sheets_Trading.js\n(FIFO Engine / Cutoff Date)"]
        end

        subgraph AI_MODULE["🤖 AI — Google Gemini"]
            AI_CFG["AI_Config.js\n(Model: gemini-1.5-flash)"]
            APP_AI_BE["App_AI.js\n(Prompt Builder / Chat)"]
            APP_NEWS_BE["App_News.js\n(Market Insight)"]
            APP_MARKET["App_MarketMail.js\n(Mail Analisi Mercato)"]
            APP_MONDAY["App_MondayBriefing.js\n(Email Lunedì)"]
            APP_DIV["App_DividendsMail.js\n(Alert Dividendi)"]
        end

        subgraph SHEETS_LAYER["📋 Google Sheets Automation"]
            SHEETS_AUTO["Sheets_Automation.js\n(Trigger Scheduler)"]
            SHEETS_BACKUP["Sheets_Backup.js\n(Nightly Backup Drive)"]
            SHEETS_ARCHIVE["Sheets_Archive.js\n(Freeze Dati Mensili)"]
            SHEETS_CACHE_FIN["Sheets_Cachefinance.js\n(Cache Prezzi)"]
            SHEETS_REAL["Sheets_RealAssets.js\n(Immobili)"]
            SHEETS_UTILS["Sheets_Utils.js\n(Helper Sheets)"]
            SHEETS_DERIV["Sheets_DerivatesEngine.js\n(⚠️ Non attivo)"]
            SHEETS_DATA["Sheets_FunzionePopolaData.js\n(Popolamento date)"]
        end

        subgraph TESTING["🧪 Testing"]
            TEST_AI["Testing_TestAIVers.js"]
        end
    end

    subgraph GOOGLE_INFRA["☁️ Google Infrastructure"]
        GS["📊 Google Sheets\n(Expenses Tracker · NW Analitico\nPortfolio · Watchlist · Config)"]
        GDRIVE["💾 Google Drive\n(Backup notturni)"]
        GMAIL_SVC["📧 Gmail\n(Report email)"]
        PROPS["🔑 Script Properties\n(GEMINI_API_KEY)"]
        TRIGGER["⏰ Time-Driven Triggers\n(3AM Backup · 23:30 Freeze)"]
    end

    subgraph EXTERNAL_API["🌐 API Esterne"]
        GEMINI["🤖 Google Gemini API\n(gemini-1.5-flash)"]
        MARKET_API["📈 Market Data\n(prezzi real-time)"]
        LEAFLET["🗺️ Leaflet.js\n(Mappe Viaggi)"]
        CHARTJS["📉 Chart.js\n(Grafici)"]
        SWA["🔔 SweetAlert2\n(Modali)"]
    end

    %% ─────────────────────────────────────────
    %% CONNECTIONS
    %% ─────────────────────────────────────────

    USER -->|"HTTPS GET"| APP_MAIN
    APP_MAIN --> HI
    HI --> HB
    HI --> JS_INIT
    JS_INIT --> JS_NAV
    JS_NAV --> JS_FORM
    JS_NAV --> JS_MOD
    JS_NAV --> JS_CHART
    JS_NAV --> JS_MARKET
    JS_NAV --> JS_NEWS
    JS_NAV --> JS_AI

    JS_FORM -->|"google.script.run"| APP_TRANS_DB
    JS_FORM -->|"google.script.run"| APP_INV_DB
    JS_CHART -->|"google.script.run"| APP_CHART2
    JS_CHART -->|"google.script.run"| APP_SAVINGS
    JS_MARKET -->|"google.script.run"| APP_LIVE
    JS_NEWS -->|"google.script.run"| APP_NEWS_BE
    JS_AI -->|"google.script.run"| APP_AI_BE

    APP_MAIN --> CORE
    APP_MAIN --> DATA_LAYER
    APP_MAIN --> PORTFOLIO_ENGINE
    APP_MAIN --> ANALYTICS
    APP_MAIN --> TAX_MODULE
    APP_MAIN --> AI_MODULE

    APP_UTILS --> APP_CACHE
    APP_CACHE --> GS
    APP_SEC --> PROPS

    DATA_LAYER --> GS
    PORTFOLIO_ENGINE --> GS
    ANALYTICS --> GS
    TAX_MODULE --> GS

    APP_AI_BE --> AI_CFG
    AI_CFG --> PROPS
    AI_CFG --> GEMINI

    APP_LIVE --> MARKET_API
    SHEETS_CACHE_FIN --> MARKET_API

    SHEETS_AUTO --> TRIGGER
    TRIGGER --> SHEETS_BACKUP
    TRIGGER --> SHEETS_ARCHIVE
    SHEETS_BACKUP --> GDRIVE

    APP_MARKET --> GMAIL_SVC
    APP_MONDAY --> GMAIL_SVC
    APP_DIV --> GMAIL_SVC

    SHEETS_TRADING --> SHEETS_TAXES
    SHEETS_TAXES --> SHEETS_TAX
    APP_TAX --> SHEETS_TAX

    JS_CHART --> CHARTJS
    JS_NAV --> LEAFLET
    JS_MOD --> SWA

    CSS_MAIN --> HB
    CSS_COMP --> HB
    CSS_MOD --> HB
    CSS_MKT --> HB
    CSS_NEWS2 --> HB

    SHEETS_REAL --> GS
    APP_REPORT --> GMAIL_SVC
    APP_REPORT --> GS

    TEST_AI --> APP_AI_BE

    %% ─────────────────────────────────────────
    %% STYLING
    %% ─────────────────────────────────────────
    classDef frontend fill:#1e3a5f,stroke:#4a90d9,color:#e8f4fd
    classDef backend fill:#1a3a2a,stroke:#4caf50,color:#e8f5e9
    classDef googleInfra fill:#3e2723,stroke:#ff8f00,color:#fff8e1
    classDef external fill:#2d1b4e,stroke:#9c27b0,color:#f3e5f5
    classDef tax fill:#4a1942,stroke:#e91e63,color:#fce4ec
    classDef ai fill:#0d2137,stroke:#00bcd4,color:#e0f7fa
    classDef testing fill:#1c1c1c,stroke:#757575,color:#eeeeee

    class HI,HB,JS_INIT,JS_NAV,JS_FORM,JS_MOD,JS_CHART,JS_MARKET,JS_NEWS,JS_AI,CSS_MAIN,CSS_COMP,CSS_MOD,CSS_MKT,CSS_NEWS2 frontend
    class APP_MAIN,APP_UTILS,APP_CACHE,APP_SEC,APP_OPT,APP_DASH,APP_TRANS_DB,APP_INV_DB,APP_DEL,APP_YEARS,APP_BANK,APP_FIXED,APP_TRIPS,APP_PENSION,APP_LIVE,APP_LASTINV,APP_WATCH,APP_SNIPER,APP_CHART2,APP_SAVINGS,APP_CHART_SAV,APP_RUNWAY,APP_BUDGET,APP_GUARD,APP_REPORT,APP_LASTTX,SHEETS_AUTO,SHEETS_BACKUP,SHEETS_ARCHIVE,SHEETS_CACHE_FIN,SHEETS_REAL,SHEETS_UTILS,SHEETS_DERIV,SHEETS_DATA backend
    class GS,GDRIVE,GMAIL_SVC,PROPS,TRIGGER googleInfra
    class GEMINI,MARKET_API,LEAFLET,CHARTJS,SWA external
    class APP_TAX,SHEETS_TAX,SHEETS_TAXES,SHEETS_TRADING tax
    class AI_CFG,APP_AI_BE,APP_NEWS_BE,APP_MARKET,APP_MONDAY,APP_DIV ai
    class TEST_AI testing