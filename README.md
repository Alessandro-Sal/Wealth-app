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

graph LR
    %% -------------------------------------
    %% FRONTEND LAYER (Browser)
    %% -------------------------------------
    subgraph Frontend ["🖥️ Frontend / UI Layer (Client-side)"]
        direction TB
        Index("Html_Index.html (Entry Point)")
        Body("Html_Body.html (Layout)")

        subgraph Styles ["🎨 CSS Stylesheets"]
            CSSMain("css_Main")
            CSSComp("css_Components / Modals")
            CSSSpec("css_Market / News")
        end

        subgraph ClientJS ["⚙️ Client Logic (JS)"]
            JSInit("Html_Script_Init / Navigation")
            JSUI("Html_Script_Module / Form")
            JSSpec("Html_Script_Chart / AI / Market / News")
        end

        Index --> |includes| Body
        Index --> |loads| Styles
        Index --> |loads| ClientJS
    end

    %% -------------------------------------
    %% BACKEND LAYER (Apps Script)
    %% -------------------------------------
    subgraph Backend ["⚙️ Backend Controllers (Server-side Apps Script)"]
        direction TB
        AppMain("App.js (Main Router / doGet)")

        subgraph Core ["🛠️ Core & Security"]
            AppUtils("App_Utils.js")
            AppSec("App_Security.js")
            AppCache("App_Cache.js")
            AIConf("AI_Config.js")
        end

        subgraph ReadOps ["🔍 Data Fetching"]
            Dash("App_DashboardData.js")
            Years("App_GetDataYEARS.js")
            Live("App_GetLivePortfolio.js")
            Analytics("App_AnalyticsChart.js")
        end

        subgraph WriteOps ["📝 Data Mutation"]
            InvDB("App_InvestmentsToDB.js")
            TransDB("App_TransactionsToDB.js")
            DelRow("App_DeleteRowToSpreadsheets.js")
        end

        subgraph AppModules ["🧩 App Modules"]
            AI("App_AI.js")
            Budget("App_BudgetGuardian / BudgetStatus")
            Tax("App_TaxOptimization.js")
            News("App_News.js")
            Misc("App_Trips / App_Pension / App_Watchlist")
        end

        subgraph Triggers ["⏰ Automations & Cron Jobs"]
            Briefing("App_MondayBriefing.js")
            Alerts("App_SniperAlert.js")
            Mails("App_MarketMail / App_DividendsMail")
        end

        AppMain -.-> Core
    end

    %% -------------------------------------
    %% DATABASE LAYER (Google Sheets)
    %% -------------------------------------
    subgraph Database ["📊 Database Layer (Google Sheets)"]
        direction TB
        SheetUtils("Sheets_Utils.js (DB Connector)")

        subgraph Managers ["🗂️ Sheet Managers"]
            Trd("Sheets_Trading.js")
            RAssets("Sheets_RealAssets.js")
            TaxSh("Sheets_Taxes & RealGains.js")
            CacheFin("Sheets_Cachefinance.js")
        end

        subgraph DBMaintenance ["🧹 Maintenance Ops"]
            Bkp("Sheets_Backup.js")
            Arch("Sheets_Archive.js")
            Auto("Sheets_Automation.js")
        end

        SheetUtils --> Managers
        SheetUtils --> DBMaintenance
    end

    %% -------------------------------------
    %% RELATIONS & DEPENDENCIES
    %% -------------------------------------
    
    %% Client to Server calls
    ClientJS ==>|google.script.run| AppMain
    ClientJS ==>|google.script.run API calls| ReadOps
    ClientJS ==>|google.script.run API calls| WriteOps
    ClientJS ==>|google.script.run API calls| AppModules

    %% Server to Core
    ReadOps -.->|uses| Core
    WriteOps -.->|uses| Core
    AppModules -.->|uses| Core
    AppModules -->|configs| AIConf

    %% Server to DB Layer
    ReadOps ==>|calls| SheetUtils
    WriteOps ==>|calls| SheetUtils
    AppModules ==>|reads/writes| SheetUtils
    Triggers ==>|scheduled query| SheetUtils