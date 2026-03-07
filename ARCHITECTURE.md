# 🏗️ Wealth-App — Architecture

> The app is a **Google Apps Script SPA** built on top of Google Sheets.  
> Below are 5 focused diagrams covering each architectural layer.

---

## 1️⃣ System Overview

```mermaid
flowchart LR
    USER("👤 User\nMobile / Browser")

    subgraph GAS["Google Apps Script"]
        FE["🖥️ Frontend\nHTML · CSS · JS"]
        BE["⚙️ Backend\nApp_*.js · Sheets_*.js"]
        AI["🤖 AI Layer\nGemini Integration"]
    end

    subgraph GINFRA["Google Infrastructure"]
        GS[("📊 Google Sheets\nDatabase")]
        GDRIVE["💾 Google Drive\nBackups"]
        GMAIL["📧 Gmail\nReports & Alerts"]
        PROPS["🔑 Script Properties\nAPI Keys"]
    end

    subgraph EXT["External APIs"]
        GEMINI["🤖 Gemini API\ngemini-1.5-flash"]
        MARKET["📈 Market Data\nReal-time prices"]
    end

    USER -->|"HTTPS GET"| FE
    FE -->|"google.script.run"| BE
    BE <-->|"Read / Write"| GS
    BE -->|"Backup nightly"| GDRIVE
    BE -->|"Send emails"| GMAIL
    AI -->|"API call"| GEMINI
    AI <-->|"Read portfolio"| GS
    BE --> AI
    BE <-->|"Live prices"| MARKET
    AI -.->|"API key"| PROPS
```

---

## 2️⃣ Frontend Architecture

```mermaid
flowchart TD
    HI["Html_Index.html\n🚪 SPA Entry Point"]

    HI --> JS_INIT["Html_Script_Init\n⚡ Bootstrap & Session"]
    JS_INIT --> JS_NAV["Html_Script_Navigation\n🧭 SPA Router"]

    JS_NAV --> JS_FORM["Html_Script_Form\n📝 Form Handler"]
    JS_NAV --> JS_MOD["Html_Script_Module\n🪟 Modals & Panels"]
    JS_NAV --> JS_CHART["Html_Script_Chart\n📊 Chart.js Renderer"]
    JS_NAV --> JS_MARKET["Html_Script_Market\n📈 Live Ticker"]
    JS_NAV --> JS_NEWS["Html_Script_News\n📰 News Feed"]
    JS_NAV --> JS_AI["Html_Script_AI\n🤖 Chat / OCR / Voice"]

    HB["Html_Body.html\n🧱 Layout Shell"]
    HB --> CSS_MAIN["css_Main\niOS-style layout"]
    HB --> CSS_COMP["css_Components\nCards & Badges"]
    HB --> CSS_MOD["css_Modals\nOverlay & Dialogs"]
    HB --> CSS_MKT["css_Market\nTicker styles"]
    HB --> CSS_NEWS["css_News\nFeed cards"]

    JS_CHART -->|uses| CHARTJS(["Chart.js"])
    JS_NAV -->|uses| LEAFLET(["Leaflet.js\n🗺️ Travel Map"])
    JS_MOD -->|uses| SWA(["SweetAlert2\n🔔 Modals"])

    JS_FORM -->|"google.script.run"| TRANS["App_TransactionsToDB"]
    JS_FORM -->|"google.script.run"| INV["App_InvestmentsToDB"]
    JS_CHART -->|"google.script.run"| DASH["App_DashboardData"]
    JS_MARKET -->|"google.script.run"| LIVE["App_GetLivePortfolio"]
    JS_AI -->|"google.script.run"| AI_BE["App_AI.js"]
```

---

## 3️⃣ Backend Modules

```mermaid
flowchart LR
    APP(["App.js\ndoGet entry"])

    subgraph CORE["🔧 Core"]
        UTILS["App_Utils"]
        CACHE["App_Cache"]
        SEC["App_Security\nPIN / Privacy"]
        OPT["App_Optimization"]
    end

    subgraph DATA["💾 Data Layer"]
        DASH["App_DashboardData\nNet Worth aggregation"]
        TRANS["App_TransactionsToDB"]
        INV["App_InvestmentsToDB"]
        DEL["App_DeleteRow"]
        BANK["App_BankBalances"]
        YEARS["App_GetDataYEARS\nHistorical data"]
        FIXED["App_Fixed_expenses\nSubs & fixed costs"]
        TRIPS["App_Trips\nTravel log"]
        PENSION["App_Pension"]
    end

    subgraph PORTFOLIO["📦 Portfolio Engine"]
        LIVE["App_GetLivePortfolio\nReal-time prices"]
        LASTINV["App_LastInvestments\n+ Search"]
        WATCH["App_Watchlist\nTarget alerts"]
        SNIPER["App_SniperAlert\nBuy / Sell alerts"]
    end

    subgraph ANALYTICS["📊 Analytics"]
        CHART["App_AnalyticsChart\nSankey / Allocation"]
        SAV["App_MonthlySavings"]
        CHART_SAV["App_MonthlyChartSavings"]
        RUNWAY["App_FinancialRunaway\nRunway & FIRE %"]
        BUDGET["App_BudgetStatus"]
        GUARD["App_BudgetGuardian\nSpending alerts"]
        REPORT["App_Reporting\nPDF Export"]
        LASTTX["App_LastTransactions\n+ Search"]
    end

    APP --> CORE
    APP --> DATA
    APP --> PORTFOLIO
    APP --> ANALYTICS
    UTILS --> CACHE
    CACHE --> GS[("📊 Google Sheets")]
    DATA --> GS
    PORTFOLIO --> GS
    ANALYTICS --> GS
```

---

## 4️⃣ Tax Engine & Sheets Automation

```mermaid
flowchart LR
    subgraph TAX["🧾 Italian Tax Engine"]
        TRADING["Sheets_Trading.js\nFIFO Engine\nCutoff Date filter"]
        TAXR["Sheets_Taxes_RealGains.js\nPlusvalenze FIFO\n26% tax calc"]
        TAXOPT_S["Sheets_TaxOptimization.js\nZainetto Fiscale\n4-year loss expiry"]
        TAXOPT_A["App_TaxOptimization.js\nOptimization logic"]

        TRADING --> TAXR
        TAXR --> TAXOPT_S
        TAXOPT_A --> TAXOPT_S
    end

    subgraph AUTOMATION["⏰ Sheets Automation"]
        AUTO["Sheets_Automation.js\nTrigger scheduler"]
        BACKUP["Sheets_Backup.js\nNightly snapshot"]
        ARCHIVE["Sheets_Archive.js\nMonthly data freeze"]
        CACHE_FIN["Sheets_Cachefinance.js\nPrice cache"]
        REAL["Sheets_RealAssets.js\nReal estate"]
        SH_UTILS["Sheets_Utils.js\nHelper functions"]
        POPOLA["Sheets_FunzionePopolaData.js\nDate population"]
    end

    subgraph TRIGGERS["☁️ Google Services"]
        T1["⏰ Trigger: 3AM\nDaily backup"]
        T2["⏰ Trigger: 23:30\nMonth-end freeze"]
        DRIVE["💾 Google Drive"]
        GS[("📊 Google Sheets")]
    end

    AUTO --> T1 & T2
    T1 --> BACKUP --> DRIVE
    T2 --> ARCHIVE --> GS
    TAX --> GS
    CACHE_FIN -->|"Live prices"| GS
    REAL --> GS
```

---

## 5️⃣ AI & Notifications Layer

```mermaid
flowchart LR
    subgraph AI["🤖 AI — Google Gemini"]
        CFG["AI_Config.js\nModel config\ngemini-1.5-flash"]
        APP_AI["App_AI.js\nChat · OCR receipts\nVoice-to-text"]
        NEWS["App_News.js\nMarket insights"]
    end

    subgraph NOTIFY["📧 Automated Emails"]
        MARKET["App_MarketMail.js\nMarket analysis report"]
        MONDAY["App_MondayBriefing.js\nWeekly briefing"]
        DIV["App_DividendsMail.js\nDividend alerts"]
        REPORT["App_Reporting.js\nPortfolio export"]
    end

    subgraph EXT["External"]
        GEMINI(["🤖 Gemini API"])
        GMAIL(["📧 Gmail"])
        PROPS(["🔑 Script Properties\nGEMINI_API_KEY"])
    end

    GS[("📊 Google Sheets\nPortfolio data")]

    CFG -->|"reads key"| PROPS
    CFG -->|"model call"| GEMINI
    APP_AI --> CFG
    NEWS --> CFG

    APP_AI <-->|"query data"| GS
    NEWS <-->|"query data"| GS

    MARKET --> GMAIL
    MONDAY --> GMAIL
    DIV --> GMAIL
    REPORT --> GMAIL

    MARKET --> CFG
    MONDAY --> CFG

    TEST["🧪 Testing_TestAIVers.js"] -.->|"test calls"| APP_AI
```
