# 📈 Wealth-App — Personal Finance & Portfolio Tracker

> An advanced, mobile-first **Single Page Application (SPA)** built entirely on **Google Apps Script**, designed to give you complete, real-time control over your Net Worth, Expenses, and Investments — all in one platform.

[![License: GPL-3.0](https://img.shields.io/badge/license-GPL--3.0-blue.svg)](LICENSE.md)
[![Platform](https://img.shields.io/badge/platform-Google%20Apps%20Script-green)](https://script.google.com)
[![Status](https://img.shields.io/badge/status-Production-orange)](https://github.com/Alessandro-Sal/Wealth-app)
[![AI](https://img.shields.io/badge/AI-Gemini%201.5%20Flash-purple)](https://deepmind.google/technologies/gemini/)

---

## 📋 Table of Contents

- [Overview](#-overview)
- [Key Features](#-key-features)
- [Tech Stack](#️-tech-stack)
- [Architecture](#️-architecture)
- [File Structure](#-file-structure)
- [Installation & Setup](#-installation--setup)
- [Configuration](#️-configuration)
- [Contributing](#-contributing)
- [License](#-license)

---

## 🌐 Overview

Wealth-App is a fully self-hosted personal finance manager that runs entirely inside **Google's ecosystem** — no external servers, no subscriptions. It leverages Google Sheets as a database, Google Apps Script as backend, and a custom HTML/CSS/JS SPA as the frontend.

Key design principles:
- **Mobile-first**: iOS 17-inspired UI with haptic feedback, smooth animations, and dark mode
- **Privacy-first**: All data stays within your Google account
- **Italy-aware**: Full support for Italian fiscal reporting, "Zainetto Fiscale", and capital gains tax (26%)
- **AI-powered**: Google Gemini integration for insights, OCR receipts, voice input, and market analysis

---

## ✨ Key Features

### 💰 Portfolio Management

- **Real-Time Tracking** — Live prices for Stocks, ETFs, and Crypto via Yahoo Finance
- **FIFO Engine** — Accurate realized/unrealized gain calculations using First-In-First-Out logic
- **Derivatives Support** — Native handling for Options and Futures (short selling, multipliers, cash flow)
- **Closed Positions** — Configurable cutoff date to hide old trades and keep the UI clean
- **Tax Optimization (Italian Law)**
  - *Zainetto Fiscale*: tracks capital loss expiry (4-year rule) and optimizes tax credits
  - Auto-calculation of 26% capital gains tax on profits
- **Drill-down Analytics** — Asset breakdown by Sector, Industry, Country, and Instrument type

### 🤖 AI Integration (Gemini 1.5 Flash)

- **Smart Input** — Voice-to-Text and Image-to-Text (OCR for receipts) to log expenses hands-free
- **Market Insights** — Generative AI analysis of macroeconomic trends and portfolio sentiment
- **Risk Analysis** — Automated stress testing and concentration reports
- **Chat Assistant** — Conversational mode to query your own financial data in natural language
- **Monday Briefing** — Weekly AI-generated email summarizing your portfolio and market outlook

### 📊 Dashboard & Analytics

- **Interactive Charts** powered by Chart.js: Sankey Flow, Asset Allocation, Monthly Trends, Savings Evolution
- **Financial Runway** — Calculates survival months based on liquid cash vs average expenses
- **FIRE Progress** — Tracks Financial Independence progress using the 4% withdrawal rule
- **Budget Guardian** — Real-time spending alerts when budget thresholds are exceeded
- **Travel Log** — Interactive map (Leaflet.js) with per-trip cost breakdown

### 📧 Automated Reports & Alerts

- **Market Mail** — Weekly market analysis email
- **Dividend Alerts** — Notified when dividends are received
- **Sniper Alerts** — Buy/sell signals based on watchlist target prices
- **Monthly Reporting** — PDF portfolio export via Gmail

### 🔒 Security & UX

- **Privacy Mode** — Blurs all sensitive values with a single tap or via PIN abort
- **PIN Lock** — Secure access control to the application
- **Automated Backups** — Nightly Google Drive snapshots + monthly data freeze
- **Pull-to-Refresh** — Native mobile-like refresh gesture
- **Real Assets** — Track non-financial assets (property, etc.)
- **Pension Module** — Monitor pension fund contributions and projections

---

## 🛠️ Tech Stack

| Layer | Technology |
|---|---|
| **Backend** | Google Apps Script (Server-side JavaScript) |
| **Frontend** | HTML5, CSS3 (iOS 17 Design System), Vanilla JavaScript |
| **Database** | Google Sheets |
| **AI** | [Google Gemini API](https://deepmind.google/technologies/gemini/) — `gemini-1.5-flash` |
| **Charts** | [Chart.js](https://www.chartjs.org/) |
| **Maps** | [Leaflet.js](https://leafletjs.com/) |
| **Modals** | [SweetAlert2](https://sweetalert2.github.io/) |
| **Market Data** | Yahoo Finance (via Sheets_Cachefinance) |
| **DevOps** | [CLASP](https://github.com/google/clasp) (local development & deployment) |

---

## 🏗️ Architecture

The app follows a clean 3-layer separation of concerns:

```
┌─────────────────────────────────────────────────────────┐
│  👤 User (Mobile / Browser)                             │
└────────────────────┬────────────────────────────────────┘
                     │ HTTPS GET
┌────────────────────▼────────────────────────────────────┐
│  Google Apps Script                                     │
│                                                         │
│  ┌─────────────┐  ┌──────────────┐  ┌───────────────┐  │
│  │  Frontend   │  │   Backend    │  │   AI Layer    │  │
│  │ Html_*.html │  │  App_*.js    │  │  App_AI.js    │  │
│  │ css_*.html  │  │  Sheets_*.js │  │  AI_Config.js │  │
│  └──────┬──────┘  └──────┬───────┘  └───────┬───────┘  │
│         │  google.script.run │               │          │
└─────────┼──────────────────┼───────────────┼──────────┘
          │                  │               │
    ┌─────▼──────────────────▼───────────────▼──────┐
    │           Google Infrastructure               │
    │  📊 Sheets  💾 Drive  📧 Gmail  🔑 Properties │
    └───────────────────────┬───────────────────────┘
                            │
                    ┌───────▼───────┐
                    │ External APIs │
                    │ Gemini · YF   │
                    └───────────────┘
```

For detailed Mermaid diagrams of each layer (Frontend, Backend modules, Tax Engine, AI layer), see **[ARCHITECTURE.md](ARCHITECTURE.md)**.

---

## 📁 File Structure

The project uses a strict naming convention to separate concerns:

```
Wealth-app/
│
├── 📄 App.js                          # doGet() entry point
│
├── 🖥️  Html_*.html                    # Frontend: SPA views & client-side scripts
│   ├── Html_Index.html                # SPA entry point
│   ├── Html_Body.html                 # Layout shell
│   ├── Html_Script_Init.html          # Bootstrap & session management
│   ├── Html_Script_Navigation.html    # SPA router
│   ├── Html_Script_Form.html          # Form handlers (add transactions/investments)
│   ├── Html_Script_Chart.html         # Chart.js renderer
│   ├── Html_Script_Market.html        # Live ticker
│   ├── Html_Script_Module.html        # Modals & panels
│   ├── Html_Script_News.html          # News feed
│   ├── Html_Script_AI.html            # Chat / OCR / Voice
│   └── Html_Script_Recap.html         # Summary views
│
├── 🎨 css_*.html                      # Styling (scoped per section)
│   ├── css_Main.html                  # iOS-style global layout
│   ├── css_Components.html            # Cards, badges, lists
│   ├── css_Modals.html                # Overlay & dialogs
│   ├── css_Market.html                # Ticker styles
│   └── css_News.html                  # Feed cards
│
├── ⚙️  App_*.js                       # Backend controllers & business logic
│   ├── App_DashboardData.js           # Net Worth aggregation
│   ├── App_GetLivePortfolio.js        # Real-time prices
│   ├── App_TransactionsToDB.js        # Write transactions to Sheets
│   ├── App_InvestmentsToDB.js         # Write investments to Sheets
│   ├── App_BankBalances.js            # Bank account balances
│   ├── App_BudgetGuardian.js          # Budget alert logic
│   ├── App_FinancialRunaway.js        # Financial runway & FIRE calculator
│   ├── App_TaxOptimization.js         # Zainetto Fiscale logic
│   ├── App_Watchlist.js               # Watchlist & target prices
│   ├── App_SniperAlert.js             # Buy/sell alert engine
│   ├── App_MondayBriefing.js          # Weekly AI briefing email
│   ├── App_MarketMail.js              # Market analysis email
│   ├── App_DividendsMail.js           # Dividend notification
│   ├── App_AI.js                      # Gemini AI: chat, OCR, voice
│   ├── App_Cache.js                   # Caching layer
│   ├── App_Security.js                # PIN & privacy mode
│   ├── App_Pension.js                 # Pension module
│   ├── App_Trips.js                   # Travel log
│   └── App_Utils.js                   # Shared utilities
│
├── 📊 Sheets_*.js                     # Data access layer (Google Sheets R/W)
│   ├── Sheets_Trading.js              # FIFO engine + cutoff date filter
│   ├── Sheets_Taxes & RealGains.js    # Capital gains & tax calculations
│   ├── Sheets_TaxOptimization.js      # Zainetto Fiscale data layer
│   ├── Sheets_Cachefinance.js         # Market price cache
│   ├── Sheets_Backup.js               # Nightly backup to Drive
│   ├── Sheets_Archive.js              # Monthly data freeze
│   ├── Sheets_Automation.js           # Trigger scheduler
│   ├── Sheets_RealAssets.js           # Real estate tracking
│   └── Sheets_Utils.js                # Helper functions
│
├── 🤖 AI_Config.js                    # Gemini model configuration
├── 🧪 Testing_TestAIVers.js           # AI integration tests
├── 📋 appsscript.json                 # GAS manifest
└── 📝 ARCHITECTURE.md                 # Detailed architecture diagrams
```

---

## 🚀 Installation & Setup

### Prerequisites

- A **Google Account**
- [CLASP](https://github.com/google/clasp) installed (optional, for local development):
  ```bash
  npm install -g @google/clasp
  clasp login
  ```

### Step 1 — Create the Google Sheet

Create a new Google Spreadsheet and set up the following required tabs:

| Sheet Name | Purpose |
|---|---|
| `Expenses Tracker` | Daily transactions |
| `NW Analitico` | Net Worth historical snapshots |
| `Portfolio` | Investment positions |
| `Watchlist` | Stocks/ETFs to monitor |
| `Config` | App configuration values |

### Step 2 — Open Apps Script

In your Google Sheet: **Extensions → Apps Script**

### Step 3 — Copy the Code

Create each file listed in the **File Structure** above and paste the corresponding code.

> 💡 **With CLASP**: clone this repo and run `clasp push` to deploy all files at once.

### Step 4 — Configure API Keys

In Apps Script: **Project Settings → Script Properties**, add:

| Key | Value |
|---|---|
| `GEMINI_API_KEY` | Your [Google AI Studio](https://aistudio.google.com/) API key |

> ⚠️ **Never hardcode API keys in the source files.**

### Step 5 — Set Triggers

In Apps Script: **Triggers (⏰ icon)**, add:

| Function | Frequency | Suggested Time |
|---|---|---|
| `createNightlyBackup` | Daily | 3:00 AM |
| `freezeMonthEndOnly` | Daily | 11:30 PM |
| `sendMondayBriefing` | Weekly (Monday) | 8:00 AM |
| `sendMarketMail` | Weekly | 7:00 AM |

### Step 6 — Deploy

1. Click **Deploy → New Deployment**
2. Type: **Web App**
3. Execute as: **Me**
4. Who has access: **Only myself** *(recommended)*
5. Click **Deploy** and copy the Web App URL

Open the URL in your mobile browser and add it to your home screen as a PWA.

---

## ⚙️ Configuration

### Cutoff Date (FIFO Engine)

In `Sheets_Trading.js`, positions closed *before* this date are hidden to keep the UI clean:

```javascript
const CUTOFF_DATE = new Date("2026-01-30");
```

### AI Model

In `AI_Config.js`, verify or update the Gemini model:

```javascript
const MODEL_NAME = "gemini-1.5-flash"; // Or "gemini-2.0-flash"
```

### Budget Thresholds & FIRE Target

Configure your personal financial targets in the `Config` sheet tab (monthly budget, FIRE number, savings rate target).

---

## 🤝 Contributing

Contributions are welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md) before submitting a pull request.

1. Fork the repository
2. Create your feature branch: `git checkout -b feature/my-feature`
3. Commit your changes: `git commit -m 'feat: add my feature'`
4. Push to the branch: `git push origin feature/my-feature`
5. Open a Pull Request

---

## 📄 License

Distributed under the **GPL-3.0 License**. See [LICENSE.md](LICENSE.md) for full details.

---

> ⚠️ **Disclaimer**: This tool is for **informational and personal use only**. Always verify tax calculations with a certified accountant. The authors are not responsible for financial losses or fiscal errors.

---

*Built with ❤️ on Google Apps Script — no servers, no subscriptions, just your data in your Google account.*
