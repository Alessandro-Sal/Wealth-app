# Welcome to the Project Contribution Guide

Thank you for investing your time in contributing to our project! Any contribution you make will be reflected on specific Italian fiscal workflows.

## ℹ️ Project Status & Scope

Please note that **this is primarily a personal project** tailored to specific workflows and Italian tax regulations.
While we welcome improvements, keep in mind that:
* Features specific to other countries' tax laws may be rejected.
* The core logic relies on `Google Apps Script` and Gemini AI models.

## 🐛 Reporting Bugs

If you find a bug or an error in the calculation logic:
1.  **Check existing issues** to avoid duplicates.
2.  Open a new Issue titled: `[BUG] - Short description`.
3.  Include steps to reproduce the error and, if possible, the JSON or data input that caused it.

## 💡 Suggesting Enhancements

Since this tool is customized for a specific workflow, please **open an issue first** to discuss any major feature changes before writing code. This saves you time!

## 🔧 Pull Request Process

We follow the standard GitHub flow. When you are ready to submit your changes:

1.  **Fork** the repository.
2.  Create a new branch for your feature or fix:
    ```bash
    git checkout -b feature/AmazingFeature
    # or
    git checkout -b fix/CalculationError
    ```
3.  **Test your changes**. Ensure the `debugAvailableModels` in `Script_DebugAI.gs` still passes and that tax formulas are accurate.
4.  Commit your changes using descriptive commit messages:
    ```bash
    git commit -m 'Fix: Corrected VAT calculation for regime forfettario'
    ```
5.  Push to your branch:
    ```bash
    git push origin feature/AmazingFeature
    ```
6.  Open a **Pull Request** targeting the `main` branch.

## 🎨 Coding Style

* Use **JavaScript (ES6+)** syntax compatible with Google Apps Script.
* Keep variable names descriptive (e.g., `calculateNetIncome` instead of `calc`).
* If modifying the AI prompts, please ensure the `MODEL_NAME` constant remains configurable.

---
*Thanks for helping make this tool better!*