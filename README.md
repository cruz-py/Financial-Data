# Financial Data Extractor
![Python](https://img.shields.io/badge/python-3.12-blue)
![License](https://img.shields.io/badge/license-MIT-green)

A lightweight desktop application that fetches and visualizes company financial statements and historical year-end prices using Alpha Vantage and Yahoo Finance.

Designed for investors and analysts who want faster access to fundamental data without manually browsing financial websites.

---

## 🚀 Why this tool exists

Checking income statements, balance sheets, and cash flows manually across multiple websites is slow and repetitive.

This tool automates that process by:

Pulling structured financial data via API

Displaying it in a clean GUI

Allowing one-click export to Excel

Simple. Fast. No clutter.

---

## ✨ Features
- Fetch **Income Statement**, **Balance Sheet**, and **Cash Flow** for any stock symbol.
- Retrieve **year-end closing prices** from Yahoo Finance.
- Export data to **Excel** (`.xlsx`), fully compatible with LibreOffice.
- User-friendly GUI built with **CustomTkinter**.
- Configurable **API key** settings.
- Dark mode support.
- Live **progress bar** and real-time log output.

---

## 🖼 Screenshots

![Main Window](screenshots/main_window.png)

---

## 🛠 Installation

### Requirements
- Python 3.10+ installed
- `pip` package manager

### Steps
1. Clone this repository:
   ```bash
   git clone https://github.com/cruz-py/Financial-Data.git
   cd financial-data-extractor
2. Create virtual environment:
    python -m venv venv
3. Activate the veritual environment:
    For windows:
        venv\Scripts\activate
    For Linux/MacOS:
        source venv/bin/activate
4. Install dependencies:
    pip install -r requirements.txt
5. Run the app:
    python Stocks_AV_GUI_v3.py

Or

6. If you prefer not to install Python, download the precompiled .exe file.

---

## 📖 Usage
1. Open the app;
2. Click Settings;
3. Insert your Alpha Vantage API Key or request one;
4. enter a stock symbol (e.g., AAPL);
5. Select the report type (Annual or Quarter);
6. Enter the number of years (1-15);
7. Click Run Analysis;
8. Once data is fetched, click Export to Excel to save locally.

🔐 Note: this tool requires a free Alpha Vantage API key. You can request one here: https://www.alphavantage.co/support/#api-key

---

## ⚠ Disclaimer

This software is provided for educational and research purposes only.

It does not constitue financial advice. Always performa your own due diligence before making investment decisions.

## 👨‍💻 About

Developed by Cruz-Py.

This project started as a personal automation tool and evolved into a shareable open-source utility.

Contributions, suggestions, and feedback are welcome.

### ❤️ Support / Donate
If you find this tool useful, consider supporting its development:

[Donate via PayPal](https://www.paypal.me/ruicruz27)

---

## 📜 License

This project is licensed under the MIT license. See the LICENSE file for details