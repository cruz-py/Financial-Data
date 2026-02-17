import customtkinter as ctk
from tkinter import messagebox, filedialog
import requests
import pandas as pd
from datetime import datetime
import yfinance as yf
import time
import os
import sys
import json
import webbrowser
import threading

# ==============================
# Developer info
# ==============================

APP_NAME = "Financial Data Extractor"
APP_VERSION = "3.1"
DEVELOPER = "Cruz-Py"
YEAR = "2026"

# ==============================
# Configuration
# ==============================
RATE_LIMIT_SLEEP = 60
NORMAL_SLEEP = 12
MAX_RETRIES = 3
CACHE_TTL_HOURS = 24

# ==============================
# App directory (fixed)
# ==============================
def get_app_directory():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

APP_DIR = get_app_directory()
CACHE_DIR = os.path.join(APP_DIR, "cache")
os.makedirs(CACHE_DIR, exist_ok=True)

# ==============================
# Settings
# ==============================
SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")

DEFAULT_SETTINGS = {
    "api_key": "",
    "api_key_validated": False,
    "save_directory": APP_DIR
}

def load_settings():
    if not os.path.exists(SETTINGS_FILE):
        return DEFAULT_SETTINGS.copy()
    try:
        with open(SETTINGS_FILE, "r") as f:
            data = json.load(f)
            return {**DEFAULT_SETTINGS, **data}
    except Exception:
        return DEFAULT_SETTINGS.copy()

def save_settings(settings):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(settings, f, indent=4)

settings = load_settings()

# ==============================
# Cache Utilities
# ==============================
def get_cache_path(symbol, function_name, period, year):
    filename = f"{symbol}_{function_name}_{period}_{year}.json"
    return os.path.join(CACHE_DIR, filename)

def is_cache_valid(path):
    if not os.path.exists(path):
        return False
    age_seconds = time.time() - os.path.getmtime(path)
    return age_seconds < CACHE_TTL_HOURS * 3600

def load_from_cache(symbol, function_name, period, year):
    path = get_cache_path(symbol, function_name, period, year)
    if not is_cache_valid(path):
        return None
    try:
        with open(path, "r") as f:
            return json.load(f)
    except Exception:
        return None

def save_to_cache(symbol, function_name, period, year, data):
    path = get_cache_path(symbol, function_name, period, year)
    try:
        with open(path, "w") as f:
            json.dump(data, f)
    except Exception:
        pass

# ==============================
# Fetch Financial Statements
# ==============================
def fetch_financial_statements(symbol, api_key, period="annual", limit=15,
                               log_callback=None, progress_callback=None):
    base_url = "https://www.alphavantage.co/query"
    function_map = {
        "income_statement": "INCOME_STATEMENT",
        "balance_sheet": "BALANCE_SHEET",
        "cash_flow": "CASH_FLOW"
    }

    financials = {}
    total_steps = len(function_map)
    completed_steps = 0
    year = str(datetime.now().year)

    for key, func in function_map.items():
        cached = load_from_cache(symbol, func, period.lower(), year)
        if cached:
            if log_callback: log_callback(f"Loaded {key.replace('_',' ')} from cache ✔️\n")
            df = pd.DataFrame(cached)
            if not df.empty:
                df = df.sort_values("fiscalDateEnding").tail(limit)
            financials[key] = df
            completed_steps += 1
            if progress_callback:
                progress_callback((completed_steps / total_steps) * 100)
            continue

        retries = 0
        while retries <= MAX_RETRIES:
            if log_callback: log_callback(f"Fetching {key.replace('_',' ')}...\n")
            params = {"function": func, "symbol": symbol, "apikey": api_key}
            try:
                response = requests.get(base_url, params=params, timeout=15)
                data = response.json()
            except Exception as e:
                if log_callback: log_callback(f"Request failed: {e}\n")
                financials[key] = pd.DataFrame()
                break

            if "Note" in data:
                retries += 1
                if log_callback: log_callback(f"⚠️ Rate limit hit. Waiting {RATE_LIMIT_SLEEP}s...\n")
                time.sleep(RATE_LIMIT_SLEEP)
                continue
            if "Error Message" in data:
                if log_callback: log_callback("API error returned by Alpha Vantage.\n")
                financials[key] = pd.DataFrame()
                break

            reports_key = "quarterlyReports" if period.lower() == "quarter" else "annualReports"
            reports = data.get(reports_key, [])
            save_to_cache(symbol, func, period.lower(), year, reports)
            df = pd.DataFrame(reports)
            if not df.empty:
                df = df.sort_values("fiscalDateEnding").tail(limit)
            financials[key] = df
            break

        completed_steps += 1
        if progress_callback:
            progress_callback((completed_steps / total_steps) * 100)
        time.sleep(NORMAL_SLEEP)

    return financials

# ==============================
# Fetch Year-End Prices
# ==============================
def fetch_year_end_closing_prices_yf(symbol, years):
    start_date = datetime(min(years), 1, 1)
    end_date = datetime(max(years)+1, 1, 1)
    data = yf.Ticker(symbol).history(start=start_date, end=end_date)
    if data.empty: return {}
    closing_prices = {}
    for year in years:
        year_data = data[data.index.year == year]
        if not year_data.empty:
            last_day = year_data.index.max()
            closing_prices[str(year)] = round(year_data.loc[last_day]["Close"], 2)
        else:
            closing_prices[str(year)] = None
    return closing_prices

# ==============================
# Save to Excel
# ==============================
def save_to_excel(financials, closing_prices, symbol):
    base_dir = settings.get("save_directory", APP_DIR)
    if not os.path.isdir(base_dir):
        messagebox.showerror("Invalid Folder", "Save directory does not exist.")
        return
    filename = os.path.join(base_dir, f"{symbol}_financials.xlsx")
    try:
        with pd.ExcelWriter(filename) as writer:
            for sheet_name, df in financials.items():
                if df.empty: continue
                if "fiscalDateEnding" in df.columns:
                    df_t = df.set_index("fiscalDateEnding").transpose()
                    df_t.columns = [c[:4] for c in df_t.columns]
                else:
                    df_t = df.transpose()
                df_t = df_t.replace([None,"None",pd.NA],0).apply(pd.to_numeric, errors="coerce").fillna(0)
                df_t.to_excel(writer, sheet_name=sheet_name.replace("_"," ").title())
            if closing_prices:
                cp_df = pd.DataFrame(closing_prices.items(), columns=["Year","Closing Price"])
                cp_df.to_excel(writer, sheet_name="Year-End Closing Prices", index=False)
        messagebox.showinfo("Saved", f"File saved:\n\n{filename}")
    except Exception as e:
        messagebox.showerror("Save Error", str(e))

# ==============================
# Thread-safe helpers
# ==============================
def safe_log(message):
    root.after(0, lambda: (result_text.insert("end", message), result_text.see("end")))

def safe_progress(value):
    # value is 0–100, convert to 0–1 for CTkProgressBar
    root.after(0, lambda: progress_bar.set(value / 100))


def safe_enable_run_button():
    root.after(0, lambda: run_button.configure(state="normal"))

# ==============================
# Run Analysis
# ==============================
def run_analysis():
    if not settings.get("api_key_validated"):
        messagebox.showerror("API Key Required","Please validate your Alpha Vantage API key in Settings.")
        return
    run_button.configure(state="disabled")
    progress_bar.set(0)
    result_text.delete("1.0","end")
    symbol = symbol_entry.get().strip().upper()
    period_input = period_var.get().strip().lower()
    years_input = years_entry.get().strip()
    try:
        years = int(years_input)
        if not (1<=years<=50): raise ValueError
    except ValueError:
        messagebox.showerror("Input Error","Please enter 1–50 years.")
        run_button.configure(state="normal")
        return
    threading.Thread(target=run_analysis_thread, args=(symbol, period_input, years), daemon=True).start()

def run_analysis_thread(symbol, period_input, years):
    try:
        current_year = datetime.now().year
        years_list = list(range(current_year - years + 1, current_year + 1))

        # Initial log
        safe_log(f"Fetching financial data for {symbol}...\n\n")

        # Fetch financial statements from Alpha Vantage
        financials = fetch_financial_statements(
            symbol,
            settings["api_key"],
            period=period_input,
            limit=years,
            log_callback=safe_log,
            progress_callback=safe_progress
        )

        # Fetch year-end closing prices from Yahoo Finance
        closing_prices = fetch_year_end_closing_prices_yf(symbol, years_list)

        # Display year-end closing prices
        safe_log("\nYear-end Closing Prices:\n")
        for y in years_list:
            safe_log(f"  {y}: {closing_prices.get(str(y), 'N/A')}\n")

        # Success message
        safe_log("\n✔️ Financial data successfully extracted!\n"
                 "Please click 'Export to Excel' to save it locally.\n")

        # Properly enable and assign command to save button on main thread
        def enable_save_button():
            save_button.configure(
                state="normal",
                command=lambda: save_to_excel(financials, closing_prices, symbol)
            )
        
        root.after(0, enable_save_button)

        # Ensure progress bar is full at the end
        safe_progress(100)

    except Exception as e:
        safe_log(f"\n❌ Error: {e}\n")

    finally:
        # Re-enable the Run Analysis button
        safe_enable_run_button()

# ==============================
# Settings window
# ==============================
def test_alpha_vantage_key(api_key):
    if not api_key: return False,"API key is empty."
    try:
        response = requests.get("https://www.alphavantage.co/query",
                                params={"function":"SYMBOL_SEARCH","keywords":"AAPL","apikey":api_key},
                                timeout=10)
        data = response.json()
    except Exception as e:
        return False,f"Network error: {e}"
    if "Note" in data: return False,"Rate limit reached."
    if "Error Message" in data: return False,"Invalid API key. Try another one."
    if "bestMatches" in data: return True,"API key is valid and saved ✔️ "
    return False,"Unexpected response."

def open_settings_window():
    win = ctk.CTkToplevel(root)
    win.title("Preferences")
    win.geometry("420x300")
    win.transient(root)
    win.grab_set()

    ctk.CTkLabel(win, text="Alpha Vantage API Key:", font=("Segoe UI",12,"bold")).pack(anchor="w", padx=10, pady=6)
    api_entry = ctk.CTkEntry(win, width=400)
    api_entry.pack(padx=10)
    api_entry.insert(0, settings.get("api_key",""))

    def open_alpha_vantage(): webbrowser.open("https://www.alphavantage.co/support/#api-key")
    link = ctk.CTkLabel(win,text="Get a free Alpha Vantage API key", text_color="green", justify="center", cursor="hand2")
    link.pack(anchor="w", padx=10, pady=(4,8))
    link.bind("<Button-1>", lambda e: open_alpha_vantage())

    status_label = ctk.CTkLabel(win, text="")
    status_label.pack(anchor="w", padx=10, pady=6)

    def test_key():
        ok,msg = test_alpha_vantage_key(api_entry.get().strip())
        status_label.configure(text=msg, text_color="green" if ok else "red")
        settings["api_key_validated"] = ok
        settings["api_key"] = api_entry.get().strip()
        save_settings(settings)

    ctk.CTkButton(win, text="Save & Test API Key", command=test_key).pack(anchor="w", padx=10)

# ==============================
# About
# ==============================

def show_about():
    messagebox.showinfo(
        "About",
        f"{APP_NAME} v{APP_VERSION}\n\n"
        f"Developed & crafted by {DEVELOPER}\n"
        f"© {YEAR}\n\n"
        "Built with Python, Tkinter & Financial APIs."
    )

# ==============================
# GUI
# ==============================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

root = ctk.CTk()
root.title(f"{APP_NAME} v{APP_VERSION}")
root.geometry("900x650")
root.minsize(750,550)

# Menu
def handle_menu(choice):
    if choice == "Settings":
        open_settings_window()
    elif choice == "About":
        show_about()

menubar = ctk.CTkOptionMenu(
    root,
    values=["Settings", "About"],
    command=handle_menu
)

menubar.set("Menu")
menubar.pack(anchor="ne", pady=5, padx=10)

# Main frame
frame = ctk.CTkFrame(root, corner_radius=15, fg_color="#1f1f1f")
frame.pack(fill="both", expand=True, padx=20, pady=10)
frame.columnconfigure(1, weight=1)
frame.rowconfigure(5, weight=1)

# Header
ctk.CTkLabel(frame, text="Financial Data Viewer", font=("Segoe UI",20,"bold")).grid(row=0,column=0,columnspan=2,pady=(0,20))

# Stock symbol
ctk.CTkLabel(frame,text="Stock Symbol:").grid(row=1,column=0,sticky="w", pady=6)
symbol_entry = ctk.CTkEntry(frame, placeholder_text="e.g. AAPL")
symbol_entry.grid(row=1,column=1,sticky="ew", pady=6)

# Report type
ctk.CTkLabel(frame,text="Report Type:").grid(row=2,column=0,sticky="w", pady=6)
period_var = ctk.StringVar(value="Annual")
ctk.CTkOptionMenu(frame, variable=period_var, values=["Annual","Quarter"]).grid(row=2,column=1,sticky="ew", pady=6)

# Number of years
ctk.CTkLabel(frame,text="Number of Years:").grid(row=3,column=0,sticky="w", pady=6)
years_entry = ctk.CTkEntry(frame, placeholder_text="e.g. 15")
years_entry.grid(row=3,column=1,sticky="ew", pady=6)

# Buttons
button_frame = ctk.CTkFrame(frame, fg_color="transparent")
button_frame.grid(row=4,column=0,columnspan=2,pady=15)
button_frame.columnconfigure((0,1), weight=1)
run_button = ctk.CTkButton(button_frame,text="Run Analysis", command=run_analysis,
                           state="normal" if settings.get("api_key_validated") else "disabled")
run_button.grid(row=0,column=0,padx=10)
save_button = ctk.CTkButton(button_frame,text="Export to Excel", state="disabled")
save_button.grid(row=0,column=1,padx=10)

# Text + progress bar frame
text_frame = ctk.CTkFrame(frame, fg_color="transparent")
text_frame.grid(row=5,column=0,columnspan=2,sticky="nsew", pady=10)
text_frame.columnconfigure(0,weight=1)
text_frame.rowconfigure(0,weight=1)
text_frame.rowconfigure(1,weight=0)

result_text = ctk.CTkTextbox(text_frame, font=("Consolas",15))
result_text.grid(row=0,column=0,sticky="nsew", pady=(0,5))
progress_bar = ctk.CTkProgressBar(
    frame,
    width=400,
    height=20,
    corner_radius=10,
    fg_color="#2b2b2b",    # background color
    progress_color="#4cd964"  # actual progress fill color
)
progress_bar.grid(row=6, column=0, columnspan=2, sticky="ew", pady=10)

# start at 0
progress_bar.set(0)

# Footer

footer = ctk.CTkLabel(
    root,
    text=f"© {YEAR} {DEVELOPER} — {APP_NAME}",
    font=("Segoe UI", 9)
)
footer.pack(side="bottom", pady=5)

root.mainloop()
