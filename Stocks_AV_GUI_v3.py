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
import re
from typing import Optional, Dict, List, Tuple
from pathlib import Path

# ==============================
# Developer info
# ==============================
APP_NAME = "AlphaFin"
APP_VERSION = "3.3"
DEVELOPER = "Cruz-Py"
YEAR = "2026"

# ==============================
# Configuration
# ==============================
RATE_LIMIT_SLEEP = 60
NORMAL_SLEEP = 12
MAX_RETRIES = 3
CACHE_TTL_HOURS = 24
CACHE_MAX_AGE_DAYS = 7  # Clean cache files older than 7 days
API_TIMEOUT = 15
SETTINGS_WINDOW_WIDTH = 420
SETTINGS_WINDOW_HEIGHT = 300
APP_WINDOW_WIDTH = 900
APP_WINDOW_HEIGHT = 650
APP_WINDOW_MIN_WIDTH = 750
APP_WINDOW_MIN_HEIGHT = 550

# Regex for valid stock symbols (1-5 uppercase letters)
VALID_SYMBOL_PATTERN = r'^[A-Z]{1,5}$'

# ==============================
# App directory
# ==============================
def get_app_directory() -> str:
    """Get the application directory."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_directory()
CACHE_DIR = os.path.join(APP_DIR, "cache")
os.makedirs(CACHE_DIR, exist_ok=True)

# ==============================
# Settings Management
# ==============================
SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")

DEFAULT_SETTINGS = {
    "api_key": "",
    "api_key_validated": False,
    "save_directory": APP_DIR
}


def load_settings() -> Dict:
    """Load settings from JSON file."""
    if not os.path.exists(SETTINGS_FILE):
        return DEFAULT_SETTINGS.copy()
    try:
        with open(SETTINGS_FILE, "r") as f:
            data = json.load(f)
            return {**DEFAULT_SETTINGS, **data}
    except Exception as e:
        print(f"Error loading settings: {e}")
        return DEFAULT_SETTINGS.copy()


def save_settings(settings: Dict) -> None:
    """Save settings to JSON file."""
    try:
        with open(SETTINGS_FILE, "w") as f:
            json.dump(settings, f, indent=4)
    except Exception as e:
        print(f"Error saving settings: {e}")


settings = load_settings()
settings_lock = threading.Lock()


def get_api_key() -> str:
    """Thread-safe API key getter."""
    with settings_lock:
        return settings.get("api_key", "")


def set_api_key(key: str, validated: bool = False) -> None:
    """Thread-safe API key setter."""
    with settings_lock:
        settings["api_key"] = key
        settings["api_key_validated"] = validated
        save_settings(settings)


# ==============================
# Cache Utilities
# ==============================
def get_cache_path(symbol: str, function_name: str, period: str, year: str) -> str:
    """Get cache file path for a specific query."""
    filename = f"{symbol}_{function_name}_{period}_{year}.json"
    return os.path.join(CACHE_DIR, filename)


def is_cache_valid(path: str) -> bool:
    """Check if cache file exists and is not expired."""
    if not os.path.exists(path):
        return False
    age_seconds = time.time() - os.path.getmtime(path)
    return age_seconds < CACHE_TTL_HOURS * 3600


def load_from_cache(symbol: str, function_name: str, period: str, year: str) -> Optional[List]:
    """Load data from cache if valid."""
    path = get_cache_path(symbol, function_name, period, year)
    if not is_cache_valid(path):
        return None
    try:
        with open(path, "r") as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading cache: {e}")
        return None


def save_to_cache(symbol: str, function_name: str, period: str, year: str, data: List) -> None:
    """Save data to cache."""
    path = get_cache_path(symbol, function_name, period, year)
    try:
        with open(path, "w") as f:
            json.dump(data, f)
    except Exception as e:
        print(f"Error saving to cache: {e}")


def clean_old_cache(max_age_days: int = CACHE_MAX_AGE_DAYS) -> None:
    """Remove cache files older than max_age_days."""
    try:
        current_time = time.time()
        max_age_seconds = max_age_days * 24 * 3600
        
        for file in os.listdir(CACHE_DIR):
            file_path = os.path.join(CACHE_DIR, file)
            if os.path.isfile(file_path):
                file_age = current_time - os.path.getmtime(file_path)
                if file_age > max_age_seconds:
                    os.remove(file_path)
    except Exception as e:
        print(f"Error cleaning cache: {e}")


# ==============================
# Input Validation
# ==============================
def is_valid_symbol(symbol: str) -> bool:
    """Validate stock symbol format."""
    return bool(re.match(VALID_SYMBOL_PATTERN, symbol.upper()))


def validate_years_count(years: int) -> bool:
    """Validate number of years."""
    return 1 <= years <= 50


# ==============================
# Alpha Vantage API
# ==============================
def alpha_vantage_request(function_name: str, symbol: str, api_key: str) -> Tuple[Optional[Dict], Optional[str]]:
    """Make request to Alpha Vantage API with error handling."""
    url = "https://www.alphavantage.co/query"
    params = {"function": function_name, "symbol": symbol, "apikey": api_key}

    try:
        response = requests.get(url, params=params, timeout=API_TIMEOUT)
        data = response.json()
        
        if "Note" in data:
            return None, "API rate limit reached. Please wait 60 seconds."
        if "Error Message" in data:
            return None, "Invalid API key or request."
        if "Information" in data:
            return None, "API request throttled. Please wait."
        
        return data, None
    except requests.Timeout:
        return None, "Request timeout. Please check your connection."
    except requests.ConnectionError:
        return None, "Connection error. Please check your internet connection."
    except Exception as e:
        return None, f"Connection error: {str(e)}"


def validate_api_key(api_key: str) -> Tuple[bool, str]:
    """
    Validate API key by making a real API call.
    Returns (is_valid, message)
    """
    if not api_key or not api_key.strip():
        return False, "API key cannot be empty."
    
    try:
        # Use TIME_SERIES_DAILY which is a reliable endpoint for validation
        url = "https://www.alphavantage.co/query"
        params = {
            "function": "TIME_SERIES_DAILY",
            "symbol": "IBM",
            "apikey": api_key.strip()
        }
        
        response = requests.get(url, params=params, timeout=API_TIMEOUT)
        data = response.json()
        
        # Check for errors
        if "Error Message" in data:
            return False, "Invalid API key. Please check and try again."
        
        # Check for rate limiting
        if "Note" in data:
            return False, "API rate limit reached. Please try again in a minute."
        
        # Check for information/throttling messages
        if "Information" in data:
            return False, "API is temporarily throttled. Please try again in a moment."
        
        # If we got time series data, the key is valid
        if "Time Series (Daily)" in data or "Meta Data" in data:
            return True, "API key is valid!"
        
        # If we got here without errors or rate limits, key is likely valid
        # Some responses might not include data if symbol is invalid, but key is valid
        return True, "API key validated successfully!"
    
    except requests.Timeout:
        return False, "Connection timeout. Please check your internet connection."
    except requests.ConnectionError:
        return False, "Connection error. Please check your internet connection."
    except Exception as e:
        return False, f"Validation error: {str(e)}"


# ==============================
# Fetch Financial Statements
# ==============================
def fetch_financial_statements(
    symbol: str,
    api_key: str,
    period: str = "annual",
    years: int = 15,
    log_callback=None,
    progress_callback=None
) -> Dict[str, pd.DataFrame]:
    """
    Fetch financial statements from Alpha Vantage.
    
    Args:
        symbol: Stock symbol
        api_key: Alpha Vantage API key
        period: "annual" or "quarter"
        years: Number of years to fetch
        log_callback: Callback function for logging
        progress_callback: Callback function for progress updates
    
    Returns:
        Dictionary with financial statement DataFrames
    """
    function_map = {
        "income_statement": "INCOME_STATEMENT",
        "balance_sheet": "BALANCE_SHEET",
        "cash_flow": "CASH_FLOW"
    }

    financials = {}
    total_steps = len(function_map)
    completed_steps = 0
    year = str(datetime.now().year)
    
    # Calculate limit: for quarters, multiply by 4; for annual, use years as-is
    if period.lower() == "quarter":
        limit = years * 4  # 4 quarters per year
    else:
        limit = years  # 1 annual report per year

    for key, func in function_map.items():
        # Try to load from cache
        cached = load_from_cache(symbol, func, period.lower(), year)
        if cached:
            if log_callback:
                log_callback(f"Loaded {key.replace('_', ' ')} from cache ✔️\n")
            
            try:
                df = pd.DataFrame(cached)
                if not df.empty and "fiscalDateEnding" in df.columns:
                    # Sort by date and get the last 'limit' entries
                    df = df.sort_values("fiscalDateEnding").tail(limit)
                financials[key] = df
            except Exception as e:
                print(f"Error processing cached data: {e}")
                financials[key] = pd.DataFrame()
            
            completed_steps += 1
            if progress_callback:
                progress_callback((completed_steps / total_steps) * 100)
            continue

        # Fetch from API with retries
        retries = 0
        while retries <= MAX_RETRIES:
            if log_callback:
                log_callback(f"Fetching {key.replace('_', ' ')}...\n")
            
            try:
                data, error = alpha_vantage_request(func, symbol, api_key)
                
                if error:
                    if log_callback:
                        log_callback(f"{error}\n")
                    financials[key] = pd.DataFrame()
                    break
                
                if "Note" in data:
                    retries += 1
                    if log_callback:
                        log_callback(f"⚠️ Rate limit hit. Waiting {RATE_LIMIT_SLEEP}s...\n")
                    time.sleep(RATE_LIMIT_SLEEP)
                    continue
                
                if "Error Message" in data:
                    if log_callback:
                        log_callback("API error returned by Alpha Vantage.\n")
                    financials[key] = pd.DataFrame()
                    break

                # Extract reports based on period
                reports_key = "quarterlyReports" if period.lower() == "quarter" else "annualReports"
                reports = data.get(reports_key, [])
                
                if not reports:
                    if log_callback:
                        log_callback(f"No {key} data available for {symbol}.\n")
                    financials[key] = pd.DataFrame()
                    break
                
                # Save to cache and process
                save_to_cache(symbol, func, period.lower(), year, reports)
                df = pd.DataFrame(reports)
                if not df.empty:
                    # Sort by fiscal date and get the last 'limit' entries
                    df = df.sort_values("fiscalDateEnding").tail(limit)
                financials[key] = df
                
                if log_callback:
                    record_count = len(df)
                    log_callback(f"✔️ Retrieved {record_count} {period} records for {key.replace('_', ' ')}\n")
                
                break

            except Exception as e:
                if log_callback:
                    log_callback(f"Request failed: {str(e)}\n")
                financials[key] = pd.DataFrame()
                break

        completed_steps += 1
        if progress_callback:
            progress_callback((completed_steps / total_steps) * 100)
        
        time.sleep(NORMAL_SLEEP)

    return financials


# ==============================
# Fetch Year-End Prices
# ==============================
def fetch_year_end_closing_prices_yf(symbol: str, years: List[int]) -> Dict[str, Optional[float]]:
    """Fetch year-end closing prices using yfinance."""
    try:
        start_date = datetime(min(years) - 1, 1, 1)
        end_date = datetime(max(years) + 1, 1, 1)
        
        data = yf.Ticker(symbol).history(start=start_date, end=end_date)
        
        if data.empty:
            return {}
        
        closing_prices = {}
        for year in years:
            year_data = data[data.index.year == year]
            if not year_data.empty:
                last_day = year_data.index.max()
                closing_prices[str(year)] = round(year_data.loc[last_day]["Close"], 2)
            else:
                closing_prices[str(year)] = None
        
        return closing_prices
    except Exception as e:
        print(f"Error fetching year-end prices: {e}")
        return {}


# ==============================
# Save to Excel
# ==============================
def save_to_excel(financials: Dict[str, pd.DataFrame], 
                  closing_prices: Dict[str, float], 
                  symbol: str) -> bool:
    """Save financial data to Excel file."""
    base_dir = settings.get("save_directory", APP_DIR)
    
    if not os.path.isdir(base_dir):
        messagebox.showerror("Invalid Folder", "Save directory does not exist.")
        return False
    
    # Validate that we have data to save
    has_data = any(not df.empty for df in financials.values())
    if not has_data:
        messagebox.showwarning("No Data", "No financial data to export. Please run analysis first.")
        return False
    
    filename = os.path.join(base_dir, f"{symbol}_financials.xlsx")
    
    try:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # Save financial statements
            for sheet_name, df in financials.items():
                if df.empty:
                    continue
                
                try:
                    if "fiscalDateEnding" in df.columns:
                        df_t = df.set_index("fiscalDateEnding").transpose()
                        df_t.columns = [c[:4] for c in df_t.columns]
                    else:
                        df_t = df.transpose()
                    
                    # Convert to numeric and handle missing values
                    df_t = df_t.replace([None, "None", pd.NA], 0)
                    df_t = df_t.apply(pd.to_numeric, errors="coerce").fillna(0)
                    
                    df_t.to_excel(writer, sheet_name=sheet_name.replace("_", " ").title())
                except Exception as e:
                    print(f"Error writing {sheet_name}: {e}")
                    continue
            
            # Save closing prices
            if closing_prices:
                cp_df = pd.DataFrame(closing_prices.items(), columns=["Year", "Closing Price"])
                cp_df.to_excel(writer, sheet_name="Year-End Closing Prices", index=False)
        
        messagebox.showinfo("Saved", f"File saved successfully:\n\n{filename}")
        return True
    
    except Exception as e:
        messagebox.showerror("Save Error", f"Failed to save file:\n\n{str(e)}")
        return False


# ==============================
# Main Application Class
# ==============================
class FinancialDataApp:
    """Main application class for Financial Data Extractor."""
    
    def __init__(self, root: ctk.CTk):
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry(f"{APP_WINDOW_WIDTH}x{APP_WINDOW_HEIGHT}")
        self.root.minsize(APP_WINDOW_MIN_WIDTH, APP_WINDOW_MIN_HEIGHT)
        
        # Analysis state
        self.financials: Optional[Dict[str, pd.DataFrame]] = None
        self.closing_prices: Optional[Dict[str, float]] = None
        self.current_symbol: Optional[str] = None
        self.analysis_running = False
        
        # Setup theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")
        
        # Build GUI
        self.setup_gui()
        
        # Load settings
        self.update_button_states()
    
    def setup_gui(self) -> None:
        """Build the user interface."""
        # Menu bar
        menu_var = ctk.StringVar(value="Menu")
        
        def handle_menu(choice):
            if choice == "Settings":
                self.open_settings_window()
            elif choice == "About":
                self.show_about()
            menu_var.set("Menu")
        
        menubar = ctk.CTkOptionMenu(
            self.root,
            variable=menu_var,
            values=["Settings", "About"],
            command=handle_menu
        )
        menubar.pack(anchor="ne", pady=5, padx=10)
        
        # Main frame
        self.main_frame = ctk.CTkFrame(self.root, corner_radius=15, fg_color="#1f1f1f")
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(5, weight=1)
        
        # Header
        ctk.CTkLabel(
            self.main_frame,
            text="AlphaFin - Financial Data Extractor",
            font=("Segoe UI", 20, "bold")
        ).grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Input fields
        ctk.CTkLabel(self.main_frame, text="Stock Symbol:").grid(
            row=1, column=0, sticky="w", pady=6
        )
        self.symbol_entry = ctk.CTkEntry(self.main_frame, placeholder_text="e.g. AAPL")
        self.symbol_entry.grid(row=1, column=1, sticky="ew", pady=6)
        
        ctk.CTkLabel(self.main_frame, text="Report Type:").grid(
            row=2, column=0, sticky="w", pady=6
        )
        self.period_var = ctk.StringVar(value="Annual")
        self.period_menu = ctk.CTkOptionMenu(
            self.main_frame,
            variable=self.period_var,
            values=["Annual", "Quarter"]
        )
        self.period_menu.grid(row=2, column=1, sticky="ew", pady=6)
        
        ctk.CTkLabel(self.main_frame, text="Number of Years:").grid(
            row=3, column=0, sticky="w", pady=6
        )
        self.years_entry = ctk.CTkEntry(self.main_frame, placeholder_text="e.g. 15")
        self.years_entry.grid(row=3, column=1, sticky="ew", pady=6)
        
        # Input validation
        vcmd = (self.root.register(self.validate_years_input), "%P")
        self.years_entry.configure(validate="key", validatecommand=vcmd)
        
        # Buttons
        button_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        button_frame.grid(row=4, column=0, columnspan=2, pady=15)
        button_frame.columnconfigure((0, 1), weight=1)
        
        self.run_button = ctk.CTkButton(
            button_frame,
            text="Run Analysis",
            command=self.run_analysis,
            state="normal" if settings.get("api_key_validated") else "disabled"
        )
        self.run_button.grid(row=0, column=0, padx=10)
        
        self.save_button = ctk.CTkButton(
            button_frame,
            text="Export to Excel",
            command=self.save_to_excel,
            state="disabled"
        )
        self.save_button.grid(row=0, column=1, padx=10)
        
        # Text output and progress
        text_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        text_frame.grid(row=5, column=0, columnspan=2, sticky="nsew", pady=10)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.result_text = ctk.CTkTextbox(text_frame, font=("Consolas", 15))
        self.result_text.grid(row=0, column=0, sticky="nsew", pady=(0, 5))
        
        self.progress_bar = ctk.CTkProgressBar(
            self.main_frame,
            width=400,
            height=20,
            corner_radius=10,
            fg_color="#2b2b2b",
            progress_color="#4cd964"
        )
        self.progress_bar.grid(row=6, column=0, columnspan=2, sticky="ew", pady=10)
        self.progress_bar.set(0)
        
        # Footer
        footer = ctk.CTkLabel(
            self.root,
            text=f"© {YEAR} {DEVELOPER} — {APP_NAME}",
            font=("Segoe UI", 9)
        )
        footer.pack(side="bottom", pady=5)
    
    def validate_years_input(self, value: str) -> bool:
        """Validate years input field."""
        if value == "":
            return True
        try:
            return validate_years_count(int(value))
        except ValueError:
            return False
    
    def update_button_states(self) -> None:
        """Update button enabled/disabled states based on app state."""
        api_key_valid = settings.get("api_key_validated", False)
        has_data = self.financials is not None and any(
            not df.empty for df in self.financials.values()
        )
        
        self.run_button.configure(state="normal" if api_key_valid else "disabled")
        self.save_button.configure(state="normal" if has_data else "disabled")
    
    def run_analysis(self) -> None:
        """Start financial data analysis."""
        if not settings.get("api_key"):
            messagebox.showwarning(
                "API Key Missing",
                "Please enter your Alpha Vantage API key in Settings."
            )
            return
        
        # Validate input
        symbol = self.symbol_entry.get().strip().upper()
        if not symbol:
            messagebox.showerror("Input Error", "Please enter a stock symbol.")
            return
        
        if not is_valid_symbol(symbol):
            messagebox.showerror(
                "Invalid Symbol",
                f"'{symbol}' is not a valid stock symbol.\nPlease use 1-5 uppercase letters."
            )
            return
        
        years_input = self.years_entry.get().strip()
        if not years_input:
            messagebox.showerror("Input Error", "Please enter the number of years.")
            return
        
        try:
            years = int(years_input)
            if not validate_years_count(years):
                raise ValueError
        except ValueError:
            messagebox.showerror("Input Error", "Please enter 1–50 years.")
            return
        
        # Update UI state
        self.analysis_running = True
        self.run_button.configure(state="disabled")
        self.save_button.configure(state="disabled")
        self.progress_bar.set(0)
        self.result_text.delete("1.0", "end")
        self.disable_inputs()
        
        # Start analysis in background thread
        period = self.period_var.get().strip().lower()
        thread = threading.Thread(
            target=self.run_analysis_thread,
            args=(symbol, period, years),
            daemon=True
        )
        thread.start()
    
    def run_analysis_thread(self, symbol: str, period: str, years: int) -> None:
        """Run analysis in a separate thread."""
        try:
            current_year = datetime.now().year
            years_list = list(range(current_year - years + 1, current_year + 1))
            
            period_display = "quarterly" if period == "quarter" else "annual"
            self.safe_log(f"Fetching {period_display} financial data for {symbol}...\n\n")
            
            # Clean old cache periodically
            clean_old_cache()
            
            # Fetch financial statements (pass years count, not limit)
            self.financials = fetch_financial_statements(
                symbol,
                get_api_key(),
                period=period,
                years=years,  # This will be multiplied by 4 for quarters inside the function
                log_callback=self.safe_log,
                progress_callback=self.safe_progress
            )
            
            # Fetch closing prices
            self.closing_prices = fetch_year_end_closing_prices_yf(symbol, years_list)
            self.current_symbol = symbol
            
            # Display closing prices
            self.safe_log("\nYear-end Closing Prices:\n")
            for y in years_list:
                price = self.closing_prices.get(str(y))
                price_str = f"${price}" if price is not None else "N/A"
                self.safe_log(f"  {y}: {price_str}\n")
            
            # Check if we got any data
            has_data = any(not df.empty for df in self.financials.values())
            if has_data:
                self.safe_log(
                    "\n✔️ Financial data successfully extracted!\n"
                    "Please click 'Export to Excel' to save it locally.\n"
                )
            else:
                self.safe_log(
                    "\n⚠️ No financial data was retrieved for this symbol.\n"
                    "Please check the symbol and try again.\n"
                )
            
            self.safe_progress(100)
        
        except Exception as e:
            self.safe_log(f"\n❌ Error: {str(e)}\n")
        
        finally:
            self.root.after(0, self.finalize_analysis)
    
    def finalize_analysis(self) -> None:
        """Finalize analysis and update UI."""
        self.analysis_running = False
        self.enable_inputs()
        self.update_button_states()
    
    def save_to_excel(self) -> None:
        """Save analysis results to Excel file."""
        if self.financials is None or self.current_symbol is None:
            messagebox.showwarning("No Data", "Please run analysis first.")
            return
        
        self.save_button.configure(state="disabled")
        
        try:
            success = save_to_excel(
                self.financials,
                self.closing_prices or {},
                self.current_symbol
            )
            if success:
                # Reset analysis state after successful save
                pass
        finally:
            self.save_button.configure(state="normal")
    
    def disable_inputs(self) -> None:
        """Disable input fields during analysis."""
        self.symbol_entry.configure(state="disabled")
        self.years_entry.configure(state="disabled")
        self.period_menu.configure(state="disabled")
    
    def enable_inputs(self) -> None:
        """Enable input fields after analysis."""
        self.symbol_entry.configure(state="normal")
        self.years_entry.configure(state="normal")
        self.period_menu.configure(state="normal")
    
    def safe_log(self, message: str) -> None:
        """Thread-safe logging to text widget."""
        self.root.after(0, lambda: (
            self.result_text.insert("end", message),
            self.result_text.see("end")
        ))
    
    def safe_progress(self, value: float) -> None:
        """Thread-safe progress bar update."""
        self.root.after(0, lambda: self.progress_bar.set(value / 100))
    
    def open_settings_window(self) -> None:
        """Open settings window."""
        settings_window = ctk.CTkToplevel(self.root)
        settings_window.title("Preferences")
        settings_window.geometry(f"{SETTINGS_WINDOW_WIDTH}x{SETTINGS_WINDOW_HEIGHT}")
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        # Center window
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (SETTINGS_WINDOW_WIDTH // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (SETTINGS_WINDOW_HEIGHT // 2)
        settings_window.geometry(f"{SETTINGS_WINDOW_WIDTH}x{SETTINGS_WINDOW_HEIGHT}+{x}+{y}")
        
        # API Key Label
        ctk.CTkLabel(
            settings_window,
            text="Alpha Vantage API Key:",
            font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=10, pady=6)
        
        # API Key Entry
        api_entry = ctk.CTkEntry(settings_window, width=400)
        api_entry.pack(padx=10, pady=(0, 5))
        api_entry.insert(0, settings.get("api_key", ""))
        
        # Link to get API key
        def open_alpha_vantage():
            webbrowser.open("https://www.alphavantage.co/support/#api-key")
        
        link = ctk.CTkLabel(
            settings_window,
            text="Get a free Alpha Vantage API key",
            text_color="green",
            justify="center",
            cursor="hand2"
        )
        link.pack(anchor="w", padx=10, pady=(4, 8))
        link.bind("<Button-1>", lambda e: open_alpha_vantage())
        
        # Status label for feedback
        status_label = ctk.CTkLabel(
            settings_window,
            text="",
            text_color="orange",
            font=("Segoe UI", 10)
        )
        status_label.pack(pady=5)
        
        # Save API Key button
        def save_api_key():
            key = api_entry.get().strip()
            if not key:
                messagebox.showerror("Invalid Key", "API key cannot be empty.")
                return
            
            status_label.configure(text="Validating API key...", text_color="orange")
            settings_window.update()
            
            is_valid, message = validate_api_key(key)
            
            if not is_valid:
                status_label.configure(text=f"❌ {message}", text_color="red")
                messagebox.showerror("Invalid Key", message)
                return
            
            set_api_key(key, validated=True)
            status_label.configure(text="✔️ API key validated!", text_color="green")
            messagebox.showinfo(
                "Saved",
                "API key validated and saved successfully!"
            )
            self.update_button_states()
            settings_window.after(1500, settings_window.destroy)
        
        ctk.CTkButton(
            settings_window,
            text="Save API Key",
            command=save_api_key
        ).pack(pady=15)
        
        # Note label
        ctk.CTkLabel(
            settings_window,
            text="Note: API key is validated automatically when saving.",
            font=("Segoe UI", 10),
            text_color="#AAAAAA",
            wraplength=380,
            justify="left"
        ).pack(pady=5, padx=10)
    
    def show_about(self) -> None:
        """Show about window."""
        about_window = ctk.CTkToplevel(self.root)
        about_window.title("About")
        about_window.geometry("400x250")
        about_window.transient(self.root)
        about_window.grab_set()
        
        # Center window
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (200)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (125)
        about_window.geometry(f"400x250+{x}+{y}")
        
        # About text
        ctk.CTkLabel(
            about_window,
            text=f"{APP_NAME} v{APP_VERSION}\n\n"
                 f"Developed & crafted by {DEVELOPER}\n"
                 f"© {YEAR}\n\n"
                 f"Built with Python, CustomTkinter & Financial APIs.",
            font=("Segoe UI", 12),
            justify="center"
        ).pack(pady=(20, 10), padx=20)
        
        # Donation link
        def open_paypal():
            webbrowser.open("https://paypal.me/ruicruz27")
        
        link = ctk.CTkLabel(
            about_window,
            text="Support development: PayPal / Donate",
            font=("Segoe UI", 12, "underline"),
            text_color="green",
            cursor="hand2"
        )
        link.pack(pady=10)
        link.bind("<Button-1>", lambda e: open_paypal())
        
        # Close button
        ctk.CTkButton(
            about_window,
            text="Close",
            command=about_window.destroy
        ).pack(pady=10)


# ==============================
# Main Entry Point
# ==============================
def main():
    """Main entry point."""
    root = ctk.CTk()
    app = FinancialDataApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()