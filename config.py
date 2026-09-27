import os
from datetime import datetime

# --- Local Project Root Directory ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- Current Year ---
CURRENT_YEAR = datetime.now().strftime("%Y")

# --- Dynamic Data Directories ---
DATA_DIR = os.path.join(BASE_DIR, "data")
PDF_DIR = os.path.join(DATA_DIR, "pdf")
REPORTS_DIR = os.path.join(DATA_DIR, "reports")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

# --- Dynamic File Paths ---
BASE_EXCEL_PATH = os.path.join(DATA_DIR, f"NNN AC.xlsx")
BASE_PDF_DIR = PDF_DIR
RECON_OUTPUT_PATH = os.path.join(REPORTS_DIR, "Temp_Reconciliation_Report.xlsx")

# --- UI Styles ---
FONT_NORM = ("Segoe UI", 11)
FONT_BOLD = ("Segoe UI", 11, "bold")
FONT_BTN1 = ("Segoe UI", 12, "bold")
FONT_CODE = ("Consolas", 9)