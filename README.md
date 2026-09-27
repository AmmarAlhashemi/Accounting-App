# 📊 Accounting Automation & Reconciliation Bridge

[![Platform](https://img.shields.io/badge/Platform-Windows-blue.svg)](https://microsoft.com)
[![Python
Version](https://img.shields.io/badge/Python-3.12.10%2B-green.svg)](https://python.org)
[![License:
MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

An advanced desktop application built with **Python & Tkinter** designed
to automate custom Excel accounting entries, streamline PDF report
generation, and perform automated PDF-to-Excel statement
reconciliations.

------------------------------------------------------------------------

## 🎯 Business Context & Customization

> **Note:** This project is a **custom-tailored software solution**
> engineered specifically for a dedicated accounting workflow.

Unlike generic off-the-shelf accounting software, this application was
custom-built to support specific business rules, ledger structures, fee
calculations (such as automatic 2% deposit fees), and statement parsing
algorithms for an individual client workspace (**NNN AC**). While the
architecture serves as an extensible showcase for Python desktop
automation, the underlying processing rules are tightly coupled to
custom Excel templates and PDF formats.

------------------------------------------------------------------------

## ⚠️ Compatibility Notice & Requirements

> **Important:** This application strictly requires a **Windows
> Operating System** with **Microsoft Excel** installed locally.

-   **Operating System:** Windows 10 / 11 (64-bit).
-   **Microsoft Office:** A local installation of **Microsoft Excel is
    mandatory**. The PDF printing and exporting engine leverages Windows
    COM Automation (`win32com.client` / Excel COM API) to dynamically
    determine page boundaries, apply print setups, and render native
    Excel sheets directly to PDF.
-   **Cross-Platform Limitations:** Linux and macOS are **not
    supported** out of the box due to reliance on Windows-native COM
    bindings.

------------------------------------------------------------------------

## 🌟 Key Features & Business Value

-   **Structured Data Entry Queue (Batch Processing):** Buffer
    transactions in a staging queue before applying changes to Excel.
    Prevents manual errors and broken formulas while automatically
    updating `SUM` formulas and row styling.
-   **Automated PDF vs. Excel Reconciliation Engine:** Extracts
    statement data from PDF files using `PyMuPDF` and cross-checks it
    against targeted Excel ledger rows. Generates a formatted 3-sheet
    audit report (`Unmatched`, `Matched`, `All Records`) complete with
    `SUMIF` formulas and color-coded rows.
-   **Selective Range PDF Export:** Allows users to choose the exact
    number of transaction blocks to print directly from the GUI.
    Configures Excel print ranges (`PrintArea`) dynamically via Excel
    COM API.
-   **Productivity-Focused UI:** Features Dark/Light theme switching,
    automated field population for recurring tasks, and comprehensive
    keyboard shortcut support (`Alt+D`, `Alt+R`, `Alt+A`, `Ctrl+S`,
    `Ctrl+P`).

------------------------------------------------------------------------

## 🏗️ Project Architecture

Built with a **Modular Architecture** to separate user interface
elements from core business logic, Excel/PDF parsers, and file
management:

``` text
Accounting-App/
│
├── app.py                   # Main entry point
├── config.py                # Environment configurations & dynamic path management
│
├── core/                    # Business Logic Layer
│   ├── excel_handler.py     # OpenPyxl operations (row insertion, styling, formulas)
│   ├── pdf_exporter.py      # Win32COM Excel print driver & PDF exporter
│   └── reconciler.py        # PyMuPDF parser & statement reconciliation logic
│
├── ui/                      # User Interface Layer
│   └── main_window.py       # Tkinter GUI layout, event handling, and theme state
│
└── data/                    # Local storage (Git-ignored)
    ├── NNN AC.xlsx     # Primary Excel ledger
    ├── pdf/                 # Generated PDF report archive
    └── reports/             # Output reconciliation reports
```

## 🛠️ Tech Stack

-   **GUI Framework:** Tkinter (Python Native)
-   **Excel Processing:** OpenPyxl, Pandas
-   **PDF Parsing & Extraction:** PyMuPDF (`fitz`)
-   **Windows COM Automation:** PyWin32 (`win32com.client`)
-   **Language:** Python 3.12.10+

## 🚀 Quick Start

### Prerequisites

-   **OS:** Windows 10 / 11 (Required for native Excel print & COM
    drivers)
-   **Software:** Microsoft Excel (Installed locally)
-   **Python:** 3.12.10 or higher

### Installation

1.  **Clone the repository:**

    Bash

    ``` bash
    git clone https://github.com/AmmarAlhashemi/Accounting-App.git
    cd Accounting-App
    ```

2.  **Create and activate a virtual environment:**

    Bash

    ``` bash
    python -m venv venv
    venv\Scripts\activate
    ```

3.  **Install dependencies:**

    Bash

    ``` bash
    pip install -r requirements.txt
    ```

4.  **Run the application:**

    Bash

    ``` bash
    python app.py
    ```

## 🛡️ Data Privacy & Dummy Data Notice

To demonstrate full application functionality safely on GitHub:
* **Demo Template Included:** The repository includes a sample Excel workbook (`data/NNN AC.xlsx`) pre-populated with **anonymized, dummy transaction data**. This allows immediate testing and demonstration without manual setup.
* **Excluded Generated Files:** All real transaction logs, exported PDF files (`data/pdf/`), and output reconciliation reports (`data/reports/`) are strictly excluded from version control via `.gitignore` to ensure privacy.

## 📝 License

Distributed under the MIT License. See `LICENSE` for more information.