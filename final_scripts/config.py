"""
Central configuration for the new-mod automation scripts.

Single source of truth (DRY) for:
  * the project root and all input/output directories
  * the concrete input/output file paths each script reads or writes
  * the per-cycle configuration values (option period, mod number, billing month)

BASE_DIR is resolved RELATIVE to this file, so the project can be moved or cloned
to any machine without editing absolute paths. This file lives in
``<project root>/final_scripts/``; ``parent.parent`` is the project root and the
data lives in the sibling ``Data`` folder.

Expected layout::

    <project root>/
        final_scripts/        <- this file lives here
            config.py
            01_..._script.py
            ...
        Data/
            input/
            output/
                coversheets/
            pr_files/
"""

from pathlib import Path

# ─── Root ─────────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent.parent / "Data"

# ─── Directories ──────────────────────────────────────────────────────────────
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
PR_DIR = BASE_DIR / "pr_files"
COVERSHEET_DIR = OUTPUT_DIR / "coversheets"

# ─── Input files ──────────────────────────────────────────────────────────────
CLIN_TABLE_FILE = INPUT_DIR / "clin_table_file.xls"
J1_PREVIOUS_FILE = INPUT_DIR / "j1_previous_file.xlsx"
BUILD_FILE_INPUT = INPUT_DIR / "build_file.xlsx"
J17_FILE = INPUT_DIR / "j17_file.xlsx"
EIS_BILLING_FILE = INPUT_DIR / "EIS Billing Detail - MAR 2026 - HHS EIS PMO 75P00120F80177.xlsx"

# ─── Output files ─────────────────────────────────────────────────────────────
BUILD_FILE_OUTPUT = OUTPUT_DIR / "build_file.xlsx"
OVERVIEW_FILE = OUTPUT_DIR / "overview_file.xlsx"
F_R_OUTPUT_FILE = OUTPUT_DIR / "f_r_output.xlsx"
J1_CURRENT_FILE = OUTPUT_DIR / "j1_current_file.xlsx"
J17_UPDATED_FILE = OUTPUT_DIR / "j17_updated_file.xlsx"

# ─── Per-cycle configuration ──────────────────────────────────────────────────
# Update these when the option period rolls over / a new mod is processed.
CURRENT_OPTION_PERIOD = 5
MOD_NUMBER = "P00078"
CURRENT_MONTH = "December"

# ─── Output documents whose names embed the mod number ────────────────────────
MFR_WALKTHROUGH_OUTPUT = OUTPUT_DIR / f"{MOD_NUMBER} MFR Walkthrough Output.docx"
SF30_OUTPUT_FILE = OUTPUT_DIR / f"{MOD_NUMBER} SF30 Output.docx"
