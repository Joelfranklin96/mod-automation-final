# Mod Automation

A Streamlit application that runs an 8-script automation pipeline for HHS contract
modification processing. Upload the required input files, run the pipeline (fully or
step-by-step), and download the generated outputs — all from a single web UI.

## Project layout

```
mod-automation-final/
├── app.py                 # Streamlit UI (entry point)
├── requirements.txt       # Python dependencies
├── final_scripts/         # The 8 pipeline scripts + shared config
│   ├── config.py          # Central paths & per-cycle settings
│   ├── 01_final_coversheet_generation_script.py
│   ├── 02_overview_file_script.py
│   ├── 03_F_and_R_script.py
│   ├── 04_build_file_script.py
│   ├── 05_J1_script.py
│   ├── 06_J17_file_script.py
│   ├── 07_MFR_walkthrough_script.py
│   └── 08_SF30_script.py
└── Data/                  # Input/output data (created automatically, not tracked in git)
    ├── input/
    ├── pr_files/
    └── output/
        └── coversheets/
```

> The `Data/` folder is git-ignored because it holds run-specific input/output files.
> The app recreates the required subfolders automatically on startup, so a fresh clone
> works out of the box.

## Requirements

- **Python 3.9+**
- **Microsoft Excel (Windows)** — Scripts 02 and 03 use `xlwings`, which drives Excel via
  COM automation to read checkbox states and copy sheets. These steps require Excel to be
  installed and will not work on machines without it (e.g. most Linux/macOS setups).

## Setup

```bash
# 1. Clone
git clone https://github.com/Joelfranklin96/mod-automation-final.git
cd mod-automation-final

# 2. Create and activate a virtual environment
python -m venv .venv
# Windows (PowerShell):
.venv\Scripts\Activate.ps1
# macOS/Linux:
# source .venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Launch the app
streamlit run app.py
```

The app opens in your browser (default http://localhost:8501).

## Per-cycle configuration

Before running a new modification cycle, update the values near the bottom of
`final_scripts/config.py`:

```python
CURRENT_OPTION_PERIOD = 5
MOD_NUMBER = "P00078"
CURRENT_MONTH = "December"
```

The EIS billing filename is also referenced in `config.py` (`EIS_BILLING_FILE`) and in
`app.py` (`EIS_BILLING_HARDCODED`) — update both to match the billing file for the current
month.

## Using the app

1. **Setup tab** — Upload the required input files (they are saved into `Data/input/` with
   the expected names), the EIS billing file, and the PR `.xlsx` files.
2. **Run Pipeline tab** — Select which of the 8 scripts to run and execute them
   sequentially.
3. **Results tab** — Download the generated files from `Data/output/`.

### Expected input files

| File | Expected name | Used by |
|------|---------------|---------|
| J1 Previous File | `j1_previous_file.xlsx` | Scripts 01, 05 |
| CLIN Table File | `clin_table_file.xls` | Scripts 01, 02 |
| Build File | `build_file.xlsx` | Script 04 |
| J17 File | `j17_file.xlsx` | Script 06 |
| EIS Billing Detail | (see `config.py`) | Script 06 |
| PR files | any `.xlsx` in `Data/pr_files/` | Scripts 01, 03 |

## The pipeline

| # | Script | Purpose |
|---|--------|---------|
| 1 | Coversheet Generation | Generates a coversheet `.xlsx` for each PR file |
| 2 | Overview File | Builds overview from coversheets (uses `xlwings` — requires Excel) |
| 3 | F&R Overview | Builds F&R overview from the overview file and PR files (uses `xlwings`) |
| 4 | Build File Processing | Creates J.1 Automated & Catalog sheets in the build file |
| 5 | J1 Update | Reads the Catalog sheet and appends entries to the J1 file |
| 6 | J17 Update | Updates the J.17 Subscription CLINs & Terms file |
| 7 | MFR Walkthrough | Generates the MFR Walkthrough Word document |
| 8 | SF30 Continuation Pages | Generates the SF30 Continuation Pages Word document |
