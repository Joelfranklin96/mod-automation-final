import streamlit as st
import subprocess
import sys
import os
import shutil
import time
from pathlib import Path
from datetime import datetime


# ─── Configuration ───────────────────────────────────────────────────────────

WORKSPACE = Path(__file__).resolve().parent
SCRIPTS_DIR = WORKSPACE / "final_scripts"
DATA_DIR = WORKSPACE / "Data"
INPUT_DIR = DATA_DIR / "input"
PR_DIR = DATA_DIR / "pr_files"
OUTPUT_DIR = DATA_DIR / "output"
COVERSHEETS_DIR = OUTPUT_DIR / "coversheets"

PIPELINE = [
    {
        "num": 1,
        "file": "01_final_coversheet_generation_script.py",
        "name": "Coversheet Generation",
        "desc": "Generates coversheet .xlsx for each PR file",
    },
    {
        "num": 2,
        "file": "02_overview_file_script.py",
        "name": "Overview File",
        "desc": "Creates overview from coversheets (uses xlwings — requires Excel)",
    },
    {
        "num": 3,
        "file": "03_F_and_R_script.py",
        "name": "F&R Overview",
        "desc": "Builds F&R overview from overview file and PR files",
    },
    {
        "num": 4,
        "file": "04_build_file_script.py",
        "name": "Build File Processing",
        "desc": "Creates J.1 Automated & Catalog sheets in the build file",
    },
    {
        "num": 5,
        "file": "05_J1_script.py",
        "name": "J1 Update",
        "desc": "Reads Catalog sheet and appends entries to J1 file",
    },
    {
        "num": 6,
        "file": "06_J17_file_script.py",
        "name": "J17 Update",
        "desc": "Updates J.17 Subscription CLINs & Terms file",
    },
    {
        "num": 7,
        "file": "07_MFR_walkthrough_script.py",
        "name": "MFR Walkthrough",
        "desc": "Generates the MFR Walkthrough Word document from j1_current_file.xlsx",
    },
    {
        "num": 8,
        "file": "08_SF30_script.py",
        "name": "SF30 Continuation Pages",
        "desc": "Generates the SF30 Continuation Pages Word document from overview_file.xlsx",
    },
]

EXPECTED_INPUT_FILES = [
    ("j1_previous_file.xlsx", "J1 Previous File", ["xlsx"], "Scripts 01, 05"),
    ("clin_table_file.xls", "CLIN Table File", ["xls"], "Scripts 01, 02"),
    ("build_file.xlsx", "Build File", ["xlsx"], "Script 04"),
    ("j17_file.xlsx", "J17 File", ["xlsx"], "Script 06"),
]

EIS_BILLING_HARDCODED = "EIS Billing Detail - FEB 2026 - HHS EIS PMO 75P00120F80177.xlsx"


# ─── Helpers ─────────────────────────────────────────────────────────────────

def ensure_dirs():
    for d in [INPUT_DIR, PR_DIR, OUTPUT_DIR, COVERSHEETS_DIR]:
        d.mkdir(parents=True, exist_ok=True)


def list_files(directory, pattern="*.xlsx"):
    if not directory.exists():
        return []
    return sorted(
        f for f in directory.glob(pattern)
        if f.is_file() and not f.name.startswith("~$")
    )


def save_uploaded(uploaded_file, dest_dir, rename_to=None):
    filename = rename_to or uploaded_file.name
    filepath = dest_dir / filename
    filepath.write_bytes(uploaded_file.getbuffer())
    return filepath


def run_script(script_filename):
    script_path = SCRIPTS_DIR / script_filename
    result = subprocess.run(
        [sys.executable, str(script_path)],
        capture_output=True,
        text=True,
        cwd=str(WORKSPACE),
    )
    return result


# ─── Page Config ─────────────────────────────────────────────────────────────

st.set_page_config(page_title="HHS Mod Automation", page_icon="⚙️", layout="wide")
ensure_dirs()

st.title("HHS Mod Automation Pipeline")
st.caption("Upload input files, run the 8-script automation pipeline, and download results.")

tab_setup, tab_run, tab_results = st.tabs(["📁 Setup", "▶ Run Pipeline", "📥 Results"])


# ═════════════════════════════════════════════════════════════════════════════
# TAB 1: SETUP — file status & uploads
# ═════════════════════════════════════════════════════════════════════════════

with tab_setup:
    st.header("Input Files")
    st.markdown("Upload the required input files. They will be saved to `Data/input/` with the expected names.")

    cols = st.columns(2)

    for idx, (expected_name, label, exts, used_by) in enumerate(EXPECTED_INPUT_FILES):
        col = cols[idx % 2]
        with col:
            exists = (INPUT_DIR / expected_name).exists()
            icon = "✅" if exists else "⬜"
            with st.expander(f"{icon}  {label}  (`{expected_name}`) — {used_by}", expanded=not exists):
                if exists:
                    mod_time = datetime.fromtimestamp((INPUT_DIR / expected_name).stat().st_mtime)
                    st.success(f"Present — last modified {mod_time:%Y-%m-%d %H:%M}")
                else:
                    st.warning("File not found in `Data/input/`.")

                uploaded = st.file_uploader(
                    f"Upload {label}",
                    type=exts,
                    key=f"inp_{expected_name}",
                )
                if uploaded is not None:
                    if st.session_state.get(f"saved_inp_{expected_name}") != uploaded.file_id:
                        save_uploaded(uploaded, INPUT_DIR, rename_to=expected_name)
                        st.session_state[f"saved_inp_{expected_name}"] = uploaded.file_id
                        st.rerun()

    st.markdown("---")

    eis_col1, eis_col2 = st.columns([1, 1])
    with eis_col1:
        eis_exists = (INPUT_DIR / EIS_BILLING_HARDCODED).exists()
        eis_icon = "✅" if eis_exists else "⬜"
        with st.expander(f"{eis_icon}  EIS Billing Detail File — Script 06", expanded=not eis_exists):
            st.caption(f"Script 06 expects this exact filename: `{EIS_BILLING_HARDCODED}`")
            if eis_exists:
                mod_time = datetime.fromtimestamp((INPUT_DIR / EIS_BILLING_HARDCODED).stat().st_mtime)
                st.success(f"Present — last modified {mod_time:%Y-%m-%d %H:%M}")
            else:
                st.warning("File not found.")

            eis_up = st.file_uploader("Upload EIS Billing File", type=["xlsx"], key="inp_eis")
            if eis_up is not None:
                if st.session_state.get("saved_eis") != eis_up.file_id:
                    save_uploaded(eis_up, INPUT_DIR)
                    st.session_state["saved_eis"] = eis_up.file_id
                    st.rerun()

    st.markdown("---")

    # PR Files
    st.header("PR Files")
    st.markdown("Upload PR `.xlsx` files. They will be saved to `Data/pr_files/`.")

    existing_pr = list_files(PR_DIR, "*.xlsx")
    if existing_pr:
        st.info(f"**{len(existing_pr)}** PR file(s) currently in `Data/pr_files/`")
        with st.expander("View existing PR files"):
            for f in existing_pr:
                st.text(f.name)
    else:
        st.warning("No PR files found in `Data/pr_files/`.")

    pr_up = st.file_uploader(
        "Upload PR Files",
        type=["xlsx"],
        accept_multiple_files=True,
        key="inp_pr",
    )
    if pr_up:
        new_pr_names = [f.name for f in pr_up]
        if st.session_state.get("saved_pr") != new_pr_names:
            for f in pr_up:
                save_uploaded(f, PR_DIR)
            st.session_state["saved_pr"] = new_pr_names
            st.rerun()

    st.markdown("---")

    # Coversheets (optional upload — Script 01 generates them)
    st.header("Coversheets (Optional)")
    st.markdown(
        "Script 01 generates coversheets automatically. "
        "Upload pre-made coversheets here only if you want to **skip Script 01**."
    )

    existing_cs = list_files(COVERSHEETS_DIR, "*.xlsx")
    if existing_cs:
        st.info(f"**{len(existing_cs)}** coversheet(s) in `Data/output/coversheets/`")
        with st.expander("View existing coversheets"):
            for f in existing_cs:
                st.text(f.name)

    cs_up = st.file_uploader(
        "Upload Coversheet Files",
        type=["xlsx"],
        accept_multiple_files=True,
        key="inp_cs",
    )
    if cs_up:
        new_cs_names = [f.name for f in cs_up]
        if st.session_state.get("saved_cs") != new_cs_names:
            for f in cs_up:
                save_uploaded(f, COVERSHEETS_DIR)
            st.session_state["saved_cs"] = new_cs_names
            st.rerun()


# ═════════════════════════════════════════════════════════════════════════════
# TAB 2: RUN PIPELINE
# ═════════════════════════════════════════════════════════════════════════════

with tab_run:
    st.header("Pipeline Execution")
    st.markdown("Select which scripts to run, then click **Run**. Scripts execute sequentially in order.")

    st.subheader("Select Scripts")
    select_col1, select_col2 = st.columns([4, 1])

    with select_col2:
        if st.button("Select All"):
            for s in PIPELINE:
                st.session_state[f"run_{s['num']}"] = True
            st.rerun()
        if st.button("Deselect All"):
            for s in PIPELINE:
                st.session_state[f"run_{s['num']}"] = False
            st.rerun()

    # Seed session state once so the checkboxes can be driven entirely via
    # st.session_state (Select All / Deselect All buttons write here too).
    for s in PIPELINE:
        st.session_state.setdefault(f"run_{s['num']}", True)

    with select_col1:
        selected = []
        for s in PIPELINE:
            checked = st.checkbox(
                f"**{s['num']:02d}** — {s['name']}:  _{s['desc']}_",
                key=f"run_{s['num']}",
            )
            if checked:
                selected.append(s)

    st.markdown("---")

    if st.button("▶  Run Selected Scripts", type="primary", use_container_width=True):
        if not selected:
            st.warning("No scripts selected. Check at least one above.")
        else:
            all_success = True
            full_log = ""

            with st.status("Running pipeline...", expanded=True) as status:
                for i, s in enumerate(selected):
                    st.write(f"**Script {s['num']:02d}** — {s['name']}...")
                    full_log += f"\n{'='*60}\n"
                    full_log += f" SCRIPT {s['num']:02d}: {s['name'].upper()}\n"
                    full_log += f"{'='*60}\n"

                    start = time.time()
                    result = run_script(s["file"])
                    elapsed = time.time() - start

                    if result.stdout:
                        full_log += result.stdout
                    if result.stderr:
                        full_log += f"\n--- stderr ---\n{result.stderr}"

                    if result.returncode == 0:
                        full_log += f"\n>> Completed in {elapsed:.1f}s\n"
                        st.write(f"  Completed in {elapsed:.1f}s")
                    else:
                        full_log += f"\n>> FAILED (exit code {result.returncode}) after {elapsed:.1f}s\n"
                        st.write(f"  **FAILED** (exit code {result.returncode})")
                        all_success = False

                if all_success:
                    status.update(label="Pipeline completed successfully!", state="complete")
                else:
                    status.update(label="Pipeline finished with errors.", state="error")

            st.session_state["pipeline_log"] = full_log

    if "pipeline_log" in st.session_state:
        with st.expander("Full Execution Log", expanded=False):
            st.code(st.session_state["pipeline_log"], language="text")


# ═════════════════════════════════════════════════════════════════════════════
# TAB 3: RESULTS
# ═════════════════════════════════════════════════════════════════════════════

with tab_results:
    st.header("Output Files")
    st.markdown("Download the files generated by the pipeline from `Data/output/`.")

    output_entries = [
        ("Coversheets", COVERSHEETS_DIR, "*.xlsx"),
        ("Overview File", OUTPUT_DIR / "overview_file.xlsx", None),
        ("F&R Output", OUTPUT_DIR / "f_r_output.xlsx", None),
        ("Build File", OUTPUT_DIR / "build_file.xlsx", None),
        ("J1 Current File", OUTPUT_DIR / "j1_current_file.xlsx", None),
        ("J17 Updated File", OUTPUT_DIR / "j17_updated_file.xlsx", None),
        ("MFR Walkthrough Output", OUTPUT_DIR, "* MFR Walkthrough Output.docx"),
        ("SF30 Output", OUTPUT_DIR, "* SF30 Output.docx"),
    ]

    has_any = False

    for label, path, pattern in output_entries:
        if pattern:
            files = list_files(path, pattern) if path.exists() else []
            if not files:
                continue
            has_any = True
            with st.expander(f"📁 {label}  ({len(files)} files)"):
                for f in files:
                    c1, c2 = st.columns([4, 1])
                    c1.write(f.name)
                    c2.download_button(
                        "Download",
                        data=f.read_bytes(),
                        file_name=f.name,
                        key=f"dl_{f.name}",
                    )
        else:
            fpath = Path(path)
            if not fpath.exists():
                continue
            has_any = True
            c1, c2 = st.columns([4, 1])
            mod_time = datetime.fromtimestamp(fpath.stat().st_mtime)
            c1.write(f"📄 **{label}** — `{fpath.name}` (modified {mod_time:%Y-%m-%d %H:%M})")
            c2.download_button(
                "Download",
                data=fpath.read_bytes(),
                file_name=fpath.name,
                key=f"dl_{fpath.name}",
            )

    if not has_any:
        st.info("No output files yet. Go to the **Run Pipeline** tab to execute scripts.")

    st.markdown("---")

    if has_any:
        st.subheader("Cleanup")
        if st.button("🗑️  Clear All Output Files"):
            if OUTPUT_DIR.exists():
                shutil.rmtree(OUTPUT_DIR)
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            COVERSHEETS_DIR.mkdir(parents=True, exist_ok=True)
            st.success("All output files cleared.")
            st.rerun()
