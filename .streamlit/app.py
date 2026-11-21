import datetime
import json
import os
from google import genai
import docx
from docx.shared import Pt
import pandas as pd
import streamlit as st

# Configure page
st.set_page_config(
    page_title="NLS | Neeta Life Sciences — Electronic Lab Notebook",
    page_icon="🧬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS for clean, modern pharmaceutical R&D interface & NLS branding
st.markdown(
    """
    <style>
    /* Tighten top padding for a crisp app feel */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }
    /* Top-left Sidebar Company Logo Card */
    .nls-brand-box {
        display: flex;
        align-items: center;
        gap: 12px;
        padding: 12px 14px;
        border-radius: 12px;
        background: linear-gradient(135deg, #0f2942 0%, #1a4b7c 100%);
        border: 1px solid rgba(56, 189, 248, 0.3);
        margin-bottom: 1rem;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    }
    .nls-logo-badge {
        background: #38bdf8;
        color: #0f172a;
        font-weight: 800;
        font-size: 1.05rem;
        letter-spacing: 1px;
        padding: 8px 10px;
        border-radius: 8px;
        line-height: 1;
    }
    .nls-brand-text {
        display: flex;
        flex-direction: column;
    }
    .nls-company-name {
        color: #ffffff;
        font-weight: 700;
        font-size: 0.95rem;
        line-height: 1.2;
    }
    .nls-subtext {
        color: #bae6fd;
        font-size: 0.72rem;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        margin-top: 2px;
    }
    /* Subtle card styling for status metrics */
    div[data-testid="stMetric"] {
        background-color: rgba(148, 163, 184, 0.08);
        border: 1px solid rgba(148, 163, 184, 0.2);
        padding: 12px 16px;
        border-radius: 10px;
    }
    /* Primary action button polish */
    div.stButton > button[kind="primary"] {
        border-radius: 8px;
        font-weight: 600;
        padding: 0.55rem 1.2rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Initialize Gemini Client using Streamlit secrets
API_KEY = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
if not API_KEY:
    st.error(
        "Gemini API Key not found! Please add it to your Streamlit secrets or environment variables."
    )
    st.stop()

client = genai.Client(api_key=API_KEY)

ARCHIVE_FILE = "batch_records_archive.json"


def load_archive():
    if os.path.exists(ARCHIVE_FILE):
        try:
            with open(ARCHIVE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []


def save_to_archive(entry):
    records = load_archive()
    for idx, rec in enumerate(records):
        if rec["batch_no"] == entry["batch_no"]:
            records[idx] = entry
            break
    else:
        records.append(entry)

    with open(ARCHIVE_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2)


def get_next_batch_number(exp_type_code):
    today_str = datetime.datetime.now().strftime("%d%m%y")
    if "batches" not in st.session_state:
        st.session_state["batches"] = {}
    key = f"{today_str}-{exp_type_code}"
    if key not in st.session_state["batches"]:
        st.session_state["batches"][key] = 1
    else:
        st.session_state["batches"][key] += 1
    seq_num = st.session_state["batches"][key]
    return f"{today_str}-{exp_type_code}-{seq_num:02d}"


def add_styled_table(doc, headers, rows):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = str(header_text)
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(10)

    for row in rows:
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val) if val else "-"
            for paragraph in row_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
    doc.add_paragraph("")


def create_or_append_docx(batch_num, exp_name, exp_category, raw_notes, record_data):
    filename = f"{batch_num}.docx"
    now_str = datetime.datetime.now().strftime("%d-%b-%Y %H:%M")

    if os.path.exists(filename):
        doc = docx.Document(filename)
        doc.add_page_break()
        doc.add_heading(f"Continuation / Multi-Day Entry — {now_str}", level=2)
    else:
        doc = docx.Document()
        doc.add_heading("NEETA LIFE SCIENCES — LABORATORY BATCH RECORD (R&D)", level=1)
        add_styled_table(
            doc,
            ["Batch No.", "Experiment Name", "Workflow Category", "Date & Time"],
            [[batch_num, exp_name, exp_category, now_str]],
        )

    p_obj = doc.add_paragraph()
    r_obj = p_obj.add_run("1. OBJECTIVE: ")
    r_obj.bold = True
    p_obj.add_run(record_data.get("objective", "Not specified"))

    p_mat = doc.add_paragraph()
    p_mat.add_run("2. MATERIALS, REAGENTS & EQUIPMENT").bold = True
    materials = record_data.get("materials", [])
    if materials:
        mat_rows = [
            [
                m.get("sr_no", idx + 1),
                m.get("material", "-"),
                m.get("quantity", "Not recorded"),
                m.get("role_specs", "-"),
            ]
            for idx, m in enumerate(materials)
        ]
        add_styled_table(
            doc,
            ["Sr. No.", "Material / Equipment", "Quantity / Qty", "Role / Remarks"],
            mat_rows,
        )

    p_proc = doc.add_paragraph()
    p_proc.add_run("3. EXPERIMENTAL PROCEDURE & PROCESS PARAMETERS").bold = True
    steps = record_data.get("procedure", [])
    if steps:
        proc_rows = [
            [
                s.get("step", idx + 1),
                s.get("operation", "-"),
                s.get("parameters", "Ambient / Standard"),
                s.get("observation", "-"),
            ]
            for idx, s in enumerate(steps)
        ]
        add_styled_table(
            doc,
            [
                "Step",
                "Unit Operation / Procedure",
                "Parameters (Temp / Time / RPM / Atm)",
                "In-Process Observation",
            ],
            proc_rows,
        )

    p_obs = doc.add_paragraph()
    p_obs.add_run("4. FINAL OBSERVATIONS, YIELD & STORAGE").bold = True
    summary_items = record_data.get("summary_table", [])
    if summary_items:
        sum_rows = [
            [item.get("parameter", "-"), item.get("details", "-")]
            for item in summary_items
        ]
        add_styled_table(doc, ["Parameter", "Recorded Details"], sum_rows)

    p_raw = doc.add_paragraph()
    p_raw.add_run("5. ORIGINAL BENCH NOTES (AUDIT TRAIL): ").bold = True
    p_raw.add_run(raw_notes)

    doc.save(filename)
    return filename


# ==============================================================================
# SIDEBAR: TOP-LEFT COMPANY BRANDING & BATCH SETUP
# ==============================================================================
with st.sidebar:
    st.markdown(
        """
        <div class="nls-brand-box">
            <div class="nls-logo-badge">NLS</div>
            <div class="nls-brand-text">
                <span class="nls-company-name">Neeta Life Sciences</span>
                <span class="nls-subtext">Process R&D • Digital ELN</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.caption("BATCH CONFIGURATION")
    exp_category = st.selectbox(
        "Workflow Category",
        [
            "Extractions (EXT)",
            "Multi-stage Chemical Synthesis (SYN)",
            "Neat Grinding Cocrystal (NGC)",
            "Solvent/Liquid-Assisted Grinding (LAG)",
            "Quick Checks (pH, Solubility, Melting Point)",
            "Analytical Sample Tracking (PXRD, DSC, NMR, HPLC, FTIR)",
        ],
    )

    code_map = {
        "Extractions (EXT)": "EXT",
        "Multi-stage Chemical Synthesis (SYN)": "SYN",
        "Neat Grinding Cocrystal (NGC)": "NGC",
        "Solvent/Liquid-Assisted Grinding (LAG)": "LAG",
        "Quick Checks (pH, Solubility, Melting Point)": "CHK",
        "Analytical Sample Tracking (PXRD, DSC, NMR, HPLC, FTIR)": "ANL",
    }
    selected_code = code_map[exp_category]

    if "current_batch" not in st.session_state:
        st.session_state["current_batch"] = ""
    if "current_exp_name" not in st.session_state:
        st.session_state["current_exp_name"] = ""
    if "bench_notes_buffer" not in st.session_state:
        st.session_state["bench_notes_buffer"] = ""

    col_b1, col_b2 = st.columns([1.3, 1])
    with col_b1:
        if st.button("⚡ Generate ID", use_container_width=True):
            st.session_state["current_batch"] = get_next_batch_number(selected_code)
    with col_b2:
        if st.button("🔄 Clear", use_container_width=True):
            st.session_state["current_batch"] = ""
            st.session_state["current_exp_name"] = ""
            st.session_state["bench_notes_buffer"] = ""
            st.rerun()

    current_batch = st.text_input(
        "Batch Number",
        value=st.session_state["current_batch"],
        placeholder="e.g., 211126-EXT-01",
    )
    st.session_state["current_batch"] = current_batch.strip()

    current_exp_name = st.text_input(
        "Experiment Name",
        value=st.session_state["current_exp_name"],
        placeholder="e.g., Curcumin-Nicotinamide Cocrystal",
    )
    st.session_state["current_exp_name"] = current_exp_name.strip()

    st.divider()
    st.caption(
        "💡 **Tip:** Reuse an existing Batch Number to append multi-day continuation entries to the same `.docx` record."
    )


# ==============================================================================
# MAIN WORKSPACE: CLEAN HEADER + INTERACTIVE TABS
# ==============================================================================
archive_records = load_archive()

display_batch = (
    st.session_state["current_batch"]
    if st.session_state["current_batch"]
    else "Unassigned"
)
display_exp_name = (
    st.session_state["current_exp_name"]
    if st.session_state["current_exp_name"]
    else "Untitled Experiment"
)

# Top Status Strip
m1, m2, m3 = st.columns([1.2, 1.8, 1])
m1.metric("Active Batch ID", display_batch)
m2.metric("Experiment Title", display_exp_name)
m3.metric("Archived Records", f"{len(archive_records)} Batches")

st.write("")

tab_record, tab_archive = st.tabs(
    ["📝 Record Bench Entry", "🔍 Search & Read-Only Archive"]
)

# ------------------------------------------------------------------------------
# TAB 1: ACTIVE BENCH RECORDING
# ------------------------------------------------------------------------------
with tab_record:
    with st.container(border=True):
        st.markdown(f"**Workflow:** `{exp_category}`")

        # Interactive Quick-Template Snippets
        st.caption("Quick-insert lab log templates:")
        q1, q2, q3, q4 = st.columns(4)
        if q1.button("+ Material Charge", use_container_width=True
