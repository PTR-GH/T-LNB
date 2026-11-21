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
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }
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
    div[data-testid="stMetric"] {
        background-color: rgba(148, 163, 184, 0.08);
        border: 1px solid rgba(148, 163, 184, 0.2);
        padding: 12px 16px;
        border-radius: 10px;
    }
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

        st.caption("Quick-insert lab log templates:")
        q1, q2, q3, q4 = st.columns(4)
        if q1.button("+ Material Charge", use_container_width=True):
            st.session_state["bench_notes_buffer"] += (
                "Charged [X] g of [Material] and [Y] mL of [Solvent] into RBF. "
            )
        if q2.button("+ Reaction / Stirring", use_container_width=True):
            st.session_state["bench_notes_buffer"] += (
                "Stirred at [RPM] rpm at [Temp] °C for [Time] h under N2 atmosphere. "
            )
        if q3.button("+ Workup & Yield", use_container_width=True):
            st.session_state["bench_notes_buffer"] += (
                "Filtered, dried under vacuum at [Temp] °C. Obtained [X] g off-white solid (Yield: [Y]%). "
            )
        if q4.button("+ Analytical / Pack", use_container_width=True):
            st.session_state["bench_notes_buffer"] += (
                "Submitted sample for PXRD/HPLC/DSC. Packed in double PP cover purged with N2. "
            )

        user_input = st.text_area(
            "Bench Notes & Process Observations",
            value=st.session_state["bench_notes_buffer"],
            height=170,
            placeholder="Enter raw observations, weights, temperatures, yields, or analytical updates...",
        )
        st.session_state["bench_notes_buffer"] = user_input

        col_act1, col_act2 = st.columns([1.5, 3])
        with col_act1:
            process_clicked = st.button(
                "✨ Generate Structured Batch Record",
                type="primary",
                use_container_width=True,
            )

    if process_clicked:
        if not st.session_state["current_batch"]:
            st.warning(
                "⚠️ Generate or enter a **Batch Number** in the left panel first."
            )
        elif not st.session_state["current_exp_name"]:
            st.warning(
                "⚠️ Enter an **Experiment Name** in the left panel before processing."
            )
        elif not user_input.strip():
            st.warning("⚠️ Enter your bench notes before processing.")
        else:
            with st.spinner("Formatting GMP/GLP Batch Record..."):
                prompt = f"""
                You are a Senior Process Research & Development Chemist with 20 years of GMP/GLP laboratory notebook documentation experience at leading Indian API pharmaceutical companies (such as Divi's Laboratories, Hetero, and Laurus Labs).

                Convert the raw bench notes below into a structured, audit-ready Batch Manufacturing Record (BMR) / R&D Lab Notebook entry.

                STRICT RULES:
                1. Return ONLY valid JSON matching the exact schema below. Do not include markdown backticks (```json), commentary, or conversational text.
                2. Do NOT use any asterisks (*), hashtags (#), or LaTeX dollar signs ($) anywhere in the values. Write chemical formulas in plain text (e.g., N2, MeOH, HCl, PP).
                3. Write in passive, formal pharmaceutical past tense (e.g., "Charged...", "Triturated...", "Purged with N2 and heat-sealed...").
                4. If quantities, temperatures, or physical appearance were not mentioned in the raw notes, state "Not recorded" or "Ambient (not specified)" professionally.

                JSON SCHEMA:
                {{
                    "objective": "One crisp technical sentence stating the purpose of the experiment.",
                    "materials": [
                        {{
                            "sr_no": "1",
                            "material": "Name of chemical, reagent, gas, or apparatus/packaging",
                            "quantity": "Exact weight/volume or 'Not recorded'",
                            "role_specs": "Role (e.g., Reactant, Inert gas, Primary packaging, Apparatus)"
                        }}
                    ],
                    "procedure": [
                        {{
                            "step": "1",
                            "operation": "Formal passive past-tense description of the unit operation",
                            "parameters": "Time, temperature, speed, atmosphere (e.g., 20 min, Ambient, N2 atmosphere)",
                            "observation": "Visual/physical observation or 'Not recorded'"
                        }}
                    ],
                    "summary_table": [
                        {{"parameter": "Physical Appearance / State", "details": "..."}},
                        {{"parameter": "Net Yield / Output", "details": "..."}},
                        {{"parameter": "Packaging & Storage Condition", "details": "..."}},
                        {{"parameter": "Analytical / Quality Remarks", "details": "..."}}
                    ]
                }}

                Raw Bench Notes:
                {user_input}
                """
                try:
                    response = client.models.generate_content(
                        model="gemini-3.8-flash", contents=prompt
                    )
                    raw_json = response.text.strip()
                    if raw_json.startswith("```"):
                        raw_json = raw_json.strip("`").replace("json", "", 1).strip()

                    record_data = json.loads(raw_json)
                    saved_file = create_or_append_docx(
                        st.session_state["current_batch"],
                        st.session_state["current_exp_name"],
                        exp_category,
                        user_input,
                        record_data,
                    )

                    save_to_archive(
                        {
                            "batch_no": st.session_state["current_batch"],
                            "experiment_name": st.session_state["current_exp_name"],
                            "workflow": exp_category,
                            "timestamp": datetime.datetime.now().strftime(
                                "%d-%b-%Y %H:%M"
                            ),
                            "raw_notes": user_input,
                            "record_data": record_data,
                        }
                    )

                    st.success(
                        f"✅ Batch Record `{saved_file}` generated and archived under **Neeta Life Sciences**."
                    )

                    with st.container(border=True):
                        st.markdown("#### 1. Objective")
                        st.write(record_data.get("objective", ""))

                        st.markdown("#### 2. Materials, Reagents & Equipment")
                        st.dataframe(
                            pd.DataFrame(record_data.get("materials", [])),
                            use_container_width=True,
                            hide_index=True,
                        )

                        st.markdown("#### 3. Experimental Procedure & Process Parameters")
                        st.dataframe(
                            pd.DataFrame(record_data.get("procedure", [])),
                            use_container_width=True,
                            hide_index=True,
                        )

                        st.markdown("#### 4. Final Observations, Yield & Storage")
                        st.dataframe(
                            pd.DataFrame(record_data.get("summary_table", [])),
                            use_container_width=True,
                            hide_index=True,
                        )

                        with open(saved_file, "rb") as f:
                            st.download_button(
                                label="📥 Download Official .docx Batch Record",
                                data=f,
                                file_name=saved_file,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            )
                except Exception as e:
                    st.error(
                        f"Unable to process record right now ({e}). Please wait a few seconds and try again."
                    )

# ------------------------------------------------------------------------------
# TAB 2: SEARCH & READ-ONLY ARCHIVE PREVIEW
# ------------------------------------------------------------------------------
with tab_archive:
    col_s1, col_s2 = st.columns([2.5, 1])
    with col_s1:
        search_query = st.text_input(
            "Search Archive by Batch Number or Experiment Name",
            placeholder="Start typing Batch ID (e.g., 211126-EXT-01) or Experiment Name...",
            key="archive_search_box",
        ).strip().lower()
    with col_s2:
        wf_filter = st.selectbox(
            "Filter by Workflow",
            ["All Workflows"] + list(code_map.keys()),
        )

    filtered_records = [
        rec
        for rec in archive_records
        if (
            not search_query
            or search_query in rec.get("batch_no", "").lower()
            or search_query in rec.get("experiment_name", "").lower()
        )
        and (
            wf_filter == "All Workflows"
            or rec.get("workflow", "") == wf_filter
        )
    ]

    if filtered_records:
        options_map = {
            f"{rec['batch_no']}  |  {rec['experiment_name']}  ({rec.get('timestamp', '')})": rec
            for rec in reversed(filtered_records)
        }
        selected_key = st.selectbox(
            f"Select Record to Inspect ({len(filtered_records)} found):",
            options=list(options_map.keys()),
        )
        selected_rec = options_map[selected_key]
        rec_body = selected_rec.get("record_data", {})

        with st.container(border=True):
            st.info(
                f"🔒 **READ-ONLY BATCH PREVIEW — NEETA LIFE SCIENCES**  \n"
                f"**Batch No:** `{selected_rec['batch_no']}`  •  "
                f"**Experiment:** {selected_rec['experiment_name']}  •  "
                f"**Workflow:** {selected_rec.get('workflow', '-')}  •  "
                f"**Timestamp:** {selected_rec.get('timestamp', '-')}"
            )

            st.markdown("#### 1. Objective")
            st.write(rec_body.get("objective", "Not specified"))

            st.markdown("#### 2. Materials, Reagents & Equipment")
            st.dataframe(
                pd.DataFrame(rec_body.get("materials", [])),
                use_container_width=True,
                hide_index=True,
            )

            st.markdown("#### 3. Experimental Procedure & Process Parameters")
            st.dataframe(
                pd.DataFrame(rec_body.get("procedure", [])),
                use_container_width=True,
                hide_index=True,
            )

            st.markdown("#### 4. Final Observations, Yield & Storage")
            st.dataframe(
                pd.DataFrame(rec_body.get("summary_table", [])),
                use_container_width=True,
                hide_index=True,
            )

            st.markdown("#### 5. Original Bench Notes (Audit Trail)")
            st.code(selected_rec.get("raw_notes", ""), language=None)
    elif archive_records:
        st.warning("No archived batch records match your current search or workflow filter.")
    else:
        st.info(
            "No batch records archived yet. Process an experiment in the **Record Bench Entry** tab to populate your searchable archive."
        )
