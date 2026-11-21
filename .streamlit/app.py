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
    page_title="T-LNB - Electronic Lab Notebook",
    page_icon="🧪",
    layout="wide",
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
    # Update existing batch entry if multi-day continuation, else append new
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
        doc.add_heading("LABORATORY BATCH RECORD (R&D / PD)", level=1)
        add_styled_table(
            doc,
            ["Batch No.", "Experiment Name", "Workflow Category", "Date & Time"],
            [[batch_num, exp_name, exp_category, now_str]],
        )

    # 1. Objective
    p_obj = doc.add_paragraph()
    r_obj = p_obj.add_run("1. OBJECTIVE: ")
    r_obj.bold = True
    p_obj.add_run(record_data.get("objective", "Not specified"))

    # 2. Materials & Equipment Table
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

    # 3. Experimental Procedure & Process Parameters Table
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

    # 4. Yield, Observations & Storage Table
    p_obs = doc.add_paragraph()
    p_obs.add_run("4. FINAL OBSERVATIONS, YIELD & STORAGE").bold = True
    summary_items = record_data.get("summary_table", [])
    if summary_items:
        sum_rows = [
            [item.get("parameter", "-"), item.get("details", "-")]
            for item in summary_items
        ]
        add_styled_table(doc, ["Parameter", "Recorded Details"], sum_rows)

    # 5. Raw Bench Notes Reference
    p_raw = doc.add_paragraph()
    p_raw.add_run("5. ORIGINAL BENCH NOTES (AUDIT TRAIL): ").bold = True
    p_raw.add_run(raw_notes)

    doc.save(filename)
    return filename


# UI Layout
st.title("🧪 T-LNB: Tulasi's Electronic Lab Notebook")
st.write(
    "GMP/GLP-style Process R&D Lab Scribe with structured tabular `.docx` batch records."
)

# Sidebar for controls
with st.sidebar:
    st.header("Experiment Control")
    exp_category = st.selectbox(
        "Select Workflow Type",
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

    st.info("Step 1: Generate a batch number and enter the experiment name below.")

    if st.button("Generate Batch Number"):
        st.session_state["current_batch"] = get_next_batch_number(selected_code)

    current_batch = st.text_input(
        "Active Batch Number",
        value=st.session_state["current_batch"],
        placeholder="Click 'Generate Batch Number' or enter existing ID...",
    )
    st.session_state["current_batch"] = current_batch.strip()

    # NEW: Experiment Name Input
    current_exp_name = st.text_input(
        "Experiment Name",
        value=st.session_state["current_exp_name"],
        placeholder="e.g., Curcumin-Nicotinamide Cocrystal Screening",
    )
    st.session_state["current_exp_name"] = current_exp_name.strip()


# Main input area
display_batch = (
    st.session_state["current_batch"]
    if st.session_state["current_batch"]
    else "NO BATCH SELECTED"
)
display_exp_name = (
    st.session_state["current_exp_name"]
    if st.session_state["current_exp_name"]
    else "Untitled Experiment"
)

st.subheader(f"Active Record: {display_batch} — {display_exp_name} | {exp_category}")

user_input = st.text_area(
    "Enter your rough bench notes, observations, weights, or analytical updates:",
    height=180,
    placeholder="e.g., Charged 5.0 g of API and 2.1 g of coformer into mortar and pestle. Ground manually for 20 min and packed in PP cover purged with N2...",
)

if st.button("Process & Record Experiment"):
    if not st.session_state["current_batch"]:
        st.warning(
            "⚠️ Please generate a Batch Number (or type an existing Batch Number) in the left sidebar first!"
        )
    elif not st.session_state["current_exp_name"]:
        st.warning(
            "⚠️ Please enter an Experiment Name in the left sidebar before processing!"
        )
    elif not user_input.strip():
        st.warning("⚠️ Please enter your bench notes before processing.")
    else:
        with st.spinner("Senior Chemist Scribe drafting batch record..."):
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

                # Save searchable record to JSON archive
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

                st.success(f"Batch Record saved to {saved_file}")

                st.markdown("#### 1. Objective")
                st.write(record_data.get("objective", ""))

                st.markdown("#### 2. Materials, Reagents & Equipment")
                st.table(pd.DataFrame(record_data.get("materials", [])))

                st.markdown("#### 3. Experimental Procedure & Process Parameters")
                st.table(pd.DataFrame(record_data.get("procedure", [])))

                st.markdown("#### 4. Final Observations, Yield & Storage")
                st.table(pd.DataFrame(record_data.get("summary_table", [])))

                with open(saved_file, "rb") as f:
                    st.download_button(
                        label="Download Word (.docx) Batch Record",
                        data=f,
                        file_name=saved_file,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
            except Exception as e:
                st.error(
                    f"Unable to process record right now ({e}). Please wait a few seconds and click 'Process & Record Experiment' again."
                )


# ==============================================================================
# NEW FEATURE: SEARCH & READ-ONLY PREVIEW OF PREVIOUS BATCH RECORDS
# ==============================================================================
st.divider()
st.subheader("🔍 Search & Preview Previous Batch Records")

archive_records = load_archive()

search_query = st.text_input(
    "Search by Batch Number or Experiment Name",
    placeholder="Type a Batch No. (e.g., 211126-EXT-01) or Experiment Name...",
    key="archive_search_box",
).strip().lower()

if search_query:
    matches = [
        rec
        for rec in archive_records
        if search_query in rec.get("batch_no", "").lower()
        or search_query in rec.get("experiment_name", "").lower()
    ]

    if matches:
        options_map = {
            f"{rec['batch_no']} — {rec['experiment_name']} ({rec.get('timestamp', '')})": rec
            for rec in matches
        }
        selected_key = st.selectbox(
            "Select matching batch record to preview:",
            options=list(options_map.keys()),
        )
        selected_rec = options_map[selected_key]
        rec_body = selected_rec.get("record_data", {})

        # Read-Only Preview Container (no editable widgets used)
        with st.container(border=True):
            st.info(
                f"🔒 **READ-ONLY ARCHIVE PREVIEW** | **Batch No:** `{selected_rec['batch_no']}` | "
                f"**Experiment:** {selected_rec['experiment_name']} | "
                f"**Workflow:** {selected_rec.get('workflow', '-')} | "
                f"**Recorded:** {selected_rec.get('timestamp', '-')}"
            )

            st.markdown("#### 1. Objective")
            st.write(rec_body.get("objective", "Not specified"))

            st.markdown("#### 2. Materials, Reagents & Equipment")
            st.table(pd.DataFrame(rec_body.get("materials", [])))

            st.markdown("#### 3. Experimental Procedure & Process Parameters")
            st.table(pd.DataFrame(rec_body.get("procedure", [])))

            st.markdown("#### 4. Final Observations, Yield & Storage")
            st.table(pd.DataFrame(rec_body.get("summary_table", [])))

            st.markdown("#### 5. Original Bench Notes (Audit Trail)")
            st.code(selected_rec.get("raw_notes", ""), language=None)
    else:
        st.warning("No matching batch records found for your search.")
elif not archive_records:
    st.caption("No archived batch records yet. Processed experiments will appear here for read-only search and preview.")
