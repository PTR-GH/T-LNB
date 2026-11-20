import datetime
import os
from google import genai
from google.genai import types
import docx
import streamlit as st

# Configure page
st.set_page_config(
    page_title="T-LNB - Electronic Lab Notebook", page_icon="🧪", layout="wide"
)

# Initialize Gemini Client using Streamlit secrets
API_KEY = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")

if not API_KEY:
    st.error(
        "Gemini API Key not found! Please add it to your Streamlit secrets or environment variables."
    )
    st.stop()

client = genai.Client(api_key=API_KEY)


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


def create_or_append_docx(batch_num, title, content_dict):
    filename = f"{batch_num}.docx"
    if os.path.exists(filename):
        doc = docx.Document(filename)
        doc.add_paragraph(
            f"--- Multi-Day Update: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')} ---"
        )
    else:
        doc = docx.Document()
        doc.add_heading(f"Batch Record: {batch_num}", level=1)
        doc.add_paragraph(f"Experiment Type / Title: {title}")
        doc.add_paragraph(
            f"Created Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"
        )

    for k, v in content_dict.items():
        doc.add_paragraph(f"{k}: {v}")

    doc.save(filename)
    return filename


# UI Layout
st.title("🧪 T-LNB: Tulasi's Electronic Lab Notebook")
st.write(
    "AI-powered daily chemistry lab scribe with multi-day tracking and automated `.docx` generation."
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

    if st.button("Generate New Batch Number"):
        st.session_state["current_batch"] = get_next_batch_number(selected_code)

    current_batch = st.session_state.get(
        "current_batch", get_next_batch_number(selected_code)
    )
    st.info(f"Active Batch: **{current_batch}**")

# Main Chat / Input Area
st.subheader("AI Lab Scribe Input")
user_input = st.text_area(
    "Enter rough bench notes (e.g., 'extracted 5g powder with 50ml MeOH, reflux 2h'):",
    height=100,
)

if st.button("Process & Record Experiment"):
    if user_input:
        with st.spinner("AI Scribe formatting your notes..."):
           prompt = f"""
You are a Senior Process Research & Development Chemist with 20 years of GMP/GLP laboratory notebook documentation experience at leading Indian API pharmaceutical companies (such as Divi's Laboratories, Hetero, and Laurus Labs).

Convert the raw bench notes below into a formal, audit-ready Standard Operating Procedure (SOP) / Batch Manufacturing Record (BMR) style experimental entry.

STRICT FORMATTING RULES:
1. Do NOT include any conversational filler, greetings, or introductory/concluding remarks (e.g., never write "Here is the summary..."). Start directly with the first section heading.
2. Do NOT use any Markdown formatting characters: NO asterisks (*), NO hashtags (#), NO dollar signs ($), and NO bullet symbols. Use plain, clean uppercase headings and numbered steps only.
3. Write in passive, formal pharmaceutical past tense (e.g., "Charged...", "Triturated...", "Purged with nitrogen and heat-sealed...").
4. Write chemical formulas in plain text (e.g., write N2, MeOH, HCl — never $N_2$).
5. Keep it concise, technical, and authentic to an actual Indian pharma R&D lab notebook.

Use this exact structure:

OBJECTIVE:
[One crisp technical sentence stating the purpose of the operation]

MATERIALS AND EQUIPMENT:
[Plain list of reagents, quantities if mentioned, and apparatus/packaging used]

EXPERIMENTAL PROCEDURE:
1. [Step-by-step formal chronological procedure with exact times, temperatures, and conditions]

OBSERVATIONS AND STORAGE:
[Physical state/appearance noted and exact storage/packaging conditions; if quantities or physical appearance were omitted in raw notes, briefly note "Quantity/appearance not recorded" in standard GMP style]

Raw Bench Notes:
{raw_notes}
"""
            response = client.models.generate_content(
                model="gemini-3.8-flash", contents=prompt
            )
            parsed_text = response.text

            content_data = {
                "Workflow Category": exp_category,
                "Raw Notes": user_input,
                "Structured Summary": parsed_text,
            }

            saved_file = create_or_append_docx(
                current_batch, exp_category, content_data
            )

            st.success(f"Successfully updated record in {saved_file}!")
            st.markdown("### Formatted Record Preview")
            st.write(parsed_text)

            with open(saved_file, "rb") as f:
                st.download_button(
                    label="Download Word (.docx) Record",
                    data=f,
                    file_name=saved_file,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
    else:
                st.warning("Please enter some bench notes before processing.")
