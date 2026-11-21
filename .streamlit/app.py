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


def create_or_append_docx(batch
