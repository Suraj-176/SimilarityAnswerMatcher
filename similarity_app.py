import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import subprocess
import threading
import shutil
import re
from pathlib import Path
import time
from openpyxl.utils import get_column_letter
import requests

st.set_page_config(page_title="Similarity Answer Matcher", layout="wide")

# Optional auth and LocalTunnel auto-start removed per user request.
# If you need these features later, the helper functions are still present below and can be re-enabled.

# Authentication helpers removed: this app no longer performs simple in-app auth.

# ----------------- LocalTunnel status UI (sidebar) -----------------
def _read_latest_url_file():
    try:
        p = Path("latest-public-url.txt")
        if p.exists():
            return p.read_text(encoding="utf-8").strip()
    except Exception:
        return None
    return None

# Sidebar LocalTunnel UI removed per user request


def _stop_localtunnel():
    """Attempt to stop a running lt process whose PID is stored in lt_pid.txt.

    Returns (True, pid) on success or (False, reason) on failure.
    """
    pid_file = Path("lt_pid.txt")
    url_file = Path("latest-public-url.txt")
    log_file = Path("lt_output.log")
    if not pid_file.exists():
        return (False, "no_pid_file")
    try:
        pid_text = pid_file.read_text(encoding="utf-8").strip()
        pid = int(pid_text)
    except Exception as e:
        return (False, f"invalid_pid: {e}")

    # Verify process command line or executable to reduce accidental kills
    try:
        proc_ok = False
        if os.name == 'nt':
            # Use wmic to query commandline for the PID
            try:
                res = subprocess.check_output(["wmic", "process", "where", f"ProcessId={pid}", "get", "CommandLine"], stderr=subprocess.DEVNULL, text=True)
                if 'lt' in (res or '').lower() or 'localtunnel' in (res or '').lower():
                    proc_ok = True
            except Exception:
                # WMIC may be missing on some systems; fall back to PowerShell Get-CimInstance or Get-Process
                try:
                    # Try PowerShell Get-CimInstance (returns CommandLine)
                    ps_cmd_str = f"Get-CimInstance Win32_Process -Filter \"ProcessId={pid}\" | Select-Object -ExpandProperty CommandLine"
                    ps_cmd = ["powershell", "-NoProfile", "-Command", ps_cmd_str]
                    res2 = subprocess.check_output(ps_cmd, stderr=subprocess.DEVNULL, text=True)
                    if 'lt' in (res2 or '').lower() or 'localtunnel' in (res2 or '').lower():
                        proc_ok = True
                except Exception:
                    try:
                        # Last-resort: use Get-Process and inspect Path or ProcessName
                        ps_cmd2_str = f"(Get-Process -Id {pid} -ErrorAction SilentlyContinue) | Select-Object -ExpandProperty Path"
                        ps_cmd2 = ["powershell", "-NoProfile", "-Command", ps_cmd2_str]
                        res3 = subprocess.check_output(ps_cmd2, stderr=subprocess.DEVNULL, text=True)
                        if 'lt' in (res3 or '').lower() or 'localtunnel' in (res3 or '').lower():
                            proc_ok = True
                    except Exception:
                        proc_ok = False
        else:
            # Try reading /proc/<pid>/cmdline
            try:
                with open(f"/proc/{pid}/cmdline", 'r', encoding='utf-8') as fh:
                    cmd = fh.read().replace('\x00', ' ')
                    if 'lt' in cmd.lower() or 'localtunnel' in cmd.lower():
                        proc_ok = True
            except Exception:
                proc_ok = False

        if not proc_ok:
            return (False, "process_verification_failed")

        # On Windows, use taskkill to ensure child processes are terminated
        if os.name == 'nt':
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            import signal
            os.kill(pid, signal.SIGTERM)
    except Exception as e:
        # proceed to cleanup files even if kill failed
        err = str(e)
        try:
            if pid_file.exists():
                pid_file.unlink()
        except Exception:
            pass
        try:
            if url_file.exists():
                url_file.unlink()
        except Exception:
            pass
        try:
            if log_file.exists():
                log_file.unlink()
        except Exception:
            pass
        # reset session state
        st.session_state["_lt_started"] = False
        st.session_state["_lt_start_info"] = None
        return (False, err)

    # cleanup files
    try:
        if pid_file.exists():
            pid_file.unlink()
    except Exception:
        pass
    try:
        if url_file.exists():
            url_file.unlink()
    except Exception:
        pass
    try:
        if log_file.exists():
            log_file.unlink()
    except Exception:
        pass

    # reset session state
    st.session_state["_lt_started"] = False
    st.session_state["_lt_start_info"] = None
    return (True, pid)


# Custom CSS for full-width, modern look
st.markdown("""
    <style>
    .main, .stApp {
        background-color: #f7fafd;
    }
    .block-container {
        padding-top: 2rem;
        max-width: 100vw !important;
        width: 100vw !important;
        margin: 0 !important;
        padding-left: 2vw !important;
        padding-right: 2vw !important;
    }
    .css-1d391kg { /* Table header */
        background: #e3f2fd !important;
    }
    .stButton>button {
        background-color: #1976d2;
        color: white;
        border-radius: 6px;
    padding: 0.6em 1.8em;
    font-weight: 600;
    min-width: 360px;
    font-size: 1rem;
    }
    .stButton>button:hover {
        background-color: #1565c0;
        color: #fff;
    }
    .stSelectbox>div>div {
        background: #e3f2fd;
    }
    .model-select-row {
        display: flex;
        align-items: center;
        gap: 1.5em;
        margin-bottom: 1.5em;
        margin-top: 1em;
    }
    .model-select-row img {
        width: 56px;
        height: 56px;
        border-radius: 50%;
        background: #fff;
        border: 2px solid #e3f2fd;
        margin-right: 1em;
    }
    .model-info {
        color: #1976d2;
        font-size: 1.1em;
        font-weight: 600;
    }
    .model-desc {
        color: #555;
        font-size: 0.98em;
        margin-bottom: 0.5em;
    }
    /* Chat bubbles full width */
    .chat-row {
        max-width: 100vw;
        width: 100vw;
        padding-left: 0;
        padding-right: 0;
    }
    .chat-bubble {
        max-width: 60vw;
        min-width: 20vw;
        width: fit-content;
        padding: 1em 1.2em;
        margin: 0.5em 0;
        border-radius: 1.2em;
        font-size: 1.1em;
        line-height: 1.5;
        box-shadow: 0 2px 8px #e3e3e3;
        word-break: break-word;
    }
    .user-bubble {
        background: #e3f2fd;
        margin-left: auto;
        margin-right: 0;
        border-bottom-right-radius: 0.3em;
        display: flex;
        align-items: center;
    }
    .ai-bubble {
        background: #fff;
        margin-right: auto;
        margin-left: 0;
        border-bottom-left-radius: 0.3em;
        display: flex;
        align-items: center;
    }
    .avatar {
        width: 36px;
        height: 36px;
        border-radius: 50%;
        margin: 0 0.7em 0 0;
        background: #fff;
        object-fit: cover;
        border: 2px solid #e3f2fd;
    }
    .ai-avatar {
        margin-right: 0.7em;
        margin-left: 0;
    }
    .user-avatar {
        margin-left: 0.7em;
        margin-right: 0;
    }
    .score-badge {
        display: inline-block;
        background: #1976d2;
        color: #fff;
        border-radius: 0.7em;
        font-size: 0.95em;
        padding: 0.2em 0.7em;
        margin-left: 0.7em;
        margin-right: 0.7em;
    }
    .match-high { background: #43a047 !important; }
    .match-medium { background: #ffa000 !important; }
    .match-low { background: #e53935 !important; }
    /* Custom CSS for column selection UI */
    .stSelectbox label {
        font-weight: 600;
        color: #1976d2;
        font-size: 0.9em;
    }
    .stSelectbox>div>div {
        background: #e3f2fd;
        border-radius: 6px;
        margin-bottom: 0.5em;
        min-height: 36px;
    }
    .stSelectbox select {
        font-size: 0.9em;
        padding: 0.3em 0.5em;
    }
    .stMarkdown b {
        color: #1565c0;
        font-size: 1.08em;
    }
    .stMarkdown small b {
        color: #1565c0;
        font-size: 0.85em;
        font-weight: 700;
    }
    </style>
""", unsafe_allow_html=True)




# --- Step 1: App Title and Instructions ---
st.markdown("""
<div style="background:#e3f2fd;padding:1.2em 1.5em;border-radius:10px;margin-bottom:1.5em;max-width:700px;">
<b>Step 1: Similarity Answer Matcher</b>
</div>
""", unsafe_allow_html=True)

# Note: external gateway warning removed per user request

# --- Step 1.5: Choose Comparison Mode ---
comparison_mode = st.radio(
    "Select comparison mode:",
    ["Compare Two Excel Files", "Compare Any Two Files", "Compare Two Columns in Same Excel File", "Compare One Column with Two Targets (Same Excel)"],
    index=0,
    horizontal=True,
    help="Choose whether to compare files or columns within the same file"
)

# --- Dynamic description below radio ---
desc = {
    "Compare Two Excel Files": "Upload two Excel files with questions in the first column and answers in the second column. The app will compare answers, show similarity, highlight differences.",
    "Compare Any Two Files": "Upload any two files (CSV, JSON, TXT, PDF, DOC/DOCX, or plain text). Select the question/answer columns in each file and the app will compare the corresponding answers, showing similarity and highlighted differences. For non-tabular files, the app will infer columns where possible.",
    "Compare Two Columns in Same Excel File": "Upload a single Excel file with at least three columns: one for questions and two for answers. Select which columns to compare. The app will compare the two answer columns, show similarity, highlight differences."
    ,
    "Compare One Column with Two Targets (Same Excel)": "Upload a single Excel file with at least three columns: one base column and two target columns. Select which column is the base and the two target columns. The app will compare the base column separately with each target and provide two similarity score sets."
}

st.info(desc.get(comparison_mode, desc["Compare Two Excel Files"]))

# Helper: modes that accept two separate uploaded files (keeps intent explicit)
IS_TWO_FILE_MODE = comparison_mode in ("Compare Two Excel Files", "Compare Any Two Files")

# --- Step 2: Matching Method Selection ---
st.markdown("""
<div style="background:#e3f2fd;padding:1.2em 1.5em;border-radius:10px;margin-bottom:1.5em;max-width:700px;">
<b>Step 2: Choose Similarity Matching Method</b>
</div>
""", unsafe_allow_html=True)

col_method, col_input = st.columns([3, 5])
with col_method:
    matching_method = st.radio(
        "Similarity Matching Method",
        ["Azure OpenAI GPT-4o"],
        index=0
    )
    # No local model selection available in this configuration
    selected_model = None
with col_input:
    if matching_method == "Azure OpenAI GPT-4o":
        api_key = st.text_input("Azure OpenAI API Key", type="password", help="Paste your Azure OpenAI API key here.")
        # Compact inline info to save vertical space
        st.markdown("""
        <div style='background:#eaf4ff;padding:0.5em;border-radius:6px;font-size:0.95em;margin-top:0.4em;'>
            Using GPT-4o for similarity matching. This may take a while for large files.
        </div>
        """, unsafe_allow_html=True)
        # Small help icon for advanced GPT request settings.
        # The explicit checkbox is removed per user request; expander will read session state key 'gpt_expand_toggle' if set elsewhere.
        st.markdown(
            "<div style='text-align:right;'><span title='Advanced GPT request settings (system prompt, template, temperature, top_p, max_tokens)' style='font-size:1.0em;'>❔</span></div>",
            unsafe_allow_html=True
        )
        # Default prompt templates
        default_system = (
            "You are a helpful assistant, Provided the similarity score by comparing text 1 and text 2, just provide only similarity score without explaination"
        )
        default_user_tpl = (
            "Compare the following two texts and provide a similarity score as a percentage. Text 1: {answer1} Text 2: {answer2}"
        )
        # Render an expander that matches the requested design
        gpt_expand = st.session_state.get('gpt_expand_toggle', False)
        with st.expander("GPT request settings (optional)", expanded=gpt_expand):
            # Place system and user prompt side-by-side to use right-side horizontal space
            left_col, right_col = st.columns([1,1])
            with left_col:
                gpt_system_prompt = st.text_area(
                    "System prompt",
                    value=st.session_state.get('gpt_system_prompt', default_system),
                    height=160,
                    key="gpt_system_prompt"
                )
            with right_col:
                gpt_user_template = st.text_area(
                    "User prompt template",
                    value=st.session_state.get('gpt_user_template', default_user_tpl),
                    height=160,
                    key="gpt_user_template"
                )
            # Compact one-row controls for generation settings below the two textareas
            tcol, pcol, mcol = st.columns([1,1,1])
            with tcol:
                gpt_temperature = st.number_input(
                    "Temperature",
                    min_value=0.0,
                    max_value=1.0,
                    value=float(st.session_state.get('gpt_temperature', 0.0)),
                    step=0.01,
                    format="%.2f",
                    key="gpt_temperature"
                )
            with pcol:
                gpt_top_p = st.number_input(
                    "Top-p",
                    min_value=0.0,
                    max_value=1.0,
                    value=float(st.session_state.get('gpt_top_p', 1.0)),
                    step=0.01,
                    format="%.2f",
                    key="gpt_top_p"
                )
            with mcol:
                gpt_max_tokens = st.number_input(
                    "Max tokens",
                    min_value=1,
                    max_value=1024,
                    value=int(st.session_state.get('gpt_max_tokens', 20)),
                    step=1,
                    key="gpt_max_tokens"
                )
    # Local model option removed; only Azure OpenAI is supported in this configuration

# --- Step 2.5: Set High Match Threshold ---
st.markdown("""
<div style="background:#e3f2fd;padding:1.2em 1.5em;border-radius:10px;margin-bottom:1.5em;max-width:700px;">
<b>Step 2.5: Set Match Quality Threshold</b>
</div>
""", unsafe_allow_html=True)
threshold = st.slider("Set High Match Threshold (%)", min_value=50, max_value=100, value=85, help="Adjust the percentage above which matches are considered 'High'.")

# Compact summary of key settings to reduce vertical noise and keep users informed
try:
    _api_flag = bool(globals().get('api_key', '') )
except Exception:
    _api_flag = False
summary_html = f"""
<div style='background:#fffef0;padding:0.8em 1em;border-radius:8px;margin-bottom:1.2em;max-width:900px;'>
<b>Mode:</b> {comparison_mode} &nbsp;|&nbsp;
<b>Method:</b> {matching_method} &nbsp;|&nbsp;
<b>Threshold:</b> {threshold}% &nbsp;|&nbsp;
<b>Azure Key:</b> {'Set' if _api_flag else 'Not Set'}
</div>
"""
st.markdown(summary_html, unsafe_allow_html=True)

# --- Step 3: Upload Excel Files (Modified based on mode) ---
st.markdown("""
<div style="background:#e3f2fd;padding:1.2em 1.5em;border-radius:10px;margin-bottom:1.5em;max-width:700px;">
<b>Step 3: Upload Excel File(s)</b>
</div>
""", unsafe_allow_html=True)

# File upload logic - "Compare Any Two Files" supports all file types, others only Excel/CSV
if IS_TWO_FILE_MODE:
    col1, col2 = st.columns(2)
    if comparison_mode == "Compare Any Two Files":
        # Support non-Excel file types for the "Compare Any Two Files" option
        with col1:
            uploaded_file1 = st.file_uploader("Upload First File (CSV / JSON / TXT / PDF / DOC)", key="file1")
        with col2:
            uploaded_file2 = st.file_uploader("Upload Second File (CSV / JSON / TXT / PDF / DOC)", key="file2")
    else:
        # Excel-specific modes only support Excel and CSV
        with col1:
            uploaded_file1 = st.file_uploader("Upload First File (Excel / CSV)", type=['xlsx', 'xls', 'csv'], key="file1")
        with col2:
            uploaded_file2 = st.file_uploader("Upload Second File (Excel / CSV)", type=['xlsx', 'xls', 'csv'], key="file2")
else:
    # Single file modes (Excel-specific) only support Excel and CSV
    uploaded_file1 = st.file_uploader("Upload File (Excel / CSV)", type=['xlsx', 'xls', 'csv'], key="single_file")
    uploaded_file2 = uploaded_file1  # Use same file for both comparisons

# --- Model Loading (with error handling) ---
@st.cache_resource
def load_main_model(selected_model):
    # Import inside the function so the app can still run parts of the UI
    # when sentence-transformers is not installed. If missing, raise a
    # clear error for the user to install the dependency.
    try:
        from sentence_transformers import SentenceTransformer
    except Exception as e:
        raise ImportError("The 'sentence_transformers' package is required for Local Model mode. Please install it (pip install -r requirements.txt or pip install sentence-transformers) and restart the app.") from e
    return SentenceTransformer(selected_model)


@st.cache_resource
def load_cross_encoder_model(selected_model):
    try:
        from sentence_transformers.cross_encoder import CrossEncoder
    except Exception:
        st.warning("Cross-encoder support is not available because the required package could not be imported. Falling back to embedding-only similarity.")
        return None
    try:
        return CrossEncoder(selected_model)
    except Exception as e:
        st.warning(f"Could not load cross-encoder model: {e}")
        return None

CROSS_ENCODER_MODELS = [
    "cross-encoder/stsb-roberta-base",
    "cross-encoder/ms-marco-MiniLM-L6-v2"
]

def get_gpt4o_similarity(answer1, answer2, api_key, system_prompt=None, user_template=None, temperature=0.0, top_p=1.0, max_tokens=20):
    url = "https://f2fdevopenai.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2024-05-01-preview"
    headers = {
        "Content-Type": "application/json",
        "api-key": api_key
    }

    # Build user prompt from template if available
    if user_template:
        try:
            prompt = user_template.format(answer1=answer1, answer2=answer2)
        except Exception:
            prompt = f"Compare the following two texts. Text 1: {answer1} Text 2: {answer2}"
    else:
        prompt = f"Compare the following two texts and provide a similarity score as a percentage. Text 1: {answer1} Text 2: {answer2}"

    sys_content = system_prompt or "You are a helpful assistant. Provide the similarity score by comparing text 1 and text 2. Respond with only the similarity score as a plain number (no explanation)."

    data = {
        "messages": [
            {"role": "system", "content": sys_content},
            {"role": "user", "content": prompt}
        ],
        "temperature": float(temperature),
        "top_p": float(top_p),
        "max_tokens": int(max_tokens)
    }
    try:
        response = requests.post(url, headers=headers, json=data, timeout=30)
        result = response.json()
        content = result['choices'][0]['message']['content']
        # Expecting just a number, e.g. "87"
        score = None
        try:
            # Allow percent sign and decimals; extract number-like substring
            import re as _re
            m = _re.search(r"\d+\.?\d*", content)
            if m:
                score = float(m.group(0))
            else:
                score = 0
        except Exception:
            score = 0
        return score, ""  # No explanation
    except Exception as e:
        return None, f"API error: {e}"

results_df = None

# If files are uploaded, read them lightly and show column selection UI immediately (no heavy models yet)
col1_q_val = col1_a_val = col2_q_val = col2_a_val = None
df1 = df2 = None
if uploaded_file1 and (uploaded_file2 or comparison_mode == "Compare Two Columns in Same Excel File"):
    # read file(s) for column selection only
    try:
        def read_uploaded_file(uploaded, sheet_name=None, allow_excel=True):
            """Read an uploaded file-like object and return a pandas DataFrame.

            Supports: CSV, JSON (list/dict), plain text (.txt, .md), and PDF files.
            If allow_excel=True, also supports Excel (.xls, .xlsx).
            For Excel files, can optionally specify a sheet name.
            For plain text, tries to split each line into columns using tab or pipe delimiters.
            """
            if uploaded is None:
                return None
            # Ensure start of file
            try:
                uploaded.seek(0)
            except Exception:
                pass

            filename = (getattr(uploaded, 'name', '') or '').lower()

            # Helper to read raw text with multiple encoding attempts
            def _read_text(u):
                try:
                    b = u.read()
                    if isinstance(b, bytes):
                        # Try multiple encodings in order of likelihood
                        encodings_to_try = ['utf-8', 'utf-8-sig', 'windows-1252', 'iso-8859-1', 'cp1252', 'latin1']
                        for encoding in encodings_to_try:
                            try:
                                return b.decode(encoding)
                            except UnicodeDecodeError:
                                continue
                        # If all fail, use utf-8 with error replacement as last resort
                        return b.decode('utf-8', errors='replace')
                    return str(b)
                finally:
                    try:
                        u.seek(0)
                    except Exception:
                        pass

            # CSV files with multiple encoding attempts
            if filename.endswith('.csv'):
                # Try multiple encodings for CSV files
                encodings_to_try = ['utf-8', 'utf-8-sig', 'windows-1252', 'iso-8859-1', 'cp1252', 'latin1']
                
                for encoding in encodings_to_try:
                    try:
                        uploaded.seek(0)
                        return pd.read_csv(uploaded, encoding=encoding)
                    except (UnicodeDecodeError, UnicodeError):
                        continue
                    except Exception:
                        # Non-encoding related error, try next encoding
                        continue
                
                # Final fallback: try with error handling
                try:
                    uploaded.seek(0)
                    return pd.read_csv(uploaded, encoding='utf-8', errors='replace', engine='python')
                except Exception:
                    try:
                        uploaded.seek(0)
                    except Exception:
                        pass
                    # Last resort: read as raw text and try to parse
                    content = _read_text(uploaded)
                    from io import StringIO
                    return pd.read_csv(StringIO(content), engine='python')

            # Excel formats (.xls, .xlsx) - only if Excel is allowed
            if allow_excel and filename.endswith(('.xls', '.xlsx')):
                try:
                    if sheet_name is not None:
                        return pd.read_excel(uploaded, sheet_name=sheet_name, engine='openpyxl')
                    else:
                        return pd.read_excel(uploaded, engine='openpyxl')
                except Exception as e:
                    # If Excel reading fails, try alternative approaches
                    try:
                        uploaded.seek(0)
                    except Exception:
                        pass
                    # For .xls files, try different engines
                    if filename.endswith('.xls'):
                        try:
                            if sheet_name is not None:
                                return pd.read_excel(uploaded, sheet_name=sheet_name, engine='xlrd')
                            else:
                                return pd.read_excel(uploaded, engine='xlrd')
                        except Exception:
                            pass
                    raise e  # Re-raise original error if alternatives fail

            # JSON files
            if filename.endswith('.json'):
                try:
                    txt = _read_text(uploaded)
                    import json
                    obj = json.loads(txt)
                    # If it's a list of dicts or dict of lists, convert to DataFrame
                    if isinstance(obj, list):
                        return pd.DataFrame(obj)
                    elif isinstance(obj, dict):
                        # try to normalize
                        try:
                            return pd.json_normalize(obj)
                        except Exception:
                            return pd.DataFrame([obj])
                except Exception:
                    raise

            # Plain text / markdown files
            if filename.endswith(('.txt', '.md')) or filename == '':
                txt = _read_text(uploaded)
                lines = [l.strip() for l in txt.splitlines() if l.strip()]
                rows = []
                for line in lines:
                    if '\t' in line:
                        parts = line.split('\t', 1)
                        rows.append(parts)
                    elif '||' in line:
                        parts = line.split('||', 1)
                        rows.append(parts)
                    elif '|' in line:
                        parts = line.split('|', 1)
                        rows.append(parts)
                    else:
                        rows.append([line])
                # Normalize to DataFrame
                max_cols = max(len(r) for r in rows) if rows else 0
                cols = [f'col{i+1}' for i in range(max_cols)]
                norm_rows = [r + ['']*(max_cols - len(r)) for r in rows]
                return pd.DataFrame(norm_rows, columns=cols)

            # PDF support: try to extract text per page and return as rows
            if filename.endswith('.pdf'):
                try:
                    try:
                        # Lazy import so PyPDF2 is only required if PDFs are used
                        import PyPDF2
                    except Exception as ie:
                        raise ImportError("PyPDF2 is required to read PDF files. Please install it (pip install PyPDF2) to enable PDF uploads.") from ie
                    from io import BytesIO
                    # _read_text returns a str or bytes; ensure bytes for PdfReader
                    raw = uploaded.read()
                    if isinstance(raw, str):
                        raw_bytes = raw.encode('utf-8')
                    else:
                        raw_bytes = raw
                    try:
                        reader = PyPDF2.PdfReader(BytesIO(raw_bytes))
                    except Exception:
                        # Some PyPDF2 versions accept a file-like object directly
                        try:
                            uploaded.seek(0)
                        except Exception:
                            pass
                        reader = PyPDF2.PdfReader(uploaded)
                    pages = []
                    for p in reader.pages:
                        try:
                            pages.append(p.extract_text() or "")
                        except Exception:
                            pages.append("")
                    # Normalize into DataFrame: one row per page
                    rows = [[pg] for pg in pages]
                    try:
                        uploaded.seek(0)
                    except Exception:
                        pass
                    return pd.DataFrame(rows, columns=['text'])
                except Exception:
                    # If PDF handling fails, fall back to other readers below
                    try:
                        uploaded.seek(0)
                    except Exception:
                        pass

            # DOC/DOCX support (basic text extraction)
            if filename.endswith(('.doc', '.docx')):
                try:
                    try:
                        # Try python-docx for .docx files
                        if filename.endswith('.docx'):
                            import docx
                            doc = docx.Document(uploaded)
                            paragraphs = [paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()]
                            return pd.DataFrame(paragraphs, columns=['text'])
                        else:
                            # For .doc files, we'd need additional libraries like python-docx2txt or textract
                            raise ImportError("Reading .doc files requires additional libraries. Please convert to .docx or use a different format.")
                    except ImportError as ie:
                        raise ImportError("python-docx is required to read DOCX files. Please install it (pip install python-docx) to enable DOCX uploads.") from ie
                except Exception:
                    # If DOC handling fails, fall back to other readers below
                    try:
                        uploaded.seek(0)
                    except Exception:
                        pass

            # Fallback: try Excel first (if allowed), then CSV with encoding detection, then raw text
            if allow_excel:
                try:
                    return pd.read_excel(uploaded, engine='openpyxl')
                except Exception:
                    try:
                        uploaded.seek(0)
                    except Exception:
                        pass
            
            # Try CSV with multiple encodings
            encodings_to_try = ['utf-8', 'utf-8-sig', 'windows-1252', 'iso-8859-1', 'cp1252', 'latin1']
            for encoding in encodings_to_try:
                try:
                    uploaded.seek(0)
                    return pd.read_csv(uploaded, encoding=encoding)
                except (UnicodeDecodeError, UnicodeError):
                    continue
                except Exception:
                    continue
            
            # Final fallback: treat as plain text with robust encoding
            try:
                uploaded.seek(0)
            except Exception:
                pass
            txt = _read_text(uploaded)
            return pd.DataFrame([l for l in txt.splitlines() if l.strip()], columns=['text'])

        # If the uploaded file has multiple sheets (only for Excel), allow the user to select which sheet(s) to use
        sheets = []
        try:
            # attempt Excel sheet detection for .xlsx and .xls files
            fn = getattr(uploaded_file1, 'name', '') or ''
            if fn.lower().endswith(('.xlsx', '.xls')):
                xls = pd.ExcelFile(uploaded_file1, engine='openpyxl')
                sheets = xls.sheet_names
        except Exception:
            sheets = []

        if not IS_TWO_FILE_MODE and sheets:
            # For single-file modes, let user pick one sheet (or multiple if desired)
            if comparison_mode == "Compare Two Columns in Same Excel File":
                selected_sheet = st.selectbox("Select sheet to compare", sheets, index=0, key="single_sheet_select")
            else:
                selected_sheet = st.selectbox("Select sheet to compare (base vs two targets)", sheets, index=0, key="single_sheet_select_2")
            df1 = read_uploaded_file(uploaded_file1, sheet_name=selected_sheet)
            st.session_state['original_df'] = df1.copy()
            # store selected sheet for downstream export
            st.session_state['selected_sheet_singlefile'] = selected_sheet
        else:
            allow_excel_files = comparison_mode != "Compare Any Two Files"
            df1 = read_uploaded_file(uploaded_file1, allow_excel=allow_excel_files)
            st.session_state['original_df'] = df1.copy()
        if IS_TWO_FILE_MODE:
            allow_excel_files = comparison_mode != "Compare Any Two Files"
            df2 = read_uploaded_file(uploaded_file2, allow_excel=allow_excel_files)
            # Only set original_df to df2 if it wasn't already set from df1
            if st.session_state.get('original_df') is None:
                st.session_state['original_df'] = df2.copy()
            # Column selection UI for two files
            st.markdown("<b>Select columns to compare:</b>", unsafe_allow_html=True)
            col_file1, col_file2 = st.columns(2)
            with col_file1:
                st.markdown("<b>File 1</b>", unsafe_allow_html=True)
                col1_q, col1_a = st.columns(2)
                with col1_q:
                    col1_q_val = st.selectbox("Question", df1.columns, index=0, key="file1_q_sel", help="Select the question column in File 1")
                with col1_a:
                    col1_a_val = st.selectbox("Answer (File 1)", df1.columns, index=1, key="file1_a_sel", help="Select the answer column in File 1")
            with col_file2:
                st.markdown("<b>File 2</b>", unsafe_allow_html=True)
                col2_q, col2_a = st.columns(2)
                with col2_q:
                    col2_q_val = st.selectbox("Question", df2.columns, index=0, key="file2_q_sel", help="Select the question column in File 2")
                with col2_a:
                    col2_a_val = st.selectbox("Answer (File 2)", df2.columns, index=1, key="file2_a_sel", help="Select the answer column in File 2")
        else:
            df2 = df1
            if df1.shape[1] < 3:
                st.error("Error: File must have at least three columns to compare. Please check your file.")
            else:
                # Two different single-file modes: keep existing UI for the old mode; add a new UI for base->two-targets
                if comparison_mode == "Compare Two Columns in Same Excel File":
                    st.markdown("<b>Select two columns to compare within the same file:</b>", unsafe_allow_html=True)
                    col_sel1, col_sel2, col_sel3 = st.columns(3)
                    with col_sel1:
                        st.markdown("<small><b>Column A Question Column</b></small>", unsafe_allow_html=True)
                        col1_q_val = st.selectbox("Question", df1.columns, index=0, key="same_file_q_sel", help="Select the question column")
                    with col_sel2:
                        st.markdown("<small><b>Column B Answer 1</b></small>", unsafe_allow_html=True)
                        col1_a_val = st.selectbox("Base Answer", df1.columns, index=1, key="same_file_a1_sel", help="First answer column to compare (Base Answer)")
                    with col_sel3:
                        st.markdown("<small><b>Column C Answer 2</b></small>", unsafe_allow_html=True)
                        col2_a_val = st.selectbox("Target Column Ans", df1.columns, index=2 if df1.shape[1] > 2 else 1, key="same_file_a2_sel", help="Second answer column to compare (Target Column Ans)")
                    col2_q_val = col1_q_val
                elif comparison_mode == "Compare One Column with Two Targets (Same Excel)":
                    st.markdown("<b>Select base column and two target columns:</b>", unsafe_allow_html=True)
                    col_base, col_target1, col_target2 = st.columns(3)
                    with col_base:
                        st.markdown("<small><b>Base Column</b></small>", unsafe_allow_html=True)
                        col1_q_val = st.selectbox("Base Answer", df1.columns, index=0, key="base_col_sel", help="Select the base answer column")
                    with col_target1:
                        st.markdown("<small><b>Target Column A</b></small>", unsafe_allow_html=True)
                        col1_a_val = st.selectbox("Target Column Ans", df1.columns, index=1, key="target_a_sel", help="Select target answer column A")
                    with col_target2:
                        st.markdown("<small><b>Target Column B</b></small>", unsafe_allow_html=True)
                        col2_a_val = st.selectbox("Target Column Ans2", df1.columns, index=2 if df1.shape[1] > 2 else 1, key="target_b_sel", help="Select target answer column B")
                    # In this mode, treat base column as question-like column for pairing
                    col2_q_val = col1_q_val
    except Exception as e:
        error_msg = str(e)
        if 'utf-8' in error_msg.lower() or 'codec' in error_msg.lower() or 'decode' in error_msg.lower():
            st.error(f"**File Encoding Error**: The uploaded file appears to use a different text encoding. Please try:\n\n"
                    f"1. Save your file as UTF-8 encoded\n"
                    f"2. Use Excel to open and re-save as CSV (UTF-8)\n"
                    f"3. If using Excel, try saving as .xlsx format instead\n\n"
                    f"Technical details: {error_msg}")
        else:
            st.error(f"Error reading uploaded file(s): {error_msg}")

# Determine if columns have been selected
cols_selected = False
if IS_TWO_FILE_MODE:
    cols_selected = all([col1_q_val is not None, col1_a_val is not None, col2_q_val is not None, col2_a_val is not None])
else:
    cols_selected = all([col1_q_val is not None, col1_a_val is not None, col2_a_val is not None])

# Ensure api_key variable exists (may be set in the Azure UI block)
api_key = globals().get('api_key', '')

# Only show Compare button when files are uploaded and columns selected
ready_to_compare = bool(uploaded_file1 and (uploaded_file2 or comparison_mode == "Compare Two Columns in Same Excel File") and cols_selected)
# Position the Compare and Cancel buttons close to each other
if ready_to_compare:
    col_left, col_compare, col_cancel, col_right = st.columns([3, 1, 1, 3])
    with col_compare:
        # Reset cancel flag when Compare is pressed
        if "cancel_requested" in st.session_state:
            st.session_state["cancel_requested"] = False
        # Disable Compare when Azure is selected but API key is missing
        disable_compare = (matching_method == "Azure OpenAI GPT-4o" and (not api_key))
        compare_clicked = st.button("Compare", help="Click to start the similarity comparison", disabled=disable_compare)
        if disable_compare:
            st.caption("Enter Azure OpenAI API Key to enable Compare")
    with col_cancel:
        # Cancel button sets a request flag; cancellation is best-effort and works for loops
        if st.button("Cancel", help="Request cancellation of running comparison"):
            st.session_state["cancel_requested"] = True
else:
    compare_clicked = False

if compare_clicked:
    # ensure cancel flag is set to False at start
    st.session_state["cancel_requested"] = False
    
    import re, string
    import difflib
    from sentence_transformers import SentenceTransformer, util
    from sentence_transformers.cross_encoder import CrossEncoder
    # --- Step 3.1: Column Selection and Validation (use earlier selections) ---
    try:
        with st.spinner("Loading models and processing files..."):
                # Ensure df1/df2 are loaded (re-read to get fresh file buffer if needed)
                allow_excel_files = comparison_mode != "Compare Any Two Files"
                if df1 is None:
                    df1 = read_uploaded_file(uploaded_file1, allow_excel=allow_excel_files)
                if IS_TWO_FILE_MODE:
                    if df2 is None:
                        df2 = read_uploaded_file(uploaded_file2, allow_excel=allow_excel_files)

                # Capture source filenames and sheet names
                file1_name = getattr(uploaded_file1, 'name', 'file1') if uploaded_file1 is not None else ''
                try:
                    fn1 = getattr(uploaded_file1, 'name', '') or ''
                    if fn1.lower().endswith('.xlsx'):
                        sheet1_name = pd.ExcelFile(uploaded_file1, engine='openpyxl').sheet_names[0]
                    else:
                        sheet1_name = ''
                except Exception:
                    sheet1_name = ''
                # Defaults for second file
                file2_name = getattr(uploaded_file2, 'name', 'file2') if uploaded_file2 is not None else ''
                sheet2_name = ''

                # Validate file format
                if IS_TWO_FILE_MODE:
                    if df1.shape[1] < 2 or df2.shape[1] < 2:
                        st.error("Error: Both files must have at least two columns (question and answer). Please check your files.")
                        st.stop()

                    # Extract selected columns from earlier UI
                    questions1 = df1[col1_q_val].astype(str).fillna("").tolist()
                    answers1 = df1[col1_a_val].astype(str).fillna("").tolist()
                    questions2 = df2[col2_q_val].astype(str).fillna("").tolist()
                    answers2 = df2[col2_a_val].astype(str).fillna("").tolist()
                    # source file 2 info
                    file2_name = getattr(uploaded_file2, 'name', 'file2') if uploaded_file2 is not None else ''
                    try:
                        fn2 = getattr(uploaded_file2, 'name', '') or ''
                        if fn2.lower().endswith('.xlsx'):
                            sheet2_name = pd.ExcelFile(uploaded_file2, engine='openpyxl').sheet_names[0]
                        else:
                            sheet2_name = ''
                    except Exception:
                        sheet2_name = ''
                    # ---- Compare Two Excel Files: run the comparison now ----
                    # Prepare lists and truncate to minimum length
                    questions1 = questions1 if 'questions1' in locals() else df1[col1_q_val].astype(str).fillna("").tolist()
                    answers1 = answers1 if 'answers1' in locals() else df1[col1_a_val].astype(str).fillna("").tolist()
                    questions2 = questions2 if 'questions2' in locals() else df2[col2_q_val].astype(str).fillna("").tolist()
                    answers2 = answers2 if 'answers2' in locals() else df2[col2_a_val].astype(str).fillna("").tolist()

                    def clean_answer(ans):
                        ans = ans.lower()
                        ans = re.sub(r'\[.*?\]', '', ans)
                        ans = re.sub(r'\bbased on the provided context,?\s*', '', ans)
                        ans = ans.translate(str.maketrans('', '', string.punctuation))
                        ans = ans.strip()
                        return ans

                    def aggressive_clean(ans):
                        ans = clean_answer(ans)
                        ans = re.sub(r'\d+', '', ans)
                        context_phrases = [
                            r'based on the provided context', r'from the context', r'from context', r'context', r'see context',
                            r'as per context', r'per context', r'per the context', r'per the provided context', r'provided context', r'according to'
                        ]
                        for phrase in context_phrases:
                            ans = re.sub(rf'\b{phrase}\b', '', ans, flags=re.IGNORECASE)
                        ans = re.sub(r'\s+', ' ', ans)
                        return ans.strip()

                    # Normalize lengths
                    min_len = min(len(questions1), len(questions2), len(answers1), len(answers2))
                    questions1 = questions1[:min_len]
                    questions2 = questions2[:min_len]
                    answers1 = answers1[:min_len]
                    answers2 = answers2[:min_len]

                    cleaned1 = [aggressive_clean(a) for a in answers1]
                    cleaned2 = [aggressive_clean(a) for a in answers2]
                    cleaned1 = list(map(str, cleaned1))
                    cleaned2 = list(map(str, cleaned2))

                    explanations = [""] * min_len

                    # Choose method: Azure GPT per-pair, or Local Model chunked
                    if matching_method == "Azure OpenAI GPT-4o" and api_key:
                        gpt_scores = []
                        gpt_explanations = []
                        progress = st.progress(0, text="Comparing answers with GPT-4o...")
                        for idx, (a1, a2) in enumerate(zip(answers1, answers2)):
                            if st.session_state.get("cancel_requested", False):
                                st.info("Comparison cancelled by user.")
                                break
                            if not a1.strip() or not a2.strip():
                                score, explanation = 0, "One or both answers are empty."
                            else:
                                score, explanation = get_gpt4o_similarity(
                                    a1,
                                    a2,
                                    api_key,
                                    system_prompt=globals().get('gpt_system_prompt', None),
                                    user_template=globals().get('gpt_user_template', None),
                                    temperature=globals().get('gpt_temperature', 0.0),
                                    top_p=globals().get('gpt_top_p', 1.0),
                                    max_tokens=globals().get('gpt_max_tokens', 20),
                                )
                            gpt_scores.append(score if score is not None else 0)
                            gpt_explanations.append(explanation)
                            progress.progress((idx+1)/min_len, text=f"Compared {idx+1}/{min_len} pairs")
                        progress.empty()
                        final_percent_sim = gpt_scores
                        explanations = gpt_explanations
                        raw_sim = [None] * min_len
                        fuzzy_scores = [None] * min_len
                    else:
                        if matching_method == "Local Model":
                            if st.session_state.get("cancel_requested", False):
                                st.info("Comparison cancelled before model loading.")
                                raise Exception("Cancelled")
                            main_model = load_main_model(selected_model)
                            cross_encoder = load_cross_encoder_model(selected_model) if selected_model in CROSS_ENCODER_MODELS else None
                            # Chunked encoding like other local paths
                            n = min_len
                            chunk_size = 64
                            progress = st.progress(0, text="Encoding and computing local similarities...")
                            sims = []
                            raw_sims = []
                            processed = 0
                            for i in range(0, n, chunk_size):
                                if st.session_state.get("cancel_requested", False):
                                    st.info("Comparison cancelled by user.")
                                    break
                                end = min(i + chunk_size, n)
                                emb1 = main_model.encode(cleaned1[i:end], convert_to_tensor=True)
                                emb2 = main_model.encode(cleaned2[i:end], convert_to_tensor=True)
                                sim_chunk = util.cos_sim(emb1, emb2).diagonal().cpu().numpy()
                                sims.extend(sim_chunk.tolist())
                                raw_sims.extend(sim_chunk.round(4).tolist())
                                processed = end
                                progress.progress(int(processed / n * 80), text=f"Encoded and compared {processed}/{n} pairs")

                            similarities = np.array(sims)
                            percent_sim_mpnet = (similarities * 100).round(2)
                            raw_sim_mpnet = np.array(raw_sims)
                            cross_scores = None
                            if cross_encoder is not None:
                                try:
                                    pairs = list(zip(cleaned1, cleaned2))
                                    cross_sim_list = []
                                    for i in range(0, n, chunk_size):
                                        if st.session_state.get("cancel_requested", False):
                                            st.info("Comparison cancelled by user during cross-encoder step.")
                                            break
                                        end = min(i + chunk_size, n)
                                        pred = cross_encoder.predict(pairs[i:end], show_progress_bar=False)
                                        cross_sim_list.extend(pred.tolist())
                                        progress.progress(80 + int(end / n * 20), text=f"Cross-encoder processed {end}/{n} pairs")
                                    cross_sim = np.array(cross_sim_list)
                                    if cross_sim.size and np.max(cross_sim) - np.min(cross_sim) > 0:
                                        cross_sim = (cross_sim - np.min(cross_sim)) / (np.max(cross_sim) - np.min(cross_sim))
                                    cross_scores = (cross_sim * 100).round(2) if cross_sim.size else percent_sim_mpnet
                                except Exception as e:
                                    st.warning(f"Cross-encoder failed: {e}")
                                    cross_scores = percent_sim_mpnet
                            else:
                                cross_scores = percent_sim_mpnet
                                try:
                                    progress.progress(100, text="Local model comparison complete")
                                except Exception:
                                    pass

                            from difflib import SequenceMatcher
                            def fuzzy_ratio(a, b):
                                return int(SequenceMatcher(None, a, b).ratio() * 100)
                            fuzzy_scores = [fuzzy_ratio(a1, a2) for a1, a2 in zip(cleaned1, cleaned2)]
                            if cross_encoder is not None and cross_scores is not None:
                                final_percent_sim = cross_scores
                            else:
                                final_percent_sim = [max(mpnet, fuzz) for mpnet, fuzz in zip(percent_sim_mpnet, fuzzy_scores)]
                            raw_sim = raw_sim_mpnet
                        else:
                            final_percent_sim = [None] * min_len
                            raw_sim = [None] * min_len
                            fuzzy_scores = [None] * min_len

                    match_quality = [
                        "High" if s and s > threshold else ("Medium" if s and s > 60 else "Low")
                        for s in final_percent_sim
                    ]

                    # Highlight differences for file1 vs file2 answers
                    def highlight_diff(a, b):
                        seqm = difflib.SequenceMatcher(None, a, b)
                        out1, out2 = '', ''
                        for opcode, a0, a1, b0, b1 in seqm.get_opcodes():
                            if opcode == 'equal':
                                out1 += a[a0:a1]
                                out2 += b[b0:b1]
                            elif opcode == 'replace':
                                out1 += f'<span style="background-color:#ffd6d6">{a[a0:a1]}</span>'
                                out2 += f'<span style="background-color:#ffd6d6">{b[b0:b1]}</span>'
                            elif opcode == 'insert':
                                out2 += f'<span style="background-color:#d6ffd6">{b[b0:b1]}</span>'
                            elif opcode == 'delete':
                                out1 += f'<span style="background-color:#ffd6d6">{a[a0:a1]}</span>'
                        return out1, out2

                    diff1, diff2 = zip(*(highlight_diff(a1, a2) for a1, a2 in zip(answers1, answers2)))

                    q_col_name = col1_q_val if col1_q_val is not None else "Question"
                    a1_col_name = col1_a_val if col1_a_val is not None else "Answer 1"
                    a2_col_name = col2_a_val if col2_a_val is not None else "Answer 2"

                    sim_col_name = f"{a1_col_name} & {a2_col_name} Similarity"

                    results_df = pd.DataFrame({
                        q_col_name: questions1,
                        a1_col_name: answers1,
                        a2_col_name: answers2,
                        "Source File 1": file1_name,
                        "Source Sheet 1": sheet1_name,
                        "Source File 2": file2_name,
                        "Source Sheet 2": sheet2_name,
                        sim_col_name: final_percent_sim
                    })

                    st.session_state['results_df'] = results_df.copy()
                    st.session_state['similarity_cols'] = [sim_col_name]
                    st.session_state['primary_sim_col'] = sim_col_name

                    st.session_state['diff_table'] = pd.DataFrame({
                        q_col_name: questions1,
                        f"{a1_col_name} (diff)": diff1,
                        f"{a2_col_name} (diff)": diff2,
                        "Source File 1": file1_name,
                        "Source Sheet 1": sheet1_name,
                        "Source File 2": file2_name,
                        "Source Sheet 2": sheet2_name,
                        sim_col_name: final_percent_sim
                    }).copy()

                    st.success(f"Compared {min_len} question-answer pairs.")
                else:
                    df2 = df1
                    # New mode: compare one base column with two target columns (no changes to other modes)
                    if comparison_mode == "Compare One Column with Two Targets (Same Excel)":
                        # Ensure at least three columns
                        if df1.shape[1] < 3:
                            st.error("Error: File must have at least three columns to compare (base + 2 targets). Please check your file.")
                            st.stop()

                        # Map selectors: base = col1_q_val, target A = col1_a_val, target B = col2_a_val
                        base_vals = df1[col1_q_val].astype(str).fillna("").tolist()
                        target_a_vals = df1[col1_a_val].astype(str).fillna("").tolist()
                        target_b_vals = df1[col2_a_val].astype(str).fillna("").tolist()

                        file1_name = getattr(uploaded_file1, 'name', 'file') if uploaded_file1 is not None else ''
                        try:
                            sheet1_name = pd.ExcelFile(uploaded_file1, engine='openpyxl').sheet_names[0]
                        except Exception:
                            sheet1_name = ''

                        min_len = min(len(base_vals), len(target_a_vals), len(target_b_vals))
                        base_vals = base_vals[:min_len]
                        target_a_vals = target_a_vals[:min_len]
                        target_b_vals = target_b_vals[:min_len]

                        def clean_answer(ans):
                            ans = ans.lower()
                            ans = re.sub(r'\[.*?\]', '', ans)
                            ans = re.sub(r'\bbased on the provided context,?\s*', '', ans)
                            ans = ans.translate(str.maketrans('', '', string.punctuation))
                            ans = ans.strip()
                            return ans

                        def aggressive_clean(ans):
                            ans = clean_answer(ans)
                            ans = re.sub(r'\d+', '', ans)
                            context_phrases = [
                                r'based on the provided context', r'from the context', r'from context', r'context', r'see context',
                                r'as per context', r'per context', r'per the context', r'per the provided context', r'provided context', r'according to'
                            ]
                            for phrase in context_phrases:
                                ans = re.sub(rf'\b{phrase}\b', '', ans, flags=re.IGNORECASE)
                            ans = re.sub(r'\s+', ' ', ans)
                            return ans.strip()

                        cleaned_base = [aggressive_clean(a) for a in base_vals]
                        cleaned_a = [aggressive_clean(a) for a in target_a_vals]
                        cleaned_b = [aggressive_clean(a) for a in target_b_vals]
                        cleaned_base = list(map(str, cleaned_base))
                        cleaned_a = list(map(str, cleaned_a))
                        cleaned_b = list(map(str, cleaned_b))

                        explanations = [""] * min_len

                        # Use GPT or local model as appropriate, computing two similarity series
                        if matching_method == "Azure OpenAI GPT-4o" and api_key:
                            gpt_scores_a = []
                            gpt_scores_b = []
                            gpt_explanations = []
                            progress = st.progress(0, text="Comparing base->targetA and base->targetB with GPT-4o...")
                            for idx, (b, a, c) in enumerate(zip(base_vals, target_a_vals, target_b_vals)):
                                if st.session_state.get("cancel_requested", False):
                                    st.info("Comparison cancelled by user.")
                                    break
                                if not b.strip() or not a.strip():
                                    s_a, e_a = 0, "Empty base or target A"
                                else:
                                    s_a, e_a = get_gpt4o_similarity(
                                        b,
                                        a,
                                        api_key,
                                        system_prompt=globals().get('gpt_system_prompt', None),
                                        user_template=globals().get('gpt_user_template', None),
                                        temperature=globals().get('gpt_temperature', 0.0),
                                        top_p=globals().get('gpt_top_p', 1.0),
                                        max_tokens=globals().get('gpt_max_tokens', 20),
                                    )
                                if not b.strip() or not c.strip():
                                    s_b, e_b = 0, "Empty base or target B"
                                else:
                                    s_b, e_b = get_gpt4o_similarity(
                                        b,
                                        c,
                                        api_key,
                                        system_prompt=globals().get('gpt_system_prompt', None),
                                        user_template=globals().get('gpt_user_template', None),
                                        temperature=globals().get('gpt_temperature', 0.0),
                                        top_p=globals().get('gpt_top_p', 1.0),
                                        max_tokens=globals().get('gpt_max_tokens', 20),
                                    )
                                gpt_scores_a.append(s_a if s_a is not None else 0)
                                gpt_scores_b.append(s_b if s_b is not None else 0)
                                gpt_explanations.append(f"A:{e_a} | B:{e_b}")
                                progress.progress((idx+1)/min_len, text=f"Compared {idx+1}/{min_len} pairs")
                            progress.empty()
                            final_percent_sim_a = gpt_scores_a
                            final_percent_sim_b = gpt_scores_b
                            explanations = gpt_explanations
                            raw_sim_a = [None] * min_len
                            raw_sim_b = [None] * min_len
                        else:
                            if matching_method == "Local Model":
                                if st.session_state.get("cancel_requested", False):
                                    st.info("Comparison cancelled before model loading.")
                                    raise Exception("Cancelled")
                                main_model = load_main_model(selected_model)
                                cross_encoder = load_cross_encoder_model(selected_model) if selected_model in CROSS_ENCODER_MODELS else None
                                # Chunked encoding and similarity computation so we can show progress and support cancellation
                                try:
                                    import torch
                                except Exception:
                                    torch = None
                                n = min_len
                                chunk_size = 64
                                progress = st.progress(0, text="Encoding and computing local similarities...")
                                sim_a_list = []
                                sim_b_list = []
                                raw_a_list = []
                                raw_b_list = []
                                processed = 0
                                for i in range(0, n, chunk_size):
                                    if st.session_state.get("cancel_requested", False):
                                        st.info("Comparison cancelled by user.")
                                        break
                                    end = min(i + chunk_size, n)
                                    # Encode chunk
                                    emb_b_chunk = main_model.encode(cleaned_base[i:end], convert_to_tensor=True)
                                    emb_a_chunk = main_model.encode(cleaned_a[i:end], convert_to_tensor=True)
                                    emb_c_chunk = main_model.encode(cleaned_b[i:end], convert_to_tensor=True)
                                    # Compute cosine similarities per chunk
                                    sim_a_chunk = util.cos_sim(emb_b_chunk, emb_a_chunk).diagonal().cpu().numpy()
                                    sim_b_chunk = util.cos_sim(emb_b_chunk, emb_c_chunk).diagonal().cpu().numpy()
                                    sim_a_list.extend(sim_a_chunk.tolist())
                                    sim_b_list.extend(sim_b_chunk.tolist())
                                    raw_a_list.extend(sim_a_chunk.round(4).tolist())
                                    raw_b_list.extend(sim_b_chunk.round(4).tolist())
                                    processed = end
                                    progress.progress(int(processed / n * 60), text=f"Encoded and compared {processed}/{n} pairs")

                                # Aggregate results
                                sim_a = np.array(sim_a_list)
                                sim_b = np.array(sim_b_list)
                                percent_a = (sim_a * 100).round(2)
                                percent_b = (sim_b * 100).round(2)
                                raw_a = np.array(raw_a_list)
                                raw_b = np.array(raw_b_list)
                                cross_scores_a = None
                                cross_scores_b = None
                                # If cross-encoder available, run chunked predictions to update progress
                                if cross_encoder is not None:
                                    try:
                                        pairs_a = list(zip(cleaned_base, cleaned_a))
                                        pairs_b = list(zip(cleaned_base, cleaned_b))
                                        cross_scores_a_list = []
                                        cross_scores_b_list = []
                                        for i in range(0, n, chunk_size):
                                            if st.session_state.get("cancel_requested", False):
                                                st.info("Comparison cancelled by user during cross-encoder step.")
                                                break
                                            end = min(i + chunk_size, n)
                                            chunk_pairs_a = pairs_a[i:end]
                                            chunk_pairs_b = pairs_b[i:end]
                                            pred_a = cross_encoder.predict(chunk_pairs_a, show_progress_bar=False)
                                            pred_b = cross_encoder.predict(chunk_pairs_b, show_progress_bar=False)
                                            cross_scores_a_list.extend(pred_a.tolist())
                                            cross_scores_b_list.extend(pred_b.tolist())
                                            progress.progress(60 + int(end / n * 30), text=f"Cross-encoder processed {end}/{n} pairs")

                                        def normalize(arr):
                                            arr = np.array(arr)
                                            if np.max(arr) - np.min(arr) > 0:
                                                arr = (arr - np.min(arr)) / (np.max(arr) - np.min(arr))
                                            return (arr * 100).round(2)

                                        cross_scores_a = normalize(cross_scores_a_list) if cross_scores_a_list else percent_a
                                        cross_scores_b = normalize(cross_scores_b_list) if cross_scores_b_list else percent_b
                                        progress.progress(100, text="Local model comparison complete")
                                    except Exception as e:
                                        st.warning(f"Cross-encoder failed: {e}")
                                        cross_scores_a = percent_a
                                        cross_scores_b = percent_b
                                else:
                                    cross_scores_a = percent_a
                                    cross_scores_b = percent_b
                                    # Ensure the progress bar completes when cross-encoder isn't used
                                    try:
                                        progress.progress(100, text="Local model comparison complete")
                                    except Exception:
                                        pass
                                from difflib import SequenceMatcher
                                def fuzzy_ratio(a, b):
                                    return int(SequenceMatcher(None, a, b).ratio() * 100)
                                fuzzy_a = [fuzzy_ratio(a,b) for a,b in zip(cleaned_base, cleaned_a)]
                                fuzzy_b = [fuzzy_ratio(a,b) for a,b in zip(cleaned_base, cleaned_b)]
                                final_percent_sim_a = [max(mpnet, fuzz) for mpnet, fuzz in zip(cross_scores_a, fuzzy_a)]
                                final_percent_sim_b = [max(mpnet, fuzz) for mpnet, fuzz in zip(cross_scores_b, fuzzy_b)]
                                raw_sim_a = raw_a
                                raw_sim_b = raw_b
                            else:
                                final_percent_sim_a = [None] * min_len
                                final_percent_sim_b = [None] * min_len
                                raw_sim_a = [None] * min_len
                                raw_sim_b = [None] * min_len

                        match_quality_a = ["High" if s and s > threshold else ("Medium" if s and s > 60 else "Low") for s in final_percent_sim_a]
                        match_quality_b = ["High" if s and s > threshold else ("Medium" if s and s > 60 else "Low") for s in final_percent_sim_b]

                        def highlight_diff(a, b):
                            seqm = difflib.SequenceMatcher(None, a, b)
                            out1, out2 = '', ''
                            for opcode, a0, a1, b0, b1 in seqm.get_opcodes():
                                if opcode == 'equal':
                                    out1 += a[a0:a1]
                                    out2 += b[b0:b1]
                                elif opcode == 'replace':
                                    out1 += f'<span style="background-color:#ffd6d6">{a[a0:a1]}</span>'
                                    out2 += f'<span style="background-color:#ffd6d6">{b[b0:b1]}</span>'
                                elif opcode == 'insert':
                                    out2 += f'<span style="background-color:#d6ffd6">{b[b0:b1]}</span>'
                                elif opcode == 'delete':
                                    out1 += f'<span style="background-color:#ffd6d6">{a[a0:a1]}</span>'
                            return out1, out2

                        diff_a1, diff_a2 = zip(*(highlight_diff(b, a) for b, a in zip(base_vals, target_a_vals)))
                        diff_b1, diff_b2 = zip(*(highlight_diff(b, c) for b, c in zip(base_vals, target_b_vals)))

                        # Build similarity column names from actual column names
                        sim1 = f"{col1_q_val} & {col1_a_val} Similarity" if col1_q_val else f"{col1_a_val} & {col1_a_val} Similarity"
                        sim2 = f"{col1_q_val} & {col2_a_val} Similarity" if col1_q_val else f"{col2_a_val} & {col2_a_val} Similarity"

                        results_df = pd.DataFrame({
                            col1_q_val: base_vals,
                            col1_a_val: target_a_vals,
                            col2_a_val: target_b_vals,
                            "Source File": file1_name,
                            "Source Sheet": sheet1_name,
                            sim1: final_percent_sim_a,
                            "Raw Similarity 1": raw_sim_a,
                            "Match Quality 1": match_quality_a,
                            sim2: final_percent_sim_b,
                            "Raw Similarity 2": raw_sim_b,
                            "Match Quality 2": match_quality_b
                        })

                        st.session_state['results_df'] = results_df.copy()
                        diff_table = pd.DataFrame({
                            col1_q_val: base_vals,
                            f"{col1_a_val} (diff)": diff_a2,
                            f"{col2_a_val} (diff)": diff_b2,
                            "Source File": file1_name,
                            "Source Sheet": sheet1_name,
                            sim1: final_percent_sim_a,
                            "Raw Similarity 1": raw_sim_a,
                            "Match Quality 1": match_quality_a,
                            sim2: final_percent_sim_b,
                            "Raw Similarity 2": raw_sim_b,
                            "Match Quality 2": match_quality_b
                        })
                        st.session_state['diff_table'] = diff_table.copy()

                        # Persist similarity column names so downstream UI shows real names
                        st.session_state['similarity_cols'] = [sim1, sim2]
                        # For downstream metrics choose the first similarity as primary
                        st.session_state['primary_sim_col'] = sim1

                        st.success(f"Compared {min_len} base->target pairs.")
                        # Ensure consistent column name variables for downstream code
                        q_col_name = col1_q_val if col1_q_val is not None else "Question"
                        a1_col_name = col1_a_val if col1_a_val is not None else "Answer 1"
                        a2_col_name = col2_a_val if col2_a_val is not None else "Answer 2"
                        similarity_cols = []
                        if 'Similarity Match 1 (%)' in results_df.columns:
                            similarity_cols.append('Similarity Match 1 (%)')
                        if 'Similarity Match 2 (%)' in results_df.columns:
                            similarity_cols.append('Similarity Match 2 (%)')
                        st.session_state['similarity_cols'] = similarity_cols
                        # Normalize names used by downstream code
                        questions1 = base_vals
                        answers1 = target_a_vals
                        answers2 = target_b_vals
                        # diff_a2 and diff_b2 hold the target-side highlighted diffs
                        try:
                            diff1 = diff_a2
                        except NameError:
                            diff1 = [''] * len(questions1)
                        try:
                            diff2 = diff_b2
                        except NameError:
                            diff2 = [''] * len(questions1)
                    else:
                        # Existing same-file behavior: compare two columns (unchanged)
                        if df1.shape[1] < 3:
                            st.error("Error: File must have at least three columns to compare two answer columns. Please check your file.")
                            st.stop()

                        questions1 = df1[col1_q_val].astype(str).fillna("").tolist()
                        answers1 = df1[col1_a_val].astype(str).fillna("").tolist()
                        questions2 = df1[col1_q_val].astype(str).fillna("").tolist()
                        answers2 = df1[col2_a_val].astype(str).fillna("").tolist()
                        file1_name = getattr(uploaded_file1, 'name', 'file') if uploaded_file1 is not None else ''
                        try:
                            sheet1_name = pd.ExcelFile(uploaded_file1, engine='openpyxl').sheet_names[0]
                        except Exception:
                            sheet1_name = ''
                        min_len = min(len(questions1), len(questions2), len(answers1), len(answers2))
                        questions1 = questions1[:min_len]
                        questions2 = questions2[:min_len]
                        answers1 = answers1[:min_len]
                        answers2 = answers2[:min_len]

                        def clean_answer(ans):
                            ans = ans.lower()
                            ans = re.sub(r'\[.*?\]', '', ans)
                            ans = re.sub(r'\bbased on the provided context,?\s*', '', ans)
                            ans = ans.translate(str.maketrans('', '', string.punctuation))
                            ans = ans.strip()
                            return ans

                        def aggressive_clean(ans):
                            ans = clean_answer(ans)
                            ans = re.sub(r'\d+', '', ans)
                            context_phrases = [
                                r'based on the provided context', r'from the context', r'from context', r'context', r'see context', r'as per context', r'per context', r'per the context', r'per the provided context', r'provided context', r'according to'
                            ]
                            for phrase in context_phrases:
                                ans = re.sub(rf'\b{phrase}\b', '', ans, flags=re.IGNORECASE)
                            ans = re.sub(r'\s+', ' ', ans)
                            return ans.strip()

                        cleaned1 = [aggressive_clean(a) for a in answers1]
                        cleaned2 = [aggressive_clean(a) for a in answers2]
                        cleaned1 = list(map(str, cleaned1))
                        cleaned2 = list(map(str, cleaned2))

                        explanations = [""] * min_len

                        if matching_method == "Azure OpenAI GPT-4o" and api_key:
                            gpt_scores = []
                            gpt_explanations = []
                            progress = st.progress(0, text="Comparing answers with GPT-4o...")
                            for idx, (a1, a2) in enumerate(zip(answers1, answers2)):
                                if st.session_state.get("cancel_requested", False):
                                    st.info("Comparison cancelled by user.")
                                    break
                                if not a1.strip() or not a2.strip():
                                    score, explanation = 0, "One or both answers are empty."
                                else:
                                    score, explanation = get_gpt4o_similarity(
                                        a1,
                                        a2,
                                        api_key,
                                        system_prompt=globals().get('gpt_system_prompt', None),
                                        user_template=globals().get('gpt_user_template', None),
                                        temperature=globals().get('gpt_temperature', 0.0),
                                        top_p=globals().get('gpt_top_p', 1.0),
                                        max_tokens=globals().get('gpt_max_tokens', 20),
                                    )
                                gpt_scores.append(score if score is not None else 0)
                                gpt_explanations.append(explanation)
                                progress.progress((idx+1)/min_len, text=f"Compared {idx+1}/{min_len} pairs")
                            progress.empty()
                            final_percent_sim = gpt_scores
                            explanations = gpt_explanations
                            raw_sim = [None] * min_len
                            fuzzy_scores = [None] * min_len
                        else:
                            if matching_method == "Local Model":
                                if st.session_state.get("cancel_requested", False):
                                    st.info("Comparison cancelled before model loading.")
                                    raise Exception("Cancelled")
                                main_model = load_main_model(selected_model)
                                cross_encoder = load_cross_encoder_model(selected_model) if selected_model in CROSS_ENCODER_MODELS else None
                                n = min_len
                                chunk_size = 64
                                progress = st.progress(0, text="Encoding and computing local similarities...")
                                sims = []
                                raw_sims = []
                                processed = 0
                                for i in range(0, n, chunk_size):
                                    if st.session_state.get("cancel_requested", False):
                                        st.info("Comparison cancelled by user.")
                                        break
                                    end = min(i + chunk_size, n)
                                    emb1 = main_model.encode(cleaned1[i:end], convert_to_tensor=True)
                                    emb2 = main_model.encode(cleaned2[i:end], convert_to_tensor=True)
                                    sim_chunk = util.cos_sim(emb1, emb2).diagonal().cpu().numpy()
                                    sims.extend(sim_chunk.tolist())
                                    raw_sims.extend(sim_chunk.round(4).tolist())
                                    processed = end
                                    progress.progress(int(processed / n * 80), text=f"Encoded and compared {processed}/{n} pairs")

                                similarities = np.array(sims)
                                percent_sim_mpnet = (similarities * 100).round(2)
                                raw_sim_mpnet = np.array(raw_sims)
                                cross_scores = None
                                if cross_encoder is not None:
                                    try:
                                        pairs = list(zip(cleaned1, cleaned2))
                                        cross_sim_list = []
                                        for i in range(0, n, chunk_size):
                                            if st.session_state.get("cancel_requested", False):
                                                st.info("Comparison cancelled by user during cross-encoder step.")
                                                break
                                            end = min(i + chunk_size, n)
                                            pred = cross_encoder.predict(pairs[i:end], show_progress_bar=False)
                                            cross_sim_list.extend(pred.tolist())
                                            progress.progress(80 + int(end / n * 20), text=f"Cross-encoder processed {end}/{n} pairs")
                                        cross_sim = np.array(cross_sim_list)
                                        if cross_sim.size and np.max(cross_sim) - np.min(cross_sim) > 0:
                                            cross_sim = (cross_sim - np.min(cross_sim)) / (np.max(cross_sim) - np.min(cross_sim))
                                        cross_scores = (cross_sim * 100).round(2) if cross_sim.size else percent_sim_mpnet
                                    except Exception as e:
                                        st.warning(f"Cross-encoder failed: {e}")
                                        cross_scores = percent_sim_mpnet
                                else:
                                    cross_scores = percent_sim_mpnet
                                    # Ensure the progress bar completes when cross-encoder isn't used
                                    try:
                                        progress.progress(100, text="Local model comparison complete")
                                    except Exception:
                                        pass
                                from difflib import SequenceMatcher
                                def fuzzy_ratio(a, b):
                                    return int(SequenceMatcher(None, a, b).ratio() * 100)
                                fuzzy_scores = [fuzzy_ratio(a, b) for a, b in zip(cleaned1, cleaned2)]
                                if cross_encoder is not None and cross_scores is not None:
                                    final_percent_sim = cross_scores
                                else:
                                    final_percent_sim = [max(mpnet, fuzz) for mpnet, fuzz in zip(percent_sim_mpnet, fuzzy_scores)]
                                raw_sim = raw_sim_mpnet
                            else:
                                final_percent_sim = [None] * min_len
                                raw_sim = [None] * min_len
                                fuzzy_scores = [None] * min_len

                        match_quality = [
                            "High" if s and s > threshold else ("Medium" if s and s > 60 else "Low")
                            for s in final_percent_sim
                        ]
                        def highlight_diff(a, b):
                            seqm = difflib.SequenceMatcher(None, a, b)
                            out1, out2 = '', ''
                            for opcode, a0, a1, b0, b1 in seqm.get_opcodes():
                                if opcode == 'equal':
                                    out1 += a[a0:a1]
                                    out2 += b[b0:b1]
                                elif opcode == 'replace':
                                    out1 += f'<span style="background-color:#ffd6d6">{a[a0:a1]}</span>'
                                    out2 += f'<span style="background-color:#ffd6d6">{b[b0:b1]}</span>'
                                elif opcode == 'insert':
                                    out2 += f'<span style="background-color:#d6ffd6">{b[b0:b1]}</span>'
                                elif opcode == 'delete':
                                    out1 += f'<span style="background-color:#ffd6d6">{a[a0:a1]}</span>'
                            return out1, out2

                        diff1, diff2 = zip(*(highlight_diff(a1, a2) for a1, a2 in zip(answers1, answers2)))

                        q_col_name = col1_q_val if col1_q_val is not None else "Question"
                        a1_col_name = col1_a_val if col1_a_val is not None else "Answer 1"
                        a2_col_name = col2_a_val if col2_a_val is not None else "Answer 2"

                        # Dynamic similarity column name (e.g., 'TruDiscovery & Open AI Similarity')
                        sim_col_name = f"{a1_col_name} & {a2_col_name} Similarity"
                        primary_sim_col = sim_col_name

                        results_df = pd.DataFrame({
                            q_col_name: questions1,
                            a1_col_name: answers1,
                            a2_col_name: answers2,
                            "Source File": file1_name,
                            "Source Sheet": sheet1_name,
                            sim_col_name: final_percent_sim
                        })

                        # Persist results and similarity column names
                        st.session_state['results_df'] = results_df.copy()
                        st.session_state['similarity_cols'] = [sim_col_name]
                        st.session_state['primary_sim_col'] = primary_sim_col

                        st.session_state['diff_table'] = pd.DataFrame({
                            q_col_name: questions1,
                            f"{a1_col_name} (diff)": diff1,
                            f"{a2_col_name} (diff)": diff2,
                            "Source File": file1_name,
                            "Source Sheet": sheet1_name,
                            sim_col_name: final_percent_sim
                        }).copy()

                        st.success(f"Compared {min_len} question-answer pairs.")
                    
                    # Normalize column and similarity names so downstream code is branch-agnostic
                    # Determine question/answer column names (exclude source/similarity/diff columns)
                    meta_cols = {'Source File','Source Sheet','Source File 1','Source Sheet 1','Source File 2','Source Sheet 2'}
                    non_sim_cols = [c for c in results_df.columns if c not in meta_cols and 'Similarity' not in c and '(diff)' not in c]
                    # Heuristic: first non-sim column is question, next two are answers/targets
                    q_col_name = non_sim_cols[0] if len(non_sim_cols) > 0 else 'Question'
                    a1_col_name = non_sim_cols[1] if len(non_sim_cols) > 1 else 'Answer 1'
                    a2_col_name = non_sim_cols[2] if len(non_sim_cols) > 2 else 'Answer 2'

                    # Ensure lists for display/diffs exist
                    try:
                        questions1 = results_df[q_col_name].astype(str).fillna("").tolist() if q_col_name in results_df.columns else []
                    except Exception:
                        questions1 = []
                    try:
                        answers1 = results_df[a1_col_name].astype(str).fillna("").tolist() if a1_col_name in results_df.columns else []
                    except Exception:
                        answers1 = []
                    try:
                        answers2 = results_df[a2_col_name].astype(str).fillna("").tolist() if a2_col_name in results_df.columns else []
                    except Exception:
                        answers2 = []

                    # Diffs if present
                    diff1_col = f"{a1_col_name} (diff)"
                    diff2_col = f"{a2_col_name} (diff)"
                    diff1 = results_df[diff1_col].tolist() if diff1_col in results_df.columns else [''] * len(questions1)
                    diff2 = results_df[diff2_col].tolist() if diff2_col in results_df.columns else [''] * len(questions1)

                    # Determine similarity columns
                    sim_cols_local = [c for c in results_df.columns if 'Similarity' in c]
                    # Prefer a stored primary_sim_col only if it exists in the current results; otherwise pick the first detected similarity column
                    stored_primary = st.session_state.get('primary_sim_col')
                    if stored_primary and stored_primary in results_df.columns:
                        primary_sim_col = stored_primary
                    else:
                        primary_sim_col = sim_cols_local[0] if sim_cols_local else None
                        st.session_state['primary_sim_col'] = primary_sim_col
                    st.session_state['similarity_cols'] = sim_cols_local

                    # final_percent_sim used in some display blocks; provide safe value
                    if primary_sim_col is not None and primary_sim_col in results_df.columns:
                        # coerce non-numeric to NaN then to None for downstream
                        try:
                            final_percent_sim = pd.to_numeric(results_df[primary_sim_col], errors='coerce').tolist()
                        except Exception:
                            final_percent_sim = results_df[primary_sim_col].tolist()
                    else:
                        final_percent_sim = [None] * len(questions1)

                    # Build filtered datasets safely only when the similarity column exists and is numeric
                    if primary_sim_col is not None and primary_sim_col in results_df.columns:
                        try:
                            numeric_sim = pd.to_numeric(results_df[primary_sim_col], errors='coerce')
                            results_below_80 = results_df[numeric_sim < 80].copy()
                            results_below_50 = results_df[numeric_sim < 50].copy()
                        except KeyError:
                            # Column disappeared between runs; clear stored primary and fallback to empty sets
                            st.session_state['primary_sim_col'] = None
                            results_below_80 = pd.DataFrame()
                            results_below_50 = pd.DataFrame()
                        except Exception:
                            results_below_80 = pd.DataFrame()
                            results_below_50 = pd.DataFrame()
                    else:
                        results_below_80 = pd.DataFrame()
                        results_below_50 = pd.DataFrame()

                    # Persist results so they survive reruns (e.g., when downloading)
                    st.session_state['results_df'] = results_df.copy()
                    st.session_state['diff_table'] = locals().get('diff_table', pd.DataFrame()).copy()
                    
                    # Display summary statistics (Total, Above threshold, Between 40-threshold, Below 40)
                    st.markdown("### Comparison Summary")
                    col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
                    with col_stat1:
                        # Count non-null similarity entries as total pairs
                        try:
                            if primary_sim_col is not None and primary_sim_col in results_df.columns:
                                numeric_sim_all = pd.to_numeric(results_df[primary_sim_col], errors='coerce')
                                total_pairs = int(numeric_sim_all.notna().sum())
                            else:
                                total_pairs = len(results_df)
                        except Exception:
                            total_pairs = len(results_df)
                        st.metric("Total Pairs", total_pairs)
                    with col_stat2:
                        # Above threshold count
                        try:
                            if primary_sim_col is not None and primary_sim_col in results_df.columns:
                                numeric_sim = pd.to_numeric(results_df[primary_sim_col], errors='coerce')
                                count_above = int((numeric_sim >= threshold).sum())
                            else:
                                count_above = 0
                        except Exception:
                            count_above = 0
                        st.metric(f"Above {threshold}%", count_above)
                    with col_stat3:
                        # Between 40% (inclusive) and threshold (exclusive)
                        try:
                            if primary_sim_col is not None and primary_sim_col in results_df.columns:
                                numeric_sim = pd.to_numeric(results_df[primary_sim_col], errors='coerce')
                                count_between = int(((numeric_sim >= 40) & (numeric_sim < threshold)).sum())
                            else:
                                count_between = 0
                        except Exception:
                            count_between = 0
                        st.metric(f"Between 40-{threshold}%", count_between)
                    with col_stat4:
                        # Below 40%
                        try:
                            if primary_sim_col is not None and primary_sim_col in results_df.columns:
                                numeric_sim = pd.to_numeric(results_df[primary_sim_col], errors='coerce')
                                count_below_40 = int((numeric_sim < 40).sum())
                            else:
                                count_below_40 = 0
                        except Exception:
                            count_below_40 = 0
                        st.metric("Below 40%", count_below_40)
                    
                    # Highlighted differences are hidden by default; user can expand to view them
                    # Build diff table using selected column names with a (diff) suffix and include source info
                    if IS_TWO_FILE_MODE:
                        diff_payload = {
                            q_col_name: questions1,
                            f"{a1_col_name} (diff)": diff1,
                            f"{a2_col_name} (diff)": diff2,
                            "Source File 1": file1_name,
                            "Source Sheet 1": sheet1_name,
                            "Source File 2": file2_name,
                            "Source Sheet 2": sheet2_name
                        }
                        if primary_sim_col is not None and primary_sim_col in results_df.columns:
                            diff_payload[primary_sim_col] = final_percent_sim
                        diff_table = pd.DataFrame(diff_payload)
                    else:
                        diff_payload = {
                            q_col_name: questions1,
                            f"{a1_col_name} (diff)": diff1,
                            f"{a2_col_name} (diff)": diff2,
                            "Source File": file1_name,
                            "Source Sheet": sheet1_name
                        }
                        if primary_sim_col is not None and primary_sim_col in results_df.columns:
                            diff_payload[primary_sim_col] = final_percent_sim
                        diff_table = pd.DataFrame(diff_payload)
                    # Comparison complete; results saved to session and shown below.
                    st.success("Comparison complete — results saved. Use the Download Options and expanders below to inspect or download results.")
    except Exception as e:
        st.error(f"Error processing files: {e}")
        # --- Error Logging for Debugging ---
        with open("error_log.txt", "a", encoding="utf-8") as logf:
            import traceback
            logf.write(traceback.format_exc() + "\n")
    # If results exist in session_state (e.g., after a rerun from download), show them
    if st.session_state.get('results_df') is not None:
        results_df = st.session_state['results_df']
        diff_table = st.session_state.get('diff_table')
        # Recompute filtered sets using detected similarity columns, with guards
        similarity_cols = [c for c in results_df.columns if 'Similarity' in c]
        primary_sim_col = similarity_cols[0] if similarity_cols else None
        if primary_sim_col is not None and primary_sim_col in results_df.columns:
            try:
                numeric_sim = pd.to_numeric(results_df[primary_sim_col], errors='coerce')
                results_below_80 = results_df[numeric_sim < 80].copy()
                results_below_50 = results_df[numeric_sim < 50].copy()
            except KeyError:
                st.session_state['primary_sim_col'] = None
                results_below_80 = pd.DataFrame()
                results_below_50 = pd.DataFrame()
            except Exception:
                results_below_80 = pd.DataFrame()
                results_below_50 = pd.DataFrame()
        else:
            results_below_80 = pd.DataFrame()
            results_below_50 = pd.DataFrame()

        st.markdown("---")
        with st.expander("Show highlighted differences", expanded=False):
            if diff_table is not None and not diff_table.empty:
                st.markdown("**Highlighted Differences:**")
                st.write("**Legend:** Red = difference, Green = insertion")
                st.write("You can scroll the table below to see highlighted differences.")
                st.write(diff_table.to_html(escape=False), unsafe_allow_html=True)
            else:
                st.info("No highlighted differences to show.")

        with st.expander("Show full results table", expanded=False):
            st.markdown("**Full Results Table:**")
            st.dataframe(results_df, use_container_width=True)

        # Download options (drop source metadata)
        st.markdown("### Download Options")
        # Let user choose sheet naming for the exported Excel (only affects the written sheet name)
        sheet_name_choice = None
        if IS_TWO_FILE_MODE:
            sheet_name_choice = st.selectbox(
                "Export sheet name:",
                ["Sheet", "Merge file+sheet names", "Use first file sheet name (original)"],
                index=0,
                help="Choose how the sheet will be named inside the exported Excel file. 'Sheet' is a fixed name."
            )
        col_dl1, col_dl2, col_dl3 = st.columns(3)
        with col_dl1:
            # Use original uploaded columns for export (preserve original shape) and append similarity column(s)
            if IS_TWO_FILE_MODE:
                # For two-file comparisons, export the paired results (both answer columns + question + similarity)
                # This avoids only exporting file1's columns when the user expects both answers side-by-side.
                exclude_meta = ['Source File','Source Sheet','Source File 1','Source Sheet 1','Source File 2','Source Sheet 2']
                export_cols = [col for col in results_df.columns if col not in exclude_meta and '(diff)' not in col]
                export_df = results_df[export_cols].copy()
            else:
                if st.session_state.get('original_df') is not None:
                    export_df = st.session_state['original_df'].copy()
                else:
                    # Fallback: use columns from results_df minus source metadata and similarity columns
                    exclude_cols = ['Source File','Source Sheet','Source File 1','Source Sheet 1','Source File 2','Source Sheet 2']
                    orig_cols = [col for col in results_df.columns if col not in exclude_cols and ('Similarity' not in col and '(diff)' not in col)]
                    export_df = results_df[orig_cols].copy()
            # Format and add all similarity columns as numeric percentages (avoid Excel 'number stored as text' warning)
            similarity_cols_export = [c for c in results_df.columns if 'Similarity' in c]
            for col in similarity_cols_export:
                # Convert to numeric and store as a fraction (e.g., 0.83 for 83%) so Excel recognizes it as a number
                numeric = pd.to_numeric(results_df[col], errors='coerce')
                # Use integer percent precision to mirror previous behaviour, then divide by 100 for Excel percent format
                export_df[col] = numeric.apply(lambda v: (int(round(v)) / 100.0) if pd.notnull(v) else None)
            output_all = io.BytesIO()
            # Determine sheet name based on user choice
            sheet_name = "Results"
            try:
                if sheet_name_choice == "Sheet":
                    sheet_name = "Sheet"
                elif sheet_name_choice == "Merge file+sheet names":
                    # Try to read both sheet names and merge them; fall back to combined file names or 'Results'
                    s1 = s2 = None
                    try:
                        if uploaded_file1 is not None:
                            fn1 = getattr(uploaded_file1, 'name', '') or ''
                            if fn1.lower().endswith('.xlsx'):
                                s1 = pd.ExcelFile(uploaded_file1, engine='openpyxl').sheet_names[0]
                            else:
                                s1 = None
                        else:
                            s1 = None
                    except Exception:
                        s1 = None
                    try:
                        if uploaded_file2 is not None:
                            fn2 = getattr(uploaded_file2, 'name', '') or ''
                            if fn2.lower().endswith('.xlsx'):
                                s2 = pd.ExcelFile(uploaded_file2, engine='openpyxl').sheet_names[0]
                            else:
                                s2 = None
                        else:
                            s2 = None
                    except Exception:
                        s2 = None
                    if s1 and s2:
                        # sanitize and shorten names to avoid Excel sheet name limits
                        def clean_sheet(n):
                            return str(n)[:25].replace('/', '_').replace('\\', '_')
                        sheet_name = f"{clean_sheet(s1)}_{clean_sheet(s2)}"
                    else:
                        # fallback to merging file basenames
                        f1 = uploaded_file1.name.rsplit('.',1)[0] if uploaded_file1 is not None else ''
                        f2 = uploaded_file2.name.rsplit('.',1)[0] if uploaded_file2 is not None else ''
                        if f1 or f2:
                            sheet_name = (f1 + '_' + f2)[:31]
                        else:
                            sheet_name = "Results"
                else:
                    # Default/original behaviour: use first available uploaded file's first sheet
                    if uploaded_file1 is not None:
                        fn1 = getattr(uploaded_file1, 'name', '') or ''
                        if fn1.lower().endswith('.xlsx'):
                            sheet_name = pd.ExcelFile(uploaded_file1, engine='openpyxl').sheet_names[0]
                    elif uploaded_file2 is not None:
                        fn2 = getattr(uploaded_file2, 'name', '') or ''
                        if fn2.lower().endswith('.xlsx'):
                            sheet_name = pd.ExcelFile(uploaded_file2, engine='openpyxl').sheet_names[0]
            except Exception:
                sheet_name = "Results"
            with pd.ExcelWriter(output_all, engine='openpyxl') as writer:
                # If single-file modes, write all sheets from uploaded_file1 and inject export_df into selected sheet
                try:
                    sheet_results = st.session_state.get('sheet_results', {})
                    if not IS_TWO_FILE_MODE and uploaded_file1 is not None:
                        # selected sheet name stored earlier when reading
                        sel = st.session_state.get('selected_sheet_singlefile')
                        # Only treat uploaded_file1 as an Excel workbook if it appears to be .xlsx
                        fn1 = getattr(uploaded_file1, 'name', '') or ''
                        if fn1.lower().endswith('.xlsx'):
                            xls = pd.ExcelFile(uploaded_file1, engine='openpyxl')
                            for s in xls.sheet_names:
                                df_orig = read_uploaded_file(uploaded_file1, sheet_name=s)
                                if sel and s == sel:
                                    # Merge similarity columns into the original sheet (preserve original columns)
                                    try:
                                        res_df = sheet_results.get('file1', {}).get(s) if sheet_results else None
                                        if res_df is None:
                                            res_df = st.session_state.get('results_df') if st.session_state.get('results_df') is not None else export_df
                                        df_write = df_orig.copy()
                                        sim_cols = [c for c in res_df.columns if 'Similarity' in c]
                                        for c in sim_cols:
                                            numeric = pd.to_numeric(res_df[c], errors='coerce')
                                            df_write[c] = numeric.apply(lambda v: (int(round(v)) / 100.0) if pd.notnull(v) else None)
                                        df_write.to_excel(writer, index=False, sheet_name=s)
                                    except Exception:
                                        df_orig.to_excel(writer, index=False, sheet_name=s)
                                    # Immediately create & write a per-sheet Summary next to this sheet
                                    try:
                                        res_df = sheet_results.get('file1', {}).get(s, st.session_state.get('results_df'))
                                        if res_df is not None:
                                            # compute summary
                                            sim_col = st.session_state.get('primary_sim_col')
                                            if not sim_col or sim_col not in res_df.columns:
                                                sim_cols = [c for c in res_df.columns if 'Similarity' in c]
                                                sim_col = sim_cols[0] if sim_cols else None
                                            if sim_col and sim_col in res_df.columns:
                                                numeric_sim = pd.to_numeric(res_df[sim_col], errors='coerce')
                                                total_pairs = int(numeric_sim.notna().sum())
                                                above_thresh = int((numeric_sim >= threshold).sum())
                                                between_40_thresh = int(((numeric_sim >= 40) & (numeric_sim < threshold)).sum())
                                                below_40 = int((numeric_sim < 40).sum())
                                            else:
                                                total_pairs = len(res_df)
                                                above_thresh = between_40_thresh = below_40 = 0
                                            summary_df = pd.DataFrame({
                                                "Metric": ["Total Pairs", f"Above {threshold}%", f"Between 40-{threshold}%", "Below 40%", "High Threshold (%)"],
                                                "Value": [total_pairs, above_thresh, between_40_thresh, below_40, threshold]
                                            })
                                            # sheet name for summary
                                            summary_name = (f"{s} Summary"[:28] + '...') if len(f"{s} Summary") > 31 else f"{s} Summary"
                                            summary_df.to_excel(writer, index=False, sheet_name=summary_name)
                                    except Exception:
                                        pass
                                else:
                                    df_orig.to_excel(writer, index=False, sheet_name=s)
                        else:
                            # uploaded_file1 is not an Excel workbook (likely CSV) - write the exported results sheet only
                            export_df.to_excel(writer, index=False, sheet_name=sel if sel else sheet_name)
                            if sel and s == sel:
                                # Merge similarity columns into the original sheet (preserve original columns)
                                try:
                                    res_df = sheet_results.get('file1', {}).get(s) if sheet_results else None
                                    if res_df is None:
                                        res_df = st.session_state.get('results_df') if st.session_state.get('results_df') is not None else export_df
                                    df_write = df_orig.copy()
                                    sim_cols = [c for c in res_df.columns if 'Similarity' in c]
                                    for c in sim_cols:
                                        numeric = pd.to_numeric(res_df[c], errors='coerce')
                                        df_write[c] = numeric.apply(lambda v: (int(round(v)) / 100.0) if pd.notnull(v) else None)
                                    df_write.to_excel(writer, index=False, sheet_name=s)
                                except Exception:
                                    df_orig.to_excel(writer, index=False, sheet_name=s)
                                # Immediately create & write a per-sheet Summary next to this sheet
                                try:
                                    res_df = sheet_results.get('file1', {}).get(s, st.session_state.get('results_df'))
                                    if res_df is not None:
                                        # compute summary
                                        sim_col = st.session_state.get('primary_sim_col')
                                        if not sim_col or sim_col not in res_df.columns:
                                            sim_cols = [c for c in res_df.columns if 'Similarity' in c]
                                            sim_col = sim_cols[0] if sim_cols else None
                                        if sim_col and sim_col in res_df.columns:
                                            numeric_sim = pd.to_numeric(res_df[sim_col], errors='coerce')
                                            total_pairs = int(numeric_sim.notna().sum())
                                            above_thresh = int((numeric_sim >= threshold).sum())
                                            between_40_thresh = int(((numeric_sim >= 40) & (numeric_sim < threshold)).sum())
                                            below_40 = int((numeric_sim < 40).sum())
                                        else:
                                            total_pairs = len(res_df)
                                            above_thresh = between_40_thresh = below_40 = 0
                                        summary_df = pd.DataFrame({
                                            "Metric": ["Total Pairs", f"Above {threshold}%", f"Between 40-{threshold}%", "Below 40%", "High Threshold (%)"],
                                            "Value": [total_pairs, above_thresh, between_40_thresh, below_40, threshold]
                                        })
                                        # sheet name for summary
                                        summary_name = (f"{s} Summary"[:28] + '...') if len(f"{s} Summary") > 31 else f"{s} Summary"
                                        summary_df.to_excel(writer, index=False, sheet_name=summary_name)
                                except Exception:
                                    pass
                            else:
                                df_orig.to_excel(writer, index=False, sheet_name=s)
                        # end for sheets
                    elif IS_TWO_FILE_MODE and uploaded_file1 is not None and uploaded_file2 is not None:
                        # For two-file comparisons produce a single slim results sheet and one Summary sheet
                        res_df = st.session_state.get('results_df', results_df)
                        if res_df is None or res_df.empty:
                            # No comparison results available; write original workbooks as-is
                            x1 = pd.ExcelFile(uploaded_file1, engine='openpyxl')
                            # If uploaded files are Excel workbooks, mirror their sheets. If CSVs, skip mirroring and write results only.
                            fn1 = getattr(uploaded_file1, 'name', '') or ''
                            fn2 = getattr(uploaded_file2, 'name', '') or ''
                            if fn1.lower().endswith('.xlsx'):
                                x1 = pd.ExcelFile(uploaded_file1, engine='openpyxl')
                                for s in x1.sheet_names:
                                    df_orig = read_uploaded_file(uploaded_file1, sheet_name=s)
                                    df_orig.to_excel(writer, index=False, sheet_name=s)
                            if fn2.lower().endswith('.xlsx'):
                                x2 = pd.ExcelFile(uploaded_file2, engine='openpyxl')
                                for s in x2.sheet_names:
                                    df_orig = read_uploaded_file(uploaded_file2, sheet_name=s)
                                    out_name = s
                                    if out_name in writer.sheets:
                                        out_name = f"{s}_file2"
                                    df_orig.to_excel(writer, index=False, sheet_name=out_name)
                        else:
                            try:
                                meta_cols = {'Source File','Source Sheet','Source File 1','Source Sheet 1','Source File 2','Source Sheet 2'}
                                non_sim = [c for c in res_df.columns if c not in meta_cols and 'Similarity' not in c and '(diff)' not in c]
                                q_col = non_sim[0] if len(non_sim) > 0 else (res_df.columns[0] if len(res_df.columns) > 0 else 'Question')
                                a1_col = non_sim[1] if len(non_sim) > 1 else (res_df.columns[1] if len(res_df.columns) > 1 else 'Answer 1')
                                a2_col = non_sim[2] if len(non_sim) > 2 else (res_df.columns[2] if len(res_df.columns) > 2 else 'Answer 2')
                                sim_cols = [c for c in res_df.columns if 'Similarity' in c]
                                sim_col = st.session_state.get('primary_sim_col') if st.session_state.get('primary_sim_col') in res_df.columns else (sim_cols[0] if sim_cols else None)

                                slim = pd.DataFrame()
                                slim['Question'] = res_df[q_col] if q_col in res_df.columns else res_df.iloc[:,0]
                                slim['Answer File1'] = res_df[a1_col] if a1_col in res_df.columns else (res_df.iloc[:,1] if res_df.shape[1] > 1 else None)
                                slim['Answer File2'] = res_df[a2_col] if a2_col in res_df.columns else (res_df.iloc[:,2] if res_df.shape[1] > 2 else None)
                                if sim_col and sim_col in res_df.columns:
                                    numeric = pd.to_numeric(res_df[sim_col], errors='coerce')
                                    slim['Similarity'] = numeric.apply(lambda v: (int(round(v)) / 100.0) if pd.notnull(v) else None)
                                # write results and summary only
                                writer_sheet_name = sheet_name if 'sheet_name' in locals() else 'Results'
                                slim.to_excel(writer, index=False, sheet_name=writer_sheet_name)
                                try:
                                    if 'Similarity' in slim.columns:
                                        numeric_sim = pd.to_numeric((slim['Similarity'] * 100).round(2), errors='coerce')
                                        total_pairs = int(numeric_sim.notna().sum())
                                        above_thresh = int((numeric_sim >= threshold).sum())
                                        between_40_thresh = int(((numeric_sim >= 40) & (numeric_sim < threshold)).sum())
                                        below_40 = int((numeric_sim < 40).sum())
                                    else:
                                        total_pairs = len(slim)
                                        above_thresh = between_40_thresh = below_40 = 0
                                    summary_df = pd.DataFrame({
                                        "Metric": ["Total Pairs", f"Above {threshold}%", f"Between 40-{threshold}%", "Below 40%", "High Threshold (%)"],
                                        "Value": [total_pairs, above_thresh, between_40_thresh, below_40, threshold]
                                    })
                                    summary_name = (f"{writer_sheet_name} Summary"[:28] + '...') if len(f"{writer_sheet_name} Summary") > 31 else f"{writer_sheet_name} Summary"
                                    summary_df.to_excel(writer, index=False, sheet_name=summary_name)
                                    st.session_state['all_results_summary_written'] = True
                                except Exception:
                                    pass
                            except Exception:
                                res_df.to_excel(writer, index=False, sheet_name=sheet_name)
                    else:
                        # Fallback: write the export_df to a single sheet
                        export_df.to_excel(writer, index=False, sheet_name=sheet_name)

                except Exception:
                    # If anything fails, fall back to single-sheet export
                    export_df.to_excel(writer, index=False, sheet_name=sheet_name)

                # After writing sheets, apply percent formatting to any similarity columns across sheets
                try:
                    for sname, ws in writer.sheets.items():
                        headers = [cell.value for cell in ws[1]]
                        for idx_h, h in enumerate(headers, start=1):
                            if h and 'Similarity' in str(h):
                                col_letter = get_column_letter(idx_h)
                                for row in range(2, ws.max_row + 1):
                                    cell = ws[f"{col_letter}{row}"]
                                    if cell.value is not None:
                                        cell.number_format = '0%'
                except Exception:
                    pass
                # Add a Summary sheet (only for the "All Results" export)
                # The user requested no Summary sheet for the "Compare Two Excel Files" mode,
                # so skip creating the global summary when that mode is active.
                if not IS_TWO_FILE_MODE and not st.session_state.get('all_results_summary_written', False):
                    try:
                        # Prefer results stored in session (if the app kept a copy), otherwise use local results_df
                        df_stats = st.session_state.get('results_df', results_df)
                        # Try to find the primary similarity column
                        sim_col = None
                        stored_primary = st.session_state.get('primary_sim_col')
                        if stored_primary and df_stats is not None and stored_primary in getattr(df_stats, 'columns', []):
                            sim_col = stored_primary
                        else:
                            sim_cols = [c for c in getattr(df_stats, 'columns', []) if 'Similarity' in c]
                            sim_col = sim_cols[0] if sim_cols else None

                        if df_stats is not None and sim_col and sim_col in df_stats.columns:
                            numeric_sim = pd.to_numeric(df_stats[sim_col], errors='coerce')
                            total_pairs = int(numeric_sim.notna().sum())
                            # Use the app's threshold slider for above/between/below counts
                            above_thresh = int((numeric_sim >= threshold).sum())
                            between_40_thresh = int(((numeric_sim >= 40) & (numeric_sim < threshold)).sum())
                            below_40 = int((numeric_sim < 40).sum())
                            # Also include below 50% as an additional metric (kept for backward compatibility)
                            below_50 = int((numeric_sim < 50).sum())
                        else:
                            # Fallback counts when similarity column isn't available
                            total_pairs = len(df_stats) if df_stats is not None else 0
                            above_thresh = between_40_thresh = below_40 = below_50 = 0
                    except Exception:
                        total_pairs = above_thresh = between_40_thresh = below_40 = below_50 = ""

                    summary_df = pd.DataFrame({
                        "Metric": [
                            "Total Pairs",
                            f"Above {threshold}%",
                            f"Between 40-{threshold}%",
                            "Below 40%",
                            "High Threshold (%)"
                        ],
                        "Value": [total_pairs, above_thresh, between_40_thresh, below_40, threshold]
                    })
                    # Determine a descriptive summary sheet name that includes the compared sheet when possible
                    try:
                        summary_sheet_base = "Summary"
                        # Prefer single-file selected sheet
                        sel_single = st.session_state.get('selected_sheet_singlefile')
                        if sel_single:
                            summary_sheet_base = f"{sel_single} Summary"
                        else:
                            # Otherwise, try to pull a sheet name from sheet_results mapping
                            sheet_results = st.session_state.get('sheet_results', {})
                            # Prefer file1 mapping
                            if sheet_results.get('file1'):
                                first_sheet = next(iter(sheet_results['file1'].keys()))
                                if first_sheet:
                                    summary_sheet_base = f"{first_sheet} Summary"
                            elif sheet_results.get('file2'):
                                first_sheet = next(iter(sheet_results['file2'].keys()))
                                if first_sheet:
                                    summary_sheet_base = f"{first_sheet} Summary"
                        # Excel sheet name limit is 31 characters
                        summary_sheet_name = (summary_sheet_base[:28] + '...') if len(summary_sheet_base) > 31 else summary_sheet_base
                    except Exception:
                        summary_sheet_name = "Summary"

                    summary_df.to_excel(writer, index=False, sheet_name=summary_sheet_name)
            output_all.seek(0)
            # Use original filename for export
            base_filename = ''
            if uploaded_file1 is not None:
                base_filename = uploaded_file1.name.rsplit('.', 1)[0]
            elif uploaded_file2 is not None:
                base_filename = uploaded_file2.name.rsplit('.', 1)[0]
            else:
                base_filename = 'exported_results'
            export_filename = f"{base_filename}_similarity.xlsx"
            st.download_button(
                label=f"📊 All Results ({len(results_df)} pairs)",
                data=output_all,
                file_name=export_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download all comparison results",
                key="download_all_results_state"
            )
        with col_dl2:
            if len(results_below_80) > 0:
                if st.session_state.get('original_df') is not None:
                    export_80 = st.session_state['original_df'].loc[results_below_80.index].copy()
                else:
                    orig_cols_80 = [col for col in results_below_80.columns if col not in [
                        'Raw Similarity','Match Quality',
                        'Source File','Source Sheet','Source File 1','Source Sheet 1','Source File 2','Source Sheet 2'
                    ] and 'Similarity' not in col and '(diff)' not in col]
                    export_80 = results_below_80[orig_cols_80].copy()
                # Add similarity columns as numeric percentages
                sim_cols_80 = [c for c in results_below_80.columns if 'Similarity' in c]
                for col in sim_cols_80:
                    numeric = pd.to_numeric(results_below_80[col], errors='coerce')
                    export_80[col] = numeric.apply(lambda v: (int(round(v)) / 100.0) if pd.notnull(v) else None)
                output_80 = io.BytesIO()
                with pd.ExcelWriter(output_80, engine='openpyxl') as writer:
                    export_80.to_excel(writer, index=False, sheet_name="Below 80%")
                    # Apply percent formatting to similarity columns
                    try:
                        ws = writer.sheets["Below 80%"]
                        for col in sim_cols_80:
                            try:
                                col_idx = list(export_80.columns).index(col) + 1
                                col_letter = get_column_letter(col_idx)
                                for row in range(2, ws.max_row + 1):
                                    cell = ws[f"{col_letter}{row}"]
                                    if cell.value is not None:
                                        cell.number_format = '0%'
                            except Exception:
                                continue
                    except Exception:
                        pass
                output_80.seek(0)
                st.download_button(
                    label=f"⚠️ Below 80% ({len(results_below_80)} pairs)",
                    data=output_80,
                    file_name="similarity_match_below_80.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Download pairs with similarity below 80%",
                    key="download_below_80_state"
                )
            else:
                st.info("No results below 80%")
        with col_dl3:
            if len(results_below_50) > 0:
                if st.session_state.get('original_df') is not None:
                    export_50 = st.session_state['original_df'].loc[results_below_50.index].copy()
                else:
                    orig_cols_50 = [col for col in results_below_50.columns if col not in [
                        'Raw Similarity','Match Quality',
                        'Source File','Source Sheet','Source File 1','Source Sheet 1','Source File 2','Source Sheet 2'
                    ] and 'Similarity' not in col and '(diff)' not in col]
                    export_50 = results_below_50[orig_cols_50].copy()
                # Add similarity columns as numeric percentages
                sim_cols_50 = [c for c in results_below_50.columns if 'Similarity' in c]
                for col in sim_cols_50:
                    numeric = pd.to_numeric(results_below_50[col], errors='coerce')
                    export_50[col] = numeric.apply(lambda v: (int(round(v)) / 100.0) if pd.notnull(v) else None)
                output_50 = io.BytesIO()
                with pd.ExcelWriter(output_50, engine='openpyxl') as writer:
                    export_50.to_excel(writer, index=False, sheet_name="Below 50%")
                    # Apply percent formatting to similarity columns
                    try:
                        ws = writer.sheets["Below 50%"]
                        for col in sim_cols_50:
                            try:
                                col_idx = list(export_50.columns).index(col) + 1
                                col_letter = get_column_letter(col_idx)
                                for row in range(2, ws.max_row + 1):
                                    cell = ws[f"{col_letter}{row}"]
                                    if cell.value is not None:
                                        cell.number_format = '0%'
                            except Exception:
                                continue
                    except Exception:
                        pass
                output_50.seek(0)
                st.download_button(
                    label=f"❌ Below 50% ({len(results_below_50)} pairs)",
                    data=output_50,
                    file_name="similarity_match_below_50.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Download pairs with similarity below 50%",
                    key="download_below_50_state"
                )
            else:
                st.info("No results below 50%")
else:
    if IS_TWO_FILE_MODE:
        st.info("Please upload both files to begin.")
    else:
        st.info("Please upload a file to begin column comparison.")
