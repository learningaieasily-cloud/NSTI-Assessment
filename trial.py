import streamlit as st
import streamlit.components.v1 as components
from datetime import date, datetime, time, timedelta
import html
from io import BytesIO
import json
import pandas as pd
import re
import os
import random
import smtplib
import subprocess
import sys
import tempfile
import time as time_module
from pathlib import Path
from email.message import EmailMessage

from lms_core import (
    DEFAULT_COURSES,
    DEFAULT_LEVEL_DISTRIBUTION,
    DEFAULT_LEVEL_MARKS,
    LEVEL_LABELS,
    SUPPORTED_LANGUAGES,
    load_courses,
    normalize_language,
    read_selected_sheets,
    select_random_questions,
    save_courses,
    standardize_question_bank,
    summarize_performance,
    summarize_question_bank,
)

try:
    from openpyxl.chart import BarChart, Reference
except ImportError:
    BarChart = None
    Reference = None

# ---------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------
st.set_page_config(
    page_title="Edunet Assessment LMS",
    page_icon="📘",
    layout="wide"
)

# ---------------------------------------------------
# CREATE ASSESSMENT FOLDER
# ---------------------------------------------------
if not os.path.exists("assessment_files"):
    os.makedirs("assessment_files")

if not os.path.exists("assessment_progress"):
    os.makedirs("assessment_progress")

if not os.path.exists("assessment_metadata"):
    os.makedirs("assessment_metadata")

# ---------------------------------------------------
# CUSTOM CSS
# ---------------------------------------------------
st.markdown(
    """
    <style>

    :root {
        --app-bg: #eef4f7;
        --surface: rgba(255, 255, 255, 0.94);
        --surface-strong: #ffffff;
        --ink: #0f172a;
        --muted: #64748b;
        --line: #dbe4ea;
        --accent: #2563eb;
        --accent-strong: #1d4ed8;
        --teal: #0f766e;
        --green: #15803d;
        --danger: #dc2626;
        --sidebar: #111827;
        --code-bg: #0c1424;
        --shadow-sm: 0 8px 22px rgba(15, 23, 42, 0.07);
        --shadow-md: 0 18px 50px rgba(15, 23, 42, 0.12);
    }

    @keyframes cardIn {
        from {
            opacity: 0;
            transform: translateY(12px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }

    @keyframes softPulse {
        0%, 100% {
            box-shadow: 0 0 0 0 rgba(37, 99, 235, 0.18);
        }
        50% {
            box-shadow: 0 0 0 7px rgba(37, 99, 235, 0);
        }
    }

    html, body, [class*="css"] {
        font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }

    [data-testid="stAppViewContainer"] {
        background:
            radial-gradient(circle at top left, rgba(20, 184, 166, 0.13), transparent 30%),
            linear-gradient(135deg, #edf7f8 0%, #f7fafc 46%, #f4f0e8 100%);
    }

    [data-testid="stHeader"] {
        background: transparent;
    }

    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1320px;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #111827 0%, #172033 62%, #0f172a 100%);
        border-right: 1px solid rgba(255, 255, 255, 0.08);
    }

    [data-testid="stSidebar"] * {
        color: #e5edf5;
    }

    [data-testid="stSidebar"] h1 {
        font-size: 1.15rem;
        letter-spacing: 0;
        margin-bottom: 1.1rem;
    }

    [data-testid="stSidebar"] .stButton>button {
        background: rgba(255, 255, 255, 0.08);
        border: 1px solid rgba(255, 255, 255, 0.12);
        box-shadow: none;
    }

    [data-testid="stSidebar"] .stButton>button:hover {
        background: rgba(37, 99, 235, 0.42);
        border-color: rgba(125, 211, 252, 0.45);
    }

    [data-testid="stSidebar"] [data-testid="column"] .stButton>button {
        width: 38px !important;
        min-width: 38px !important;
        max-width: 38px !important;
        height: 38px !important;
        min-height: 38px !important;
        padding: 0 !important;
        border-radius: 50% !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        line-height: 1 !important;
        white-space: nowrap !important;
        font-size: 13px !important;
    }

    [data-testid="stSidebar"] [data-testid="column"] .stButton {
        display: flex;
        justify-content: center;
    }

    .dashboard-card {
        background: var(--surface);
        border-radius: 8px;
        padding: 24px;
        margin-bottom: 20px;
        box-shadow: var(--shadow-sm);
        border: 1px solid rgba(219, 228, 234, 0.9);
        animation: cardIn 0.45s ease both;
        backdrop-filter: blur(18px);
    }

    .dashboard-card h1,
    .dashboard-card h2,
    .dashboard-card h3 {
        color: var(--ink);
        letter-spacing: 0;
        margin: 0;
    }

    .dashboard-card p {
        color: var(--muted);
        margin: 0.45rem 0 0;
    }

    .hero-card {
        background:
            linear-gradient(135deg, rgba(15, 118, 110, 0.95), rgba(37, 99, 235, 0.92)),
            #0f766e;
        color: white;
        box-shadow: var(--shadow-md);
    }

    .hero-card h1,
    .hero-card p {
        color: white;
    }

    .metric-card {
        background: var(--surface-strong);
        border-radius: 8px;
        padding: 20px;
        border: 1px solid var(--line);
        margin-bottom: 15px;
        box-shadow: var(--shadow-sm);
        transition: transform 180ms ease, border-color 180ms ease, box-shadow 180ms ease;
    }

    .metric-card:hover {
        transform: translateY(-2px);
        border-color: rgba(37, 99, 235, 0.45);
        box-shadow: var(--shadow-md);
    }

    .metric-pill {
        display: inline-block;
        margin-top: 10px;
        background: rgba(15, 118, 110, 0.10);
        color: var(--teal);
        padding: 6px 12px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 700;
        letter-spacing: 0;
    }

    .assessment-topbar {
        display: grid;
        grid-template-columns: 1fr auto;
        gap: 18px;
        align-items: end;
        margin-bottom: 18px;
        animation: cardIn 0.4s ease both;
    }

    .assessment-kicker {
        color: var(--teal);
        font-size: 12px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 6px;
    }

    .assessment-title {
        color: var(--ink);
        font-size: clamp(2rem, 4vw, 3.1rem);
        line-height: 1.04;
        font-weight: 800;
        letter-spacing: 0;
        margin: 0;
    }

    .score-strip {
        display: flex;
        gap: 12px;
        flex-wrap: wrap;
        justify-content: flex-end;
    }

    .score-card {
        min-width: 108px;
        background: var(--surface-strong);
        border: 1px solid var(--line);
        border-radius: 8px;
        padding: 13px 15px;
        box-shadow: var(--shadow-sm);
    }

    .score-card span {
        display: block;
        color: var(--muted);
        font-size: 11px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }

    .score-card strong {
        display: block;
        color: var(--ink);
        font-size: 24px;
        line-height: 1.1;
        margin-top: 4px;
    }

    .panel-card {
        background: var(--surface);
        border: 1px solid rgba(219, 228, 234, 0.95);
        border-radius: 8px;
        box-shadow: var(--shadow-sm);
        padding: 18px;
        margin-bottom: 18px;
        animation: cardIn 0.5s ease both;
    }

    .panel-title {
        color: var(--ink);
        font-size: 14px;
        font-weight: 800;
        margin: 0 0 14px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .panel-title::before {
        content: "";
        width: 8px;
        height: 8px;
        border-radius: 999px;
        background: var(--teal);
        animation: softPulse 2.4s ease infinite;
    }

    .question-text {
        font-size: 17px;
        font-weight: 400;
        color: var(--ink);
        line-height: 1.7;
        margin-bottom: 20px;
        -webkit-user-select: none !important;
        user-select: none !important;
    }

    .question-pre {
        white-space: pre-wrap;
        font-size: 16px;
        font-family: Consolas, "Cascadia Code", monospace;
        line-height: 1.7;
        color: #172033;
        background: linear-gradient(180deg, #f9fbfc, #f1f6f8);
        padding: 18px;
        border-radius: 8px;
        border: 1px solid var(--line);
        overflow-x: auto;
        margin: 0;
        -webkit-user-select: none !important;
        user-select: none !important;
        -webkit-touch-callout: none !important;
    }

    .stButton>button {
        background: linear-gradient(135deg, var(--accent), #0f766e);
        color: white;
        border-radius: 8px;
        padding: 0.72rem 1.15rem;
        border: 1px solid rgba(255, 255, 255, 0.16);
        font-weight: 800;
        width: 100%;
        box-shadow: 0 10px 24px rgba(37, 99, 235, 0.20);
        transition: transform 160ms ease, box-shadow 160ms ease, filter 160ms ease;
    }

    .stButton>button:hover {
        background: linear-gradient(135deg, var(--accent-strong), #0d9488);
        color: white;
        transform: translateY(-1px);
        box-shadow: 0 14px 28px rgba(15, 118, 110, 0.23);
        filter: saturate(1.04);
    }

    .stButton>button:active {
        transform: translateY(0);
    }

    .stButton>button:disabled {
        background: #cbd5e1;
        color: #64748b;
        border-color: #cbd5e1;
        box-shadow: none;
    }

    div[role="radiogroup"] > label {
        display: flex !important;
        align-items: center !important;
        gap: 14px !important;
        width: 100% !important;
        max-width: 100% !important;
        box-sizing: border-box !important;
        margin-bottom: 14px !important;
        padding: 14px 15px !important;
        border-radius: 8px;
        border: 1px solid var(--line);
        background-color: var(--surface-strong);
        font-size: 15px;
        font-weight: 500;
        color: var(--ink);
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.04);
        transition: transform 150ms ease, border-color 150ms ease, background 150ms ease;
    }

    div[role="radiogroup"] > label:hover {
        background-color: #f7fbfc;
        border-color: rgba(15, 118, 110, 0.38);
        transform: translateX(2px);
    }

    [data-baseweb="input"] input,
    [data-baseweb="select"] > div,
    textarea {
        border-radius: 8px !important;
        border-color: var(--line) !important;
        transition: box-shadow 160ms ease, border-color 160ms ease;
    }

    [data-baseweb="input"] input:focus,
    textarea:focus {
        border-color: rgba(37, 99, 235, 0.7) !important;
        box-shadow: 0 0 0 4px rgba(37, 99, 235, 0.11) !important;
    }

    textarea {
        font-family: Consolas, "Cascadia Code", monospace !important;
        font-size: 15px !important;
        line-height: 1.6 !important;
    }

    textarea[aria-label="Code Editor"],
    textarea[aria-label="Logic Editor"] {
        background: var(--code-bg) !important;
        color: #dbeafe !important;
        border-color: #1f2a44 !important;
        box-shadow: inset 44px 0 0 rgba(255, 255, 255, 0.035);
    }

    textarea[aria-label="Code Editor"] {
        min-height: 460px;
    }

    textarea[aria-label="Code Editor"]::selection,
    textarea[aria-label="Logic Editor"]::selection {
        background: rgba(96, 165, 250, 0.35);
    }

    pre {
        border-radius: 8px !important;
    }

    .run-output {
        background: var(--code-bg);
        color: #f9fafb;
        padding: 16px;
        border-radius: 8px;
        white-space: pre-wrap;
        font-family: Consolas, "Cascadia Code", monospace;
        font-size: 14px;
        line-height: 1.6;
        min-height: 80px;
        border: 1px solid #1f2a44;
    }

    [data-testid="stProgress"] > div > div > div {
        background: linear-gradient(90deg, var(--teal), var(--accent));
    }

    [data-testid="stAlert"] {
        border-radius: 8px;
        border: 1px solid rgba(219, 228, 234, 0.9);
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.05);
    }

    @media (max-width: 900px) {
        .assessment-topbar {
            grid-template-columns: 1fr;
        }

        .score-strip {
            justify-content: stretch;
        }

        .score-card {
            flex: 1;
        }
    }

    </style>
    """,
    unsafe_allow_html=True
)

# ---------------------------------------------------
# SESSION STATE
# ---------------------------------------------------
default_values = {
    "page": "login",
    "student_id": "",
    "student_email": "",
    "pending_student_email": "",
    "student_verification_code": "",
    "student_verification_expires_at": "",
    "student_verification_sent": False,
    "student_verification_notice": "",
    "admin_email": "",
    "pending_admin_email": "",
    "admin_verification_code": "",
    "admin_verification_expires_at": "",
    "admin_verification_sent": False,
    "admin_verification_notice": "",
    "full_name": "",
    "selected_course": "",
    "selected_nsti": "",
    "location": "",
    "assessment_date": date.today(),
    "selected_assessment": "",
    "selected_sheet": "",
    "assessment_df": pd.DataFrame(),
    "current_question": 0,
    "score": 0,
    "answers": {},
    "code_answers": {},
    "correct_questions": set(),
    "code_drafts": {},
    "run_outputs": {},
    "test_results": {},
    "test_passed": {},
    "question_time_spent": {},
    "active_question_index": None,
    "active_level_filter": "All Levels",
    "question_started_at": "",
    "code_run_counts": {},
    "code_failed_attempts": {},
    "submission_timestamps": {},
    "proctoring_violations": [],
    "processed_proctor_events": [],
    "last_proctor_alert": "",
    "last_answer_feedback": {},
    "code_snapshots": {},
    "selection_warnings": [],
    "timer_started_at": "",
    "timer_ends_at": "",
    "timer_duration_minutes": 0,
    "assessment_submitted": False,
    "submitted_at": ""
}

for key, value in default_values.items():

    if key not in st.session_state:

        st.session_state[key] = value

# ---------------------------------------------------
# FUNCTIONS
# ---------------------------------------------------
def logout():

    for key in list(st.session_state.keys()):

        del st.session_state[key]

def is_valid_email(email):

    pattern = r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'

    return re.match(pattern, email)

def is_valid_name(value):

    return re.fullmatch(r"[A-Za-z ]+", str(value).strip()) is not None

def is_valid_admin_email_prefix(value):

    return re.fullmatch(r"[A-Za-z0-9._%+-]+", str(value).strip()) is not None

def get_config_value(name, default=""):

    value = os.getenv(name)

    if value:

        return value

    try:

        return st.secrets.get(name, default)

    except Exception:

        return default

def generate_verification_code():

    return f"{random.randint(100000, 999999)}"

def send_verification_email(recipient_email, verification_code, role_label):

    smtp_host = get_config_value("SMTP_HOST")
    smtp_port = int(get_config_value("SMTP_PORT", "587"))
    smtp_username = get_config_value("SMTP_USERNAME")
    smtp_password = get_config_value("SMTP_PASSWORD")
    smtp_from = get_config_value(
        "SMTP_FROM_EMAIL",
        smtp_username or "no-reply@nsti-assessment.local"
    )
    smtp_use_tls = str(
        get_config_value("SMTP_USE_TLS", "true")
    ).lower() != "false"

    if not smtp_host or not smtp_username or not smtp_password:

        return False, (
            "Email service is not configured. For local testing, use this "
            f"verification code: {verification_code}"
        )

    message = EmailMessage()
    message["Subject"] = "NSTI Assessment LMS verification code"
    message["From"] = smtp_from
    message["To"] = recipient_email
    message.set_content(
        f"""
Your NSTI Assessment LMS {role_label} verification code is:

{verification_code}

This code is valid for 10 minutes. If you did not request this code, you can ignore this email.
        """.strip()
    )

    try:

        with smtplib.SMTP(smtp_host, smtp_port, timeout=15) as server:

            if smtp_use_tls:

                server.starttls()

            server.login(smtp_username, smtp_password)
            server.send_message(message)

        return True, "Verification code sent to your email."

    except Exception as e:

        return False, f"Could not send verification email: {e}"

def send_student_verification_email(student_email, verification_code):

    return send_verification_email(
        student_email,
        verification_code,
        "student"
    )

def send_admin_verification_email(admin_email, verification_code):

    return send_verification_email(
        admin_email,
        verification_code,
        "administrator"
    )

def safe_file_name(value):

    return re.sub(
        r"[^A-Za-z0-9_.-]+",
        "_",
        str(value).strip()
    )

def progress_file_path(student_id, student_email, assessment_file):

    file_name = "_".join(
        [
            safe_file_name(student_id),
            safe_file_name(student_email),
            safe_file_name(assessment_file)
        ]
    )

    return os.path.join(
        "assessment_progress",
        f"{file_name}.json"
    )

def metadata_file_path(assessment_file):

    return os.path.join(
        "assessment_metadata",
        f"{safe_file_name(assessment_file)}.json"
    )

def save_assessment_metadata(assessment_file, metadata):

    with open(
        metadata_file_path(assessment_file),
        "w",
        encoding="utf-8"
    ) as metadata_file:

        json.dump(metadata, metadata_file, indent=2, default=str)

def load_assessment_metadata(assessment_file):

    path = metadata_file_path(assessment_file)

    if not os.path.exists(path):

        return {}

    try:

        with open(path, "r", encoding="utf-8") as metadata_file:

            return json.load(metadata_file)

    except Exception:

        return {}

def delete_assessment_artifacts(assessment_file, delete_progress=False):

    deleted_items = []
    assessment_path = os.path.join("assessment_files", assessment_file)

    if os.path.exists(assessment_path):

        os.remove(assessment_path)
        deleted_items.append(assessment_file)

    metadata_path = metadata_file_path(assessment_file)

    if os.path.exists(metadata_path):

        os.remove(metadata_path)
        deleted_items.append(os.path.basename(metadata_path))

    if delete_progress:

        for progress_file in Path("assessment_progress").glob("*.json"):

            try:

                with open(progress_file, "r", encoding="utf-8") as file:

                    progress_data = json.load(file)

                if progress_data.get("assessment") == assessment_file:

                    progress_file.unlink()
                    deleted_items.append(progress_file.name)

            except Exception:

                continue

    return deleted_items

def parse_datetime(value):

    if not value:

        return None

    try:

        return datetime.fromisoformat(value)

    except ValueError:

        return None

def format_datetime(value):

    parsed_value = parse_datetime(value)

    if parsed_value is None:

        return "Not set"

    return parsed_value.strftime("%d-%m-%Y %I:%M %p")

def get_assessment_availability(metadata):

    now = datetime.now()
    start_at = parse_datetime(metadata.get("start_at"))
    end_at = parse_datetime(metadata.get("end_at"))

    if start_at and now < start_at:

        return False, f"Starts on {format_datetime(metadata.get('start_at'))}"

    if end_at and now > end_at:

        return False, f"Ended on {format_datetime(metadata.get('end_at'))}"

    return True, "Available"

def format_duration(minutes):

    try:

        minutes = int(minutes)

    except (TypeError, ValueError):

        return "No time limit"

    if minutes <= 0:

        return "No time limit"

    hours = minutes // 60
    remaining_minutes = minutes % 60

    if hours and remaining_minutes:

        return f"{hours} hr {remaining_minutes} min"

    if hours:

        return f"{hours} hr"

    return f"{remaining_minutes} min"

def format_remaining_time(seconds):

    if seconds is None:

        return "No time limit"

    seconds = max(0, int(seconds))
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    remaining_seconds = seconds % 60

    if hours:

        return f"{hours:02d}:{minutes:02d}:{remaining_seconds:02d}"

    return f"{minutes:02d}:{remaining_seconds:02d}"

def now_iso():

    return datetime.now().isoformat(timespec="seconds")

def get_query_param_value(name, default=""):

    try:

        value = st.query_params.get(name, default)

    except Exception:

        try:

            value = st.experimental_get_query_params().get(name, default)

        except Exception:

            return default

    if isinstance(value, list):

        return value[0] if value else default

    return value

def update_question_time():

    question_index = st.session_state.get("active_question_index")
    started_at = parse_datetime(st.session_state.get("question_started_at", ""))

    if question_index is None or started_at is None:

        return

    elapsed_seconds = int((datetime.now() - started_at).total_seconds())

    if elapsed_seconds < 0 or elapsed_seconds > 12 * 60 * 60:

        elapsed_seconds = 0

    key = str(question_index)
    current_total = int(
        st.session_state.question_time_spent.get(key, 0) or 0
    )
    st.session_state.question_time_spent[key] = current_total + elapsed_seconds
    st.session_state.question_started_at = now_iso()

def begin_question_tracking(question_index):

    active_question_index = st.session_state.get("active_question_index")

    if active_question_index is None:

        st.session_state.active_question_index = question_index
        st.session_state.question_started_at = now_iso()
        return

    if active_question_index != question_index:

        update_question_time()
        st.session_state.active_question_index = question_index
        st.session_state.question_started_at = now_iso()

def navigate_to_question(question_index):

    update_question_time()
    st.session_state.current_question = question_index
    st.session_state.active_question_index = question_index
    st.session_state.question_started_at = now_iso()

def get_question_level(row):

    return str(get_row_value(row, "Level", "Level 1")).strip() or "Level 1"

def get_level_filter_options(df):

    available_levels = [
        level
        for level in LEVEL_LABELS
        if level in set(df["Level"].astype(str))
    ]

    return ["All Levels"] + available_levels

def get_visible_question_indices(df, level_filter):

    if level_filter == "All Levels":

        return list(range(len(df)))

    return [
        index
        for index, row in df.iterrows()
        if get_question_level(row) == level_filter
    ]

def get_next_visible_question(current_index, visible_indices):

    if current_index not in visible_indices:

        return visible_indices[0] if visible_indices else None

    visible_position = visible_indices.index(current_index)

    if visible_position + 1 >= len(visible_indices):

        return None

    return visible_indices[visible_position + 1]

def get_previous_visible_question(current_index, visible_indices):

    if current_index not in visible_indices:

        return visible_indices[0] if visible_indices else None

    visible_position = visible_indices.index(current_index)

    if visible_position == 0:

        return None

    return visible_indices[visible_position - 1]

def get_completed_levels(df):

    completed_levels = set()

    for level in LEVEL_LABELS:

        level_indices = get_visible_question_indices(df, level)

        if level_indices and all(
            question_index in st.session_state.answers
            for question_index in level_indices
        ):

            completed_levels.add(level)

    return completed_levels

def get_first_unanswered_question(df, level_filter):

    visible_indices = get_visible_question_indices(df, level_filter)

    for question_index in visible_indices:

        if question_index not in st.session_state.answers:

            return question_index

    return visible_indices[0] if visible_indices else None

def get_next_incomplete_level(df, current_level="All Levels"):

    completed_levels = get_completed_levels(df)

    if current_level in LEVEL_LABELS:

        start_index = LEVEL_LABELS.index(current_level) + 1

    else:

        start_index = 0

    ordered_levels = LEVEL_LABELS[start_index:] + LEVEL_LABELS[:start_index]

    for level in ordered_levels:

        level_indices = get_visible_question_indices(df, level)

        if level_indices and level not in completed_levels:

            return level

    return None

def render_level_selector(df):

    available_levels = [
        level
        for level in LEVEL_LABELS
        if get_visible_question_indices(df, level)
    ]
    completed_levels = get_completed_levels(df)

    st.sidebar.markdown("**Question Level**")

    if st.sidebar.button(
        "All Levels",
        key="level_filter_all",
        type=(
            "primary"
            if st.session_state.active_level_filter == "All Levels"
            else "secondary"
        )
    ):

        st.session_state.active_level_filter = "All Levels"
        target_question = get_first_unanswered_question(df, "All Levels")

        if target_question is not None:

            navigate_to_question(target_question)

        st.rerun()

    for level in available_levels:

        is_completed = level in completed_levels
        label = f"{level} - Completed" if is_completed else level

        if st.sidebar.button(
            label,
            key=f"level_filter_{level}",
            disabled=is_completed,
            type=(
                "primary"
                if st.session_state.active_level_filter == level
                and not is_completed
                else "secondary"
            )
        ):

            st.session_state.active_level_filter = level
            target_question = get_first_unanswered_question(df, level)

            if target_question is not None:

                navigate_to_question(target_question)

            st.rerun()

    if (
        st.session_state.active_level_filter in completed_levels
        and st.session_state.active_level_filter != "All Levels"
    ):

        next_level = get_next_incomplete_level(
            df,
            st.session_state.active_level_filter
        )
        st.session_state.active_level_filter = next_level or "All Levels"

    return st.session_state.active_level_filter

def get_auto_advance_target(df, current_index, visible_indices, current_level):

    next_question = get_next_visible_question(current_index, visible_indices)

    if next_question is not None:

        return current_level, next_question

    if current_level != "All Levels":

        next_level = get_next_incomplete_level(df, current_level)

        if next_level:

            return next_level, get_first_unanswered_question(df, next_level)

    return "All Levels", None

def increment_counter(counter_name, question_index, amount=1):

    key = str(question_index)
    counter = st.session_state[counter_name]
    counter[key] = int(counter.get(key, 0) or 0) + amount

def format_seconds(total_seconds):

    total_seconds = int(total_seconds or 0)
    minutes = total_seconds // 60
    seconds = total_seconds % 60

    if minutes:

        return f"{minutes}m {seconds}s"

    return f"{seconds}s"

def format_question_time_spent(time_spent):

    if not time_spent:

        return ""

    items = []

    for key in sorted(time_spent, key=lambda value: int(value)):

        items.append(f"Q{int(key) + 1}: {format_seconds(time_spent[key])}")

    return "; ".join(items)

def format_question_timestamps(timestamps):

    if not timestamps:

        return ""

    items = []

    for key in sorted(timestamps, key=lambda value: int(value)):

        items.append(f"Q{int(key) + 1}: {timestamps[key]}")

    return "; ".join(items)

def log_proctoring_violation(
    event_type,
    question_index,
    details="",
    event_id=None
):

    event_id = event_id or f"{event_type}_{now_iso()}"

    if event_id in st.session_state.processed_proctor_events:

        return

    st.session_state.processed_proctor_events.append(event_id)
    st.session_state.processed_proctor_events = (
        st.session_state.processed_proctor_events[-100:]
    )
    st.session_state.proctoring_violations.append(
        {
            "event_id": event_id,
            "event_type": event_type,
            "question_index": question_index,
            "details": details,
            "timestamp": now_iso()
        }
    )
    alert_messages = {
        "copy": "Copy-paste disabled.",
        "cut": "Copy-paste disabled.",
        "paste": "Copy-paste disabled.",
        "clipboard_shortcut": "Copy-paste disabled.",
        "right_click": "Right-click is disabled.",
        "window_focus_loss": "Change of tab is not allowed.",
        "tab_switch": "Change of tab is not allowed."
    }
    st.session_state.last_proctor_alert = alert_messages.get(
        event_type,
        "Suspicious activity detected."
    )

def process_proctoring_query_event():

    event_id = get_query_param_value("proctor_event_id")

    if not event_id:

        return

    event_type = get_query_param_value("proctor_event_type", "unknown")
    question_index = get_query_param_value(
        "proctor_question",
        st.session_state.current_question
    )
    details = get_query_param_value("proctor_details", "")

    try:

        question_index = int(question_index)

    except (TypeError, ValueError):

        question_index = st.session_state.current_question

    log_proctoring_violation(
        event_type,
        question_index,
        details,
        event_id
    )
    save_assessment_progress()

    try:

        for param_name in [
            "proctor_event_id",
            "proctor_event_type",
            "proctor_question",
            "proctor_details"
        ]:

            if param_name in st.query_params:

                del st.query_params[param_name]

    except Exception:

        pass

def get_inserted_text(current_text, previous_text):

    current_text = str(current_text or "")
    previous_text = str(previous_text or "")

    if current_text == previous_text or len(current_text) <= len(previous_text):

        return ""

    prefix_length = 0
    max_prefix_length = min(len(current_text), len(previous_text))

    while (
        prefix_length < max_prefix_length
        and current_text[prefix_length] == previous_text[prefix_length]
    ):

        prefix_length += 1

    suffix_length = 0
    max_suffix_length = min(
        len(current_text) - prefix_length,
        len(previous_text) - prefix_length
    )

    while (
        suffix_length < max_suffix_length
        and current_text[-(suffix_length + 1)]
        == previous_text[-(suffix_length + 1)]
    ):

        suffix_length += 1

    end_index = len(current_text) - suffix_length if suffix_length else len(current_text)
    return current_text[prefix_length:end_index]

def is_probable_paste(current_text, previous_text):

    inserted_text = get_inserted_text(current_text, previous_text)
    stripped_insert = inserted_text.strip()

    if not stripped_insert:

        return False

    code_markers = [
        "def ",
        "class ",
        "import ",
        "return ",
        "print(",
        "console.",
        "function ",
        "=>",
        "for ",
        "while ",
        "if ",
        "{",
        "}",
        ";",
        "</",
        "<script",
        "SELECT ",
        "CALCULATE(",
        "SUM("
    ]
    marker_hits = sum(
        1
        for marker in code_markers
        if marker.lower() in stripped_insert.lower()
    )

    return (
        len(stripped_insert) > 60
        or ("\n" in inserted_text and len(stripped_insert) > 8)
        or ("\t" in inserted_text and len(stripped_insert) > 12)
        or ("    " in inserted_text and len(stripped_insert) > 12)
        or (len(stripped_insert) > 24 and marker_hits >= 1)
        or (len(stripped_insert) > 16 and marker_hits >= 2)
    )

def render_proctoring_guard(metadata, question_index):

    settings = {
        "enabled": True,
        "disable_clipboard": True,
        "prevent_right_click": True,
        "detect_focus_loss": True
    }
    settings.update(metadata.get("proctoring", {}))

    if not settings.get("enabled", False):

        return

    disable_clipboard = bool(settings.get("disable_clipboard", True))
    prevent_right_click = bool(settings.get("prevent_right_click", True))
    detect_focus_loss = bool(settings.get("detect_focus_loss", True))

    components.html(
        f"""
        <script>
        (function() {{
            const parentWindow = window.parent || window;
            const doc = parentWindow.document || document;
            const hookVersion = "2026-05-19-v7";
            parentWindow.__edunetProctorContext = {{
                question: {int(question_index)}
            }};

            if (parentWindow.__edunetProctorController) {{
                try {{
                    parentWindow.__edunetProctorController.abort();
                }} catch (error) {{}}
            }}
            const controller = new AbortController();
            const listenerOptions = {{
                capture: true,
                signal: controller.signal
            }};
            parentWindow.__edunetProctorController = controller;
            parentWindow.__edunetProctorVersion = hookVersion;

            function showAlert(message, type) {{
                const now = Date.now();
                const key = "__edunetAlertLast_" + (type || message);
                if (
                    parentWindow.__edunetAlertOpen
                    || (
                        parentWindow[key]
                        && now - parentWindow[key] < 8000
                    )
                ) {{
                    return;
                }}
                parentWindow[key] = now;
                parentWindow.__edunetAlertOpen = true;
                try {{
                    parentWindow.alert(message);
                }} catch (error) {{
                    alert(message);
                }} finally {{
                    setTimeout(function() {{
                        parentWindow.__edunetAlertOpen = false;
                    }}, 1200);
                }}
            }}

            function sendEvent(type, details) {{
                const now = Date.now();
                const throttleKey = "__edunetProctorLast_" + type;
                if (
                    parentWindow[throttleKey]
                    && now - parentWindow[throttleKey] < 8000
                ) {{
                    return;
                }}
                parentWindow[throttleKey] = now;
                const params = new URLSearchParams(parentWindow.location.search);
                params.set("proctor_event_id", String(now) + "_" + type);
                params.set("proctor_event_type", type);
                params.set(
                    "proctor_question",
                    String(parentWindow.__edunetProctorContext.question)
                );
                params.set("proctor_details", details || "");
                parentWindow.location.search = params.toString();
            }}

            function blockAction(event, type, message, details) {{
                if (event && event.preventDefault) {{
                    event.preventDefault();
                }}
                if (event && event.stopPropagation) {{
                    event.stopPropagation();
                }}
                showAlert(message, type);
                sendEvent(type, details || message);
                return false;
            }}

            const blockClipboard = {str(disable_clipboard).lower()};
            const blockRightClick = {str(prevent_right_click).lower()};
            const watchFocus = {str(detect_focus_loss).lower()};

            if (blockClipboard) {{
                ["copy", "cut", "paste"].forEach(function(eventName) {{
                    doc.addEventListener(eventName, function(event) {{
                        return blockAction(
                            event,
                            eventName,
                            "Copy-paste disabled.",
                            "Clipboard action blocked during assessment"
                        );
                    }}, listenerOptions);
                }});

                doc.addEventListener("keydown", function(event) {{
                    const key = String(event.key || "").toLowerCase();
                    const blockedCombo = (
                        (event.ctrlKey || event.metaKey)
                        && ["c", "x", "v", "insert"].includes(key)
                    ) || (
                        event.shiftKey && key === "insert"
                    );
                    if (blockedCombo) {{
                        return blockAction(
                            event,
                            "clipboard_shortcut",
                            "Copy-paste disabled.",
                            "Clipboard shortcut blocked during assessment"
                        );
                    }}
                }}, listenerOptions);

                doc.addEventListener("beforeinput", function(event) {{
                    const inputType = String(event.inputType || "").toLowerCase();
                    if (
                        inputType === "insertfrompaste"
                        || inputType === "insertfromdrop"
                        || inputType === "insertfromyank"
                        || inputType.indexOf("paste") >= 0
                    ) {{
                        return blockAction(
                            event,
                            "paste",
                            "Copy-paste disabled.",
                            "Paste/drop input blocked during assessment"
                        );
                    }}
                }}, listenerOptions);

                doc.addEventListener("drop", function(event) {{
                    return blockAction(
                        event,
                        "paste",
                        "Copy-paste disabled.",
                        "Drag/drop paste blocked during assessment"
                    );
                }}, listenerOptions);

                doc.addEventListener("selectstart", function(event) {{
                    const target = event.target;
                    const insideQuestion = (
                        target
                        && target.closest
                        && target.closest(".question-text, .question-pre")
                    );
                    if (!insideQuestion) {{
                        return true;
                    }}
                    return blockAction(
                        event,
                        "copy",
                        "Copy-paste disabled.",
                        "Question text selection blocked during assessment"
                    );
                }}, listenerOptions);

                const hardenTextareas = function() {{
                    doc.querySelectorAll("textarea").forEach(function(textarea) {{
                        textarea.setAttribute("autocomplete", "off");
                        textarea.setAttribute("autocorrect", "off");
                        textarea.setAttribute("spellcheck", "false");
                        textarea.setAttribute("data-proctor-locked", "true");
                        textarea.onpaste = function(event) {{
                            return blockAction(
                                event,
                                "paste",
                                "Copy-paste disabled.",
                                "Direct textarea paste blocked"
                            );
                        }};
                        textarea.oncopy = function(event) {{
                            return blockAction(
                                event,
                                "copy",
                                "Copy-paste disabled.",
                                "Direct textarea copy blocked"
                            );
                        }};
                        textarea.oncut = function(event) {{
                            return blockAction(
                                event,
                                "cut",
                                "Copy-paste disabled.",
                                "Direct textarea cut blocked"
                            );
                        }};
                    }});
                }};
                hardenTextareas();
                const observer = new MutationObserver(hardenTextareas);
                observer.observe(doc.body, {{
                    childList: true,
                    subtree: true
                }});
                controller.signal.addEventListener("abort", function() {{
                    observer.disconnect();
                }});
            }}

            if (blockRightClick) {{
                doc.addEventListener("contextmenu", function(event) {{
                    return blockAction(
                        event,
                        "right_click",
                        "Right-click is disabled.",
                        "Right click blocked during assessment"
                    );
                }}, listenerOptions);
            }}

            if (watchFocus) {{
                function reportFocusLoss(type, details) {{
                    const now = Date.now();
                    if (
                        parentWindow.__edunetAlertOpen
                        || (
                            parentWindow.__edunetFocusLast
                            && now - parentWindow.__edunetFocusLast < 8000
                        )
                    ) {{
                        return;
                    }}
                    parentWindow.__edunetFocusLast = now;
                    showAlert(
                        "Change of tab is not allowed.",
                        type
                    );
                    sendEvent(type, details);
                }}
                parentWindow.addEventListener("blur", function() {{
                    reportFocusLoss(
                        "window_focus_loss",
                        "Assessment window lost focus"
                    );
                }}, {{
                    signal: controller.signal
                }});
                doc.addEventListener("visibilitychange", function() {{
                    if (doc.hidden) {{
                        reportFocusLoss(
                            "tab_switch",
                            "Browser tab became hidden"
                        );
                    }}
                }}, {{
                    signal: controller.signal
                }});
            }}
        }})();
        </script>
        """,
        height=0
    )

def render_live_timer(remaining_seconds, duration_minutes):

    if remaining_seconds is None:

        st.info("No assessment time limit is set.")
        return

    total_seconds = max(1, int(duration_minutes or 0) * 60)

    if total_seconds <= 1:

        total_seconds = max(1, int(remaining_seconds))

    components.html(
        f"""
        <style>
            @keyframes timerPulse {{
                0%, 100% {{ box-shadow: 0 0 0 0 rgba(45, 212, 191, 0.30); }}
                50% {{ box-shadow: 0 0 0 10px rgba(45, 212, 191, 0); }}
            }}
            @keyframes clockSweep {{
                from {{ transform: rotate(0deg); }}
                to {{ transform: rotate(360deg); }}
            }}
            .timer-card {{
                background: linear-gradient(160deg, #111827 0%, #172033 56%, #0f766e 140%);
                color: #f9fafb;
                padding: 15px 16px;
                border-radius: 10px;
                border: 1px solid rgba(148, 163, 184, 0.25);
                font-family: Inter, Arial, sans-serif;
                box-sizing: border-box;
            }}
            .timer-head {{
                display: flex;
                align-items: center;
                gap: 10px;
                margin-bottom: 12px;
            }}
            .clock-icon {{
                position: relative;
                width: 34px;
                height: 34px;
                border: 2px solid #5eead4;
                border-radius: 50%;
                background: rgba(45, 212, 191, 0.10);
                animation: timerPulse 2.2s ease-in-out infinite;
                flex: 0 0 auto;
            }}
            .clock-icon::before {{
                content: "";
                position: absolute;
                width: 2px;
                height: 10px;
                background: #f8fafc;
                left: 15px;
                top: 7px;
                border-radius: 999px;
                transform-origin: 1px 10px;
                animation: clockSweep 12s linear infinite;
            }}
            .clock-icon::after {{
                content: "";
                position: absolute;
                width: 8px;
                height: 2px;
                background: #93c5fd;
                left: 15px;
                top: 16px;
                border-radius: 999px;
                transform-origin: 0 1px;
            }}
            .timer-label {{
                color: #cbd5e1;
                font-size: 11px;
                font-weight: 800;
                letter-spacing: 0.08em;
                text-transform: uppercase;
            }}
            #timer-value {{
                font-size: 25px;
                line-height: 1.1;
                font-weight: 800;
                letter-spacing: 0;
                margin-top: 2px;
            }}
            .timer-meta {{
                display: flex;
                justify-content: space-between;
                gap: 10px;
                color: #cbd5e1;
                font-size: 12px;
                margin-top: 10px;
            }}
            .timer-meta strong {{
                color: #f8fafc;
            }}
            .timer-track {{
                background: rgba(148, 163, 184, 0.30);
                height: 8px;
                border-radius: 999px;
                margin-top: 13px;
                overflow: hidden;
            }}
            #timer-bar {{
                background: linear-gradient(90deg, #2dd4bf, #60a5fa);
                height: 8px;
                width: {max(0, min(100, int((remaining_seconds / total_seconds) * 100)))}%;
                border-radius: 999px;
                transition: width 0.4s ease, background 0.4s ease;
            }}
        </style>
        <div class="timer-card">
            <div class="timer-head">
                <div class="clock-icon"></div>
                <div>
                    <div class="timer-label">Time Remaining</div>
                    <div id="timer-value">{format_remaining_time(remaining_seconds)}</div>
                </div>
            </div>
            <div class="timer-meta">
                <span>Total <strong>{format_duration(duration_minutes)}</strong></span>
                <span>Live timer</span>
            </div>
            <div class="timer-track">
                <div id="timer-bar"></div>
            </div>
        </div>
        <script>
            let remaining = {int(remaining_seconds)};
            const total = {int(total_seconds)};
            const value = document.getElementById("timer-value");
            const bar = document.getElementById("timer-bar");

            function formatTime(seconds) {{
                seconds = Math.max(0, seconds);
                const hours = Math.floor(seconds / 3600);
                const minutes = Math.floor((seconds % 3600) / 60);
                const secs = seconds % 60;
                if (hours > 0) {{
                    return String(hours).padStart(2, "0") + ":" +
                           String(minutes).padStart(2, "0") + ":" +
                           String(secs).padStart(2, "0");
                }}
                return String(minutes).padStart(2, "0") + ":" +
                       String(secs).padStart(2, "0");
            }}

            setInterval(() => {{
                remaining -= 1;
                value.textContent = formatTime(remaining);
                const percent = Math.max(0, Math.min(100, (remaining / total) * 100));
                bar.style.width = percent + "%";
                if (remaining <= 300) {{
                    bar.style.background = "#ef4444";
                }} else if (remaining <= 600) {{
                    bar.style.background = "#f59e0b";
                }}
            }}, 1000);
        </script>
        """,
        height=155
    )

def render_visible_test_cases(test_cases):

    visible_test_cases = [
        test_case
        for test_case in test_cases
        if test_case.get("visible", True)
    ]

    if not visible_test_cases:

        return

    st.markdown("**Test Cases**")

    for index, test_case in enumerate(visible_test_cases, start=1):

        test_input = html.escape(test_case["input"] or "(no input)")
        expected_output = html.escape(test_case["expected"])

        st.markdown(
            f"""
            <div style="
                background:#ffffff;
                border:1px solid #e5e7eb;
                border-radius:12px;
                padding:14px;
                margin-bottom:12px;">
                <div style="font-weight:700; margin-bottom:8px;">
                    Test Case {index}
                </div>
                <div style="font-size:13px; color:#4b5563; margin-bottom:4px;">
                    Input
                </div>
                <pre style="
                    white-space:pre-wrap;
                    background:#f8fafc;
                    padding:10px;
                    border-radius:8px;
                    border:1px solid #e5e7eb;
                    margin:0 0 10px 0;">{test_input}</pre>
                <div style="font-size:13px; color:#4b5563; margin-bottom:4px;">
                    Expected Output
                </div>
                <pre style="
                    white-space:pre-wrap;
                    background:#f8fafc;
                    padding:10px;
                    border-radius:8px;
                    border:1px solid #e5e7eb;
                    margin:0;">{expected_output}</pre>
            </div>
            """,
            unsafe_allow_html=True
        )

def render_revealed_hidden_test_cases(test_results):

    hidden_results = [
        result
        for result in test_results
        if not result.get("visible", True)
    ]

    if not hidden_results:

        return

    all_hidden_passed = all(
        result.get("passed", False)
        for result in hidden_results
    )

    if not all_hidden_passed:

        st.warning(
            "Hidden test case(s) did not pass. Details will be shown after they pass."
        )
        return

    st.markdown("**Hidden Test Cases**")

    for index, result in enumerate(hidden_results, start=1):

        test_input = html.escape(result.get("input") or "(no input)")
        expected_output = html.escape(result.get("expected", ""))
        actual_output = html.escape(normalize_output(result.get("actual", "")))

        st.markdown(
            f"""
            <div style="
                background:#ecfdf5;
                border:1px solid #86efac;
                border-radius:12px;
                padding:14px;
                margin-bottom:12px;">
                <div style="font-weight:800; margin-bottom:8px; color:#166534;">
                    Hidden Test Case {index}: Passed
                </div>
                <div style="font-size:13px; color:#166534; margin-bottom:4px;">
                    Input
                </div>
                <pre style="
                    white-space:pre-wrap;
                    background:#ffffff;
                    padding:10px;
                    border-radius:8px;
                    border:1px solid #bbf7d0;
                    margin:0 0 10px 0;">{test_input}</pre>
                <div style="font-size:13px; color:#166534; margin-bottom:4px;">
                    Expected Output
                </div>
                <pre style="
                    white-space:pre-wrap;
                    background:#ffffff;
                    padding:10px;
                    border-radius:8px;
                    border:1px solid #bbf7d0;
                    margin:0 0 10px 0;">{expected_output}</pre>
                <div style="font-size:13px; color:#166534; margin-bottom:4px;">
                    Actual Output
                </div>
                <pre style="
                    white-space:pre-wrap;
                    background:#ffffff;
                    padding:10px;
                    border-radius:8px;
                    border:1px solid #bbf7d0;
                    margin:0;">{actual_output}</pre>
            </div>
            """,
            unsafe_allow_html=True
        )

def start_assessment_timer(metadata):

    duration_minutes = int(metadata.get("duration_minutes", 0) or 0)

    if duration_minutes <= 0:

        return

    if st.session_state.timer_started_at and st.session_state.timer_ends_at:

        return

    now = datetime.now()
    ends_at = now + timedelta(minutes=duration_minutes)

    st.session_state.timer_started_at = now.isoformat(timespec="seconds")
    st.session_state.timer_ends_at = ends_at.isoformat(timespec="seconds")
    st.session_state.timer_duration_minutes = duration_minutes

    save_assessment_progress()

def get_remaining_seconds():

    if not st.session_state.timer_ends_at:

        return None

    ends_at = parse_datetime(st.session_state.timer_ends_at)

    if ends_at is None:

        return None

    return max(0, int((ends_at - datetime.now()).total_seconds()))

def convert_keys_to_int(value):

    if not isinstance(value, dict):

        return {}

    return {
        int(key): item
        for key, item in value.items()
        if str(key).isdigit()
    }

def convert_keys_to_str(value):

    if not isinstance(value, dict):

        return {}

    return {
        str(key): item
        for key, item in value.items()
    }

def save_assessment_progress():

    if (
        st.session_state.student_id == ""
        or st.session_state.student_email == ""
        or st.session_state.selected_assessment == ""
        or st.session_state.assessment_df.empty
    ):

        return

    if st.session_state.get("page") == "take_assessment":

        update_question_time()

    progress_path = progress_file_path(
        st.session_state.student_id,
        st.session_state.student_email,
        st.session_state.selected_assessment
    )

    progress_data = {
        "student_id": st.session_state.student_id,
        "student_email": st.session_state.student_email,
        "full_name": st.session_state.full_name,
        "course": st.session_state.selected_course,
        "nsti": st.session_state.selected_nsti,
        "location": st.session_state.location,
        "assessment_date": str(st.session_state.assessment_date),
        "assessment": st.session_state.selected_assessment,
        "current_question": st.session_state.current_question,
        "answers": st.session_state.answers,
        "code_answers": st.session_state.code_answers,
        "code_drafts": st.session_state.code_drafts,
        "run_outputs": st.session_state.run_outputs,
        "test_results": st.session_state.test_results,
        "test_passed": st.session_state.test_passed,
        "question_time_spent": st.session_state.question_time_spent,
        "active_question_index": st.session_state.active_question_index,
        "active_level_filter": st.session_state.active_level_filter,
        "question_started_at": st.session_state.question_started_at,
        "code_run_counts": st.session_state.code_run_counts,
        "code_failed_attempts": st.session_state.code_failed_attempts,
        "submission_timestamps": st.session_state.submission_timestamps,
        "proctoring_violations": st.session_state.proctoring_violations,
        "last_proctor_alert": st.session_state.last_proctor_alert,
        "last_answer_feedback": st.session_state.last_answer_feedback,
        "code_snapshots": st.session_state.code_snapshots,
        "selection_warnings": st.session_state.selection_warnings,
        "timer_started_at": st.session_state.timer_started_at,
        "timer_ends_at": st.session_state.timer_ends_at,
        "timer_duration_minutes": st.session_state.timer_duration_minutes,
        "assessment_submitted": st.session_state.assessment_submitted,
        "submitted_at": st.session_state.submitted_at,
        "assessment_records": (
            st.session_state.assessment_df
            .where(pd.notna(st.session_state.assessment_df), None)
            .to_dict("records")
        )
    }

    with open(progress_path, "w", encoding="utf-8") as progress_file:

        json.dump(
            progress_data,
            progress_file,
            indent=2,
            default=str
        )

def load_assessment_progress(student_id, student_email, assessment_file):

    progress_path = progress_file_path(
        student_id,
        student_email,
        assessment_file
    )

    if not os.path.exists(progress_path):

        return None

    try:

        with open(progress_path, "r", encoding="utf-8") as progress_file:

            progress_data = json.load(progress_file)

        progress_data["answers"] = convert_keys_to_int(
            progress_data.get("answers", {})
        )
        progress_data["code_answers"] = convert_keys_to_int(
            progress_data.get("code_answers", {})
        )
        progress_data["code_drafts"] = convert_keys_to_int(
            progress_data.get("code_drafts", {})
        )
        progress_data["run_outputs"] = convert_keys_to_int(
            progress_data.get("run_outputs", {})
        )
        progress_data["test_results"] = convert_keys_to_int(
            progress_data.get("test_results", {})
        )
        progress_data["test_passed"] = convert_keys_to_int(
            progress_data.get("test_passed", {})
        )
        progress_data["question_time_spent"] = convert_keys_to_str(
            progress_data.get("question_time_spent", {})
        )
        progress_data["code_run_counts"] = convert_keys_to_str(
            progress_data.get("code_run_counts", {})
        )
        progress_data["code_failed_attempts"] = convert_keys_to_str(
            progress_data.get("code_failed_attempts", {})
        )
        progress_data["submission_timestamps"] = convert_keys_to_str(
            progress_data.get("submission_timestamps", {})
        )

        return progress_data

    except Exception:

        return None

def apply_assessment_progress(progress_data):

    st.session_state.selected_course = progress_data.get(
        "course",
        st.session_state.selected_course
    )
    st.session_state.current_question = progress_data.get(
        "current_question",
        0
    )
    st.session_state.answers = progress_data.get("answers", {})
    st.session_state.code_answers = progress_data.get("code_answers", {})
    st.session_state.code_drafts = progress_data.get("code_drafts", {})
    st.session_state.run_outputs = progress_data.get("run_outputs", {})
    st.session_state.test_results = progress_data.get("test_results", {})
    st.session_state.test_passed = progress_data.get("test_passed", {})
    st.session_state.question_time_spent = progress_data.get(
        "question_time_spent",
        {}
    )
    st.session_state.active_question_index = progress_data.get(
        "active_question_index",
        None
    )
    st.session_state.active_level_filter = progress_data.get(
        "active_level_filter",
        "All Levels"
    )
    st.session_state.question_started_at = progress_data.get(
        "question_started_at",
        ""
    )
    st.session_state.code_run_counts = progress_data.get(
        "code_run_counts",
        {}
    )
    st.session_state.code_failed_attempts = progress_data.get(
        "code_failed_attempts",
        {}
    )
    st.session_state.submission_timestamps = progress_data.get(
        "submission_timestamps",
        {}
    )
    st.session_state.proctoring_violations = progress_data.get(
        "proctoring_violations",
        []
    )
    st.session_state.last_proctor_alert = progress_data.get(
        "last_proctor_alert",
        ""
    )
    st.session_state.last_answer_feedback = progress_data.get(
        "last_answer_feedback",
        {}
    )
    st.session_state.code_snapshots = progress_data.get(
        "code_snapshots",
        {}
    )
    st.session_state.selection_warnings = progress_data.get(
        "selection_warnings",
        []
    )
    st.session_state.timer_started_at = progress_data.get(
        "timer_started_at",
        ""
    )
    st.session_state.timer_ends_at = progress_data.get(
        "timer_ends_at",
        ""
    )
    st.session_state.timer_duration_minutes = progress_data.get(
        "timer_duration_minutes",
        0
    )
    st.session_state.assessment_submitted = progress_data.get(
        "assessment_submitted",
        False
    )
    st.session_state.submitted_at = progress_data.get(
        "submitted_at",
        ""
    )

    records = progress_data.get("assessment_records", [])

    if records:

        st.session_state.assessment_df = pd.DataFrame(records)

def get_row_value(row, column_name, default_value=""):

    if column_name in row.index and pd.notna(row[column_name]):

        return row[column_name]

    normalized_target = re.sub(
        r"[^a-z0-9]+",
        "",
        str(column_name).strip().lower()
    )

    for existing_column in row.index:

        normalized_existing = re.sub(
            r"[^a-z0-9]+",
            "",
            str(existing_column).strip().lower()
        )

        if (
            normalized_existing == normalized_target
            and pd.notna(row[existing_column])
        ):

            return row[existing_column]

    return default_value

def get_first_row_value(row, column_names, default_value=None):

    for column_name in column_names:

        value = get_row_value(row, column_name, None)

        if value is not None:

            return value

    return default_value

def is_coding_question(row):

    question_type = str(
        get_row_value(row, "Question Type", "")
    ).strip().lower()

    if question_type in ["mcq", "multiple choice", "objective"]:

        return False

    if question_type in ["coding", "code", "programming"]:

        return True

    level = str(get_row_value(row, "Level", "")).strip().lower()

    if level in ["3", "l3", "level 3", "level3"]:

        return True

    return len(get_mcq_options(row)) < 2

def get_mcq_options(row):

    option_aliases = [
        ["Option 1", "Option1", "A"],
        ["Option 2", "Option2", "B"],
        ["Option 3", "Option3", "C"],
        ["Option 4", "Option4", "D"]
    ]
    options = []

    for aliases in option_aliases:

        value = get_first_row_value(row, aliases, None)

        if value is not None and str(value).strip() != "":

            options.append(value)

    if options:

        return options

    packed_options = get_first_row_value(
        row,
        ["Options", "Choices", "Answer Options"],
        None
    )

    if packed_options is None:

        return []

    text = str(packed_options).strip()

    if text == "":

        return []

    try:

        parsed = json.loads(text)

        if isinstance(parsed, list):

            return [
                option
                for option in parsed
                if str(option).strip() != ""
            ]

    except Exception:

        pass

    for delimiter in ["\n", "|", ";"]:

        if delimiter in text:

            return [
                option.strip()
                for option in text.split(delimiter)
                if option.strip() != ""
            ]

    return [
        option.strip()
        for option in text.split(",")
        if option.strip() != ""
    ]

def is_answer_correct(question_index, row):

    if question_index not in st.session_state.answers:

        return False

    if is_coding_question(row):

        code_answer = st.session_state.code_answers.get(
            question_index,
            {}
        )

        test_results = code_answer.get("test_results", [])

        if test_results:

            return all(
                result.get("passed", False)
                for result in test_results
            )

        return st.session_state.test_passed.get(
            question_index,
            False
        )

    selected_answer = st.session_state.answers.get(question_index)
    correct_answer = get_row_value(row, "Correct Answer", "")

    return str(selected_answer).strip() == str(correct_answer).strip()

def get_question_marks(row):

    marks = get_row_value(row, "Marks", 5)

    try:

        return int(float(marks))

    except (TypeError, ValueError):

        return 5

def split_packed_values(value):

    text = str(value or "").strip()

    if text == "":

        return []

    for delimiter in ["|||", "\n---\n", "\n\n", "|"]:

        if delimiter in text:

            return [
                item.strip()
                for item in text.split(delimiter)
                if item.strip() != ""
            ]

    return [text]

def parse_packed_test_cases(test_cases_value, expected_output_value):

    if test_cases_value is None and expected_output_value is None:

        return []

    test_cases_text = str(test_cases_value or "").strip()
    expected_output_text = str(expected_output_value or "").strip()

    if test_cases_text == "" and expected_output_text == "":

        return []

    try:

        parsed = json.loads(test_cases_text)

        if isinstance(parsed, list):

            packed_cases = []

            for item in parsed:

                if isinstance(item, dict):

                    packed_cases.append(
                        {
                            "input": str(
                                item.get("input")
                                or item.get("test_input")
                                or item.get("case")
                                or ""
                            ),
                            "expected": str(
                                item.get("expected")
                                or item.get("expected_output")
                                or item.get("output")
                                or ""
                            ),
                            "visible": bool(item.get("visible", True))
                        }
                    )

            return [
                item
                for item in packed_cases
                if item["expected"].strip() != ""
            ]

    except Exception:

        pass

    inputs = split_packed_values(test_cases_text)
    outputs = split_packed_values(expected_output_text)

    if len(outputs) == 1 and len(inputs) > 1:

        outputs = outputs * len(inputs)

    packed_cases = []

    for index, expected in enumerate(outputs):

        packed_cases.append(
            {
                "input": inputs[index] if index < len(inputs) else "",
                "expected": expected,
                "visible": True
            }
        )

    return packed_cases

def get_assessment_stats(df):

    attempted = len(st.session_state.answers)
    correct = 0
    earned_marks = 0
    total_marks = 0

    for _, row in df.iterrows():

        total_marks += get_question_marks(row)

    for question_index in st.session_state.answers.keys():

        if question_index < len(df):

            row = df.iloc[question_index]

            if is_answer_correct(question_index, row):

                correct += 1
                earned_marks += get_question_marks(row)

    total_questions = len(df)
    incorrect = attempted - correct
    pending = total_questions - attempted

    if total_questions == 0:

        percentage = 0

    else:

        percentage = int((earned_marks / total_marks) * 100)

    return (
        attempted,
        correct,
        incorrect,
        pending,
        percentage,
        earned_marks,
        total_marks
    )

def get_test_cases(row):

    test_cases = []

    packed_cases = parse_packed_test_cases(
        get_first_row_value(
            row,
            ["Test Cases", "Test Case", "Test Inputs"],
            None
        ),
        get_first_row_value(
            row,
            ["Expected Output", "Expected Outputs"],
            None
        )
    )

    test_cases.extend(packed_cases)

    single_input = get_first_row_value(
        row,
        ["Test Input", "Input", "Sample Input"],
        None
    )
    single_output = get_first_row_value(
        row,
        ["Expected Output", "Test Output", "Output", "Sample Output"],
        None
    )

    if single_output is not None and not packed_cases:

        test_cases.append(
            {
                "input": str(single_input or ""),
                "expected": str(single_output),
                "visible": True
            }
        )

    for index in range(1, 6):

        test_input = get_first_row_value(
            row,
            [
                f"Test Input {index}",
                f"Input {index}",
                f"Sample Input {index}"
            ],
            None
        )
        expected_output = get_first_row_value(
            row,
            [
                f"Expected Output {index}",
                f"Test Output {index}",
                f"Output {index}",
                f"Sample Output {index}"
            ],
            None
        )

        if expected_output is not None:

            test_cases.append(
                {
                    "input": str(test_input or ""),
                    "expected": str(expected_output),
                    "visible": True
                }
            )

    hidden_single_input = get_first_row_value(
        row,
        [
            "Hidden Test Input",
            "Hidden Input",
            "Hidden Test Case Input",
            "Private Test Input",
            "Private Input",
            "Backend Test Input",
            "Backend Input"
        ],
        None
    )
    hidden_single_output = get_first_row_value(
        row,
        [
            "Hidden Expected Output",
            "Hidden Test Output",
            "Hidden Output",
            "Hidden Test Case Output",
            "Hidden Test Case Expected Output",
            "Private Expected Output",
            "Private Test Output",
            "Private Output",
            "Backend Expected Output",
            "Backend Test Output",
            "Backend Output"
        ],
        None
    )

    if hidden_single_output is not None:

        test_cases.append(
            {
                "input": str(hidden_single_input or ""),
                "expected": str(hidden_single_output),
                "visible": False
            }
        )

    for index in range(1, 6):

        hidden_input = get_first_row_value(
            row,
            [
                f"Hidden Test Input {index}",
                f"Hidden Input {index}",
                f"Hidden Test Case Input {index}",
                f"Private Test Input {index}",
                f"Private Input {index}",
                f"Backend Test Input {index}",
                f"Backend Input {index}"
            ],
            None
        )
        hidden_output = get_first_row_value(
            row,
            [
                f"Hidden Expected Output {index}",
                f"Hidden Test Output {index}",
                f"Hidden Output {index}",
                f"Hidden Test Case Expected Output {index}",
                f"Hidden Test Case Output {index}",
                f"Private Expected Output {index}",
                f"Private Test Output {index}",
                f"Private Output {index}",
                f"Backend Expected Output {index}",
                f"Backend Test Output {index}",
                f"Backend Output {index}"
            ],
            None
        )

        if hidden_output is not None:

            test_cases.append(
                {
                    "input": str(hidden_input or ""),
                    "expected": str(hidden_output),
                    "visible": False
                }
            )

    return test_cases

def get_logic_scaffold(row):

    prefix = get_row_value(row, "Logic Prefix", None)
    suffix = get_row_value(row, "Logic Suffix", None)
    placeholder = get_row_value(
        row,
        "Logic Placeholder",
        "# Write only the required logic here"
    )

    if prefix is None and suffix is None:

        return None

    return {
        "prefix": str(prefix or ""),
        "suffix": str(suffix or ""),
        "placeholder": str(placeholder or "")
    }

def build_code_from_logic(logic_code, scaffold):

    if not scaffold:

        return logic_code

    parts = [
        scaffold.get("prefix", "").rstrip(),
        logic_code.strip("\n"),
        scaffold.get("suffix", "").lstrip()
    ]

    return "\n".join(
        part
        for part in parts
        if part != ""
    )

def generate_default_starter_code(row, language):

    language = str(language).lower()
    question_text = str(get_row_value(row, "Question", "")).lower()

    if language != "python":

        return ""

    if "space-separated" in question_text and (
        "number" in question_text or "integer" in question_text
    ):

        if "largest" in question_text or "maximum" in question_text:

            return "\n".join(
                [
                    "numbers = list(map(int, input().split()))",
                    "",
                    "# Write your logic here",
                    "answer = None",
                    "",
                    "print(answer)"
                ]
            )

        if "smallest" in question_text or "minimum" in question_text:

            return "\n".join(
                [
                    "numbers = list(map(int, input().split()))",
                    "",
                    "# Write your logic here",
                    "answer = None",
                    "",
                    "print(answer)"
                ]
            )

        if "sum" in question_text or "total" in question_text:

            return "\n".join(
                [
                    "numbers = list(map(int, input().split()))",
                    "",
                    "# Write your logic here",
                    "answer = 0",
                    "",
                    "print(answer)"
                ]
            )

        return "\n".join(
            [
                "numbers = list(map(int, input().split()))",
                "",
                "# Write your logic here",
                "answer = None",
                "",
                "print(answer)"
            ]
        )

    if "word" in question_text or "string" in question_text:

        return "\n".join(
            [
                "text = input().strip()",
                "",
                "# Write your logic here",
                "answer = None",
                "",
                "print(answer)"
            ]
        )

    return "\n".join(
        [
            "# Read input",
            "data = input().strip()",
            "",
            "# Write your logic here",
            "answer = None",
            "",
            "print(answer)"
        ]
    )

def get_code_suggestions(row, language):

    question_text = str(get_row_value(row, "Question", "")).lower()
    language = str(language).lower()
    suggestions = []

    if language == "python":

        suggestions.extend(
            [
                "Use `split()` to read space-separated values.",
                "Convert numeric inputs with `int()` or `float()` before calculations.",
                "Print only the final answer expected by the test cases."
            ]
        )

    if "list" in question_text or "array" in question_text:

        suggestions.append(
            "Loop through the values once and store only what you need."
        )

    if "maximum" in question_text or "largest" in question_text:

        suggestions.append("Track the current best value while iterating.")

    if "count" in question_text or "frequency" in question_text:

        suggestions.append("A dictionary/map is useful for counting values.")

    if "string" in question_text or "word" in question_text:

        suggestions.append(
            "Normalize whitespace/case if the question expects text comparison."
        )

    return suggestions[:5]

def is_saved_answer_correct(question_index, row, answers, code_answers):

    key = str(question_index)

    if key not in answers and question_index not in answers:

        return False

    if is_coding_question(row):

        code_answer = (
            code_answers.get(key)
            or code_answers.get(question_index)
            or {}
        )

        test_results = code_answer.get("test_results", [])

        return bool(test_results) and all(
            result.get("passed", False)
            for result in test_results
        )

    selected_answer = answers.get(key) or answers.get(question_index)
    correct_answer = get_row_value(row, "Correct Answer", "")

    return str(selected_answer).strip() == str(correct_answer).strip()

def calculate_saved_progress_stats(progress_data):

    records = progress_data.get("assessment_records", [])
    df = pd.DataFrame(records)

    if df.empty:

        return 0, 0, 0, 0, 0, 0, 0

    answers = progress_data.get("answers", {})
    code_answers = progress_data.get("code_answers", {})

    attempted = len(answers)
    correct = 0
    earned_marks = 0
    total_marks = 0

    for index, row in df.iterrows():

        marks = get_question_marks(row)
        total_marks += marks

        if is_saved_answer_correct(index, row, answers, code_answers):

            correct += 1
            earned_marks += marks

    total_questions = len(df)
    incorrect = attempted - correct
    pending = total_questions - attempted
    percentage = int((earned_marks / total_marks) * 100) if total_marks else 0

    return (
        attempted,
        correct,
        incorrect,
        pending,
        percentage,
        earned_marks,
        total_marks
    )

def has_saved_answer(question_index, answers):

    key = str(question_index)

    return key in answers or question_index in answers

def build_progress_analysis(progress_data, metadata=None):

    assessment_file = progress_data.get("assessment", "")
    metadata = metadata or load_assessment_metadata(assessment_file)
    records = progress_data.get("assessment_records", [])
    df = pd.DataFrame(records)
    answers = progress_data.get("answers", {})
    code_answers = progress_data.get("code_answers", {})
    course = (
        progress_data.get("course")
        or metadata.get("course")
        or "AIMD"
    )
    assessment_name = (
        metadata.get("assessment_name")
        or Path(str(assessment_file)).stem.replace("_", " ")
    )
    institution_name = progress_data.get("nsti", "")
    submitted_status = (
        "Yes"
        if progress_data.get("assessment_submitted", False)
        else "No"
    )
    level_scores = {
        level: {
            "earned": 0,
            "total": 0,
            "correct": 0,
            "questions": 0,
            "attempted": 0
        }
        for level in LEVEL_LABELS
    }
    skills = set()
    question_records = []

    if df.empty:

        return {
            "level_scores": level_scores,
            "skills": [],
            "question_records": question_records
        }

    for question_index, row in df.iterrows():

        level = get_question_level(row)

        if level not in level_scores:

            level_scores[level] = {
                "earned": 0,
                "total": 0,
                "correct": 0,
                "questions": 0,
                "attempted": 0
            }

        skill = str(get_row_value(row, "Skill", "General")).strip()
        skill = skill or "General"
        marks = get_question_marks(row)
        attempted = has_saved_answer(question_index, answers)
        is_correct = is_saved_answer_correct(
            question_index,
            row,
            answers,
            code_answers
        )
        earned = marks if is_correct else 0

        level_scores[level]["total"] += marks
        level_scores[level]["earned"] += earned
        level_scores[level]["questions"] += 1
        level_scores[level]["attempted"] += 1 if attempted else 0
        level_scores[level]["correct"] += 1 if is_correct else 0
        skills.add(skill)

        question_records.append(
            {
                "Student ID": progress_data.get("student_id", ""),
                "Student Name": progress_data.get("full_name", ""),
                "Email": progress_data.get("student_email", ""),
                "Course": course,
                "Institution Name": institution_name,
                "Assessment": assessment_file,
                "Assessment Name": assessment_name,
                "Question Number": question_index + 1,
                "Level": level,
                "Skill": skill,
                "Question Type": (
                    "Coding" if is_coding_question(row) else "MCQ"
                ),
                "Attempted": 1 if attempted else 0,
                "Correct": 1 if is_correct else 0,
                "Marks Obtained": earned,
                "Total Marks": marks,
                "Submitted": submitted_status
            }
        )

    return {
        "level_scores": level_scores,
        "skills": sorted(skills),
        "question_records": question_records
    }

def format_level_score(level_scores, level):

    values = level_scores.get(level, {})
    earned = int(values.get("earned", 0) or 0)
    total = int(values.get("total", 0) or 0)

    return f"{earned}/{total}"

def calculate_level_percentage(level_scores, level):

    values = level_scores.get(level, {})
    total = float(values.get("total", 0) or 0)

    if total <= 0:

        return 0

    return round((float(values.get("earned", 0) or 0) / total) * 100, 2)

def load_all_progress_records():

    records = []

    for progress_file in Path("assessment_progress").glob("*.json"):

        try:

            with open(progress_file, "r", encoding="utf-8") as file:

                progress_data = json.load(file)

            (
                attempted,
                correct,
                incorrect,
                pending,
                percentage,
                earned_marks,
                total_marks
            ) = calculate_saved_progress_stats(progress_data)

            assessment_file = progress_data.get("assessment", "")
            metadata = load_assessment_metadata(assessment_file)
            performance = summarize_performance(
                progress_data,
                is_saved_answer_correct,
                get_question_marks
            )
            progress_analysis = build_progress_analysis(
                progress_data,
                metadata
            )
            level_scores = progress_analysis["level_scores"]
            run_counts = progress_data.get("code_run_counts", {})
            failed_attempts = progress_data.get("code_failed_attempts", {})
            violations = progress_data.get("proctoring_violations", [])
            institution_name = progress_data.get("nsti", "")

            records.append(
                {
                    "Student ID": progress_data.get("student_id", ""),
                    "Student Name": progress_data.get("full_name", ""),
                    "Email": progress_data.get("student_email", ""),
                    "Course": (
                        progress_data.get("course")
                        or metadata.get("course")
                        or "AIMD"
                    ),
                    "Institution Name": institution_name,
                    "NSTI": institution_name,
                    "Location": progress_data.get("location", ""),
                    "Assessment Date": progress_data.get(
                        "assessment_date",
                        ""
                    ),
                    "Assessment": assessment_file,
                    "Assessment Name": (
                        metadata.get("assessment_name")
                        or Path(str(assessment_file)).stem.replace("_", " ")
                    ),
                    "Submitted": (
                        "Yes"
                        if progress_data.get("assessment_submitted", False)
                        else "No"
                    ),
                    "Submitted At": progress_data.get("submitted_at", ""),
                    "Final Submission Status": (
                        "Submitted"
                        if progress_data.get("assessment_submitted", False)
                        else "In Progress"
                    ),
                    "Total Questions": len(
                        progress_data.get("assessment_records", [])
                    ),
                    "Attempted": attempted,
                    "Correct": correct,
                    "Incorrect": incorrect,
                    "Pending": pending,
                    "Marks Obtained": earned_marks,
                    "Total Marks": total_marks,
                    "Percentage": percentage,
                    "Skill-wise Performance": performance["skill_summary"],
                    "Level-wise Performance": performance["level_summary"],
                    "Coding Accuracy": performance["coding_accuracy"],
                    "Code Runs": sum(
                        int(value or 0)
                        for value in run_counts.values()
                    ),
                    "Failed Coding Attempts": sum(
                        int(value or 0)
                        for value in failed_attempts.values()
                    ),
                    "Time Spent Per Question": format_question_time_spent(
                        progress_data.get("question_time_spent", {})
                    ),
                    "Question Submission Timestamps": (
                        format_question_timestamps(
                            progress_data.get(
                                "submission_timestamps",
                                {}
                            )
                        )
                    ),
                    "Proctoring Violations": len(violations),
                    "Proctoring Details": "; ".join(
                        [
                            (
                                f"Q{int(item.get('question_index', 0)) + 1} "
                                f"{item.get('event_type', '')} "
                                f"at {item.get('timestamp', '')}"
                            )
                            for item in violations
                        ]
                    ),
                    "Level 1 Score": format_level_score(
                        level_scores,
                        "Level 1"
                    ),
                    "Level 1 Percentage": calculate_level_percentage(
                        level_scores,
                        "Level 1"
                    ),
                    "Level 2 Score": format_level_score(
                        level_scores,
                        "Level 2"
                    ),
                    "Level 2 Percentage": calculate_level_percentage(
                        level_scores,
                        "Level 2"
                    ),
                    "Level 3 Score": format_level_score(
                        level_scores,
                        "Level 3"
                    ),
                    "Level 3 Percentage": calculate_level_percentage(
                        level_scores,
                        "Level 3"
                    ),
                    "Skills Covered": ", ".join(
                        progress_analysis["skills"]
                    )
                }
            )

        except Exception:

            continue

    return records

def load_all_question_performance_records():

    records = []

    for progress_file in Path("assessment_progress").glob("*.json"):

        try:

            with open(progress_file, "r", encoding="utf-8") as file:

                progress_data = json.load(file)

            assessment_file = progress_data.get("assessment", "")
            metadata = load_assessment_metadata(assessment_file)
            progress_analysis = build_progress_analysis(
                progress_data,
                metadata
            )
            records.extend(progress_analysis["question_records"])

        except Exception:

            continue

    return records

def calculate_analysis_percentage(df):

    if df.empty or "Total Marks" not in df.columns:

        return pd.Series(dtype=float)

    earned_marks = pd.to_numeric(
        df["Marks Obtained"],
        errors="coerce"
    ).fillna(0)
    total_marks = pd.to_numeric(
        df["Total Marks"],
        errors="coerce"
    )

    return (
        earned_marks
        .div(total_marks.where(total_marks != 0))
        .fillna(0)
        .mul(100)
        .round(2)
    )

def build_grouped_question_performance(df, group_column):

    if df.empty or group_column not in df.columns:

        return pd.DataFrame(
            columns=[
                group_column,
                "Questions",
                "Attempted",
                "Correct",
                "Marks Obtained",
                "Total Marks",
                "Average Percentage"
            ]
        )

    grouped_df = (
        df
        .groupby(group_column, dropna=False)
        .agg(
            **{
                "Questions": ("Question Number", "count"),
                "Attempted": ("Attempted", "sum"),
                "Correct": ("Correct", "sum"),
                "Marks Obtained": ("Marks Obtained", "sum"),
                "Total Marks": ("Total Marks", "sum")
            }
        )
        .reset_index()
    )
    grouped_df["Average Percentage"] = calculate_analysis_percentage(
        grouped_df
    )

    return grouped_df

def build_report_workbook(records, question_records=None):

    output = BytesIO()
    report_df = pd.DataFrame(records)
    question_df = pd.DataFrame(question_records or [])
    report_columns = [
        "Student ID",
        "Student Name",
        "Email",
        "Course",
        "Institution Name",
        "NSTI",
        "Location",
        "Assessment Date",
        "Assessment",
        "Assessment Name",
        "Submitted",
        "Submitted At",
        "Final Submission Status",
        "Total Questions",
        "Attempted",
        "Correct",
        "Incorrect",
        "Pending",
        "Marks Obtained",
        "Total Marks",
        "Percentage",
        "Skill-wise Performance",
        "Level-wise Performance",
        "Coding Accuracy",
        "Code Runs",
        "Failed Coding Attempts",
        "Time Spent Per Question",
        "Question Submission Timestamps",
        "Proctoring Violations",
        "Proctoring Details",
        "Level 1 Score",
        "Level 1 Percentage",
        "Level 2 Score",
        "Level 2 Percentage",
        "Level 3 Score",
        "Level 3 Percentage",
        "Skills Covered"
    ]
    question_columns = [
        "Student ID",
        "Student Name",
        "Email",
        "Course",
        "Institution Name",
        "Assessment",
        "Assessment Name",
        "Question Number",
        "Level",
        "Skill",
        "Question Type",
        "Attempted",
        "Correct",
        "Marks Obtained",
        "Total Marks",
        "Submitted"
    ]

    if report_df.empty:

        report_df = pd.DataFrame(columns=report_columns)

    else:

        if "Institution Name" not in report_df.columns:

            report_df["Institution Name"] = report_df.get("NSTI", "")

        for column in report_columns:

            if column not in report_df.columns:

                report_df[column] = ""

        report_df = report_df[report_columns]

    if question_df.empty:

        question_df = pd.DataFrame(columns=question_columns)

    else:

        for column in question_columns:

            if column not in question_df.columns:

                question_df[column] = ""

        question_df = question_df[question_columns]

    submitted_df = report_df[report_df["Submitted"] == "Yes"]

    if submitted_df.empty:

        institution_df = pd.DataFrame(
            columns=[
                "Institution Name",
                "Students Submitted",
                "Average Marks",
                "Average Percentage",
                "Highest Marks",
                "Lowest Marks"
            ]
        )
        course_df = pd.DataFrame(
            columns=[
                "Course",
                "Students Submitted",
                "Average Marks",
                "Average Percentage",
                "Average Code Runs",
                "Average Proctoring Violations"
            ]
        )
        test_df = pd.DataFrame(
            columns=[
                "Course",
                "Assessment Name",
                "Students Submitted",
                "Average Marks",
                "Average Percentage",
                "Highest Marks",
                "Lowest Marks"
            ]
        )

    else:

        institution_df = (
            submitted_df
            .groupby("Institution Name", dropna=False)
            .agg(
                **{
                    "Students Submitted": ("Student ID", "count"),
                    "Average Marks": ("Marks Obtained", "mean"),
                    "Average Percentage": ("Percentage", "mean"),
                    "Highest Marks": ("Marks Obtained", "max"),
                    "Lowest Marks": ("Marks Obtained", "min")
                }
            )
            .reset_index()
        )
        institution_df[
            ["Average Marks", "Average Percentage"]
        ] = institution_df[
            ["Average Marks", "Average Percentage"]
        ].round(2)

        course_df = (
            submitted_df
            .groupby("Course", dropna=False)
            .agg(
                **{
                    "Students Submitted": ("Student ID", "count"),
                    "Average Marks": ("Marks Obtained", "mean"),
                    "Average Percentage": ("Percentage", "mean"),
                    "Average Code Runs": ("Code Runs", "mean"),
                    "Average Proctoring Violations": (
                        "Proctoring Violations",
                        "mean"
                    )
                }
            )
            .reset_index()
        )
        course_df[
            [
                "Average Marks",
                "Average Percentage",
                "Average Code Runs",
                "Average Proctoring Violations"
            ]
        ] = course_df[
            [
                "Average Marks",
                "Average Percentage",
                "Average Code Runs",
                "Average Proctoring Violations"
            ]
        ].round(2)

        test_df = (
            submitted_df
            .groupby(["Course", "Assessment Name"], dropna=False)
            .agg(
                **{
                    "Students Submitted": ("Student ID", "count"),
                    "Average Marks": ("Marks Obtained", "mean"),
                    "Average Percentage": ("Percentage", "mean"),
                    "Highest Marks": ("Marks Obtained", "max"),
                    "Lowest Marks": ("Marks Obtained", "min")
                }
            )
            .reset_index()
        )
        test_df[
            ["Average Marks", "Average Percentage"]
        ] = test_df[
            ["Average Marks", "Average Percentage"]
        ].round(2)

    submitted_question_df = question_df[question_df["Submitted"] == "Yes"]

    if submitted_question_df.empty:

        level_df = pd.DataFrame(
            columns=[
                "Level",
                "Questions",
                "Attempted",
                "Correct",
                "Marks Obtained",
                "Total Marks",
                "Average Percentage"
            ]
        )
        skill_df = pd.DataFrame(
            columns=[
                "Skill",
                "Questions",
                "Attempted",
                "Correct",
                "Marks Obtained",
                "Total Marks",
                "Average Percentage"
            ]
        )

    else:

        level_df = (
            submitted_question_df
            .groupby("Level", dropna=False)
            .agg(
                **{
                    "Questions": ("Question Number", "count"),
                    "Attempted": ("Attempted", "sum"),
                    "Correct": ("Correct", "sum"),
                    "Marks Obtained": ("Marks Obtained", "sum"),
                    "Total Marks": ("Total Marks", "sum")
                }
            )
            .reset_index()
        )
        level_df["Average Percentage"] = calculate_analysis_percentage(
            level_df
        )

        skill_df = (
            submitted_question_df
            .groupby("Skill", dropna=False)
            .agg(
                **{
                    "Questions": ("Question Number", "count"),
                    "Attempted": ("Attempted", "sum"),
                    "Correct": ("Correct", "sum"),
                    "Marks Obtained": ("Marks Obtained", "sum"),
                    "Total Marks": ("Total Marks", "sum")
                }
            )
            .reset_index()
        )
        skill_df["Average Percentage"] = calculate_analysis_percentage(
            skill_df
        )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        report_df.to_excel(
            writer,
            sheet_name="Student Report",
            index=False
        )
        question_df.to_excel(
            writer,
            sheet_name="Question Performance",
            index=False
        )
        course_df.to_excel(
            writer,
            sheet_name="Course Analysis",
            index=False
        )
        test_df.to_excel(
            writer,
            sheet_name="Test Analysis",
            index=False
        )
        institution_df.to_excel(
            writer,
            sheet_name="Institution Analysis",
            index=False
        )
        level_df.to_excel(
            writer,
            sheet_name="Level Analysis",
            index=False
        )
        skill_df.to_excel(
            writer,
            sheet_name="Skill Analysis",
            index=False
        )

        workbook = writer.book

        def add_bar_chart(
            sheet_name,
            source_df,
            category_column,
            value_column,
            title,
            anchor
        ):

            if (
                BarChart is None
                or Reference is None
                or source_df.empty
                or category_column not in source_df.columns
                or value_column not in source_df.columns
            ):

                return

            sheet = workbook[sheet_name]
            chart = BarChart()
            chart.title = title
            chart.y_axis.title = value_column
            chart.x_axis.title = category_column
            value_col = source_df.columns.get_loc(value_column) + 1
            category_col = source_df.columns.get_loc(category_column) + 1
            data = Reference(
                sheet,
                min_col=value_col,
                min_row=1,
                max_row=len(source_df) + 1
            )
            categories = Reference(
                sheet,
                min_col=category_col,
                min_row=2,
                max_row=len(source_df) + 1
            )
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(categories)
            chart.height = 8
            chart.width = 16
            sheet.add_chart(chart, anchor)

        add_bar_chart(
            "Course Analysis",
            course_df,
            "Course",
            "Average Percentage",
            "Course-wise Average Percentage",
            "H2"
        )
        add_bar_chart(
            "Test Analysis",
            test_df,
            "Assessment Name",
            "Average Percentage",
            "Test-wise Average Percentage",
            "J2"
        )
        add_bar_chart(
            "Institution Analysis",
            institution_df,
            "Institution Name",
            "Average Percentage",
            "Institution-wise Average Percentage",
            "H2"
        )
        add_bar_chart(
            "Level Analysis",
            level_df,
            "Level",
            "Average Percentage",
            "Level-wise Task Performance",
            "I2"
        )
        add_bar_chart(
            "Skill Analysis",
            skill_df,
            "Skill",
            "Average Percentage",
            "Skill-based Performance",
            "I2"
        )

        for sheet in workbook.worksheets:

            sheet.freeze_panes = "A2"

            for column_cells in sheet.columns:

                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )
                sheet.column_dimensions[
                    column_cells[0].column_letter
                ].width = min(max(max_length + 2, 12), 34)

    output.seek(0)

    return output

def build_test_report_workbook(records, question_records=None):

    output = BytesIO()
    report_df = pd.DataFrame(records)
    question_df = pd.DataFrame(question_records or [])
    summary_columns = [
        "Sr. No.",
        "Student Name",
        "Institution Name",
        "Level 1 score",
        "Level 2 score",
        "Level 3 score",
        "Total",
        "Overall Skills checked"
    ]

    if report_df.empty:

        student_summary_df = pd.DataFrame(columns=summary_columns)

    else:

        if "Institution Name" not in report_df.columns:

            report_df["Institution Name"] = report_df.get("NSTI", "")

        report_df = report_df.sort_values(
            ["Institution Name", "Student Name"],
            na_position="last"
        ).reset_index(drop=True)
        marks_obtained = pd.to_numeric(
            report_df.get(
                "Marks Obtained",
                pd.Series([0] * len(report_df))
            ),
            errors="coerce"
        ).fillna(0)
        total_marks = pd.to_numeric(
            report_df.get(
                "Total Marks",
                pd.Series([0] * len(report_df))
            ),
            errors="coerce"
        ).fillna(0)
        student_summary_df = pd.DataFrame(
            {
                "Sr. No.": range(1, len(report_df) + 1),
                "Student Name": report_df.get("Student Name", ""),
                "Institution Name": report_df.get("Institution Name", ""),
                "Level 1 score": report_df.get("Level 1 Score", ""),
                "Level 2 score": report_df.get("Level 2 Score", ""),
                "Level 3 score": report_df.get("Level 3 Score", ""),
                "Total": (
                    marks_obtained.astype(int).astype(str)
                    + "/"
                    + total_marks.astype(int).astype(str)
                ),
                "Overall Skills checked": report_df.get(
                    "Skills Covered",
                    ""
                )
            }
        )

    if report_df.empty:

        institution_df = pd.DataFrame(
            columns=[
                "Institution Name",
                "Students Tracked",
                "Students Submitted",
                "Average Marks",
                "Average Percentage",
                "Highest Marks",
                "Lowest Marks"
            ]
        )

    else:

        institution_df = (
            report_df
            .groupby("Institution Name", dropna=False)
            .agg(
                **{
                    "Students Tracked": ("Student Name", "count"),
                    "Students Submitted": (
                        "Submitted",
                        lambda values: int((values == "Yes").sum())
                    ),
                    "Average Marks": ("Marks Obtained", "mean"),
                    "Average Percentage": ("Percentage", "mean"),
                    "Highest Marks": ("Marks Obtained", "max"),
                    "Lowest Marks": ("Marks Obtained", "min")
                }
            )
            .reset_index()
        )
        institution_df[
            ["Average Marks", "Average Percentage"]
        ] = institution_df[
            ["Average Marks", "Average Percentage"]
        ].round(2)

    if question_df.empty:

        skill_df = pd.DataFrame(
            columns=[
                "Skill",
                "Questions",
                "Attempted",
                "Correct",
                "Marks Obtained",
                "Total Marks",
                "Average Percentage"
            ]
        )

    else:

        skill_df = build_grouped_question_performance(
            question_df,
            "Skill"
        )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        student_summary_df.to_excel(
            writer,
            sheet_name="Student Summary",
            index=False
        )
        institution_df.to_excel(
            writer,
            sheet_name="Institution Performance",
            index=False
        )
        skill_df.to_excel(
            writer,
            sheet_name="Skill Performance",
            index=False
        )

        workbook = writer.book

        def add_report_chart(
            sheet_name,
            source_df,
            category_column,
            value_column,
            title,
            anchor
        ):

            if (
                BarChart is None
                or Reference is None
                or source_df.empty
                or category_column not in source_df.columns
                or value_column not in source_df.columns
            ):

                return

            sheet = workbook[sheet_name]
            chart = BarChart()
            chart.title = title
            chart.y_axis.title = value_column
            chart.x_axis.title = category_column
            value_col = source_df.columns.get_loc(value_column) + 1
            category_col = source_df.columns.get_loc(category_column) + 1
            data = Reference(
                sheet,
                min_col=value_col,
                min_row=1,
                max_row=len(source_df) + 1
            )
            categories = Reference(
                sheet,
                min_col=category_col,
                min_row=2,
                max_row=len(source_df) + 1
            )
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(categories)
            chart.height = 9
            chart.width = 17
            sheet.add_chart(chart, anchor)

        add_report_chart(
            "Institution Performance",
            institution_df,
            "Institution Name",
            "Average Percentage",
            "Institution-wise Performance",
            "I2"
        )
        add_report_chart(
            "Institution Performance",
            institution_df,
            "Institution Name",
            "Students Submitted",
            "Institution-wise Submissions",
            "I20"
        )
        add_report_chart(
            "Skill Performance",
            skill_df,
            "Skill",
            "Average Percentage",
            "Skill-wise Performance",
            "I2"
        )
        add_report_chart(
            "Skill Performance",
            skill_df,
            "Skill",
            "Correct",
            "Skill-wise Correct Answers",
            "I20"
        )

        for sheet in workbook.worksheets:

            sheet.freeze_panes = "A2"

            for column_cells in sheet.columns:

                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )
                sheet.column_dimensions[
                    column_cells[0].column_letter
                ].width = min(max(max_length + 2, 12), 38)

    output.seek(0)

    return output

def normalize_output(value):

    lines = str(value).replace("\r\n", "\n").replace("\r", "\n").split("\n")

    cleaned_lines = [
        line.rstrip()
        for line in lines
    ]

    return "\n".join(cleaned_lines).strip()

def output_matches(actual_output, expected_output):

    actual = normalize_output(actual_output)
    expected = normalize_output(expected_output)

    return actual == expected or actual.endswith(expected)

def is_static_check_language(language):

    return str(language).strip().lower() in ["html", "css", "powerbi"]

def run_static_code_check(language, code, test_case):

    required_text = str(test_case.get("input", "") or "")
    required_tokens = [
        token.strip()
        for token in re.split(r"\|\|\||\||,|\n", required_text)
        if token.strip() != ""
    ]

    normalized_code = str(code).lower()
    missing_tokens = [
        token
        for token in required_tokens
        if token.lower() not in normalized_code
    ]
    passed = len(missing_tokens) == 0

    return {
        "test": test_case.get("test", 1),
        "input": required_text,
        "expected": test_case.get("expected", "PASS"),
        "actual": "PASS" if passed else f"Missing: {', '.join(missing_tokens)}",
        "passed": passed,
        "visible": test_case.get("visible", True)
    }

def execute_student_code(language, code, program_input=""):

    language = language.lower()

    if language in ["html", "css", "powerbi"]:

        return {
            "stdout": (
                f"{language.title()} answer saved. Static checks are used "
                "for this question type."
            ),
            "stderr": "",
            "returncode": 0
        }

    with tempfile.TemporaryDirectory() as temp_dir:

        temp_path = Path(temp_dir)

        try:

            if language == "python":

                code_file = temp_path / "main.py"
                code_file.write_text(code, encoding="utf-8")
                command = [sys.executable, str(code_file)]

                result = subprocess.run(
                    command,
                    input=program_input,
                    capture_output=True,
                    text=True,
                    timeout=10
                )

            elif language == "java":

                code_file = temp_path / "Main.java"
                code_file.write_text(code, encoding="utf-8")

                compile_result = subprocess.run(
                    ["javac", str(code_file)],
                    capture_output=True,
                    text=True,
                    timeout=5
                )

                if compile_result.returncode != 0:

                    return {
                        "stdout": compile_result.stdout,
                        "stderr": compile_result.stderr,
                        "returncode": compile_result.returncode
                    }

                result = subprocess.run(
                    ["java", "-cp", str(temp_path), "Main"],
                    input=program_input,
                    capture_output=True,
                    text=True,
                    timeout=10
                )

            elif language == "javascript":

                code_file = temp_path / "main.js"
                code_file.write_text(code, encoding="utf-8")

                result = subprocess.run(
                    ["node", str(code_file)],
                    input=program_input,
                    capture_output=True,
                    text=True,
                    timeout=10
                )

            elif language == "c":

                code_file = temp_path / "main.c"
                exe_file = temp_path / "main.exe"
                code_file.write_text(code, encoding="utf-8")

                compile_result = subprocess.run(
                    ["gcc", str(code_file), "-o", str(exe_file)],
                    capture_output=True,
                    text=True,
                    timeout=5
                )

                if compile_result.returncode != 0:

                    return {
                        "stdout": compile_result.stdout,
                        "stderr": compile_result.stderr,
                        "returncode": compile_result.returncode
                    }

                result = subprocess.run(
                    [str(exe_file)],
                    input=program_input,
                    capture_output=True,
                    text=True,
                    timeout=10
                )

            elif language == "c++":

                code_file = temp_path / "main.cpp"
                exe_file = temp_path / "main.exe"
                code_file.write_text(code, encoding="utf-8")

                compile_result = subprocess.run(
                    ["g++", str(code_file), "-o", str(exe_file)],
                    capture_output=True,
                    text=True,
                    timeout=5
                )

                if compile_result.returncode != 0:

                    return {
                        "stdout": compile_result.stdout,
                        "stderr": compile_result.stderr,
                        "returncode": compile_result.returncode
                    }

                result = subprocess.run(
                    [str(exe_file)],
                    input=program_input,
                    capture_output=True,
                    text=True,
                    timeout=10
                )

            else:

                return {
                    "stdout": "",
                    "stderr": "Selected language is not supported yet.",
                    "returncode": 1
                }

            return {
                "stdout": result.stdout,
                "stderr": result.stderr,
                "returncode": result.returncode
            }

        except FileNotFoundError:

            return {
                "stdout": "",
                "stderr": (
                    f"{language.title()} runner is not installed or not "
                    "available in PATH on this system."
                ),
                "returncode": 1
            }

        except subprocess.TimeoutExpired:

            return {
                "stdout": "",
                "stderr": (
                    "Execution stopped because the code took longer than "
                    "10 seconds. If your program uses input(), make sure "
                    "the test input or Program Input has the required values."
                ),
                "returncode": 1
            }

        except Exception as e:

            return {
                "stdout": "",
                "stderr": f"Error while running code: {e}",
                "returncode": 1
            }

def format_execution_result(result):

    output = result["stdout"]
    error = result["stderr"]

    if output.strip() == "" and error.strip() == "":

        return "Code ran successfully with no output."

    if error.strip() != "":

        return error

    return output

def run_student_code(language, code, program_input=""):

    return format_execution_result(
        execute_student_code(language, code, program_input)
    )

def run_code_tests(language, code, test_cases):

    results = []

    for index, test_case in enumerate(test_cases, start=1):

        if is_static_check_language(language):

            static_result = run_static_code_check(
                language,
                code,
                {
                    **test_case,
                    "test": index
                }
            )
            results.append(static_result)
            continue

        result = execute_student_code(
            language,
            code,
            test_case["input"]
        )

        actual_output = result["stderr"] or result["stdout"]

        passed = (
            result["returncode"] == 0
            and output_matches(actual_output, test_case["expected"])
        )

        results.append(
            {
                "test": index,
                "input": test_case["input"],
                "expected": test_case["expected"],
                "actual": actual_output,
                "passed": passed,
                "visible": test_case.get("visible", True)
            }
        )

    return results

def format_test_results(test_results):

    if not test_results:

        return "No test cases found. Custom run completed."

    lines = []

    for result in test_results:

        status = "PASS" if result["passed"] else "FAIL"

        if not result.get("visible", True):

            lines.extend(
                [
                    f"Hidden Test Case {result['test']}: {status}",
                    ""
                ]
            )
            continue

        lines.extend(
            [
                f"Test Case {result['test']}: {status}",
                f"Input: {result['input'] or '(no input)'}",
                f"Expected Output: {normalize_output(result['expected'])}",
                f"Actual Output: {normalize_output(result['actual'])}",
                ""
            ]
        )

    passed_count = sum(
        1
        for result in test_results
        if result["passed"]
    )

    lines.append(
        f"Passed {passed_count}/{len(test_results)} test cases."
    )

    return "\n".join(lines)

# ---------------------------------------------------
# NSTI LIST
# ---------------------------------------------------
nsti_list = [
    "NSTI (W) Allahabad",
    "NSTI (W) Hyderabad",
    "NSTI (W) Kolkata",
    "NSTI (W) Patna",
    "NSTI Bangalore (General)",
    "NSTI Bangalore (Women)",
    "NSTI Calicut",
    "NSTI Chennai",
    "NSTI Haldwani",
    "NSTI Howrah",
    "NSTI Hyderabad (Ramanthapur)",
    "NSTI Indore (Women)",
    "NSTI Jaipur (Women)",
    "NSTI Jodhpur",
    "NSTI Kanpur",
    "NSTI Ludhiana",
    "NSTI Mumbai (General)",
    "NSTI Mumbai (Women)",
    "NSTI Noida (Women)",
    "NSTI Srinagar",
    "NSTI Trivandrum",
    "NSTI Vadodara (Women)"
]

# ===================================================
# LOGIN PAGE
# ===================================================
if st.session_state.page == "login":

    st.markdown(
        """
        <div class='dashboard-card hero-card'>
            <h1>Edunet Assessment LMS</h1>
            <p>Login to continue.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    role = st.radio(
        "Select Role",
        ["Student", "Administrator"]
    )

    # ---------------------------------------------------
    # STUDENT LOGIN
    # ---------------------------------------------------
    if role == "Student":

        student_email = st.text_input("Student Email")

        if st.button("Send Verification Code"):

            student_email = student_email.strip().lower()

            if not is_valid_email(student_email):

                st.error("Please enter valid Student Email.")

            else:

                verification_code = generate_verification_code()
                sent, notice = send_student_verification_email(
                    student_email,
                    verification_code
                )

                st.session_state.pending_student_email = student_email
                st.session_state.student_verification_code = verification_code
                st.session_state.student_verification_expires_at = (
                    datetime.now() + timedelta(minutes=10)
                ).isoformat(timespec="seconds")
                st.session_state.student_verification_sent = sent
                st.session_state.student_verification_notice = notice

                st.rerun()

        if st.session_state.pending_student_email:

            if st.session_state.student_verification_sent:

                st.success(st.session_state.student_verification_notice)

            else:

                st.warning(st.session_state.student_verification_notice)

            st.caption(
                f"Verification pending for {st.session_state.pending_student_email}"
            )

            entered_code = st.text_input(
                "Enter Verification Code",
                max_chars=6
            )

            verify_col, resend_col = st.columns(2)

            with verify_col:

                if st.button("Verify Student Email"):

                    expires_at = datetime.fromisoformat(
                        st.session_state.student_verification_expires_at
                    )

                    if datetime.now() > expires_at:

                        st.error(
                            "Verification code expired. Please send a new code."
                        )

                    elif (
                        entered_code.strip()
                        == st.session_state.student_verification_code
                    ):

                        st.session_state.student_email = (
                            st.session_state.pending_student_email
                        )
                        st.session_state.student_id = (
                            st.session_state.pending_student_email
                        )
                        st.session_state.pending_student_email = ""
                        st.session_state.student_verification_code = ""
                        st.session_state.student_verification_expires_at = ""
                        st.session_state.student_verification_sent = False
                        st.session_state.student_verification_notice = ""
                        st.session_state.page = "student_details"

                        st.rerun()

                    else:

                        st.error("Invalid verification code.")

            with resend_col:

                if st.button("Resend Code"):

                    verification_code = generate_verification_code()
                    sent, notice = send_student_verification_email(
                        st.session_state.pending_student_email,
                        verification_code
                    )
                    st.session_state.student_verification_code = (
                        verification_code
                    )
                    st.session_state.student_verification_expires_at = (
                        datetime.now() + timedelta(minutes=10)
                    ).isoformat(timespec="seconds")
                    st.session_state.student_verification_sent = sent
                    st.session_state.student_verification_notice = notice

                    st.rerun()

    # ---------------------------------------------------
    # ADMIN LOGIN
    # ---------------------------------------------------
    elif role == "Administrator":

        email_col1, email_col2 = st.columns([1.4, 1])

        with email_col1:

            admin_prefix = st.text_input(
                "Administrator Email",
                placeholder="name"
            )

        with email_col2:

            st.text_input(
                "Domain",
                value="@edunetfoundation.org",
                disabled=True,
                label_visibility="hidden"
            )

        if st.button("Send Admin Verification Code"):

            admin_prefix = admin_prefix.strip().lower()

            if not admin_prefix:

                st.error("Please enter Administrator Email.")

            elif "@" in admin_prefix or not is_valid_admin_email_prefix(
                admin_prefix
            ):

                st.error(
                    "Please enter only the part before @edunetfoundation.org."
                )

            else:

                admin_email = f"{admin_prefix}@edunetfoundation.org"

                verification_code = generate_verification_code()
                sent, notice = send_admin_verification_email(
                    admin_email,
                    verification_code
                )

                st.session_state.pending_admin_email = admin_email
                st.session_state.admin_verification_code = verification_code
                st.session_state.admin_verification_expires_at = (
                    datetime.now() + timedelta(minutes=10)
                ).isoformat(timespec="seconds")
                st.session_state.admin_verification_sent = sent
                st.session_state.admin_verification_notice = notice

                st.rerun()

        if st.session_state.pending_admin_email:

            if st.session_state.admin_verification_sent:

                st.success(st.session_state.admin_verification_notice)

            else:

                st.warning(st.session_state.admin_verification_notice)

            st.caption(
                f"Verification pending for {st.session_state.pending_admin_email}"
            )

            entered_code = st.text_input(
                "Enter Admin Verification Code",
                max_chars=6
            )

            verify_col, resend_col = st.columns(2)

            with verify_col:

                if st.button("Verify Admin Email"):

                    expires_at = datetime.fromisoformat(
                        st.session_state.admin_verification_expires_at
                    )

                    if datetime.now() > expires_at:

                        st.error(
                            "Verification code expired. Please send a new code."
                        )

                    elif (
                        entered_code.strip()
                        == st.session_state.admin_verification_code
                    ):

                        st.session_state.admin_email = (
                            st.session_state.pending_admin_email
                        )
                        st.session_state.pending_admin_email = ""
                        st.session_state.admin_verification_code = ""
                        st.session_state.admin_verification_expires_at = ""
                        st.session_state.admin_verification_sent = False
                        st.session_state.admin_verification_notice = ""
                        st.session_state.page = "admin_dashboard"

                        st.rerun()

                    else:

                        st.error("Invalid verification code.")

            with resend_col:

                if st.button("Resend Admin Code"):

                    verification_code = generate_verification_code()
                    sent, notice = send_admin_verification_email(
                        st.session_state.pending_admin_email,
                        verification_code
                    )
                    st.session_state.admin_verification_code = (
                        verification_code
                    )
                    st.session_state.admin_verification_expires_at = (
                        datetime.now() + timedelta(minutes=10)
                    ).isoformat(timespec="seconds")
                    st.session_state.admin_verification_sent = sent
                    st.session_state.admin_verification_notice = notice

                    st.rerun()

# ===================================================
# STUDENT DETAILS PAGE
# ===================================================
elif st.session_state.page == "student_details":

    with st.sidebar:

        st.title("Student")

        st.write(
            f"Email: {st.session_state.student_email}"
        )

        if st.button("Back to Login"):

            st.session_state.page = "login"

            st.rerun()

        st.button(
            "Logout",
            on_click=logout
        )

    st.markdown(
        """
        <div class='dashboard-card'>
            <h2>Student Details</h2>
        </div>
        """,
        unsafe_allow_html=True
    )

    courses = load_courses()

    with st.form("student_form"):

        full_name = st.text_input("Student Full Name")

        selected_course = st.selectbox(
            "Select Course",
            ["Select Course"] + courses
        )

        selected_nsti = st.selectbox(
            "Select Institution",
            ["Select Institution"] + nsti_list
        )

        location = st.text_input("Location")

        st.text_input(
            "Student Email",
            value=st.session_state.student_email,
            disabled=True
        )

        assessment_date = st.date_input(
            "Assessment Date",
            value=date.today()
        )

        next_button = st.form_submit_button(
            "Save & Continue"
        )

    if next_button:

        full_name = " ".join(full_name.strip().split())

        if full_name == "":

            st.error("Please enter Student Full Name.")

        elif not is_valid_name(full_name):

            st.error(
                "Student Full Name should contain letters and spaces only."
            )

        elif selected_course == "Select Course":

            st.error("Please select Course.")

        elif selected_nsti == "Select Institution":

            st.error("Please select Institution.")

        elif location.strip() == "":

            st.error("Please enter Location.")

        else:

            st.session_state.full_name = full_name
            st.session_state.selected_course = selected_course
            st.session_state.selected_nsti = selected_nsti
            st.session_state.location = location
            st.session_state.assessment_date = assessment_date

            st.session_state.page = "student_assessment_selection"

            st.rerun()

# ===================================================
# ADMIN DASHBOARD
# ===================================================
elif st.session_state.page == "admin_dashboard":

    with st.sidebar:

        st.title("Administrator")

        st.write(st.session_state.admin_email)

        if st.button("Back to Login"):

            st.session_state.page = "login"

            st.rerun()

    st.markdown(
        """
        <div class='dashboard-card hero-card'>
            <h1>Administrator Dashboard</h1>
            <p>Manage courses, assessments, schedules, analytics, and reports.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    courses = load_courses()
    assessment_files = [
        file
        for file in os.listdir("assessment_files")
        if file.endswith(".xlsx")
    ]
    report_records = load_all_progress_records()
    submitted_count = sum(
        1
        for record in report_records
        if record.get("Submitted") == "Yes"
    )

    metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)

    with metric_col1:

        st.markdown(
            f"""
            <div class='metric-card'>
                <h3>{len(courses)}</h3>
                <div class='metric-pill'>Courses</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with metric_col2:

        st.markdown(
            f"""
            <div class='metric-card'>
                <h3>{len(assessment_files)}</h3>
                <div class='metric-pill'>Assessments</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with metric_col3:

        st.markdown(
            f"""
            <div class='metric-card'>
                <h3>{len(report_records)}</h3>
                <div class='metric-pill'>Students Tracked</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with metric_col4:

        st.markdown(
            f"""
            <div class='metric-card'>
                <h3>{submitted_count}</h3>
                <div class='metric-pill'>Submitted</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("### Quick Actions")

    col1, col2, col3, col4 = st.columns(4)

    with col1:

        if st.button("Upload Assessment"):

            st.session_state.page = "upload_assessment"

            st.rerun()

    with col2:

        if st.button("View Assessments"):

            st.session_state.page = "view_assessments"

            st.rerun()

    with col3:

        if st.button("Manage Courses"):

            st.session_state.page = "course_management"

            st.rerun()

    with col4:

        if st.button("Download Reports"):

            st.session_state.page = "download_reports"

            st.rerun()

    if assessment_files:

        course_counts = {}

        for file in assessment_files:

            metadata = load_assessment_metadata(file)
            course = metadata.get("course", "AIMD")
            course_counts[course] = course_counts.get(course, 0) + 1

        st.markdown("### Assessment Mix")
        st.bar_chart(pd.DataFrame.from_dict(
            course_counts,
            orient="index",
            columns=["Assessments"]
        ))

    st.button(
        "Logout",
        on_click=logout
    )

# ===================================================
# COURSE MANAGEMENT PAGE
# ===================================================
elif st.session_state.page == "course_management":

    with st.sidebar:

        st.title("Administrator")

        if st.button("Back to Dashboard"):

            st.session_state.page = "admin_dashboard"

            st.rerun()

        st.button(
            "Logout",
            on_click=logout
        )

    st.markdown(
        """
        <div class='dashboard-card'>
            <h2>Course Management</h2>
            <p>Create the course list used by admins and students.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    courses = load_courses()

    col_add, col_reset = st.columns([2, 1])

    with col_add:

        new_course = st.text_input(
            "Add Course",
            placeholder="Example: AIMD"
        )

        if st.button("Add Course"):

            course_name = new_course.strip().upper()

            if not course_name:

                st.error("Please enter a course name.")

            elif course_name in courses:

                st.warning("Course already exists.")

            else:

                courses.append(course_name)
                save_courses(courses)
                st.success("Course added.")
                st.rerun()

    with col_reset:

        if st.button("Restore Defaults"):

            save_courses(DEFAULT_COURSES)
            st.success("Default courses restored.")
            st.rerun()

    if courses:

        st.markdown("### Active Courses")

        for course in courses:

            course_col, action_col = st.columns([4, 1])

            with course_col:

                st.markdown(
                    f"""
                    <div class='metric-card'>
                        <h3>{html.escape(course)}</h3>
                        <div class='metric-pill'>Active</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            with action_col:

                if st.button(
                    "Remove",
                    key=f"remove_course_{course}",
                    disabled=len(courses) <= 1
                ):

                    save_courses(
                        [
                            item
                            for item in courses
                            if item != course
                        ]
                    )
                    st.success("Course removed.")
                    st.rerun()

# ===================================================
# UPLOAD ASSESSMENT PAGE
# ===================================================
elif st.session_state.page == "upload_assessment":

    with st.sidebar:

        st.title("Administrator")

        if st.button("Back to Dashboard"):

            st.session_state.page = "admin_dashboard"

            st.rerun()

        st.button(
            "Logout",
            on_click=logout
        )

    st.markdown(
        """
        <div class='dashboard-card'>
            <h2>Upload Assessment</h2>
        </div>
        """,
        unsafe_allow_html=True
    )

    courses = load_courses()

    assessment_name = st.text_input("Assessment Name")

    selected_course = st.selectbox(
        "Course",
        courses,
        index=0
    )

    schedule_col1, schedule_col2, duration_col = st.columns(3)

    with schedule_col1:

        start_date = st.date_input(
            "Assessment Start Date",
            value=date.today()
        )

        start_time = st.time_input(
            "Assessment Start Time",
            value=time(9, 0)
        )

    with schedule_col2:

        end_date = st.date_input(
            "Assessment End Date",
            value=date.today()
        )

        end_time = st.time_input(
            "Assessment End Time",
            value=time(18, 0)
        )

    with duration_col:

        duration_minutes = st.number_input(
            "Duration in Minutes",
            min_value=1,
            max_value=360,
            value=60,
            step=5
        )

    uploaded_file = st.file_uploader(
        "Upload Excel Question Bank",
        type=["xlsx"]
    )

    st.info(
        """
        Expected columns include Question ID, Course, Skill, Level,
        Question Type, Question, Options or Option 1-4, Correct Answer,
        Test Cases, Expected Output, Marks, and Difficulty.
        Multiple sheets are supported.
        """
    )

    if uploaded_file is not None:

        if BarChart is None or Reference is None:
            st.error(
                "The openpyxl package is not installed in this environment. "
                "Excel upload and report generation require openpyxl."
            )
            st.stop()

        workbook_bytes = uploaded_file.getvalue()

        try:
            excel_file = pd.ExcelFile(
                BytesIO(workbook_bytes),
                engine="openpyxl"
            )
        except ImportError:
            st.error(
                "openpyxl is required to read Excel files. "
                "Please install openpyxl and redeploy the app."
            )
            st.stop()
        except Exception as e:
            st.error(f"Could not read workbook: {e}")
            st.stop()

        sheet_names = excel_file.sheet_names

        selected_sheets = st.multiselect(
            "Select Sheet(s)",
            sheet_names,
            default=sheet_names[:1]
        )

        if selected_sheets:

            question_bank = read_selected_sheets(
                workbook_bytes,
                selected_sheets,
                selected_course
            )
            summary = summarize_question_bank(question_bank)

            metric_col1, metric_col2, metric_col3, metric_col4, metric_col5 = (
                st.columns(5)
            )

            with metric_col1:
                st.metric("Level 1", summary["level_counts"]["Level 1"])
            with metric_col2:
                st.metric("Level 2", summary["level_counts"]["Level 2"])
            with metric_col3:
                st.metric("Level 3", summary["level_counts"]["Level 3"])
            with metric_col4:
                st.metric("Skills", summary["total_skills"])
            with metric_col5:
                st.metric("Unique IDs", summary["unique_question_ids"])

            if not question_bank.empty:

                st.dataframe(
                    question_bank.head(10),
                    use_container_width=True
                )

            st.markdown("### Level Distribution")

            level_count_cols = st.columns(3)
            level_counts = {}
            level_marks = {}

            for index, level in enumerate(LEVEL_LABELS):

                with level_count_cols[index]:

                    available = summary["level_counts"].get(level, 0)
                    level_counts[level] = int(
                        st.number_input(
                            f"{level} Questions",
                            min_value=0,
                            value=min(
                                DEFAULT_LEVEL_DISTRIBUTION[level],
                                available
                            ),
                            step=1,
                            key=f"create_count_{level}"
                        )
                    )
                    level_marks[level] = int(
                        st.number_input(
                            f"{level} Marks",
                            min_value=1,
                            value=DEFAULT_LEVEL_MARKS[level],
                            step=1,
                            key=f"create_marks_{level}"
                        )
                    )

            st.markdown("### Skill Requirements")
            st.caption(
                "Keep a skill count at 0 to let level-wise randomization fill it."
            )

            skill_source_df = (
                question_bank
                .groupby("Skill", dropna=False)["Source Sheet"]
                .apply(
                    lambda values: ", ".join(
                        sorted(
                            {
                                str(value)
                                for value in values
                                if str(value).strip()
                            }
                        )
                    )
                )
                .reset_index(name="Source Sheet(s)")
                .sort_values(["Source Sheet(s)", "Skill"])
            )
            skill_editor_df = skill_source_df.copy()
            skill_editor_df["Questions Required"] = 0
            skill_editor_key = (
                "skill_requirements_editor_"
                + safe_file_name(
                    "_".join(selected_sheets)
                    + "_"
                    + "_".join(summary["skills"])
                )[:120]
            )
            edited_skill_df = st.data_editor(
                skill_editor_df,
                disabled=["Skill", "Source Sheet(s)"],
                use_container_width=True,
                hide_index=True,
                key=skill_editor_key
            )
            skill_counts = {}

            for _, row in edited_skill_df.iterrows():

                try:

                    required_questions = int(
                        float(row["Questions Required"] or 0)
                    )

                except (TypeError, ValueError):

                    required_questions = 0

                if required_questions > 0:

                    skill_counts[str(row["Skill"]).strip()] = (
                        required_questions
                    )

            allowed_languages = st.multiselect(
                "Allowed Programming Languages",
                SUPPORTED_LANGUAGES,
                default=["Python"]
            )

            with st.expander("Proctoring Settings", expanded=True):

                proctoring_enabled = st.checkbox(
                    "Enable Proctoring",
                    value=True
                )
                proctor_col1, proctor_col2, proctor_col3 = st.columns(3)

                with proctor_col1:
                    disable_clipboard = st.checkbox(
                        "Disable Copy/Paste",
                        value=True
                    )

                with proctor_col2:
                    prevent_right_click = st.checkbox(
                        "Prevent Right-Click",
                        value=True
                    )

                with proctor_col3:
                    detect_focus_loss = st.checkbox(
                        "Detect Tab Switching",
                        value=True
                    )

            if st.button("Create Assessment"):

                start_at = datetime.combine(start_date, start_time)
                end_at = datetime.combine(end_date, end_time)
                normalized_columns = {
                    re.sub(
                        r"[^a-z0-9]+",
                        "",
                        str(column).strip().lower()
                    )
                    for column in question_bank.columns
                }

                if assessment_name.strip() == "":

                    st.error("Please enter Assessment Name.")

                elif question_bank.empty:

                    st.error("No valid questions found in selected sheets.")

                elif "question" not in normalized_columns:

                    st.error("Missing required column: Question")

                elif sum(level_counts.values()) <= 0:

                    st.error("Please select at least one question.")

                elif not allowed_languages:

                    st.error(
                        "Please select at least one programming language."
                    )

                elif end_at <= start_at:

                    st.error(
                        "Assessment end date/time must be after start date/time."
                    )

                else:

                    saved_file_name = (
                        f"{safe_file_name(assessment_name)}_"
                        f"{safe_file_name(selected_course)}.xlsx"
                    )
                    save_path = os.path.join(
                        "assessment_files",
                        saved_file_name
                    )

                    question_bank.to_excel(save_path, index=False)

                    save_assessment_metadata(
                        saved_file_name,
                        {
                            "assessment_name": assessment_name.strip(),
                            "course": selected_course,
                            "selected_sheets": selected_sheets,
                            "question_bank_summary": summary,
                            "level_counts": level_counts,
                            "level_marks": level_marks,
                            "skill_counts": skill_counts,
                            "allowed_languages": allowed_languages,
                            "proctoring": {
                                "enabled": proctoring_enabled,
                                "disable_clipboard": disable_clipboard,
                                "prevent_right_click": prevent_right_click,
                                "detect_focus_loss": detect_focus_loss
                            },
                            "start_at": start_at.isoformat(timespec="seconds"),
                            "end_at": end_at.isoformat(timespec="seconds"),
                            "duration_minutes": int(duration_minutes)
                        }
                    )

                    st.success("Assessment created successfully.")
                    st.dataframe(
                        question_bank.head(),
                        use_container_width=True
                    )

# ===================================================
# VIEW ASSESSMENTS PAGE
# ===================================================
elif st.session_state.page == "view_assessments":

    with st.sidebar:

        st.title("Administrator")

        if st.button("Back to Dashboard"):

            st.session_state.page = "admin_dashboard"

            st.rerun()

        st.button(
            "Logout",
            on_click=logout
        )

    st.markdown(
        """
        <div class='dashboard-card'>
            <h2>Uploaded Assessments</h2>
        </div>
        """,
        unsafe_allow_html=True
    )

    files = os.listdir("assessment_files")

    excel_files = [
        file for file in files
        if file.endswith(".xlsx")
    ]

    if len(excel_files) == 0:

        st.warning("No assessments uploaded.")

    else:

        selected_for_delete = st.multiselect(
            "Select assessments for bulk delete",
            sorted(excel_files),
            placeholder="Choose one or more assessments"
        )
        delete_progress_records = st.checkbox(
            "Also delete related student progress/results",
            value=False
        )

        bulk_col1, bulk_col2 = st.columns([1, 2])

        with bulk_col1:

            confirm_bulk_delete = st.checkbox(
                "Confirm bulk delete",
                key="confirm_bulk_delete"
            )

        with bulk_col2:

            if st.button(
                "Delete Selected Assessments",
                disabled=(
                    not selected_for_delete
                    or not confirm_bulk_delete
                )
            ):

                try:

                    total_deleted = 0

                    for selected_file in selected_for_delete:

                        total_deleted += len(
                            delete_assessment_artifacts(
                                selected_file,
                                delete_progress_records
                            )
                        )

                    st.success(
                        f"Deleted {len(selected_for_delete)} assessment(s). "
                        f"Removed {total_deleted} file(s)/record(s)."
                    )

                    st.rerun()

                except Exception as e:

                    st.error(f"Bulk delete failed: {e}")

        st.markdown("<br>", unsafe_allow_html=True)

        for file in sorted(excel_files):

            col1, col2, col3 = st.columns([4, 1, 1])
            metadata = load_assessment_metadata(file)

            with col1:

                st.markdown(
                    f"""
                    <div class='metric-card'>
                        <h3>{file.replace('.xlsx', '')}</h3>
                        <div class='metric-pill'>
                            Available
                        </div>
                        <div style="margin-top:12px; line-height:1.8;">
                            Starts: <b>{format_datetime(metadata.get("start_at"))}</b><br>
                            Ends: <b>{format_datetime(metadata.get("end_at"))}</b><br>
                            Duration: <b>{format_duration(metadata.get("duration_minutes"))}</b><br>
                            Course: <b>{html.escape(str(metadata.get("course", "AIMD")))}</b><br>
                            Sheets: <b>{html.escape(", ".join(metadata.get("selected_sheets", [metadata.get("sheet_name", "Sheet1")])))}</b>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            with col2:

                if st.button(
                    "Edit",
                    key=f"edit_assessment_{file}"
                ):

                    st.session_state.edit_assessment_file = file
                    st.rerun()

            with col3:

                confirm_delete = st.checkbox(
                    "Confirm",
                    key=f"confirm_delete_{file}"
                )

                if st.button(
                    "Delete",
                    key=f"delete_assessment_{file}",
                    disabled=not confirm_delete
                ):

                    file_path = os.path.join(
                        "assessment_files",
                        file
                    )

                    try:

                        delete_assessment_artifacts(file)

                        st.success(
                            f"{file} deleted successfully."
                        )

                        st.rerun()

                    except Exception as e:

                        st.error(
                            f"Could not delete assessment: {e}"
                        )

            if st.session_state.get("edit_assessment_file") == file:

                with st.expander(
                    f"Edit {file.replace('.xlsx', '')}",
                    expanded=True
                ):

                    try:

                        edit_path = os.path.join(
                            "assessment_files",
                            file
                        )
                        edit_df = pd.read_excel(edit_path)
                        edited_df = st.data_editor(
                            edit_df,
                            num_rows="dynamic",
                            use_container_width=True,
                            key=f"data_editor_{file}"
                        )

                        start_at = parse_datetime(
                            metadata.get("start_at")
                        ) or datetime.combine(date.today(), time(9, 0))
                        end_at = parse_datetime(
                            metadata.get("end_at")
                        ) or datetime.combine(date.today(), time(18, 0))

                        edit_col1, edit_col2 = st.columns(2)

                        with edit_col1:

                            edited_start_date = st.date_input(
                                "Start Date",
                                value=start_at.date(),
                                key=f"edit_start_date_{file}"
                            )
                            edited_start_time = st.time_input(
                                "Start Time",
                                value=start_at.time(),
                                key=f"edit_start_time_{file}"
                            )

                        with edit_col2:

                            edited_end_date = st.date_input(
                                "End Date",
                                value=end_at.date(),
                                key=f"edit_end_date_{file}"
                            )
                            edited_end_time = st.time_input(
                                "End Time",
                                value=end_at.time(),
                                key=f"edit_end_time_{file}"
                            )

                        duration_options = [
                            15, 30, 45, 60, 75, 90, 120, 150, 180
                        ]

                        try:

                            current_duration = int(
                                metadata.get("duration_minutes", 60)
                            )

                        except (TypeError, ValueError):

                            current_duration = 60

                        if current_duration not in duration_options:

                            current_duration = 60

                        edited_duration = st.selectbox(
                            "Duration",
                            duration_options,
                            index=duration_options.index(current_duration),
                            format_func=lambda value: format_duration(value),
                            key=f"edit_duration_{file}"
                        )

                        edit_courses = load_courses()
                        current_course = metadata.get("course", "AIMD")

                        if current_course not in edit_courses:

                            edit_courses.append(current_course)

                        edited_course = st.selectbox(
                            "Course",
                            edit_courses,
                            index=edit_courses.index(current_course),
                            key=f"edit_course_{file}"
                        )

                        edit_summary = summarize_question_bank(edited_df)
                        st.caption(
                            "Question bank summary: "
                            f"L1 {edit_summary['level_counts']['Level 1']} | "
                            f"L2 {edit_summary['level_counts']['Level 2']} | "
                            f"L3 {edit_summary['level_counts']['Level 3']} | "
                            f"{edit_summary['total_skills']} skills"
                        )

                        current_level_counts = metadata.get(
                            "level_counts",
                            DEFAULT_LEVEL_DISTRIBUTION
                        )
                        current_level_marks = metadata.get(
                            "level_marks",
                            DEFAULT_LEVEL_MARKS
                        )
                        edited_level_counts = {}
                        edited_level_marks = {}
                        level_edit_cols = st.columns(3)

                        for level_index, level in enumerate(LEVEL_LABELS):

                            with level_edit_cols[level_index]:

                                edited_level_counts[level] = int(
                                    st.number_input(
                                        f"{level} Questions",
                                        min_value=0,
                                        value=int(
                                            current_level_counts.get(
                                                level,
                                                0
                                            )
                                        ),
                                        step=1,
                                        key=f"edit_count_{file}_{level}"
                                    )
                                )
                                edited_level_marks[level] = int(
                                    st.number_input(
                                        f"{level} Marks",
                                        min_value=1,
                                        value=int(
                                            current_level_marks.get(
                                                level,
                                                DEFAULT_LEVEL_MARKS[level]
                                            )
                                        ),
                                        step=1,
                                        key=f"edit_marks_{file}_{level}"
                                    )
                                )

                        current_languages = [
                            language
                            for language in metadata.get(
                                "allowed_languages",
                                ["Python"]
                            )
                            if language in SUPPORTED_LANGUAGES
                        ] or ["Python"]

                        edited_languages = st.multiselect(
                            "Allowed Programming Languages",
                            SUPPORTED_LANGUAGES,
                            default=current_languages,
                            key=f"edit_languages_{file}"
                        )

                        current_proctoring = metadata.get("proctoring", {})

                        with st.expander(
                            "Edit Proctoring Settings",
                            expanded=False
                        ):

                            edited_proctoring_enabled = st.checkbox(
                                "Enable Proctoring",
                                value=current_proctoring.get(
                                    "enabled",
                                    True
                                ),
                                key=f"edit_proctor_enabled_{file}"
                            )
                            edited_disable_clipboard = st.checkbox(
                                "Disable Copy/Paste",
                                value=current_proctoring.get(
                                    "disable_clipboard",
                                    True
                                ),
                                key=f"edit_disable_clipboard_{file}"
                            )
                            edited_prevent_right_click = st.checkbox(
                                "Prevent Right-Click",
                                value=current_proctoring.get(
                                    "prevent_right_click",
                                    True
                                ),
                                key=f"edit_prevent_right_click_{file}"
                            )
                            edited_detect_focus_loss = st.checkbox(
                                "Detect Tab Switching",
                                value=current_proctoring.get(
                                    "detect_focus_loss",
                                    True
                                ),
                                key=f"edit_detect_focus_loss_{file}"
                            )

                        save_col, cancel_col = st.columns(2)

                        with save_col:

                            if st.button(
                                "Save Assessment Changes",
                                key=f"save_edit_{file}"
                            ):

                                if datetime.combine(
                                    edited_end_date,
                                    edited_end_time
                                ) <= datetime.combine(
                                    edited_start_date,
                                    edited_start_time
                                ):

                                    st.error(
                                        "Assessment end date/time must be after start date/time."
                                    )

                                else:

                                    edited_df.to_excel(
                                        edit_path,
                                        index=False
                                    )
                                    metadata.update(
                                        {
                                            "course": edited_course,
                                            "question_bank_summary": (
                                                edit_summary
                                            ),
                                            "level_counts": (
                                                edited_level_counts
                                            ),
                                            "level_marks": (
                                                edited_level_marks
                                            ),
                                            "allowed_languages": (
                                                edited_languages
                                            ),
                                            "proctoring": {
                                                "enabled": (
                                                    edited_proctoring_enabled
                                                ),
                                                "disable_clipboard": (
                                                    edited_disable_clipboard
                                                ),
                                                "prevent_right_click": (
                                                    edited_prevent_right_click
                                                ),
                                                "detect_focus_loss": (
                                                    edited_detect_focus_loss
                                                )
                                            },
                                            "start_at": datetime.combine(
                                                edited_start_date,
                                                edited_start_time
                                            ).isoformat(timespec="seconds"),
                                            "end_at": datetime.combine(
                                                edited_end_date,
                                                edited_end_time
                                            ).isoformat(timespec="seconds"),
                                            "duration_minutes": int(
                                                edited_duration
                                            )
                                        }
                                    )
                                    save_assessment_metadata(file, metadata)
                                    st.session_state.edit_assessment_file = ""
                                    st.success(
                                        "Assessment updated successfully."
                                    )
                                    st.rerun()

                        with cancel_col:

                            if st.button(
                                "Close Editor",
                                key=f"close_edit_{file}"
                            ):

                                st.session_state.edit_assessment_file = ""
                                st.rerun()

                    except Exception as e:

                        st.error(f"Could not edit assessment: {e}")

# ===================================================
# STUDENT ASSESSMENT SELECTION
# ===================================================
elif st.session_state.page == "student_assessment_selection":

    with st.sidebar:

        st.title("Student Dashboard")

        st.write(
            st.session_state.full_name
        )
        st.write(
            f"Course: {st.session_state.selected_course}"
        )

        if st.button("Back to Student Details"):

            st.session_state.page = "student_details"

            st.rerun()

        st.button(
            "Logout",
            on_click=logout
        )

    st.markdown(
        """
        <div class='dashboard-card'>
            <h2>Select Assessment</h2>
        </div>
        """,
        unsafe_allow_html=True
    )

    files = os.listdir("assessment_files")

    excel_files = [
        file for file in files
        if file.endswith(".xlsx")
    ]

    course_assessments = []

    for file in excel_files:

        metadata = load_assessment_metadata(file)

        if metadata.get("course", "AIMD") == st.session_state.selected_course:

            course_assessments.append(file)

    if len(course_assessments) == 0:

        st.warning("No assessments available for your selected course.")

    else:

        for file in course_assessments:

            col1, col2 = st.columns([4, 1])
            metadata = load_assessment_metadata(file)
            is_available, availability_message = (
                get_assessment_availability(metadata)
            )
            existing_progress = load_assessment_progress(
                st.session_state.student_id,
                st.session_state.student_email,
                file
            )

            with col1:

                st.markdown(
                    f"""
                    <div class='metric-card'>
                        <h3>{file.replace('.xlsx', '')}</h3>
                        <div class='metric-pill'>
                            {availability_message}
                        </div>
                        <div style="margin-top:12px; line-height:1.8;">
                            Starts: <b>{format_datetime(metadata.get("start_at"))}</b><br>
                            Ends: <b>{format_datetime(metadata.get("end_at"))}</b><br>
                            Duration: <b>{format_duration(metadata.get("duration_minutes"))}</b><br>
                            Course: <b>{html.escape(str(metadata.get("course", "AIMD")))}</b>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            with col2:

                if st.button(
                    "Resume" if existing_progress else "Start",
                    key=file,
                    disabled=not is_available
                ):

                    if existing_progress:

                        st.session_state.selected_assessment = file
                        apply_assessment_progress(existing_progress)

                    else:

                        file_path = (
                            f"assessment_files/{file}"
                        )

                        workbook_path = Path(file_path)
                        workbook_bytes = workbook_path.read_bytes()
                        workbook = pd.ExcelFile(
                            BytesIO(workbook_bytes),
                            engine="openpyxl"
                        )
                        configured_sheets = metadata.get(
                            "selected_sheets",
                            []
                        )

                        if (
                            configured_sheets
                            and all(
                                sheet in workbook.sheet_names
                                for sheet in configured_sheets
                            )
                        ):

                            df = read_selected_sheets(
                                workbook_bytes,
                                configured_sheets,
                                metadata.get(
                                    "course",
                                    st.session_state.selected_course
                                )
                            )

                        elif len(workbook.sheet_names) > 1:

                            df = read_selected_sheets(
                                workbook_bytes,
                                workbook.sheet_names,
                                metadata.get(
                                    "course",
                                    st.session_state.selected_course
                                )
                            )

                        else:

                            df = pd.read_excel(file_path)

                        df = df.dropna(how="all")
                        df = df.reset_index(drop=True)
                        df = standardize_question_bank(
                            df,
                            default_course=metadata.get(
                                "course",
                                st.session_state.selected_course
                            )
                        )
                        df, selection_warnings = select_random_questions(
                            df,
                            metadata.get(
                                "level_counts",
                                DEFAULT_LEVEL_DISTRIBUTION
                            ),
                            metadata.get("skill_counts", {}),
                            metadata.get(
                                "allowed_languages",
                                ["Python"]
                            ),
                            metadata.get(
                                "level_marks",
                                DEFAULT_LEVEL_MARKS
                            ),
                            seed=random.randint(1, 1000000)
                        )

                        st.session_state.assessment_df = df
                        st.session_state.selected_assessment = file
                        st.session_state.current_question = 0
                        st.session_state.score = 0
                        st.session_state.answers = {}
                        st.session_state.code_answers = {}
                        st.session_state.code_drafts = {}
                        st.session_state.run_outputs = {}
                        st.session_state.test_results = {}
                        st.session_state.test_passed = {}
                        st.session_state.question_time_spent = {}
                        st.session_state.active_question_index = 0
                        st.session_state.active_level_filter = "All Levels"
                        st.session_state.question_started_at = now_iso()
                        st.session_state.code_run_counts = {}
                        st.session_state.code_failed_attempts = {}
                        st.session_state.submission_timestamps = {}
                        st.session_state.proctoring_violations = []
                        st.session_state.processed_proctor_events = []
                        st.session_state.last_answer_feedback = {}
                        st.session_state.code_snapshots = {}
                        st.session_state.selection_warnings = (
                            selection_warnings
                        )
                        st.session_state.timer_started_at = ""
                        st.session_state.timer_ends_at = ""
                        st.session_state.timer_duration_minutes = 0
                        st.session_state.assessment_submitted = False
                        st.session_state.submitted_at = ""

                        save_assessment_progress()

                    st.session_state.page = "take_assessment"

                    st.rerun()

# ===================================================
# TAKE ASSESSMENT PAGE
# ===================================================
elif st.session_state.page == "take_assessment":

    df = st.session_state.assessment_df

    total_questions = len(df)

    # Safety Check
    if total_questions == 0:

        st.error("No questions available.")

        st.stop()

    current_q = st.session_state.current_question

    # Prevent Index Error
    if current_q >= total_questions:

        current_q = 0
        st.session_state.current_question = 0

    level_filter_options = get_level_filter_options(df)

    if st.session_state.active_level_filter not in level_filter_options:

        st.session_state.active_level_filter = "All Levels"

    selected_level_filter = render_level_selector(df)
    visible_question_indices = get_visible_question_indices(
        df,
        selected_level_filter
    )

    if not visible_question_indices:

        visible_question_indices = list(range(total_questions))

    if current_q not in visible_question_indices:

        current_q = visible_question_indices[0]
        st.session_state.current_question = current_q
        st.session_state.active_question_index = current_q
        st.session_state.question_started_at = now_iso()

    row = df.iloc[current_q]

    coding_question = is_coding_question(row)
    test_cases = get_test_cases(row) if coding_question else []
    assessment_metadata = load_assessment_metadata(
        st.session_state.selected_assessment
    )
    begin_question_tracking(current_q)
    process_proctoring_query_event()
    render_proctoring_guard(assessment_metadata, current_q)
    start_assessment_timer(assessment_metadata)
    remaining_seconds = get_remaining_seconds()
    time_expired = remaining_seconds == 0
    assessment_submitted = st.session_state.assessment_submitted

    (
        attempted,
        correct,
        incorrect,
        pending,
        percentage,
        earned_marks,
        total_marks
    ) = (
        get_assessment_stats(df)
    )

    with st.sidebar:
        st.title("Assessment Dashboard")

        if st.button("Back to Assessments"):

            save_assessment_progress()
            st.session_state.page = "student_assessment_selection"

            st.rerun()

        render_live_timer(
            remaining_seconds,
            st.session_state.timer_duration_minutes
            or assessment_metadata.get("duration_minutes", 0)
        )

        st.write(
            f"Assessment Ends: "
            f"{format_datetime(assessment_metadata.get('end_at'))}"
        )

        st.markdown("**Question Map**")
        st.markdown(
            """
            <div style="display:flex; gap:10px; flex-wrap:wrap; font-size:12px; margin-bottom:10px;">
                <span><b style="color:#10b981;">●</b> Submitted</span>
                <span><b style="color:#ef4444;">●</b> Pending</span>
                <span><b style="color:#818cf8;">●</b> Current</span>
            </div>
            """,
            unsafe_allow_html=True
        )

        question_map_items = []

        for question_index in visible_question_indices:

            if question_index == current_q:

                bg = "#6366f1"
                border = "2px solid #c7d2fe"
                shadow = "0 0 0 3px rgba(99, 102, 241, 0.24)"
                status = "Current"

            elif question_index in st.session_state.answers:

                bg = "#10b981"
                border = "2px solid rgba(255, 255, 255, 0.24)"
                shadow = "none"
                status = "Submitted"

            else:

                bg = "#ef4444"
                border = "2px solid rgba(255, 255, 255, 0.22)"
                shadow = "none"
                status = "Pending"

            question_map_items.append(
                f'<div title="Q{question_index + 1}: {status}" '
                f'style="width:34px;height:34px;border-radius:50%;'
                f'display:flex;align-items:center;justify-content:center;'
                f'box-sizing:border-box;background:{bg};color:#ffffff;'
                f'border:{border};box-shadow:{shadow};font-size:13px;'
                f'font-weight:800;line-height:1;text-align:center;'
                f'user-select:none;">{question_index + 1}</div>'
            )

        question_map_html = (
            '<div style="display:grid;grid-template-columns:repeat(5, 34px);'
            'gap:10px 13px;align-items:center;margin-top:12px;'
            'margin-bottom:12px;">'
            + "".join(question_map_items)
            + "</div>"
        )

        st.markdown(question_map_html, unsafe_allow_html=True)

    assessment_title = Path(
        str(st.session_state.selected_assessment)
    ).stem.replace("_", " ")
    visible_question_number = visible_question_indices.index(current_q) + 1
    visible_total_questions = len(visible_question_indices)

    st.markdown(
        f"""
        <div class="assessment-topbar">
            <div>
                <div class="assessment-kicker">
                    Question {visible_question_number} of {visible_total_questions}
                    ({html.escape(selected_level_filter)})
                </div>
                <h1 class="assessment-title">
                    {html.escape(assessment_title)}
                </h1>
            </div>
            <div class="score-strip">
                <div class="score-card">
                    <span>Score</span>
                    <strong>{earned_marks}/{total_marks}</strong>
                </div>
                <div class="score-card">
                    <span>Attempted</span>
                    <strong>{attempted}/{total_questions}</strong>
                </div>
                <div class="score-card">
                    <span>Pending</span>
                    <strong>{pending}</strong>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    progress_value = attempted / total_questions

    st.progress(progress_value)

    st.caption(
        f"{attempted} of {total_questions} questions attempted"
    )

    if st.session_state.last_proctor_alert:

        st.error(st.session_state.last_proctor_alert)

    if st.session_state.last_answer_feedback:

        feedback = st.session_state.last_answer_feedback
        feedback_message = feedback.get("message", "")

        if feedback.get("status") == "correct":

            st.success(feedback_message)

        elif feedback.get("status") == "incorrect":

            st.error(feedback_message)

        else:

            st.info(feedback_message)

    if st.session_state.selection_warnings:

        with st.expander("Question Selection Notes", expanded=False):

            for warning in st.session_state.selection_warnings:

                st.warning(warning)

    nav_index = (
        visible_question_indices.index(current_q)
        if current_q in visible_question_indices
        else 0
    )

    if (
        "question_nav_select" in st.session_state
        and st.session_state.question_nav_select not in visible_question_indices
    ):

        del st.session_state["question_nav_select"]

    selected_nav = st.selectbox(
        "Go to question",
        options=visible_question_indices,
        index=nav_index,
        format_func=lambda x: (
            f"Question {visible_question_indices.index(x) + 1} "
            f"({get_question_level(df.iloc[x])})"
        ),
        key="question_nav_select",
        label_visibility="collapsed"
    )

    if selected_nav != current_q:
        navigate_to_question(selected_nav)
        save_assessment_progress()
        st.rerun()

    if time_expired and not assessment_submitted:

        st.session_state.assessment_submitted = True
        st.session_state.submitted_at = datetime.now().isoformat(
            timespec="seconds"
        )

        save_assessment_progress()

        st.error(
            "Assessment time is over. Assessment has been automatically submitted."
        )

        st.success(
            f"""
            Assessment Completed.

            Final Score:
            {earned_marks}/{total_marks} ({percentage}%)
            """
        )

        st.stop()

    if assessment_submitted:

        st.success(
            f"Assessment submitted successfully. Final Score: "
            f"{earned_marks}/{total_marks} ({percentage}%)."
        )

    st.markdown("<br>", unsafe_allow_html=True)

    left, right = st.columns([1, 1])

    # ---------------------------------------------------
    # QUESTION SECTION
    # ---------------------------------------------------
    with left:
        question_text = str(row["Question"])

        st.markdown(
            f"""
        <div class="panel-card">
            <h3 class="panel-title">Problem</h3>
            <div class="question-text">
                <pre class="question-pre">{html.escape(question_text)}</pre>
            </div>
        </div>
        """,
            unsafe_allow_html=True
        )

        input_format = get_first_row_value(
            row,
            ["Input Format", "Input"],
            None
        )
        output_format = get_first_row_value(
            row,
            ["Output Format"],
            None
        )

        if coding_question and (input_format or output_format):

            st.markdown(
                """
                <div class="panel-card">
                    <h3 class="panel-title">Input / Output Format</h3>
                </div>
                """,
                unsafe_allow_html=True
            )

            if input_format:

                st.markdown("**Input Format**")
                st.code(str(input_format))

            if output_format:

                st.markdown("**Output Format**")
                st.code(str(output_format))

        render_visible_test_cases(test_cases)

        latest_test_results = st.session_state.test_results.get(
            current_q,
            []
        )
        render_revealed_hidden_test_cases(latest_test_results)

    # ---------------------------------------------------
    # ANSWER SECTION
    # ---------------------------------------------------
    coding_controls_rendered = False

    with right:

        if coding_question:

            st.markdown(
                """
                <div class="panel-card">
                    <h3 class="panel-title">Workspace</h3>
                </div>
                """,
                unsafe_allow_html=True
            )

            language_options = assessment_metadata.get(
                "allowed_languages",
                ["Python"]
            )

            language_options = [
                language
                for language in SUPPORTED_LANGUAGES
                if language in language_options
            ] or ["Python"]

            default_language = str(
                get_row_value(row, "Language", "Python")
            ).strip()

            if default_language not in language_options:

                default_language = language_options[0]

            saved_code_draft = st.session_state.code_drafts.get(
                current_q,
                {}
            )
            saved_code_answer = st.session_state.code_answers.get(
                current_q,
                {}
            )

            saved_language = (
                saved_code_draft.get("language")
                or saved_code_answer.get("language")
                or default_language
            )

            if saved_language not in language_options:

                saved_language = default_language

            language_key = f"language_{current_q}"

            if language_key not in st.session_state:

                st.session_state[language_key] = saved_language

            selected_language = st.selectbox(
                "Language",
                language_options,
                index=language_options.index(saved_language),
                key=language_key
            )

            starter_code_value = get_row_value(row, "Starter Code", "")

            if starter_code_value is None:

                starter_code = generate_default_starter_code(
                    row,
                    selected_language
                )

            else:

                starter_code = str(starter_code_value).strip()

                if starter_code == "":

                    starter_code = generate_default_starter_code(
                        row,
                        selected_language
                    )

            if (
                "code" in saved_code_draft
                and str(saved_code_draft.get("code", "")).strip() != ""
            ):

                saved_code = saved_code_draft.get("code", "")

            elif (
                "code" in saved_code_answer
                and str(saved_code_answer.get("code", "")).strip() != ""
            ):

                saved_code = saved_code_answer.get("code", "")

            else:

                saved_code = starter_code

            scaffold = get_logic_scaffold(row)

            if scaffold:

                st.markdown("**Default Code**")

                st.code(
                    build_code_from_logic(
                        scaffold.get("placeholder", ""),
                        scaffold
                    ),
                    language=selected_language.lower()
                )

                logic_editor_key = f"logic_editor_{current_q}"
                saved_logic = (
                    saved_code_draft.get("logic")
                    or saved_code_answer.get("logic")
                    or ""
                )

                if logic_editor_key not in st.session_state:

                    st.session_state[logic_editor_key] = saved_logic

                code_value = st.text_area(
                    "Logic Editor",
                    height=280,
                    key=logic_editor_key,
                    placeholder=scaffold.get("placeholder", "")
                )

                snapshot_key = f"logic_{current_q}"
                previous_snapshot = st.session_state.code_snapshots.get(
                    snapshot_key,
                    saved_logic
                )
                pasted_block_detected = is_probable_paste(
                    code_value,
                    previous_snapshot
                )

                if pasted_block_detected:

                    st.session_state[logic_editor_key] = previous_snapshot
                    log_proctoring_violation(
                        "paste",
                        current_q,
                        "Large pasted block detected in logic editor"
                    )
                    st.error("Copy-paste disabled.")
                    save_assessment_progress()
                    st.rerun()

                st.session_state.code_snapshots[snapshot_key] = code_value

                final_code_value = build_code_from_logic(
                    code_value,
                    scaffold
                )

                with st.expander("Generated full code", expanded=False):

                    st.code(
                        final_code_value,
                        language=selected_language.lower()
                    )

            else:

                code_editor_key = f"code_editor_{current_q}"

                if (
                    code_editor_key not in st.session_state
                    or str(st.session_state[code_editor_key]).strip() == ""
                ):

                    st.session_state[code_editor_key] = saved_code

                code_value = st.text_area(
                    "Code Editor",
                    height=460,
                    key=code_editor_key
                )

                snapshot_key = f"code_{current_q}"
                previous_snapshot = st.session_state.code_snapshots.get(
                    snapshot_key,
                    saved_code
                )
                pasted_block_detected = is_probable_paste(
                    code_value,
                    previous_snapshot
                )

                if pasted_block_detected:

                    st.session_state[code_editor_key] = previous_snapshot
                    log_proctoring_violation(
                        "paste",
                        current_q,
                        "Large pasted block detected in code editor"
                    )
                    st.error("Copy-paste disabled.")
                    save_assessment_progress()
                    st.rerun()

                st.session_state.code_snapshots[snapshot_key] = code_value

                final_code_value = code_value

            st.session_state.code_drafts[current_q] = {
                "language": selected_language,
                "code": final_code_value,
                "logic": code_value if scaffold else ""
            }

            save_assessment_progress()

            if test_cases:

                hidden_test_count = len(
                    [
                        test_case
                        for test_case in test_cases
                        if not test_case.get("visible", True)
                    ]
                )

                st.info(
                    f"{len(test_cases)} test case(s) will be checked, "
                    f"including {hidden_test_count} hidden test case(s)."
                )

            suggestions = get_code_suggestions(row, selected_language)

            if suggestions:

                with st.expander("Suggestions", expanded=False):

                    for suggestion in suggestions:

                        st.write(f"- {suggestion}")

            program_input = st.text_area(
                "Program Input",
                height=90,
                key=f"program_input_{current_q}",
                placeholder=(
                    "Optional custom input. Test cases use the input "
                    "configured by the admin."
                )
            )

            previous_visible_question = get_previous_visible_question(
                current_q,
                visible_question_indices
            )
            next_visible_question = get_next_visible_question(
                current_q,
                visible_question_indices
            )
            control_prev, control_next, run_col, submit_col = st.columns(4)
            coding_controls_rendered = True

            with control_prev:

                if st.button(
                    "Previous",
                    key=f"coding_previous_{current_q}",
                    disabled=previous_visible_question is None
                ):

                    navigate_to_question(previous_visible_question)
                    save_assessment_progress()
                    st.rerun()

            with control_next:

                if st.button(
                    "Next",
                    key=f"coding_next_{current_q}",
                    disabled=next_visible_question is None
                ):

                    navigate_to_question(next_visible_question)
                    save_assessment_progress()
                    st.rerun()

            with run_col:

                if st.button(
                    "Run Code",
                    key=f"run_code_{current_q}",
                    disabled=time_expired or assessment_submitted
                ):

                    st.session_state.code_answers[current_q] = {
                        "language": selected_language,
                        "code": final_code_value,
                        "logic": code_value if scaffold else ""
                    }
                    increment_counter(
                        "code_run_counts",
                        current_q
                    )

                    if test_cases:

                        test_results = run_code_tests(
                            selected_language,
                            final_code_value,
                            test_cases
                        )

                        st.session_state.test_results[current_q] = test_results
                        st.session_state.test_passed[current_q] = all(
                            result["passed"]
                            for result in test_results
                        )
                        if not st.session_state.test_passed[current_q]:

                            increment_counter(
                                "code_failed_attempts",
                                current_q
                            )
                        st.session_state.run_outputs[current_q] = (
                            format_test_results(test_results)
                        )

                    else:

                        st.session_state.test_results[current_q] = []
                        st.session_state.test_passed[current_q] = True
                        st.session_state.run_outputs[current_q] = (
                            run_student_code(
                                selected_language,
                                final_code_value,
                                program_input
                            )
                        )

                    save_assessment_progress()

                    st.rerun()

            with submit_col:

                latest_run = st.session_state.code_answers.get(
                    current_q,
                    {}
                )
                existing_submission = st.session_state.answers.get(
                    current_q,
                    {}
                )
                submitted_current_code = (
                    isinstance(existing_submission, dict)
                    and existing_submission.get("code") == final_code_value
                    and existing_submission.get("language") == selected_language
                )

                can_submit_code = (
                    not time_expired
                    and not assessment_submitted
                    and not submitted_current_code
                    and st.session_state.test_passed.get(current_q, False)
                    and latest_run.get("code") == final_code_value
                    and latest_run.get("language") == selected_language
                )

                if st.button(
                    "Submitted" if submitted_current_code else "Submit Code",
                    key=f"submit_code_{current_q}",
                    disabled=submitted_current_code or not can_submit_code
                ):

                    st.session_state.answers[current_q] = {
                        "language": selected_language,
                        "code": final_code_value,
                        "logic": code_value if scaffold else ""
                    }
                    st.session_state.last_answer_feedback = {
                        "status": "submitted",
                        "message": f"Question {current_q + 1}: Code submitted successfully."
                    }

                    st.session_state.code_answers[current_q] = {
                        "language": selected_language,
                        "code": final_code_value,
                        "logic": code_value if scaffold else "",
                        "test_results": st.session_state.test_results.get(
                            current_q,
                            []
                        )
                    }
                    st.session_state.submission_timestamps[str(current_q)] = (
                        now_iso()
                    )

                    save_assessment_progress()

                    target_level, target_question = get_auto_advance_target(
                        df,
                        current_q,
                        visible_question_indices,
                        selected_level_filter
                    )

                    if target_question is not None:

                        st.session_state.active_level_filter = target_level
                        navigate_to_question(target_question)
                        save_assessment_progress()
                        st.rerun()

                    st.rerun()

            if (
                current_q in st.session_state.answers
                and isinstance(st.session_state.answers[current_q], dict)
                and st.session_state.answers[current_q].get("code")
                == final_code_value
                and st.session_state.answers[current_q].get("language")
                == selected_language
            ):

                st.success("Code submitted successfully.")

            output_text = st.session_state.run_outputs.get(
                current_q,
                "Run code to see output or errors here."
            )

            st.markdown("**Output**")
            st.code(output_text)

            if test_cases and not st.session_state.test_passed.get(
                current_q,
                False
            ):

                st.warning(
                    "Submit Code will be enabled after all test cases pass."
                )

        else:

            options = get_mcq_options(row)

            if len(options) < 2:

                st.error(
                    "This MCQ needs at least two options in the question bank."
                )
                st.stop()

            question_key = f"question_{current_q}"
            mcq_already_submitted = current_q in st.session_state.answers

            if (
                mcq_already_submitted
                and question_key not in st.session_state
            ):

                st.session_state[question_key] = (
                    st.session_state.answers[current_q]
                )

            selected_option = st.radio(
                "Choose Option",
                options,
                key=question_key,
                disabled=mcq_already_submitted
            )
            next_visible_question = get_next_visible_question(
                current_q,
                visible_question_indices
            )

            if mcq_already_submitted:

                st.info("This MCQ has already been submitted.")

            if st.button(
                "Submit Answer",
                disabled=(
                    time_expired
                    or assessment_submitted
                    or mcq_already_submitted
                )
            ):

                correct_answer = get_row_value(row, "Correct Answer", "")

                if selected_option == correct_answer:

                    st.success("Correct Answer")
                    st.session_state.last_answer_feedback = {
                        "status": "correct",
                        "message": (
                            f"Question {current_q + 1}: Correct answer."
                        )
                    }

                    if current_q not in st.session_state.answers:

                        st.session_state.score += 1

                else:

                    st.error(
                        f"""
                        Incorrect Answer.

                        Correct Answer:
                        {correct_answer}
                        """
                    )
                    st.session_state.last_answer_feedback = {
                        "status": "incorrect",
                        "message": (
                            f"Question {current_q + 1}: Incorrect answer. "
                            f"Correct answer: {correct_answer}"
                        )
                    }

                st.session_state.answers[current_q] = selected_option
                st.session_state.submission_timestamps[str(current_q)] = (
                    now_iso()
                )

                save_assessment_progress()

                time_module.sleep(1.2)

                target_level, target_question = get_auto_advance_target(
                    df,
                    current_q,
                    visible_question_indices,
                    selected_level_filter
                )

                if target_question is not None:

                    st.session_state.active_level_filter = target_level
                    navigate_to_question(target_question)
                    save_assessment_progress()
                    st.rerun()

                st.rerun()

    (
        attempted,
        correct,
        incorrect,
        pending,
        percentage,
        earned_marks,
        total_marks
    ) = (
        get_assessment_stats(df)
    )

    st.markdown("<br>", unsafe_allow_html=True)

    if not coding_controls_rendered:

        previous_visible_question = get_previous_visible_question(
            current_q,
            visible_question_indices
        )
        next_visible_question = get_next_visible_question(
            current_q,
            visible_question_indices
        )
        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "Previous",
                disabled=previous_visible_question is None
            ):

                navigate_to_question(previous_visible_question)
                save_assessment_progress()

                st.rerun()

        with col2:

            if st.button(
                "Next",
                disabled=next_visible_question is None
            ):

                navigate_to_question(next_visible_question)
                save_assessment_progress()

                st.rerun()

    is_last_visible_question = current_q == visible_question_indices[-1]

    if is_last_visible_question or pending == 0:

        if pending > 0 and not assessment_submitted:

            st.warning(
                "Answer all questions to submit the assessment. "
                "Switch the Question Level filter to continue with other levels."
            )

        if st.button(
            "Submit Assessment",
            disabled=assessment_submitted or pending > 0
        ):

            update_question_time()
            st.session_state.assessment_submitted = True
            st.session_state.submitted_at = now_iso()

            save_assessment_progress()

            st.success(
                f"""
                Assessment Completed Successfully.

                Final Score:
                {earned_marks}/{total_marks} ({percentage}%)
                """
            )

    # ---------------------------------------------------
    # ASSESSMENT SUMMARY
    # ---------------------------------------------------
    st.markdown("**Assessment Summary**")

    summary_values = [
        ("Total", total_questions),
        ("Attempted", attempted),
        ("Correct", correct),
        ("Incorrect", incorrect),
        ("Pending", pending),
        ("Marks", f"{earned_marks}/{total_marks}"),
        ("Score", f"{percentage}%")
    ]

    summary_cols = st.columns(7)

    for summary_col, (label, value) in zip(summary_cols, summary_values):

        with summary_col:

            st.markdown(
                f"""
                <div style="
                    background:#ffffff;
                    border:1px solid #e5e7eb;
                    border-radius:10px;
                    padding:12px;
                    text-align:center;
                    min-height:76px;">
                    <div style="
                        color:#6b7280;
                        font-size:12px;
                        font-weight:600;
                        margin-bottom:8px;">
                        {label}
                    </div>
                    <div style="
                        color:#111827;
                        font-size:24px;
                        font-weight:700;
                        line-height:1;">
                        {value}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.markdown("<br>", unsafe_allow_html=True)

    st.button(
        "Logout",
        on_click=logout
    )

# ===================================================
# DOWNLOAD REPORTS PAGE
# ===================================================
elif st.session_state.page == "download_reports":

    with st.sidebar:

        st.title("Administrator")

        if st.button("Back to Dashboard"):

            st.session_state.page = "admin_dashboard"

            st.rerun()

        st.button(
            "Logout",
            on_click=logout
        )

    st.markdown(
        """
        <div class='dashboard-card'>
            <h2>Assessment Reports</h2>
        </div>
        """,
        unsafe_allow_html=True
    )

    report_records = load_all_progress_records()

    if not report_records:

        st.warning("No student assessment progress found yet.")

    else:

        report_df = pd.DataFrame(report_records)
        question_performance_df = pd.DataFrame(
            load_all_question_performance_records()
        )

        uploaded_test_rows = []

        for assessment_file in os.listdir("assessment_files"):

            if not assessment_file.endswith(".xlsx"):

                continue

            metadata = load_assessment_metadata(assessment_file)
            uploaded_test_rows.append(
                {
                    "Assessment": assessment_file,
                    "Assessment Name": (
                        metadata.get("assessment_name")
                        or Path(str(assessment_file)).stem.replace("_", " ")
                    ),
                    "Course": metadata.get("course", "AIMD")
                }
            )

        uploaded_tests_df = pd.DataFrame(uploaded_test_rows)
        available_courses = sorted(
            set(report_df["Course"].dropna().unique())
            | (
                set(uploaded_tests_df["Course"].dropna().unique())
                if not uploaded_tests_df.empty
                else set()
            )
        )

        selected_course_filter = st.selectbox(
            "Select Course",
            available_courses
        )

        course_report_df = report_df[
            report_df["Course"] == selected_course_filter
        ].copy()
        course_question_df = question_performance_df.copy()

        if not course_question_df.empty:

            course_question_df = course_question_df[
                course_question_df["Course"] == selected_course_filter
            ]

        submitted_count = int(
            (course_report_df["Submitted"] == "Yes").sum()
        )
        average_percentage = (
            course_report_df.loc[
                course_report_df["Submitted"] == "Yes",
                "Percentage"
            ].mean()
        )

        report_col1, report_col2, report_col3 = st.columns(3)

        with report_col1:

            st.metric("Students Tracked", len(course_report_df))

        with report_col2:

            st.metric("Submitted", submitted_count)

        with report_col3:

            st.metric(
                "Average Score",
                (
                    "0%"
                    if pd.isna(average_percentage)
                    else f"{average_percentage:.2f}%"
                )
            )

        st.markdown("### Tests Under Selected Course")

        report_test_files = (
            course_report_df[
                ["Assessment", "Assessment Name"]
            ]
            .drop_duplicates()
        )
        uploaded_test_files = (
            uploaded_tests_df[
                uploaded_tests_df["Course"] == selected_course_filter
            ][["Assessment", "Assessment Name"]]
            if not uploaded_tests_df.empty
            else pd.DataFrame(columns=["Assessment", "Assessment Name"])
        )
        test_files = (
            pd.concat(
                [uploaded_test_files, report_test_files],
                ignore_index=True
            )
            .drop_duplicates(subset=["Assessment"])
            .sort_values("Assessment Name")
        )

        if test_files.empty:

            st.warning("No test reports found for this course yet.")

        else:

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            for _, test_row in test_files.iterrows():

                assessment_file = test_row["Assessment"]
                assessment_name = test_row["Assessment Name"]
                test_report_df = course_report_df[
                    course_report_df["Assessment"] == assessment_file
                ].copy()
                test_question_df = (
                    course_question_df[
                        course_question_df["Assessment"] == assessment_file
                    ].copy()
                    if not course_question_df.empty
                    else pd.DataFrame()
                )
                test_submitted_count = int(
                    (test_report_df["Submitted"] == "Yes").sum()
                )
                test_average = test_report_df.loc[
                    test_report_df["Submitted"] == "Yes",
                    "Percentage"
                ].mean()
                test_col, status_col, download_col = st.columns([4, 2, 2])

                with test_col:

                    st.markdown(
                        f"""
                        <div class='metric-card'>
                            <h3>{html.escape(str(assessment_name))}</h3>
                            <div class='metric-pill'>
                                {html.escape(str(selected_course_filter))}
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                with status_col:

                    st.metric(
                        "Submitted",
                        test_submitted_count
                    )
                    st.caption(
                        (
                            "Average: 0%"
                            if pd.isna(test_average)
                            else f"Average: {test_average:.2f}%"
                        )
                    )

                with download_col:

                    if test_report_df.empty:

                        st.button(
                            "No Report Yet",
                            key=f"no_report_{assessment_file}",
                            disabled=True
                        )

                    else:

                        test_workbook = build_test_report_workbook(
                            test_report_df.to_dict("records"),
                            test_question_df.to_dict("records")
                        )
                        st.download_button(
                            "Download Report",
                            data=test_workbook,
                            file_name=(
                                f"{safe_file_name(selected_course_filter)}_"
                                f"{safe_file_name(assessment_name)}_"
                                f"Report_{timestamp}.xlsx"
                            ),
                            mime=(
                                "application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet"
                            ),
                            key=f"download_report_{assessment_file}"
                        ),

        st.markdown("### Course Analysis Preview")

        submitted_course_df = course_report_df[
            course_report_df["Submitted"] == "Yes"
        ]
        submitted_question_df = (
            course_question_df[
                course_question_df["Submitted"] == "Yes"
            ]
            if not course_question_df.empty
            else pd.DataFrame()
        )

        preview_col1, preview_col2 = st.columns(2)

        with preview_col1:

            institution_chart_df = (
                submitted_course_df
                .groupby("Institution Name", dropna=False)["Percentage"]
                .mean()
                .round(2)
                .reset_index(name="Average Percentage")
            )

            st.markdown("**Institution-wise Performance**")

            if institution_chart_df.empty:

                st.info("No submitted institution data yet.")

            else:

                st.bar_chart(
                    institution_chart_df,
                    x="Institution Name",
                    y="Average Percentage"
                )

        with preview_col2:

            skill_chart_df = build_grouped_question_performance(
                submitted_question_df,
                "Skill"
            )

            st.markdown("**Skill-wise Performance**")

            if skill_chart_df.empty:

                st.info("No skill data yet.")

            else:

                st.bar_chart(
                    skill_chart_df.sort_values(
                        "Average Percentage",
                        ascending=False
                    ),
                    x="Skill",
                    y="Average Percentage"
                )
