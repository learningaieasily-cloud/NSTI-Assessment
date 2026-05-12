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
from pathlib import Path
from email.message import EmailMessage

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
    "selected_nsti": "",
    "location": "",
    "assessment_date": date.today(),
    "selected_assessment": "",
    "selected_sheet": "",
    "assessment_df": pd.DataFrame(),
    "current_question": 0,
    "score": 0,
    "correct_questions": set(),
    "code_drafts": {},
    "run_outputs": {},
    "test_results": {},
    "test_passed": {},
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

def save_assessment_progress():

    if (
        st.session_state.student_id == ""
        or st.session_state.student_email == ""
        or st.session_state.selected_assessment == ""
        or st.session_state.assessment_df.empty
    ):

        return

    progress_path = progress_file_path(
        st.session_state.student_id,
        st.session_state.student_email,
        st.session_state.selected_assessment
    )

    progress_data = {
        "student_id": st.session_state.student_id,
        "student_email": st.session_state.student_email,
        "full_name": st.session_state.full_name,
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

        return progress_data

    except Exception:

        return None

def apply_assessment_progress(progress_data):

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

    if question_type in ["coding", "code", "programming"]:

        return True

    mcq_columns = [
        "Option 1",
        "Option 2",
        "Option 3",
        "Option 4",
        "Correct Answer"
    ]

    return not all(column in row.index for column in mcq_columns)

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

    if single_output is not None:

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

            records.append(
                {
                    "Student ID": progress_data.get("student_id", ""),
                    "Student Name": progress_data.get("full_name", ""),
                    "Email": progress_data.get("student_email", ""),
                    "NSTI": progress_data.get("nsti", ""),
                    "Location": progress_data.get("location", ""),
                    "Assessment Date": progress_data.get(
                        "assessment_date",
                        ""
                    ),
                    "Assessment": progress_data.get("assessment", ""),
                    "Submitted": (
                        "Yes"
                        if progress_data.get("assessment_submitted", False)
                        else "No"
                    ),
                    "Submitted At": progress_data.get("submitted_at", ""),
                    "Total Questions": len(
                        progress_data.get("assessment_records", [])
                    ),
                    "Attempted": attempted,
                    "Correct": correct,
                    "Incorrect": incorrect,
                    "Pending": pending,
                    "Marks Obtained": earned_marks,
                    "Total Marks": total_marks,
                    "Percentage": percentage
                }
            )

        except Exception:

            continue

    return records

def build_report_workbook(records):

    output = BytesIO()
    report_df = pd.DataFrame(records)

    if report_df.empty:

        report_df = pd.DataFrame(
            columns=[
                "Student ID",
                "Student Name",
                "Email",
                "NSTI",
                "Location",
                "Assessment Date",
                "Assessment",
                "Submitted",
                "Submitted At",
                "Total Questions",
                "Attempted",
                "Correct",
                "Incorrect",
                "Pending",
                "Marks Obtained",
                "Total Marks",
                "Percentage"
            ]
        )

    submitted_df = report_df[report_df["Submitted"] == "Yes"]

    if submitted_df.empty:

        nsti_df = pd.DataFrame(
            columns=[
                "NSTI",
                "Students Submitted",
                "Average Marks",
                "Average Percentage",
                "Highest Marks",
                "Lowest Marks"
            ]
        )

    else:

        nsti_df = (
            submitted_df
            .groupby("NSTI", dropna=False)
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

        nsti_df["Average Marks"] = nsti_df["Average Marks"].round(2)
        nsti_df["Average Percentage"] = (
            nsti_df["Average Percentage"].round(2)
        )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        report_df.to_excel(
            writer,
            sheet_name="Student Report",
            index=False
        )
        nsti_df.to_excel(
            writer,
            sheet_name="NSTI Analysis",
            index=False
        )

        workbook = writer.book

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

        analysis_sheet = workbook["NSTI Analysis"]

        if len(nsti_df) > 0:

            marks_chart = BarChart()
            marks_chart.title = "NSTI-wise Average Marks"
            marks_chart.y_axis.title = "Average Marks"
            marks_chart.x_axis.title = "NSTI"
            marks_data = Reference(
                analysis_sheet,
                min_col=3,
                min_row=1,
                max_row=len(nsti_df) + 1
            )
            marks_categories = Reference(
                analysis_sheet,
                min_col=1,
                min_row=2,
                max_row=len(nsti_df) + 1
            )
            marks_chart.add_data(marks_data, titles_from_data=True)
            marks_chart.set_categories(marks_categories)
            marks_chart.height = 8
            marks_chart.width = 16
            analysis_sheet.add_chart(marks_chart, "H2")

            percentage_chart = BarChart()
            percentage_chart.title = "NSTI-wise Average Percentage"
            percentage_chart.y_axis.title = "Average %"
            percentage_chart.x_axis.title = "NSTI"
            percentage_data = Reference(
                analysis_sheet,
                min_col=4,
                min_row=1,
                max_row=len(nsti_df) + 1
            )
            percentage_chart.add_data(
                percentage_data,
                titles_from_data=True
            )
            percentage_chart.set_categories(marks_categories)
            percentage_chart.height = 8
            percentage_chart.width = 16
            analysis_sheet.add_chart(percentage_chart, "H20")

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

def execute_student_code(language, code, program_input=""):

    language = language.lower()

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
            <h1>NSTI Assessment LMS</h1>
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

    with st.form("student_form"):

        full_name = st.text_input("Student Full Name")

        selected_nsti = st.selectbox(
            "Select NSTI",
            ["Select NSTI"] + nsti_list
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

        elif selected_nsti == "Select NSTI":

            st.error("Please select NSTI.")

        elif location.strip() == "":

            st.error("Please enter Location.")

        else:

            st.session_state.full_name = full_name
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
        <div class='dashboard-card'>
            <h2>Administrator Dashboard</h2>
        </div>
        """,
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        if st.button("Upload Assessment"):

            st.session_state.page = "upload_assessment"

            st.rerun()

    with col2:

        if st.button("View Assessments"):

            st.session_state.page = "view_assessments"

            st.rerun()

    with col3:

        if st.button("Download Reports"):

            st.session_state.page = "download_reports"

            st.rerun()

    st.button(
        "Logout",
        on_click=logout
    )

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

    assessment_name = st.text_input(
        "Assessment Name"
    )

    schedule_col1, schedule_col2 = st.columns(2)

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

    duration_minutes = st.selectbox(
        "Assessment Duration in Minutes",
        [15, 30, 45, 60, 75, 90, 120, 150, 180],
        index=3,
        format_func=lambda value: format_duration(value)
    )

    uploaded_file = st.file_uploader(
        "Upload Excel File",
        type=["xlsx"]
    )

    if uploaded_file is not None:

        if BarChart is None or Reference is None:
            st.error(
                "The openpyxl package is not installed in this environment. "
                "Excel upload and report generation require openpyxl. "
                "Please add openpyxl to requirements.txt and redeploy."
            )
            st.stop()

        try:
            excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
        except ImportError:
            st.error(
                "openpyxl is required to read Excel files. "
                "Please install openpyxl and redeploy the app."
            )
            st.stop()

        sheet_names = excel_file.sheet_names

        selected_sheet = st.selectbox(
            "Select Sheet",
            sheet_names
        )

        st.info(
            """
            Required Excel Columns:

            Question

            For MCQ questions:
            Option 1
            Option 2
            Option 3
            Option 4
            Correct Answer

            For coding questions:
            Question Type = Coding
            Language
            Starter Code
            Logic Prefix
            Logic Suffix
            Test Input 1
            Expected Output 1
            Hidden Test Input 1
            Hidden Expected Output 1

            Hidden output column aliases are also supported:
            Hidden Test Output 1
            Hidden Output 1
            Private Test Output 1
            """
        )

        if st.button("Upload Assessment"):

            if assessment_name.strip() == "":

                st.error("Please enter Assessment Name.")

            elif datetime.combine(end_date, end_time) <= datetime.combine(
                start_date,
                start_time
            ):

                st.error(
                    "Assessment end date/time must be after start date/time."
                )

            else:

                try:

                    df = pd.read_excel(
                        uploaded_file,
                        sheet_name=selected_sheet,
                        header=2
                    )

                    # Remove empty rows
                    df = df.dropna(how="all")

                    # Reset index
                    df = df.reset_index(drop=True)

                    mcq_columns = [
                        "Option 1",
                        "Option 2",
                        "Option 3",
                        "Option 4",
                        "Correct Answer"
                    ]

                    has_question_column = "Question" in df.columns

                    has_mcq_columns = all(
                        col in df.columns
                        for col in mcq_columns
                    )

                    has_coding_column = (
                        "Question Type" in df.columns
                        and df["Question Type"]
                        .astype(str)
                        .str.lower()
                        .str.contains("coding|code|programming")
                        .any()
                    )

                    if not has_question_column:

                        st.error(
                            "Missing required column: Question"
                        )

                    elif not has_mcq_columns and not has_coding_column:

                        st.error(
                            "For coding assessments, add a Question Type "
                            "column with value Coding. For MCQ assessments, "
                            "include Option 1 to Option 4 and Correct Answer."
                        )

                    else:

                        saved_file_name = (
                            f"{assessment_name}_{selected_sheet}.xlsx"
                        )

                        save_path = os.path.join(
                            "assessment_files",
                            saved_file_name
                        )

                        df.to_excel(
                            save_path,
                            index=False
                        )

                        save_assessment_metadata(
                            saved_file_name,
                            {
                                "assessment_name": assessment_name,
                                "sheet_name": selected_sheet,
                                "start_at": datetime.combine(
                                    start_date,
                                    start_time
                                ).isoformat(timespec="seconds"),
                                "end_at": datetime.combine(
                                    end_date,
                                    end_time
                                ).isoformat(timespec="seconds"),
                                "duration_minutes": int(duration_minutes)
                            }
                        )

                        st.success(
                            "Assessment uploaded successfully."
                        )

                        st.dataframe(df.head())

                except Exception as e:

                    st.error(f"Error: {e}")

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
                            Duration: <b>{format_duration(metadata.get("duration_minutes"))}</b>
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

    if len(excel_files) == 0:

        st.warning("No assessments available.")

    else:

        for file in excel_files:

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
                            Duration: <b>{format_duration(metadata.get("duration_minutes"))}</b>
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

                        df = pd.read_excel(file_path)

                        # Remove empty rows
                        df = df.dropna(how="all")

                        # Reset index
                        df = df.reset_index(drop=True)

                        # Randomly select 30 questions
                        if len(df) > 30:

                            df = df.sample(
                                n=30,
                                random_state=random.randint(1, 10000)
                            )

                            df = df.reset_index(drop=True)

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

    row = df.iloc[current_q]

    coding_question = is_coding_question(row)
    test_cases = get_test_cases(row) if coding_question else []
    assessment_metadata = load_assessment_metadata(
        st.session_state.selected_assessment
    )
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

        for question_index in range(total_questions):

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
                f'style="width:28px;height:28px;border-radius:50%;'
                f'display:flex;align-items:center;justify-content:center;'
                f'box-sizing:border-box;background:{bg};color:#ffffff;'
                f'border:{border};box-shadow:{shadow};font-size:11px;'
                f'font-weight:800;line-height:1;text-align:center;'
                f'user-select:none;">{question_index + 1}</div>'
            )

        question_map_html = (
            '<div style="display:grid;grid-template-columns:repeat(5, 28px);'
            'gap:8px 10px;align-items:center;margin-top:10px;'
            'margin-bottom:12px;">'
            + "".join(question_map_items)
            + "</div>"
        )

        st.markdown(question_map_html, unsafe_allow_html=True)

    assessment_title = Path(
        str(st.session_state.selected_assessment)
    ).stem.replace("_", " ")

    st.markdown(
        f"""
        <div class="assessment-topbar">
            <div>
                <div class="assessment-kicker">
                    Question {current_q + 1} of {total_questions}
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

        render_visible_test_cases(test_cases)

        latest_test_results = st.session_state.test_results.get(
            current_q,
            []
        )
        render_revealed_hidden_test_cases(latest_test_results)

    # ---------------------------------------------------
    # ANSWER SECTION
    # ---------------------------------------------------
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

            language_options = [
                "Python",
                "Java",
                "JavaScript",
                "C",
                "C++"
            ]

            default_language = str(
                get_row_value(row, "Language", "Python")
            ).strip()

            if default_language not in language_options:

                default_language = "Python"

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

            run_col, submit_col = st.columns(2)

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

                can_submit_code = (
                    not time_expired
                    and not assessment_submitted
                    and st.session_state.test_passed.get(current_q, False)
                    and latest_run.get("code") == final_code_value
                    and latest_run.get("language") == selected_language
                )

                if st.button(
                    "Submit Code",
                    key=f"submit_code_{current_q}",
                    disabled=not can_submit_code
                ):

                    st.session_state.answers[current_q] = {
                        "language": selected_language,
                        "code": final_code_value,
                        "logic": code_value if scaffold else ""
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

                    save_assessment_progress()

                    if current_q < total_questions - 1:

                        st.session_state.current_question += 1
                        save_assessment_progress()
                        st.rerun()

                    else:

                        save_assessment_progress()
                        st.success("Assessment Completed Successfully.")

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

            options = [
                row["Option 1"],
                row["Option 2"],
                row["Option 3"],
                row["Option 4"]
            ]

            question_key = f"question_{current_q}"

            if (
                current_q in st.session_state.answers
                and question_key not in st.session_state
            ):

                st.session_state[question_key] = (
                    st.session_state.answers[current_q]
                )

            selected_option = st.radio(
                "Choose Option",
                options,
                key=question_key
            )

            if st.button(
                "Submit Answer",
                disabled=time_expired or assessment_submitted
            ):

                correct_answer = row["Correct Answer"]

                if selected_option == correct_answer:

                    st.success("Correct Answer")

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

                st.session_state.answers[current_q] = selected_option

                save_assessment_progress()

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

    col1, col2, col3 = st.columns(3)

    # ---------------------------------------------------
    # PREVIOUS BUTTON
    # ---------------------------------------------------
    with col1:

        if current_q > 0:

            if st.button("Previous"):

                st.session_state.current_question -= 1
                save_assessment_progress()

                st.rerun()

    # ---------------------------------------------------
    # NEXT / FINISH BUTTON
    # ---------------------------------------------------
    with col2:

        if current_q < total_questions - 1:

            if st.button("Next"):

                st.session_state.current_question += 1
                save_assessment_progress()

                st.rerun()

        else:

            if pending > 0 and not assessment_submitted and current_q < total_questions - 1:

                st.warning(
                    "Answer all questions to submit the assessment."
                )

            if st.button(
                "Submit Assessment",
                disabled=assessment_submitted or (pending > 0 and current_q < total_questions - 1)
            ):

                st.session_state.assessment_submitted = True
                st.session_state.submitted_at = datetime.now().isoformat(
                    timespec="seconds"
                )

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

# ---------------------------------------------------
    # QUESTION NAVIGATION BUBBLES
    # Rendered as pure HTML via components.html so we get
    # full CSS control: green = answered, red = pending,
    # indigo ring = current question.
    # Clicking a bubble sets a hidden selectbox which
    # Streamlit picks up to trigger navigation.
    # ---------------------------------------------------
    st.markdown("")

    # Legend (Changed display:none to display:flex so it becomes visible)
    st.markdown(
        """
        <div style="display:flex; gap:18px; margin-bottom:15px; font-size:13px;">
            <span>
                <span style="
                    display:inline-block;
                    width:14px; height:14px;
                    background:#10b981;
                    border-radius:50%;
                    vertical-align:middle;
                    margin-right:5px;">
                </span>
                Submitted
            </span>
            <span>
                <span style="
                    display:inline-block;
                    width:14px; height:14px;
                    background:#ef4444;
                    border-radius:50%;
                    vertical-align:middle;
                    margin-right:5px;">
                </span>
                Pending
            </span>
            <span>
                <span style="
                    display:inline-block;
                    width:14px; height:14px;
                    background:#6366f1;
                    border-radius:50%;
                    vertical-align:middle;
                    margin-right:5px;">
                </span>
                Current
            </span>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Build bubble HTML
    bubble_html = """
    <div style="display:flex; flex-wrap:wrap; gap:8px; margin-bottom:8px; align-items:flex-start; font-family: sans-serif;">
    """

    for i in range(total_questions):

        if i == current_q:
            # Current question — indigo fill
            bg = "#6366f1"
            border = "3px solid #312e81"
            status = "Current"
        elif i in st.session_state.answers:
            # Answered/submitted
            bg = "#10b981"
            border = "3px solid transparent"
            status = "Submitted"
        else:
            # Pending
            bg = "#ef4444"
            border = "3px solid transparent"
            status = "Pending"

        bubble_html += f"""
        <div
            title="Q{i + 1}: {status}"
            onclick="
                var sel = window.parent.document.querySelectorAll('select');
                for (var s = 0; s < sel.length; s++) {{
                    if (sel[s].getAttribute('aria-label') === 'question_nav_select' ||
                        sel[s].id.includes('question_nav_select_legacy') ||
                        sel[s].id.includes('question_nav_select')) {{
                        sel[s].value = '{i}';
                        sel[s].dispatchEvent(new Event('change', {{bubbles: true}}));
                        break;
                    }}
                }}
                // Fallback: try the last select on page
                var allSel = window.parent.document.querySelectorAll('select');
                if (allSel.length > 0) {{
                    var last = allSel[allSel.length - 1];
                    last.value = '{i}';
                    last.dispatchEvent(new Event('change', {{bubbles: true}}));
                }}
            "
            style="
                background: {bg};
                color: white;
                border: {border};
                border-radius: 50%;
                width: 36px;
                height: 36px;
                display: inline-flex;
                align-items: center;
                justify-content: center;
                font-size: 12px;
                font-weight: 700;
                line-height: 1;
                cursor: pointer;
                box-sizing: border-box;
                user-select: none;
                transition: transform 0.1s;
            "
            onmouseover="this.style.transform='scale(1.15)'"
            onmouseout="this.style.transform='scale(1.0)'"
        >
            {i + 1}
        </div>
        """

    bubble_html += "</div>"

    # Calculate iframe height dynamically: ~46px per row of 15 bubbles + padding buffer
    bubble_rows = max(1, (total_questions + 14) // 15)
    bubble_height = (bubble_rows * 46) + 12 

    components.html(bubble_html, height=bubble_height, scrolling=False)

    # Hidden selectbox — receives click events from the bubbles above
    nav_index = current_q if current_q < total_questions else 0

    selected_nav = st.selectbox(
        "question_nav_select_legacy",
        options=list(range(total_questions)),
        index=nav_index,
        format_func=lambda x: f"Question {x + 1}",
        key="question_nav_select_legacy",
        label_visibility="collapsed"
    )

    if selected_nav != current_q:
        st.session_state.current_question = selected_nav
        save_assessment_progress()
        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

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

        submitted_count = int(
            (report_df["Submitted"] == "Yes").sum()
        )
        average_percentage = (
            report_df.loc[
                report_df["Submitted"] == "Yes",
                "Percentage"
            ].mean()
        )

        report_col1, report_col2, report_col3 = st.columns(3)

        with report_col1:

            st.metric("Students Tracked", len(report_df))

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

        st.dataframe(report_df, use_container_width=True)

        report_workbook = build_report_workbook(report_records)

        st.download_button(
            "Download Excel Report",
            data=report_workbook,
            file_name=(
                "NSTI_Assessment_Report_"
                f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )
