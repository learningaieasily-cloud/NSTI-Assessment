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
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    from openpyxl.chart import BarChart, Reference
except ImportError:
    BarChart = None
    Reference = None

# ---------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------
st.set_page_config(
    page_title="NSTI Assessment LMS",
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

    [data-testid="stAppViewContainer"] {
        background: linear-gradient(180deg, #eef2ff 0%, #f8fafc 55%, #ffffff 100%);
    }

    .dashboard-card {
        background: #ffffff;
        border-radius: 24px;
        padding: 24px;
        margin-bottom: 20px;
        box-shadow: 0px 10px 30px rgba(0,0,0,0.05);
        border: 1px solid #e2e8f0;
    }

    .hero-card {
        background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);
        color: white;
    }

    .metric-card {
        background: #f8fafc;
        border-radius: 20px;
        padding: 20px;
        border: 1px solid #e2e8f0;
        margin-bottom: 15px;
    }

    .metric-pill {
        display: inline-block;
        margin-top: 10px;
        background: rgba(99,102,241,0.12);
        color: #4f46e5;
        padding: 6px 12px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 600;
    }

    .stButton>button {
        background-color: #6366f1;
        color: white;
        border-radius: 12px;
        padding: 0.7rem 1.5rem;
        border: none;
        font-weight: 600;
        width: 100%;
    }

    .stButton>button:hover {
        background-color: #4f46e5;
    }

    /* Question Styling */
    .question-text {
        font-size: 18px;
        font-weight: 400;
        color: #111827;
        line-height: 1.7;
        margin-bottom: 20px;
    }

    /* Increase spacing between options */
    div[role="radiogroup"] > label {
        display: flex !important;
        align-items: center !important;
        gap: 14px !important;
        width: 100% !important;
        max-width: 100% !important;
        box-sizing: border-box !important;
        margin-bottom: 18px !important;
        padding: 14px !important;
        border-radius: 10px;
        border: 1px solid #e5e7eb;
        background-color: #ffffff;
        font-size: 15px;
        font-weight: 400;
    }

    /* Hover effect */
    div[role="radiogroup"] > label:hover {
        background-color: #f8fafc;
    }

    textarea {
        font-family: Consolas, monospace !important;
        font-size: 15px !important;
        line-height: 1.6 !important;
    }

    .run-output {
        background: #111827;
        color: #f9fafb;
        padding: 16px;
        border-radius: 12px;
        white-space: pre-wrap;
        font-family: Consolas, monospace;
        font-size: 14px;
        line-height: 1.6;
        min-height: 80px;
        border: 1px solid #374151;
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
    "admin_email": "",
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

    st.rerun()

def is_valid_email(email):

    pattern = r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'

    return re.match(pattern, email)

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
        <div style="
            background:#111827;
            color:#f9fafb;
            padding:16px 18px;
            border-radius:12px;
            border:1px solid #374151;
            margin-bottom:14px;
            font-family:Arial, sans-serif;">
            <div style="font-size:13px; color:#cbd5e1; margin-bottom:6px;">
                Assessment Timer
            </div>
            <div style="display:flex; justify-content:space-between; gap:16px; align-items:center;">
                <div>
                    <div id="timer-value" style="font-size:30px; font-weight:700;">
                        {format_remaining_time(remaining_seconds)}
                    </div>
                    <div style="font-size:13px; color:#cbd5e1;">
                        Time remaining
                    </div>
                </div>
                <div style="text-align:right;">
                    <div style="font-size:20px; font-weight:700;">
                        {format_duration(duration_minutes)}
                    </div>
                    <div style="font-size:13px; color:#cbd5e1;">
                        Total duration
                    </div>
                </div>
            </div>
            <div style="background:#374151; height:8px; border-radius:999px; margin-top:14px; overflow:hidden;">
                <div id="timer-bar" style="
                    background:#22c55e;
                    height:8px;
                    width:{max(0, min(100, int((remaining_seconds / total_seconds) * 100)))}%;
                    border-radius:999px;">
                </div>
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
        height=150
    )

def render_visible_test_cases(test_cases):

    if not test_cases:

        return

    st.markdown("**Test Cases**")

    for index, test_case in enumerate(test_cases, start=1):

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

    single_input = get_row_value(row, "Test Input", None)
    single_output = get_row_value(row, "Expected Output", None)

    if single_output is not None:

        test_cases.append(
            {
                "input": str(single_input or ""),
                "expected": str(single_output)
            }
        )

    for index in range(1, 6):

        test_input = get_row_value(row, f"Test Input {index}", None)
        expected_output = get_row_value(row, f"Expected Output {index}", None)

        if expected_output is not None:

            test_cases.append(
                {
                    "input": str(test_input or ""),
                    "expected": str(expected_output)
                }
            )

    return test_cases

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
                "passed": passed
            }
        )

    return results

def format_test_results(test_results):

    if not test_results:

        return "No test cases found. Custom run completed."

    lines = []

    for result in test_results:

        status = "PASS" if result["passed"] else "FAIL"

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

        student_id = st.text_input("Student ID")

        student_email = st.text_input("Student Email")

        if st.button("Login as Student"):

            if student_id.strip() == "":

                st.error("Please enter Student ID.")

            elif not is_valid_email(student_email):

                st.error("Please enter valid Student Email.")

            else:

                st.session_state.student_id = student_id
                st.session_state.student_email = student_email
                st.session_state.page = "student_details"

                st.rerun()

    # ---------------------------------------------------
    # ADMIN LOGIN
    # ---------------------------------------------------
    elif role == "Administrator":

        admin_email = st.text_input("Administrator Email")

        if st.button("Login as Administrator"):

            if admin_email.lower().endswith("@edunetfoundation.org"):

                st.session_state.admin_email = admin_email
                st.session_state.page = "admin_dashboard"

                st.rerun()

            else:

                st.error(
                    "Only @edunetfoundation.org emails allowed."
                )

# ===================================================
# STUDENT DETAILS PAGE
# ===================================================
elif st.session_state.page == "student_details":

    with st.sidebar:

        st.title("Student")

        st.write(
            f"Student ID: {st.session_state.student_id}"
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

        if full_name.strip() == "":

            st.error("Please enter Student Full Name.")

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
            Test Input 1
            Expected Output 1
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

                confirm_delete = st.checkbox(
                    "Confirm",
                    key=f"confirm_delete_{file}"
                )

            with col3:

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

                        os.remove(file_path)

                        metadata_path = metadata_file_path(file)

                        if os.path.exists(metadata_path):

                            os.remove(metadata_path)

                        st.success(
                            f"{file} deleted successfully."
                        )

                        st.rerun()

                    except Exception as e:

                        st.error(
                            f"Could not delete assessment: {e}"
                        )

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

    with st.sidebar:
        st.title("Assessment Dashboard")

        if st.button("Back to Assessments"):

            save_assessment_progress()
            st.session_state.page = "student_assessment_selection"

            st.rerun()

        st.write(
            f"Time Remaining: {format_remaining_time(remaining_seconds)}"
        )
        st.write(
            f"Assessment Ends: "
            f"{format_datetime(assessment_metadata.get('end_at'))}"
        )

    # ---------------------------------------------------
    # CALCULATIONS
    # ---------------------------------------------------
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

    # ---------------------------------------------------
    # QUESTION TRACKER
    # ---------------------------------------------------
    st.write(
        f"Current Question: "
        f"{current_q + 1}/{total_questions}"
    )

    # ---------------------------------------------------
    # PROGRESS BAR
    # ---------------------------------------------------
    progress_value = attempted / total_questions

    st.progress(progress_value)

    st.write(
        f"{attempted} of {total_questions} Questions Attempted"
    )

    render_live_timer(
        remaining_seconds,
        st.session_state.timer_duration_minutes
        or assessment_metadata.get("duration_minutes", 0)
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
        <div class="question-text">
<pre style="
white-space: pre-wrap;
font-size:16px;
font-family: Consolas, monospace;
line-height:1.6;
background-color:#f8fafc;
padding:18px;
border-radius:12px;
border:1px solid #e5e7eb;
">
{question_text}
</pre>
        </div>
        """,
        unsafe_allow_html=True
    )

        render_visible_test_cases(test_cases)

    # ---------------------------------------------------
    # ANSWER SECTION
    # ---------------------------------------------------
    with right:

        if coding_question:

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

                starter_code = ""

            else:

                starter_code = str(starter_code_value)

            if "code" in saved_code_draft:

                saved_code = saved_code_draft.get("code", "")

            elif "code" in saved_code_answer:

                saved_code = saved_code_answer.get("code", "")

            else:

                saved_code = starter_code

            code_editor_key = f"code_editor_{current_q}"

            if code_editor_key not in st.session_state:

                st.session_state[code_editor_key] = saved_code

            code_value = st.text_area(
                "Code Editor",
                value=saved_code,
                height=460,
                key=code_editor_key
            )

            st.session_state.code_drafts[current_q] = {
                "language": selected_language,
                "code": code_value
            }

            save_assessment_progress()

            if test_cases:

                st.info(
                    f"{len(test_cases)} test case(s) will be checked."
                )

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
                        "code": code_value
                    }

                    if test_cases:

                        test_results = run_code_tests(
                            selected_language,
                            code_value,
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
                                code_value,
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
                    and latest_run.get("code") == code_value
                    and latest_run.get("language") == selected_language
                )

                if st.button(
                    "Submit Code",
                    key=f"submit_code_{current_q}",
                    disabled=not can_submit_code
                ):

                    st.session_state.answers[current_q] = {
                        "language": selected_language,
                        "code": code_value
                    }

                    st.session_state.code_answers[current_q] = {
                        "language": selected_language,
                        "code": code_value,
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
    st.markdown("### Question Status")

    # Legend
    st.markdown(
        """
        <div style="display:flex; gap:18px; margin-bottom:10px; font-size:13px;">
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
    <div style="display:flex; flex-wrap:wrap; gap:8px; margin-bottom:8px;">
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
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 12px;
                font-weight: 700;
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

    # Calculate iframe height: ~44px per row of 15 bubbles + padding
    bubble_rows = max(1, (total_questions + 14) // 15)
    bubble_height = bubble_rows * 52 + 20

    components.html(bubble_html, height=bubble_height, scrolling=False)

    # Hidden selectbox — receives click events from the bubbles above
    nav_index = current_q if current_q < total_questions else 0

    selected_nav = st.selectbox(
        "question_nav_select",
        options=list(range(total_questions)),
        index=nav_index,
        format_func=lambda x: f"Question {x + 1}",
        key="question_nav_select",
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