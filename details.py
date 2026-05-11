import streamlit as st
from datetime import date
import pandas as pd
import re
import os
import random

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
    "answers": {}
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

    st.markdown(
        """
        <div class='dashboard-card'>
            <h2>Administrator Dashboard</h2>
        </div>
        """,
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        if st.button("Upload Assessment"):

            st.session_state.page = "upload_assessment"

            st.rerun()

    with col2:

        if st.button("View Assessments"):

            st.session_state.page = "view_assessments"

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

    uploaded_file = st.file_uploader(
        "Upload Excel File",
        type=["xlsx"]
    )

    if uploaded_file is not None:

        excel_file = pd.ExcelFile(uploaded_file)

        sheet_names = excel_file.sheet_names

        selected_sheet = st.selectbox(
            "Select Sheet",
            sheet_names
        )

        st.info(
            """
            Required Excel Columns:

            Question No
            Question
            Option 1
            Option 2
            Option 3
            Option 4
            Correct Answer
            """
        )

        if st.button("Upload Assessment"):

            if assessment_name.strip() == "":

                st.error("Please enter Assessment Name.")

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

                    required_columns = [
                        "Question No",
                        "Question",
                        "Option 1",
                        "Option 2",
                        "Option 3",
                        "Option 4",
                        "Correct Answer"
                    ]

                    missing_columns = [
                        col for col in required_columns
                        if col not in df.columns
                    ]

                    if missing_columns:

                        st.error(
                            f"Missing columns: {missing_columns}"
                        )

                    else:

                        save_path = (
                            f"assessment_files/"
                            f"{assessment_name}_{selected_sheet}.xlsx"
                        )

                        df.to_excel(
                            save_path,
                            index=False
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

        for file in excel_files:

            st.markdown(
                f"""
                <div class='metric-card'>
                    <h3>{file.replace('.xlsx', '')}</h3>
                    <div class='metric-pill'>
                        Available
                    </div>
                </div>
                """,
                unsafe_allow_html=True
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

            with col1:

                st.markdown(
                    f"""
                    <div class='metric-card'>
                        <h3>{file.replace('.xlsx', '')}</h3>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            with col2:

                if st.button(
                    "Start",
                    key=file
                ):

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

    with st.sidebar:
        st.title("Assessment Dashboard")

    # ---------------------------------------------------
    # CALCULATIONS
    # ---------------------------------------------------
        attempted = len(st.session_state.answers)

        correct = st.session_state.score

        incorrect = attempted - correct

        pending = total_questions - attempted

        percentage = int(
            (correct / total_questions) * 100
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

    st.markdown("<br>", unsafe_allow_html=True)

    left, right = st.columns([2, 1])

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

    # ---------------------------------------------------
    # OPTIONS SECTION
    # ---------------------------------------------------
    with right:

        options = [
            row["Option 1"],
            row["Option 2"],
            row["Option 3"],
            row["Option 4"]
        ]

        selected_option = st.radio(
            "Choose Option",
            options,
            key=f"question_{current_q}"
        )

        if st.button("Submit Answer"):

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

    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    # ---------------------------------------------------
    # PREVIOUS BUTTON
    # ---------------------------------------------------
    with col1:

        if current_q > 0:

            if st.button("Previous"):

                st.session_state.current_question -= 1

                st.rerun()

    # ---------------------------------------------------
    # NEXT / FINISH BUTTON
    # ---------------------------------------------------
    with col2:

        if current_q < total_questions - 1:

            if st.button("Next"):

                st.session_state.current_question += 1

                st.rerun()

        else:

            if st.button("Finish Assessment"):

                st.success(
                    f"""
                    Assessment Completed Successfully.

                    Final Score:
                    {st.session_state.score}/{total_questions}
                    """
                )

    # ---------------------------------------------------
    # ASSESSMENT SUMMARY
    # ---------------------------------------------------
    st.markdown(
        f"""
        <div style="
            background-color:white;
            padding:18px;
            border-radius:16px;
            border:1px solid #e5e7eb;
            margin-top:24px;
        ">

        <div style="font-size:16px;
                    font-weight:600;
                    margin-bottom:10px;
                    color:#111827;">
            Assessment Summary
        </div>

        <div style="line-height:2;
                    font-size:15px;
                    color:#374151;">

        Total Questions: <b>{total_questions}</b><br>

        Attempted: <b>{attempted}</b><br>

        Correct: <b>{correct}</b><br>

        Incorrect: <b>{incorrect}</b><br>

        Pending: <b>{pending}</b><br>

        Score: <b>{percentage}%</b>

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
