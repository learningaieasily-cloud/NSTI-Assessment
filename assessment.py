# ---------------------------------------------------
# ADD THESE IMPORTS AT THE TOP
# ---------------------------------------------------

import pandas as pd
import os

# ---------------------------------------------------
# CREATE ASSESSMENT DIRECTORY
# ADD THIS BELOW NSTI LIST
# ---------------------------------------------------

if not os.path.exists("assessment_files"):
    os.makedirs("assessment_files")

# ===================================================
# UPDATE INSIDE ADMIN DASHBOARD PAGE
# ===================================================
# ADD THIS BELOW YOUR EXISTING METRIC CARDS
# INSIDE:
# elif st.session_state.page == "admin_dashboard":

st.markdown("<br>", unsafe_allow_html=True)

col_a, col_b = st.columns(2)

with col_a:
    if st.button("Create / Upload Assessment"):
        st.session_state.page = "upload_assessment"
        st.rerun()

with col_b:
    if st.button("View Uploaded Assessments"):
        st.session_state.page = "view_assessments"
        st.rerun()

# ===================================================
# PAGE : UPLOAD ASSESSMENT
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
            <p style='color:#475569;'>
                Upload an Excel file containing questions,
                options, and correct answers.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    with st.container():

        st.markdown(
            '<div class="dashboard-card">',
            unsafe_allow_html=True
        )

        assessment_name = st.text_input(
            "Assessment Name"
        )

        uploaded_file = st.file_uploader(
            "Upload Assessment Excel File",
            type=["xlsx"]
        )

        st.info(
            "Excel file must contain columns:\n\n"
            "Question, Option1, Option2, "
            "Option3, Option4, CorrectAnswer"
        )

        if st.button("Upload Assessment"):

            if assessment_name.strip() == "":
                st.error("Please enter assessment name.")

            elif uploaded_file is None:
                st.error("Please upload an Excel file.")

            else:

                try:

                    df = pd.read_excel(uploaded_file)

                    required_columns = [
                        "Question",
                        "Option1",
                        "Option2",
                        "Option3",
                        "Option4",
                        "CorrectAnswer"
                    ]

                    missing_columns = [
                        col for col in required_columns
                        if col not in df.columns
                    ]

                    if missing_columns:

                        st.error(
                            f"Missing columns: "
                            f"{', '.join(missing_columns)}"
                        )

                    else:

                        file_path = (
                            f"assessment_files/"
                            f"{assessment_name}.xlsx"
                        )

                        df.to_excel(
                            file_path,
                            index=False
                        )

                        st.success(
                            "Assessment uploaded successfully."
                        )

                        st.dataframe(df.head())

                except Exception as e:
                    st.error(f"Error: {e}")

        st.markdown(
            '</div>',
            unsafe_allow_html=True
        )

# ===================================================
# PAGE : VIEW ASSESSMENTS
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

            <p style='color:#475569;'>
                View all uploaded assessments.
            </p>
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

        st.warning("No assessments uploaded yet.")

    else:

        for file in excel_files:

            assessment_name = file.replace(".xlsx", "")

            col1, col2 = st.columns([4, 1])

            with col1:

                st.markdown(
                    f"""
                    <div class='metric-card'>
                        <h3>{assessment_name}</h3>
                        <div class='metric-pill'>
                            Assessment Available
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

# ===================================================
# UPDATE STUDENT DETAILS PAGE
# ===================================================
# INSIDE:
# elif st.session_state.page == "student_details":

# REPLACE THIS:
#
# st.info(
#     "Assessment page will be added in the next step."
# )
#
# WITH THIS:

st.success("Details submitted successfully.")

st.session_state.page = "student_assessment_selection"

st.rerun()

# ===================================================
# PAGE : STUDENT ASSESSMENT SELECTION
# ===================================================

elif st.session_state.page == "student_assessment_selection":

    with st.sidebar:

        st.title("Student Dashboard")

        st.write(
            f"Student: "
            f"{st.session_state.full_name}"
        )

        st.button(
            "Logout",
            on_click=logout
        )

    st.markdown(
        """
        <div class='dashboard-card'>
            <h2>Available Assessments</h2>

            <p style='color:#475569;'>
                Select an assessment to begin.
            </p>
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

            assessment_name = file.replace(".xlsx", "")

            col1, col2 = st.columns([4, 1])

            with col1:

                st.markdown(
                    f"""
                    <div class='metric-card'>
                        <h3>{assessment_name}</h3>

                        <div class='metric-pill'>
                            Ready to Attempt
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            with col2:

                if st.button(
                    "Start Assessment",
                    key=f"start_{assessment_name}"
                ):

                    st.session_state.selected_assessment = file
                    st.session_state.current_question = 0
                    st.session_state.score = 0
                    st.session_state.answers = {}

                    st.session_state.page = "take_assessment"

                    st.rerun()

# ===================================================
# PAGE : TAKE ASSESSMENT
# ===================================================

elif st.session_state.page == "take_assessment":

    with st.sidebar:

        st.title("Assessment")

        st.write(
            f"Assessment: "
            f"{st.session_state.selected_assessment.replace('.xlsx', '')}"
        )

        st.write(
            f"Question Number: "
            f"{st.session_state.current_question + 1}"
        )

        st.button(
            "Logout",
            on_click=logout
        )

    file_path = (
        f"assessment_files/"
        f"{st.session_state.selected_assessment}"
    )

    df = pd.read_excel(file_path)

    total_questions = len(df)

    current_q = st.session_state.current_question

    row = df.iloc[current_q]

    st.markdown(
        f"""
        <div class='dashboard-card'>
            <h2>
                Question {current_q + 1} of {total_questions}
            </h2>
        </div>
        """,
        unsafe_allow_html=True
    )

    left, right = st.columns([2, 1], gap="large")

    # ---------------------------------------------------
    # LEFT SIDE - QUESTION
    # ---------------------------------------------------
    with left:

        st.markdown(
            f"""
            <div class='dashboard-card'>
                <h3>{row['Question']}</h3>
            </div>
            """,
            unsafe_allow_html=True
        )

    # ---------------------------------------------------
    # RIGHT SIDE - OPTIONS
    # ---------------------------------------------------
    with right:

        options = [
            row["Option1"],
            row["Option2"],
            row["Option3"],
            row["Option4"]
        ]

        selected_option = st.radio(
            "Select an Option",
            options,
            key=f"question_{current_q}"
        )

        if st.button("Submit Answer"):

            correct_answer = row["CorrectAnswer"]

            st.session_state.answers[current_q] = selected_option

            if selected_option == correct_answer:

                st.success("Correct Answer")

                st.session_state.score += 1

            else:

                st.error(
                    f"Incorrect Answer.\n\n"
                    f"Correct Answer: {correct_answer}"
                )

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
                        f"Assessment Completed Successfully.\n\n"
                        f"Final Score: "
                        f"{st.session_state.score}/{total_questions}"
                    )