from __future__ import annotations

import json
import random
import re
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd


APP_DATA_DIR = Path("app_data")
COURSES_FILE = APP_DATA_DIR / "courses.json"

DEFAULT_COURSES = ["AIMD", "ACT", "EY", "S4F"]
LEVEL_LABELS = ["Level 1", "Level 2", "Level 3"]
DEFAULT_LEVEL_DISTRIBUTION = {
    "Level 1": 35,
    "Level 2": 35,
    "Level 3": 30,
}
DEFAULT_LEVEL_MARKS = {
    "Level 1": 1,
    "Level 2": 2,
    "Level 3": 5,
}
SUPPORTED_LANGUAGES = [
    "Python",
    "Java",
    "JavaScript",
    "HTML",
    "CSS",
    "PowerBI",
    "C",
    "C++",
]

GENERIC_SKILL_VALUES = {
    "",
    "general",
    "na",
    "n/a",
    "none",
    "misc",
    "miscellaneous",
    "skill",
}

QUESTION_BANK_COLUMNS = [
    "Question ID",
    "Course",
    "Skill",
    "Level",
    "Question Type",
    "Question",
    "Options",
    "Correct Answer",
    "Test Cases",
    "Expected Output",
    "Marks",
    "Difficulty",
]


def safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def normalize_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def first_existing_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    normalized_columns = {
        normalize_key(column): column
        for column in df.columns
    }

    for alias in aliases:
        column = normalized_columns.get(normalize_key(alias))
        if column is not None:
            return column

    return None


def get_cell(row: pd.Series, aliases: list[str], default: Any = "") -> Any:
    normalized_aliases = {normalize_key(alias) for alias in aliases}

    for column in row.index:
        if normalize_key(column) in normalized_aliases and pd.notna(row[column]):
            return row[column]

    return default


def normalize_level(value: Any, question_type: Any = "") -> str:
    raw = str(value or "").strip().lower()
    question_type_raw = str(question_type or "").strip().lower()

    if raw in {"1", "l1", "level1", "level 1", "basic", "beginner"}:
        return "Level 1"

    if raw in {
        "2",
        "l2",
        "level2",
        "level 2",
        "intermediate",
        "medium",
    }:
        return "Level 2"

    if raw in {
        "3",
        "l3",
        "level3",
        "level 3",
        "coding",
        "programming",
        "advanced",
    }:
        return "Level 3"

    if question_type_raw in {"coding", "code", "programming"}:
        return "Level 3"

    return "Level 1"


def normalize_question_type(value: Any, level: Any = "") -> str:
    raw = str(value or "").strip().lower()

    if raw in {"coding", "code", "programming"}:
        return "Coding"

    if normalize_level(level) == "Level 3":
        return "Coding"

    return "MCQ"


def normalize_language(value: Any) -> str:
    raw = str(value or "").strip().lower()
    mapping = {
        "py": "Python",
        "python": "Python",
        "java": "Java",
        "js": "JavaScript",
        "javascript": "JavaScript",
        "html": "HTML",
        "css": "CSS",
        "powerbi": "PowerBI",
        "power bi": "PowerBI",
        "dax": "PowerBI",
        "c": "C",
        "cpp": "C++",
        "c++": "C++",
    }

    return mapping.get(raw, "Python")


def detect_assessment_domain(source_sheet: Any, language: Any = "") -> str:
    source = str(source_sheet or "").strip().lower()
    normalized_language = normalize_language(language)

    if "javascript" in source or "java script" in source:
        return "JavaScript"

    if "java" in source:
        return "Java"

    if "html" in source:
        return "HTML"

    if "css" in source:
        return "CSS"

    if "power" in source or "dax" in source:
        return "PowerBI"

    if "aiml" in source or "ai/ml" in source or "machine learning" in source:
        return "Python AIML"

    if "python" in source:
        return "Python"

    return normalized_language or "Assessment"


def infer_assessment_skill(
    row: pd.Series,
    existing_skill: Any,
    source_sheet_col: str | None,
    question_col: str | None,
    language_col: str | None,
) -> str:
    cleaned_skill = str(existing_skill or "").strip()

    if normalize_key(cleaned_skill) not in GENERIC_SKILL_VALUES:
        return cleaned_skill

    source_sheet = row.get(source_sheet_col, "") if source_sheet_col else ""
    question_text = row.get(question_col, "") if question_col else ""
    language = row.get(language_col, "") if language_col else ""
    domain = detect_assessment_domain(source_sheet, language)
    search_text = f"{source_sheet} {question_text}".lower()
    rules = {
        "Java": [
            (["class", "object", "inherit", "polymorphism", "interface"], "Java OOP"),
            (["loop", "for ", "while"], "Java Loops"),
            (["array", "list", "collection", "map"], "Java Collections"),
            (["string", "character"], "Java Strings"),
            (["exception", "try", "catch"], "Java Exception Handling"),
            (["method", "function"], "Java Methods"),
            (["data type", "integer", "boolean", "char"], "Java Data Types"),
        ],
        "JavaScript": [
            (["dom", "document", "element"], "JavaScript DOM"),
            (["event", "click", "listener"], "JavaScript Events"),
            (["array", "map", "filter", "reduce"], "JavaScript Arrays"),
            (["object", "json"], "JavaScript Objects"),
            (["function", "arrow"], "JavaScript Functions"),
            (["async", "promise", "await"], "JavaScript Async"),
            (["variable", "let", "const", "var"], "JavaScript Variables"),
        ],
        "HTML": [
            (["form", "input", "label"], "HTML Forms"),
            (["table", "row", "cell"], "HTML Tables"),
            (["semantic", "section", "article", "header", "footer"], "Semantic HTML"),
            (["image", "audio", "video", "media"], "HTML Media"),
            (["link", "anchor", "href"], "HTML Links"),
            (["heading", "paragraph", "list"], "HTML Structure"),
        ],
        "CSS": [
            (["selector", "class", "id"], "CSS Selectors"),
            (["box", "margin", "padding", "border"], "CSS Box Model"),
            (["flex", "flexbox"], "CSS Flexbox"),
            (["grid"], "CSS Grid"),
            (["responsive", "media query"], "Responsive CSS"),
            (["position", "absolute", "relative"], "CSS Positioning"),
            (["animation", "transition"], "CSS Animations"),
        ],
        "Python": [
            (["loop", "for ", "while"], "Python Loops"),
            (["function", "def "], "Python Functions"),
            (["list", "tuple", "set"], "Python Collections"),
            (["dictionary", "dict"], "Python Dictionaries"),
            (["string"], "Python Strings"),
            (["file", "csv"], "Python File Handling"),
            (["class", "object"], "Python OOP"),
        ],
        "Python AIML": [
            (["model", "train", "prediction"], "Model Training"),
            (["accuracy", "precision", "recall", "evaluation"], "Model Evaluation"),
            (["data cleaning", "preprocess", "missing"], "Data Preprocessing"),
            (["numpy", "array"], "NumPy"),
            (["pandas", "dataframe"], "Pandas"),
            (["supervised", "classification", "regression"], "Machine Learning Basics"),
        ],
        "PowerBI": [
            (["dax", "measure", "calculate", "sum("], "PowerBI DAX"),
            (["relationship", "model"], "PowerBI Data Modeling"),
            (["visual", "chart", "dashboard"], "PowerBI Visualizations"),
            (["filter", "slicer"], "PowerBI Filters"),
            (["power query", "transform"], "Power Query"),
        ],
    }

    for keywords, skill in rules.get(domain, []):
        if any(keyword in search_text for keyword in keywords):
            return skill

    return f"{domain} Fundamentals"


def load_courses() -> list[str]:
    APP_DATA_DIR.mkdir(exist_ok=True)

    if not COURSES_FILE.exists():
        save_courses(DEFAULT_COURSES)
        return DEFAULT_COURSES.copy()

    try:
        courses = json.loads(COURSES_FILE.read_text(encoding="utf-8"))
    except Exception:
        courses = DEFAULT_COURSES

    cleaned = []
    for course in courses:
        course_name = str(course).strip().upper()
        if course_name and course_name not in cleaned:
            cleaned.append(course_name)

    if not cleaned:
        cleaned = DEFAULT_COURSES.copy()

    return cleaned


def save_courses(courses: list[str]) -> None:
    APP_DATA_DIR.mkdir(exist_ok=True)
    cleaned = []

    for course in courses:
        course_name = str(course).strip().upper()
        if course_name and course_name not in cleaned:
            cleaned.append(course_name)

    COURSES_FILE.write_text(
        json.dumps(cleaned, indent=2),
        encoding="utf-8",
    )


def read_excel_sheet(workbook_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    candidates: list[pd.DataFrame] = []

    for header in (0, 1, 2):
        try:
            df = pd.read_excel(
                BytesIO(workbook_bytes),
                sheet_name=sheet_name,
                header=header,
                engine="openpyxl",
            )
            df = df.dropna(how="all").reset_index(drop=True)
            candidates.append(df)
        except Exception:
            continue

    if not candidates:
        return pd.DataFrame()

    def score(candidate: pd.DataFrame) -> int:
        columns = {normalize_key(column) for column in candidate.columns}
        expected = {normalize_key(column) for column in QUESTION_BANK_COLUMNS}
        legacy = {
            "question",
            "option1",
            "option2",
            "option3",
            "option4",
            "correctanswer",
        }
        return len(columns & expected) + len(columns & legacy)

    best = max(candidates, key=score)

    if score(best) == 0:
        return candidates[0]

    return best


def read_selected_sheets(
    workbook_bytes: bytes,
    sheet_names: list[str],
    course: str = "",
) -> pd.DataFrame:
    frames = []

    for sheet_name in sheet_names:
        df = read_excel_sheet(workbook_bytes, sheet_name)

        if df.empty:
            continue

        df["Source Sheet"] = sheet_name

        if course:
            course_column = first_existing_column(df, ["Course"])
            if course_column is None:
                df["Course"] = course
            else:
                df[course_column] = df[course_column].fillna(course)

        frames.append(df)

    if not frames:
        return pd.DataFrame()

    return standardize_question_bank(
        pd.concat(frames, ignore_index=True),
        default_course=course,
    )


def standardize_question_bank(
    df: pd.DataFrame,
    default_course: str = "",
) -> pd.DataFrame:
    df = df.copy().dropna(how="all").reset_index(drop=True)

    if df.empty:
        return df

    source_sheet_col = first_existing_column(df, ["Source Sheet"])
    question_id_col = first_existing_column(
        df,
        ["Question ID", "Question No", "Question Number", "QID", "ID"],
    )
    question_col = first_existing_column(
        df,
        ["Question", "Problem Statement", "Problem", "Prompt"],
    )
    course_col = first_existing_column(df, ["Course"])
    skill_col = first_existing_column(df, ["Skill", "Topic", "Sub Skill"])
    level_col = first_existing_column(df, ["Level", "Question Level"])
    type_col = first_existing_column(df, ["Question Type", "Type"])
    marks_col = first_existing_column(df, ["Marks", "Score"])
    language_col = first_existing_column(df, ["Language", "Programming Language"])

    if question_id_col is None:
        ids = []
        for index, row in df.iterrows():
            source = (
                str(row.get(source_sheet_col, "Sheet")).strip()
                if source_sheet_col
                else "Sheet"
            )
            ids.append(f"{source}-{index + 1}")
        df["Question ID"] = ids
    elif question_id_col != "Question ID":
        df["Question ID"] = df[question_id_col]

    if course_col is None:
        df["Course"] = default_course
    elif course_col != "Course":
        df["Course"] = df[course_col]

    df["Course"] = (
        df["Course"]
        .fillna(default_course)
        .astype(str)
        .str.strip()
        .replace("", default_course)
        .str.upper()
    )

    if skill_col is None:
        df["Skill"] = "General"
    elif skill_col != "Skill":
        df["Skill"] = df[skill_col]

    df["Skill"] = [
        infer_assessment_skill(
            row,
            row.get("Skill", "General"),
            source_sheet_col,
            question_col,
            language_col,
        )
        for _, row in df.iterrows()
    ]

    level_values = (
        df[level_col].tolist()
        if level_col is not None
        else [""] * len(df)
    )
    type_values = (
        df[type_col].tolist()
        if type_col is not None
        else [""] * len(df)
    )

    df["Level"] = [
        normalize_level(level_value, type_value)
        for level_value, type_value in zip(level_values, type_values)
    ]
    df["Question Type"] = [
        normalize_question_type(type_value, level_value)
        for type_value, level_value in zip(type_values, df["Level"])
    ]

    if marks_col is None:
        df["Marks"] = df["Level"].map(DEFAULT_LEVEL_MARKS)
    elif marks_col != "Marks":
        df["Marks"] = df[marks_col]

    df["Marks"] = [
        safe_int(value, DEFAULT_LEVEL_MARKS.get(level, 1))
        for value, level in zip(df["Marks"], df["Level"])
    ]

    if language_col is None:
        df["Language"] = "Python"
    elif language_col != "Language":
        df["Language"] = df[language_col]

    df["Language"] = df["Language"].apply(normalize_language)

    if "Source Sheet" not in df.columns:
        df["Source Sheet"] = "Sheet1"

    df["Question ID"] = (
        df["Question ID"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    missing_ids = df["Question ID"] == ""
    if missing_ids.any():
        df.loc[missing_ids, "Question ID"] = [
            f"{df.loc[index, 'Source Sheet']}-{index + 1}"
            for index in df.index[missing_ids]
        ]

    return df


def summarize_question_bank(df: pd.DataFrame) -> dict[str, Any]:
    if df.empty:
        return {
            "total_questions": 0,
            "unique_question_ids": 0,
            "total_skills": 0,
            "level_counts": {level: 0 for level in LEVEL_LABELS},
            "skills": [],
            "courses": [],
        }

    question_bank = standardize_question_bank(df)

    level_counts = {
        level: int((question_bank["Level"] == level).sum())
        for level in LEVEL_LABELS
    }

    return {
        "total_questions": int(len(question_bank)),
        "unique_question_ids": int(question_bank["Question ID"].nunique()),
        "total_skills": int(question_bank["Skill"].nunique()),
        "level_counts": level_counts,
        "skills": sorted(question_bank["Skill"].dropna().unique().tolist()),
        "courses": sorted(question_bank["Course"].dropna().unique().tolist()),
    }


def apply_level_marks(
    df: pd.DataFrame,
    level_marks: dict[str, int],
) -> pd.DataFrame:
    df = df.copy()

    for level, marks in level_marks.items():
        df.loc[df["Level"] == level, "Marks"] = safe_int(
            marks,
            DEFAULT_LEVEL_MARKS.get(level, 1),
        )

    return df


def select_random_questions(
    question_bank: pd.DataFrame,
    level_counts: dict[str, int],
    skill_counts: dict[str, int] | None = None,
    allowed_languages: list[str] | None = None,
    level_marks: dict[str, int] | None = None,
    seed: int | None = None,
) -> tuple[pd.DataFrame, list[str]]:
    warnings: list[str] = []

    if question_bank.empty:
        return pd.DataFrame(), ["No questions found in the selected bank."]

    rng = random.Random(seed)
    df = standardize_question_bank(question_bank)

    if allowed_languages:
        allowed = {normalize_language(language) for language in allowed_languages}
        coding_mask = df["Question Type"].str.lower().eq("coding")
        df = df[
            (~coding_mask)
            | (df["Language"].apply(normalize_language).isin(allowed))
        ].copy()

    df = df.drop_duplicates(subset=["Question ID"]).reset_index(drop=True)

    if level_marks:
        df = apply_level_marks(df, level_marks)

    clean_level_counts = {
        level: max(0, safe_int(level_counts.get(level, 0), 0))
        for level in LEVEL_LABELS
    }
    clean_skill_counts = {
        str(skill).strip(): max(0, safe_int(count, 0))
        for skill, count in (skill_counts or {}).items()
        if str(skill).strip() and safe_int(count, 0) > 0
    }

    selected_indices: list[int] = []
    selected_ids: set[str] = set()
    remaining_by_level = clean_level_counts.copy()

    def can_take(index: int) -> bool:
        question_id = str(df.at[index, "Question ID"])
        level = df.at[index, "Level"]
        return (
            question_id not in selected_ids
            and remaining_by_level.get(level, 0) > 0
        )

    def take(index: int) -> None:
        question_id = str(df.at[index, "Question ID"])
        level = df.at[index, "Level"]
        selected_indices.append(index)
        selected_ids.add(question_id)
        remaining_by_level[level] = max(0, remaining_by_level.get(level, 0) - 1)

    for skill, required_count in clean_skill_counts.items():
        chosen_for_skill = 0
        candidate_indices = df.index[
            df["Skill"].astype(str).str.lower() == skill.lower()
        ].tolist()
        rng.shuffle(candidate_indices)

        while chosen_for_skill < required_count:
            available = [index for index in candidate_indices if can_take(index)]

            if not available:
                break

            available.sort(
                key=lambda index: remaining_by_level.get(df.at[index, "Level"], 0),
                reverse=True,
            )
            best_remaining = remaining_by_level.get(df.at[available[0], "Level"], 0)
            best_pool = [
                index
                for index in available
                if remaining_by_level.get(df.at[index, "Level"], 0) == best_remaining
            ]
            chosen_index = rng.choice(best_pool)
            take(chosen_index)
            chosen_for_skill += 1

        if chosen_for_skill < required_count:
            warnings.append(
                f"Skill '{skill}' needed {required_count} question(s), "
                f"but only {chosen_for_skill} could be selected with the "
                "current level and language rules."
            )

    for level, remaining_count in list(remaining_by_level.items()):
        if remaining_count <= 0:
            continue

        candidate_indices = df.index[
            (df["Level"] == level)
            & (~df["Question ID"].astype(str).isin(selected_ids))
        ].tolist()
        rng.shuffle(candidate_indices)
        chosen = candidate_indices[:remaining_count]

        for index in chosen:
            take(index)

        if len(chosen) < remaining_count:
            warnings.append(
                f"{level} needed {remaining_count} more question(s), "
                f"but only {len(chosen)} were available."
            )

    selected_df = df.loc[selected_indices].sample(
        frac=1,
        random_state=rng.randint(1, 1_000_000),
    )

    return selected_df.reset_index(drop=True), warnings


def summarize_performance(
    progress_data: dict[str, Any],
    is_correct_callback,
    marks_callback,
) -> dict[str, Any]:
    records = progress_data.get("assessment_records", [])
    df = pd.DataFrame(records)

    if df.empty:
        return {
            "skill_summary": "",
            "level_summary": "",
            "coding_accuracy": "0/0",
            "coding_correct": 0,
            "coding_total": 0,
        }

    df = standardize_question_bank(df)
    answers = progress_data.get("answers", {})
    code_answers = progress_data.get("code_answers", {})

    skill_totals: dict[str, dict[str, int]] = {}
    level_totals: dict[str, dict[str, int]] = {}
    coding_total = 0
    coding_correct = 0

    for index, row in df.iterrows():
        skill = str(row.get("Skill", "General"))
        level = str(row.get("Level", "Level 1"))
        marks = marks_callback(row)
        correct = is_correct_callback(index, row, answers, code_answers)
        question_type = str(row.get("Question Type", "")).lower()

        skill_bucket = skill_totals.setdefault(skill, {"earned": 0, "total": 0})
        level_bucket = level_totals.setdefault(level, {"earned": 0, "total": 0})
        skill_bucket["total"] += marks
        level_bucket["total"] += marks

        if correct:
            skill_bucket["earned"] += marks
            level_bucket["earned"] += marks

        if question_type == "coding":
            coding_total += 1
            if correct:
                coding_correct += 1

    skill_summary = "; ".join(
        f"{skill}: {values['earned']}/{values['total']}"
        for skill, values in sorted(skill_totals.items())
    )
    level_summary = "; ".join(
        f"{level}: {values['earned']}/{values['total']}"
        for level, values in sorted(level_totals.items())
    )

    return {
        "skill_summary": skill_summary,
        "level_summary": level_summary,
        "coding_accuracy": f"{coding_correct}/{coding_total}",
        "coding_correct": coding_correct,
        "coding_total": coding_total,
    }
