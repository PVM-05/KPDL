# =====================================================================
# COURSE-GRADE PREPROCESSING AND HARMONIZATION
# =====================================================================
#
# Despite the filename, this is an executable Python script. It is released
# with synthetic source-formatted data to illustrate the programmatic
# course-grade harmonization stage described in the accompanying data article.
#
# Public input columns:
#   StudentID    : synthetic or anonymized student identifier
#   CourseLabel  : source-formatted course label
#   SemesterID   : semester identifier
#   Grade4Point  : final course grade on the institutional 0-4 scale
#   Credits      : number of credits assigned to the course
#
# Implemented steps:
#   1. Validate and normalize the source-formatted grade table.
#   2. Map equivalent synthetic course labels to common harmonized fields.
#   3. Aggregate mapped records using credit-weighted grades.
#   4. Compute credit-weighted GPA for each student.
#   5. Pivot the long table into a student-by-course matrix.
#   6. Export aggregate audit summaries for the public mock workflow.
#
# Scope limitation:
#   This public script does not contain institution-specific course names or
#   the complete internal mapping. It does not reproduce the private database
#   extraction, administrative-record merge, manual reconciliation,
#   anonymization, or final 56-variable coding steps.
#
# Example:
#   python 05_preprocessing_pseudocode.py \
#       --input 06_mock_raw_course_records.csv \
#       --output-dir mock_outputs
# =====================================================================

from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd


REQUIRED_COLUMNS = [
    "StudentID",
    "CourseLabel",
    "SemesterID",
    "Grade4Point",
    "Credits",
]

# Synthetic labels are used in the public version so that the workflow can be
# inspected without disclosing institution-specific curriculum labels.
COURSE_MAPPING = {
    "HARMONIZED_COURSE_01": ["Source_Course_A", "Source_Course_B"],
    "HARMONIZED_COURSE_02": ["Source_Course_C", "Source_Course_D"],
    "HARMONIZED_COURSE_03": ["Source_Course_E", "Source_Course_F"],
}


def normalize_text(value: object) -> object:
    """Strip leading/trailing spaces and collapse repeated whitespace."""
    if pd.isna(value):
        return value
    return re.sub(r"\s+", " ", str(value)).strip()


def read_input(path: Path) -> pd.DataFrame:
    """Read a CSV or Excel source-formatted grade table."""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path)
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(path)
    raise ValueError(
        f"Unsupported input format: {path.suffix}. Use CSV or Excel."
    )


def validate_input(df: pd.DataFrame) -> pd.DataFrame:
    """Validate required fields and numeric ranges."""
    data = df.copy()
    data.columns = [str(column).strip() for column in data.columns]

    missing_columns = [
        column for column in REQUIRED_COLUMNS if column not in data.columns
    ]
    if missing_columns:
        raise ValueError(f"Missing required columns: {missing_columns}")

    data = data[REQUIRED_COLUMNS].copy()
    data["CourseLabel"] = data["CourseLabel"].apply(normalize_text)

    for column in ["SemesterID", "Grade4Point", "Credits"]:
        data[column] = pd.to_numeric(data[column], errors="coerce")

    missing_required = data[
        ["StudentID", "CourseLabel", "Grade4Point", "Credits"]
    ].isna().any(axis=1)
    if missing_required.any():
        raise ValueError(
            f"The input contains {int(missing_required.sum())} rows "
            "with missing required values."
        )

    invalid_grades = ~data["Grade4Point"].between(0, 4, inclusive="both")
    if invalid_grades.any():
        raise ValueError(
            f"The input contains {int(invalid_grades.sum())} grades "
            "outside the 0-4 scale."
        )

    invalid_credits = data["Credits"] <= 0
    if invalid_credits.any():
        raise ValueError(
            f"The input contains {int(invalid_credits.sum())} non-positive "
            "credit values."
        )

    return data


def build_course_map(
    mapping: dict[str, list[str]]
) -> dict[str, str]:
    """Create a normalized source-label-to-harmonized-field map."""
    course_map: dict[str, str] = {}

    for harmonized_field, source_labels in mapping.items():
        target = str(normalize_text(harmonized_field))

        for source_label in source_labels:
            label = str(normalize_text(source_label))
            if label in course_map and course_map[label] != target:
                raise ValueError(
                    f"Source label {label!r} is assigned to more than "
                    "one harmonized field."
                )
            course_map[label] = target

    return course_map


def harmonize_course_records(
    df: pd.DataFrame,
    course_map: dict[str, str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Map course labels and aggregate equivalent records by credit weight."""
    data = df.copy()
    data["HarmonizedField"] = data["CourseLabel"].map(course_map)
    data["MappingStatus"] = data["HarmonizedField"].notna().map(
        {True: "mapped", False: "retained_original_label"}
    )
    data["HarmonizedField"] = data["HarmonizedField"].fillna(
        data["CourseLabel"]
    )

    data["WeightedGradePoints"] = (
        data["Grade4Point"] * data["Credits"]
    )

    harmonized = (
        data.groupby(
            ["StudentID", "HarmonizedField"],
            as_index=False,
        )
        .agg(
            SemesterID=("SemesterID", "min"),
            Credits=("Credits", "sum"),
            WeightedGradePoints=("WeightedGradePoints", "sum"),
        )
    )

    harmonized["Grade4Point"] = (
        harmonized["WeightedGradePoints"] / harmonized["Credits"]
    )
    harmonized = harmonized.drop(columns=["WeightedGradePoints"])

    return data, harmonized


def compute_gpa(df: pd.DataFrame) -> pd.DataFrame:
    """Compute credit-weighted GPA from the validated source records."""
    data = df.copy()
    data["WeightedGradePoints"] = (
        data["Grade4Point"] * data["Credits"]
    )

    gpa = (
        data.groupby("StudentID", as_index=False)
        .agg(
            TotalCredits=("Credits", "sum"),
            TotalGradePoints=("WeightedGradePoints", "sum"),
        )
    )

    gpa["GPA"] = gpa["TotalGradePoints"] / gpa["TotalCredits"]
    return gpa[["StudentID", "TotalCredits", "GPA"]]


def create_field_summary(
    mapped_source: pd.DataFrame,
) -> pd.DataFrame:
    """Create aggregate counts without exposing student-level identifiers."""
    return (
        mapped_source.groupby(
            ["HarmonizedField", "MappingStatus"],
            as_index=False,
        )
        .agg(
            SourceLabelCount=("CourseLabel", "nunique"),
            GradeRecordCount=("CourseLabel", "size"),
            StudentCount=("StudentID", "nunique"),
        )
        .sort_values(
            ["MappingStatus", "GradeRecordCount"],
            ascending=[True, False],
        )
        .reset_index(drop=True)
    )


def create_preprocessing_summary(
    source: pd.DataFrame,
    mapped_source: pd.DataFrame,
    harmonized: pd.DataFrame,
    pivot: pd.DataFrame,
) -> pd.DataFrame:
    """Create a compact summary of the public mock workflow."""
    mapped_mask = mapped_source["MappingStatus"].eq("mapped")

    rows = [
        (
            "Input",
            "Input grade records",
            len(source),
            "Source-formatted course-grade rows.",
        ),
        (
            "Input",
            "Input students",
            source["StudentID"].nunique(),
            "Distinct synthetic student identifiers.",
        ),
        (
            "Input",
            "Distinct source course labels",
            source["CourseLabel"].nunique(),
            "Distinct source-formatted labels.",
        ),
        (
            "Mapping",
            "Mapped source labels",
            mapped_source.loc[mapped_mask, "CourseLabel"].nunique(),
            "Labels assigned to a common harmonized field.",
        ),
        (
            "Mapping",
            "Source labels retained unchanged",
            mapped_source.loc[~mapped_mask, "CourseLabel"].nunique(),
            "Labels without a mapping in the public mock example.",
        ),
        (
            "Mapping",
            "Mapped grade records",
            int(mapped_mask.sum()),
            "Grade rows covered by the public mock mapping.",
        ),
        (
            "Mapping",
            "Grade records retained under original labels",
            int((~mapped_mask).sum()),
            "Unmapped rows retained rather than deleted.",
        ),
        (
            "Output",
            "Harmonized student-course records",
            len(harmonized),
            "Rows after equivalent records are combined by student and field.",
        ),
        (
            "Output",
            "Student rows in pivot table",
            len(pivot),
            "Student rows in the intermediate wide matrix.",
        ),
        (
            "Output",
            "Course or course-group fields in pivot table",
            max(len(pivot.columns) - 1, 0),
            "Excludes the StudentID column.",
        ),
    ]

    return pd.DataFrame(
        rows,
        columns=["Stage", "Measure", "Value", "Notes"],
    )


def run(input_path: Path, output_dir: Path) -> None:
    """Run the public mock harmonization workflow."""
    output_dir.mkdir(parents=True, exist_ok=True)

    source = validate_input(read_input(input_path))
    course_map = build_course_map(COURSE_MAPPING)

    mapped_source, harmonized = harmonize_course_records(
        source,
        course_map,
    )
    gpa = compute_gpa(source)

    pivot = (
        harmonized.pivot_table(
            index="StudentID",
            columns="HarmonizedField",
            values="Grade4Point",
            aggfunc="first",
        )
        .reset_index()
    )
    pivot.columns.name = None

    field_summary = create_field_summary(mapped_source)
    preprocessing_summary = create_preprocessing_summary(
        source,
        mapped_source,
        harmonized,
        pivot,
    )

    harmonized.to_csv(
        output_dir / "harmonized_long_mock_output.csv",
        index=False,
        encoding="utf-8-sig",
    )
    pivot.to_csv(
        output_dir / "student_course_matrix_mock_output.csv",
        index=False,
        encoding="utf-8-sig",
    )
    gpa.to_csv(
        output_dir / "student_gpa_mock_output.csv",
        index=False,
        encoding="utf-8-sig",
    )
    field_summary.to_csv(
        output_dir / "harmonized_field_summary_mock_output.csv",
        index=False,
        encoding="utf-8-sig",
    )
    preprocessing_summary.to_csv(
        output_dir / "preprocessing_summary_mock_output.csv",
        index=False,
        encoding="utf-8-sig",
    )

    print(f"Completed. Outputs were written to: {output_dir.resolve()}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Illustrate course-grade harmonization using synthetic "
            "source-formatted records."
        )
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=Path("06_mock_raw_course_records.csv"),
        help="Input CSV or Excel file.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("mock_outputs"),
        help="Directory for generated outputs.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    arguments = parse_args()
    run(arguments.input, arguments.output_dir)
