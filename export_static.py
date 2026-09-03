from __future__ import annotations

import json
import os
import re
from datetime import datetime

import pandas as pd


# ---------------------------------------------------------------------------
# File configuration
# ---------------------------------------------------------------------------

CORE_EXCEL = "CP_Members.xlsx"
FFCS_EXCEL = "FFCS_Members.xlsx"
FFCS_ROSTER_FALLBACK = "cccp202627.xlsx"
MEETING_EXCEL = "Meeting_Attendance.xlsx"

TEMPLATE_FILE = "index_template.html"
OUTPUT_HTML = "index.html"

STARTER_COL_RE = re.compile(r"^Starters\s+(\d+)$", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def clean_text(value, default="N/A") -> str:
    """Convert an Excel value into a clean display string."""
    if value is None:
        return default

    try:
        if pd.isna(value):
            return default
    except (TypeError, ValueError):
        pass

    text = str(value).strip()

    if not text or text.lower() in {"nan", "none", "null"}:
        return default

    return text


def pick(row: pd.Series, *candidates, default="N/A"):
    """Return the first non-empty value found among candidate column names."""
    for column in candidates:
        if column in row.index:
            value = row[column]

            try:
                if pd.notna(value) and str(value).strip():
                    return value
            except (TypeError, ValueError):
                pass

    return default


def starter_columns(df: pd.DataFrame) -> list[str]:
    """
    Return all columns matching:
        Starters 1
        Starters 2
        Starters 3
        ...
    """
    matches = []

    for column in df.columns:
        column_name = str(column).strip()

        if STARTER_COL_RE.match(column_name):
            matches.append(column_name)

    # Sort numerically rather than alphabetically.
    # This prevents Starters 10 appearing before Starters 2.
    matches.sort(
        key=lambda c: int(STARTER_COL_RE.match(c).group(1))
    )

    return matches


def find_column(df: pd.DataFrame, keywords: list[str]):
    """
    Find a likely column using case-insensitive keyword matching.
    """
    normalized = {
        str(column).strip().lower(): column
        for column in df.columns
    }

    # First try exact keyword matches.
    for keyword in keywords:
        keyword_lower = keyword.lower()

        if keyword_lower in normalized:
            return normalized[keyword_lower]

    # Then try partial matches.
    for column_lower, original_column in normalized.items():
        if any(keyword.lower() in column_lower for keyword in keywords):
            return original_column

    return None


def clean_username(value) -> str:
    """
    Normalize a CodeChef username/URL.

    Examples:
        codechef.com/users/example -> example
        https://www.codechef.com/users/example/ -> example
        example -> example
    """
    text = clean_text(value)

    if text == "N/A":
        return text

    text = text.strip().rstrip("/")

    # If the value is a URL, keep only the final path component.
    if "/" in text:
        text = text.split("/")[-1]

    return text


def dataframe_to_records(df: pd.DataFrame) -> list[dict]:
    """
    Convert a DataFrame to JSON-safe records.
    """
    if df.empty:
        return []

    # Replace NaN/NaT with None before JSON conversion.
    safe_df = df.copy()

    safe_df = safe_df.where(pd.notna(safe_df), None)

    return json.loads(
        safe_df.to_json(
            orient="records",
            date_format="iso"
        )
    )


# ---------------------------------------------------------------------------
# Loading cohort data
# ---------------------------------------------------------------------------

def load_cohort(
    excel_file: str,
    member_type: str,
    fallback_roster: str | None = None,
) -> pd.DataFrame:
    """
    Load a Core or FFCS Excel file.

    If the main Excel file does not exist and a fallback roster is supplied,
    the fallback roster is used to create a basic member list.
    """

    # -----------------------------------------------------------------------
    # Load primary workbook
    # -----------------------------------------------------------------------

    if os.path.exists(excel_file):
        try:
            df = pd.read_excel(excel_file)
        except Exception as exc:
            print(f"⚠️ Could not read {excel_file}: {exc}")
            return pd.DataFrame()

    # -----------------------------------------------------------------------
    # Load fallback roster
    # -----------------------------------------------------------------------

    elif fallback_roster and os.path.exists(fallback_roster):
        try:
            raw = pd.read_excel(fallback_roster)
        except Exception as exc:
            print(f"⚠️ Could not read fallback roster {fallback_roster}: {exc}")
            return pd.DataFrame()

        if raw.empty:
            return pd.DataFrame()

        name_col = find_column(
            raw,
            ["name", "student name", "full name"]
        )

        reg_col = find_column(
            raw,
            [
                "registration number",
                "register number",
                "reg no",
                "reg. no",
                "roll number",
                "roll no",
                "register",
            ]
        )

        user_col = find_column(
            raw,
            [
                "codechef id",
                "codechef username",
                "codechef",
                "username",
                "handle",
                "profile",
            ]
        )

        phone_col = find_column(
            raw,
            ["phone number", "phone", "mobile number", "mobile", "contact"],
        )

        df = pd.DataFrame({
            "Name": (
                raw[name_col]
                if name_col is not None
                else ["N/A"] * len(raw)
            ),
            "Register number": (
                raw[reg_col]
                if reg_col is not None
                else ["N/A"] * len(raw)
            ),
            "Phone number": (
                raw[phone_col]
                if phone_col is not None
                else ["N/A"] * len(raw)
            ),
            "Username": (
                raw[user_col].apply(clean_username)
                if user_col is not None
                else ["N/A"] * len(raw)
            ),
        })

    else:
        print(f"ℹ️ {excel_file} not found.")
        return pd.DataFrame()

    if df.empty:
        return pd.DataFrame()

    # -----------------------------------------------------------------------
    # Normalize column names
    # -----------------------------------------------------------------------

    df = df.copy()

    # Strip accidental whitespace from Excel headers.
    df.columns = [str(column).strip() for column in df.columns]

    # -----------------------------------------------------------------------
    # Detect contest columns
    # -----------------------------------------------------------------------

    starter_cols = starter_columns(df)

    # Convert contest values to integers.
    for column in starter_cols:
        df[column] = pd.to_numeric(
            df[column],
            errors="coerce"
        ).fillna(0)

        df[column] = df[column].astype(int)

    # -----------------------------------------------------------------------
    # Build normalized records
    # -----------------------------------------------------------------------

    records = []

    for _, row in df.iterrows():

        name = clean_text(
            pick(
                row,
                "Name",
                "Student Name",
                "Full Name",
                default="N/A",
            )
        )

        reg_no = clean_text(
            pick(
                row,
                "Registration Number",
                "Register number",
                "Register Number",
                "Reg No",
                "Reg. No",
                "Roll Number",
                "Roll No",
                default="N/A",
            )
        )

        phone_number = clean_text(
            pick(
                row,
                "Phone number",
                "Phone Number",
                "Phone",
                "Mobile number",
                "Mobile Number",
                "Mobile",
                "Contact",
                default="N/A",
            )
        )

        codechef_id = clean_username(
            pick(
                row,
                "CodeChef ID",
                "CodeChef Username",
                "Username",
                "Handle",
                default="N/A",
            )
        )

        # Total problems solved across all tracked contests.
        total_solved = sum(
            int(row[column])
            for column in starter_cols
        )

        # A contest counts as participated when at least one problem
        # was solved in that contest.
        contests_participated = sum(
            1
            for column in starter_cols
            if int(row[column]) > 0
        )

        contests_tracked = len(starter_cols)

        attendance_rate = (
            round(
                100 * contests_participated / contests_tracked,
                1,
            )
            if contests_tracked
            else 0.0
        )

        attendance_status = (
            "Present"
            if contests_participated > 0
            else "Absent"
        )

        per_contest = {
            column: int(row[column])
            for column in starter_cols
        }

        records.append({
            "name": name,
            "regNo": reg_no,
            "phoneNumber": phone_number,
            "codechefId": codechef_id,
            "memberType": member_type,
            "attendanceStatus": attendance_status,
            "totalProblemsSolved": total_solved,
            "contestsParticipated": contests_participated,
            "contestsTracked": contests_tracked,
            "attendanceRate": attendance_rate,
            "perContest": per_contest,
        })

    return pd.DataFrame(records)


# ---------------------------------------------------------------------------
# Dataset generation
# ---------------------------------------------------------------------------

def build_dataset() -> dict:
    """
    Build the complete dataset consumed by index_template.html.
    """

    print("📊 Loading member data...")

    core_df = load_cohort(
        CORE_EXCEL,
        "Core",
    )

    ffcs_df = load_cohort(
        FFCS_EXCEL,
        "FFCS",
    )

    # -----------------------------------------------------------------------
    # Combine cohorts
    # -----------------------------------------------------------------------

    frames = [
        df
        for df in (core_df, ffcs_df)
        if not df.empty
    ]

    if frames:
        all_df = pd.concat(
            frames,
            ignore_index=True,
        )
    else:
        all_df = pd.DataFrame()

    # -----------------------------------------------------------------------
    # Determine all tracked contests
    # -----------------------------------------------------------------------

    contest_numbers: set[int] = set()

    for df in (core_df, ffcs_df):

        if df.empty or "perContest" not in df.columns:
            continue

        for contest_data in df["perContest"]:

            if not isinstance(contest_data, dict):
                continue

            for column in contest_data.keys():

                match = STARTER_COL_RE.match(
                    str(column).strip()
                )

                if match:
                    contest_numbers.add(
                        int(match.group(1))
                    )

    contest_list = sorted(contest_numbers)

    contest_labels = [
        f"Starters {number}"
        for number in contest_list
    ]

    # -----------------------------------------------------------------------
    # KPI calculations
    # -----------------------------------------------------------------------

    total_members = len(all_df)

    if not all_df.empty:
        total_active_solvers = int(
            (
                all_df["totalProblemsSolved"] > 0
            ).sum()
        )
    else:
        total_active_solvers = 0

    core_count = len(core_df)

    if not core_df.empty:
        core_active = int(
            (
                core_df["attendanceStatus"] == "Present"
            ).sum()
        )
    else:
        core_active = 0

    core_turnout_rate = (
        round(
            100 * core_active / core_count,
            1,
        )
        if core_count
        else 0.0
    )

    ffcs_count = len(ffcs_df)

    if not ffcs_df.empty:
        ffcs_active = int(
            (
                ffcs_df["attendanceStatus"] == "Present"
            ).sum()
        )
    else:
        ffcs_active = 0

    ffcs_turnout_rate = (
        round(
            100 * ffcs_active / ffcs_count,
            1,
        )
        if ffcs_count
        else 0.0
    )

    # -----------------------------------------------------------------------
    # Top performer for each contest
    # -----------------------------------------------------------------------

    top_performers = {}

    for label in contest_labels:

        if all_df.empty:
            continue

        values = all_df["perContest"].apply(
            lambda data: int(data.get(label, 0))
            if isinstance(data, dict)
            else 0
        )

        if values.empty:
            continue

        max_value = int(values.max())

        if max_value <= 0:
            continue

        best_index = values.idxmax()

        top_performers[label] = {
            "name": str(
                all_df.loc[best_index, "name"]
            ),
            "memberType": str(
                all_df.loc[best_index, "memberType"]
            ),
            "solved": max_value,
        }

    # -----------------------------------------------------------------------
    # Overall top performer
    # -----------------------------------------------------------------------

    overall_top_performer = None

    if (
        not all_df.empty
        and all_df["totalProblemsSolved"].max() > 0
    ):
        best_index = all_df[
            "totalProblemsSolved"
        ].idxmax()

        overall_top_performer = {
            "name": str(
                all_df.loc[best_index, "name"]
            ),
            "memberType": str(
                all_df.loc[best_index, "memberType"]
            ),
            "solved": int(
                all_df.loc[
                    best_index,
                    "totalProblemsSolved"
                ]
            ),
        }

    leaderboard = []
    if not all_df.empty:
        ranked = all_df.sort_values(
            by=["totalProblemsSolved", "contestsParticipated", "name"],
            ascending=[False, False, True],
        ).head(3)
        leaderboard = [
            {
                "rank": rank,
                "name": str(row["name"]),
                "memberType": str(row["memberType"]),
                "contestsParticipated": int(row["contestsParticipated"]),
                "solved": int(row["totalProblemsSolved"]),
            }
            for rank, (_, row) in enumerate(ranked.iterrows(), start=1)
            if int(row["totalProblemsSolved"]) > 0
        ]

    # -----------------------------------------------------------------------
    # KPI object
    # -----------------------------------------------------------------------

    kpis = {
        "totalMembers": total_members,
        "totalActiveSolvers": total_active_solvers,

        "coreCount": core_count,
        "coreTurnoutRate": core_turnout_rate,

        "ffcsCount": ffcs_count,
        "ffcsTurnoutRate": ffcs_turnout_rate,

        "contestsTracked": len(contest_list),

        "overallTopPerformer": overall_top_performer,

        "topPerformersByContest": top_performers,

        "leaderboard": leaderboard,
    }

    # -----------------------------------------------------------------------
    # Member records
    # -----------------------------------------------------------------------

    members = dataframe_to_records(all_df)

    # -----------------------------------------------------------------------
    # Meeting attendance
    # -----------------------------------------------------------------------

    meeting_records = []
    meeting_columns: list[str] = []

    if os.path.exists(MEETING_EXCEL):

        try:
            meeting_df = pd.read_excel(MEETING_EXCEL)

            if not meeting_df.empty:

                meeting_df = meeting_df.copy()

                meeting_df.columns = [
                    str(column).strip()
                    for column in meeting_df.columns
                ]

                meeting_columns = list(
                    meeting_df.columns
                )

                meeting_records = dataframe_to_records(
                    meeting_df
                )

        except Exception as exc:
            print(
                f"⚠️ Could not read {MEETING_EXCEL}: {exc}"
            )

    # -----------------------------------------------------------------------
    # Final dataset
    # -----------------------------------------------------------------------

    return {
        "generatedAt": datetime.now().isoformat(
            timespec="seconds"
        ),

        "contestList": contest_list,

        "contestLabels": contest_labels,

        "kpis": kpis,

        "members": members,

        "meeting": {
            "columns": meeting_columns,
            "records": meeting_records,
        },
    }


# ---------------------------------------------------------------------------
# HTML compilation
# ---------------------------------------------------------------------------

def compile_html(dataset: dict) -> str:
    """
    Inject DASHBOARD_DATA into index_template.html.
    """

    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError(
            f"{TEMPLATE_FILE} not found. "
            "Keep index_template.html in the same folder "
            "as export_static.py."
        )

    with open(
        TEMPLATE_FILE,
        "r",
        encoding="utf-8",
    ) as file:
        template = file.read()

    marker = "/*__DASHBOARD_DATA__*/"

    if marker not in template:
        raise ValueError(
            "Could not find the DASHBOARD_DATA placeholder "
            "in index_template.html.\n\n"
            "The template must contain exactly:\n\n"
            "/*__DASHBOARD_DATA__*/"
        )

    payload_json = json.dumps(
        dataset,
        ensure_ascii=False,
        separators=(",", ":"),
    )

    injected = template.replace(
        marker,
        f"const DASHBOARD_DATA = {payload_json};",
        1,
    )

    return injected


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=" * 60)
    print("CodeChef Contest Tracking Hub - Static Export")
    print("=" * 60)

    dataset = build_dataset()

    html = compile_html(dataset)

    with open(
        OUTPUT_HTML,
        "w",
        encoding="utf-8",
    ) as file:
        file.write(html)

    kpis = dataset["kpis"]

    print()
    print(f"✅ Generated: {OUTPUT_HTML}")
    print(f"   Core members:       {kpis['coreCount']}")
    print(f"   FFCS members:       {kpis['ffcsCount']}")
    print(f"   Contests tracked:   {kpis['contestsTracked']}")
    print(
        f"   Active solvers:     "
        f"{kpis['totalActiveSolvers']}"
    )

    if dataset["meeting"]["records"]:
        print(
            f"   Meeting records:    "
            f"{len(dataset['meeting']['records'])}"
        )
    else:
        print("   Meeting records:    0")

    print()
    print("🎉 Static dashboard ready.")
    print("=" * 60)


if __name__ == "__main__":
    main()

