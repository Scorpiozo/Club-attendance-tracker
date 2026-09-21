"""
member_ops.py
--------------
Member-management layer for the live NiceGUI app (app.py).

Adds, on top of the existing Excel-file architecture (no database):
  - Single "Add Member" + CSV/Excel bulk import with column mapping,
    duplicate detection, and a single Batch applied to the whole import.
  - Batch + attendance-range + member-type + search multi-filtering.
  - A lightweight attendance % (from Meeting_Attendance.xlsx "Meeting *"
    columns) used to decide review-eligibility (0%-30% inclusive).
  - A Review workflow (task/description/status/notes/outcome) persisted
    to Review_Records.xlsx, plus permanent member deletion.

Existing files (CP_Members.xlsx, FFCS_Members.xlsx, Meeting_Attendance.xlsx,
ContestAttendance.py's scraping/export pipeline) are read/updated using the
same pandas + openpyxl approach already used elsewhere in this project and
are never restructured.
"""
from __future__ import annotations

import io
import os
import re
from datetime import datetime

import pandas as pd

import ContestAttendance as ca

CORE_EXCEL_FILE = "CP_Members.xlsx"
FFCS_EXCEL_FILE = "FFCS_Members.xlsx"
MEETING_EXCEL_FILE = "Meeting_Attendance.xlsx"
REVIEW_EXCEL_FILE = "Review_Records.xlsx"

COHORT_FILES = {"Core": CORE_EXCEL_FILE, "FFCS": FFCS_EXCEL_FILE}

# Canonical roster columns every cohort file should carry. "Register
# number" and "CodeChef ID" used to be kept as separate alias columns
# duplicating "Registration Number" and "Username" -- consolidated now.
ROSTER_COLUMNS = [
    "Name", "Registration Number", "Phone number", "Username",
    "Member Type", "Batch", "Attendance Status",
]

# Fields the bulk-import mapping step asks about. Batch is deliberately
# excluded here: the whole import gets ONE batch value chosen by the user.
IMPORT_TARGET_FIELDS = ["Name", "Registration Number", "Phone number", "CodeChef ID"]

REVIEW_STATUSES = ["Open", "In Progress", "Completed", "Failed", "Incomplete"]
REVIEW_COLUMNS = [
    "Review Date", "Member Type", "Username", "Name", "Registration Number",
    "Task", "Task Description", "Status", "Notes", "Outcome",
]

ATTENDANCE_LOW, ATTENDANCE_HIGH = 0, 30  # inclusive review-eligible band


# --------------------------------------------------------------------------
# Roster load / save (thin wrappers so every caller stays in sync)
# --------------------------------------------------------------------------

def load_cohort(member_type: str) -> pd.DataFrame:
    excel_file = COHORT_FILES[member_type]
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
    else:
        df = pd.DataFrame(columns=ROSTER_COLUMNS)
    return ca._standardize_frame(df, member_type)


def save_cohort(member_type: str, df: pd.DataFrame) -> None:
    df.to_excel(COHORT_FILES[member_type], index=False)


def load_all_members() -> pd.DataFrame:
    core = load_cohort("Core")
    ffcs = load_cohort("FFCS")
    if core.empty and ffcs.empty:
        return pd.DataFrame(columns=ROSTER_COLUMNS)
    return pd.concat([core, ffcs], ignore_index=True, sort=False)


def known_batches() -> list[str]:
    df = load_all_members()
    if df.empty or "Batch" not in df.columns:
        return [ca.DEFAULT_BATCH_LABEL]
    batches = sorted({str(b).strip() for b in df["Batch"].dropna() if str(b).strip()})
    return batches or [ca.DEFAULT_BATCH_LABEL]


# --------------------------------------------------------------------------
# Meeting-sheet sync (so new/removed members stay consistent everywhere)
# --------------------------------------------------------------------------

def load_meeting_data() -> pd.DataFrame:
    if os.path.exists(MEETING_EXCEL_FILE):
        return pd.read_excel(MEETING_EXCEL_FILE)
    df_all = load_all_members()
    base_cols = [c for c in ["Name", "Registration Number", "Phone number", "Username", "Member Type"] if c in df_all.columns]
    return df_all[base_cols].copy() if not df_all.empty else pd.DataFrame(
        columns=["Name", "Registration Number", "Phone number", "Username", "Member Type"]
    )


def meeting_columns(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if str(c).startswith("Meeting ")]


def _sync_meeting_sheet_add(new_rows: pd.DataFrame) -> None:
    """Append newly-added members to Meeting_Attendance.xlsx so they show up
    next time attendance is marked. Meetings held BEFORE they joined are
    left blank (not 0/absent) — they weren't members yet, so it would be
    wrong (and skews their attendance %, wrongly flagging them for review)
    to count those as absences. Attendance % only ever averages meetings
    that actually have a value for that member."""
    m_df = load_meeting_data()
    cols = meeting_columns(m_df)
    add_rows = new_rows[["Name", "Registration Number", "Phone number", "Username", "Member Type"]].copy()
    for c in cols:
        add_rows[c] = pd.NA
    if m_df.empty:
        m_df = add_rows
    else:
        existing_usernames = set(m_df["Username"].astype(str).str.strip())
        add_rows = add_rows[~add_rows["Username"].astype(str).str.strip().isin(existing_usernames)]
        if not add_rows.empty:
            m_df = pd.concat([m_df, add_rows], ignore_index=True, sort=False)
    m_df.to_excel(MEETING_EXCEL_FILE, index=False)


def _sync_meeting_sheet_remove(username: str) -> None:
    if not os.path.exists(MEETING_EXCEL_FILE):
        return
    m_df = pd.read_excel(MEETING_EXCEL_FILE)
    if m_df.empty:
        return
    m_df = m_df[m_df["Username"].astype(str).str.strip() != str(username).strip()]
    m_df.to_excel(MEETING_EXCEL_FILE, index=False)


# --------------------------------------------------------------------------
# Attendance % (drives the review-eligibility filter)
# --------------------------------------------------------------------------

def attendance_lookup() -> dict[str, float | None]:
    """Username -> attendance % (mean of 'Meeting *' columns * 100), or
    None if no meetings have been recorded yet for that member."""
    m_df = load_meeting_data()
    cols = meeting_columns(m_df)
    result: dict[str, float | None] = {}
    if m_df.empty:
        return result
    for _, row in m_df.iterrows():
        username = str(row.get("Username", "")).strip()
        if not username:
            continue
        if not cols:
            result[username] = None
            continue
        values = pd.to_numeric(row[cols], errors="coerce").dropna()
        result[username] = float(values.mean() * 100) if len(values) else None
    return result


def with_attendance(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy of df with 'Attendance %' (meetings) and
    'Contest Attendance %' (CodeChef Starters rounds) merged in."""
    lookup = attendance_lookup()
    out = df.copy()
    out["Attendance %"] = out["Username"].astype(str).str.strip().map(lookup)
    out["Contest Attendance %"] = out.apply(contest_attendance_pct, axis=1)
    return out


def contest_attendance_pct(row) -> float | None:
    """% of tracked CodeChef 'Starters N' rounds a member actually
    participated in (entered/solved >=1 problem), i.e. the original
    "absentee for 0 contests" signal, generalized from a 0/nonzero flag
    into a proper percentage so it can share the same 0%-30% review band
    as meeting attendance instead of only catching total no-shows."""
    starter_cols = [c for c in row.index if re.match(r"^Starters\s+\d+$", str(c))]
    if not starter_cols:
        return None
    values = pd.to_numeric(pd.Series([row[c] for c in starter_cols]), errors="coerce").dropna()
    if len(values) == 0:
        return None
    participated = int((values > 0).sum())
    return participated / len(starter_cols) * 100


def is_review_eligible(attendance_pct) -> bool:
    if attendance_pct is None or pd.isna(attendance_pct):
        return False
    return ATTENDANCE_LOW <= attendance_pct <= ATTENDANCE_HIGH


# --------------------------------------------------------------------------
# Filtering (Batch + attendance range + member type + search, combined)
# --------------------------------------------------------------------------

def filter_members(
    df: pd.DataFrame,
    batch: str | None = None,
    member_type: str | None = None,
    attendance_min: float | None = None,
    attendance_max: float | None = None,
    contest_attendance_min: float | None = None,
    contest_attendance_max: float | None = None,
    search: str | None = None,
) -> pd.DataFrame:
    out = with_attendance(df)

    if batch and batch != "All":
        out = out[out["Batch"].astype(str) == batch]

    if member_type and member_type != "All":
        out = out[out["Member Type"].astype(str) == member_type]

    if attendance_min is not None or attendance_max is not None:
        lo = attendance_min if attendance_min is not None else 0
        hi = attendance_max if attendance_max is not None else 100
        out = out[out["Attendance %"].apply(lambda v: v is not None and not pd.isna(v) and lo <= v <= hi)]

    if contest_attendance_min is not None or contest_attendance_max is not None:
        lo = contest_attendance_min if contest_attendance_min is not None else 0
        hi = contest_attendance_max if contest_attendance_max is not None else 100
        out = out[out["Contest Attendance %"].apply(lambda v: v is not None and not pd.isna(v) and lo <= v <= hi)]

    if search:
        s = search.strip().lower()
        if s:
            mask = (
                out["Name"].astype(str).str.lower().str.contains(s, na=False)
                | out["Registration Number"].astype(str).str.lower().str.contains(s, na=False)
                | out["Username"].astype(str).str.lower().str.contains(s, na=False)
            )
            out = out[mask]

    return out


# --------------------------------------------------------------------------
# Add Member (single)
# --------------------------------------------------------------------------

def add_single_member(
    member_type: str,
    name: str,
    registration_number: str,
    phone: str,
    codechef_id: str,
    batch: str,
) -> tuple[bool, str]:
    df = load_cohort(member_type)

    reg = str(registration_number).strip()
    handle = str(codechef_id).strip()
    if not name.strip() or not reg:
        return False, "Name and Registration Number are required."

    existing = df["Registration Number"].astype(str).str.strip().str.lower()
    if reg.lower() in set(existing):
        return False, f"A member with Registration Number '{reg}' already exists."

    new_row = {
        "Name": name.strip(),
        "Registration Number": reg,
        "Phone number": str(phone).strip() or "N/A",
        "Username": handle or "N/A",
        "Member Type": member_type,
        "Batch": batch.strip() or ca.DEFAULT_BATCH_LABEL,
        "Attendance Status": "Unknown",
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True, sort=False)
    save_cohort(member_type, df)
    _sync_meeting_sheet_add(pd.DataFrame([new_row]))
    return True, f"Added {name.strip()} to the {member_type} roster."


# --------------------------------------------------------------------------
# Bulk import: parse -> guess column mapping -> preview/dupe-check -> commit
# --------------------------------------------------------------------------

def parse_uploaded_table(filename: str, content: bytes) -> pd.DataFrame:
    buf = io.BytesIO(content)
    if filename.lower().endswith(".csv"):
        return pd.read_csv(buf)
    return pd.read_excel(buf)


def guess_column_mapping(raw_df: pd.DataFrame) -> dict[str, str | None]:
    """Best-effort auto-detect of raw_df columns -> IMPORT_TARGET_FIELDS.
    Returns None for any field it isn't confident about, so the caller can
    prompt the user to map that column manually."""
    cols_lower = {str(c).lower(): c for c in raw_df.columns}

    def find(*keywords):
        matches = [orig for low, orig in cols_lower.items() if any(k in low for k in keywords)]
        # Confident only when exactly one column matches.
        return matches[0] if len(matches) == 1 else None

    return {
        "Name": find("name"),
        "Registration Number": find("reg", "roll"),
        "Phone number": find("phone", "mobile", "contact"),
        "CodeChef ID": find("codechef", "username", "handle"),
    }

def apply_mapping(raw_df: pd.DataFrame, mapping: dict[str, str], member_type: str, batch: str) -> pd.DataFrame:
    """Build a standardized roster dataframe from raw_df using a confirmed
    (possibly user-adjusted) column mapping. IMPORT_TARGET_FIELDS uses
    "CodeChef ID" only as the user-facing label for this mapping step;
    the resulting column is the canonical "Username"."""
    out = pd.DataFrame()
    for field in IMPORT_TARGET_FIELDS:
        src = mapping.get(field)
        out[field] = raw_df[src].astype(str).str.strip() if src and src in raw_df.columns else "N/A"

    out = out.rename(columns={"CodeChef ID": "Username"})
    out["Username"] = out["Username"].str.rstrip("/").str.split("/").str[-1]
    out["Member Type"] = member_type
    out["Batch"] = batch.strip() or ca.DEFAULT_BATCH_LABEL
    out["Attendance Status"] = "Unknown"
    return out


def detect_duplicates(new_df: pd.DataFrame, member_type: str) -> pd.DataFrame:
    """Flag rows in new_df whose Registration Number already exists in
    EITHER cohort roster (a member shouldn't be double-entered)."""
    existing = load_all_members()
    existing_regs = set(existing["Registration Number"].astype(str).str.strip().str.lower())
    within_batch_seen = set()

    def check(reg):
        r = str(reg).strip().lower()
        if r in existing_regs or r in within_batch_seen:
            within_batch_seen.add(r)
            return True
        within_batch_seen.add(r)
        return False

    out = new_df.copy()
    out["Duplicate"] = out["Registration Number"].apply(check)
    return out


def commit_bulk_import(preview_df: pd.DataFrame, member_type: str, skip_duplicates: bool = True) -> tuple[int, int]:
    """Appends non-duplicate (unless skip_duplicates=False) rows to the
    target cohort roster and syncs the meeting sheet. Returns
    (imported_count, skipped_count)."""
    rows = preview_df.copy()
    if skip_duplicates and "Duplicate" in rows.columns:
        skipped = int(rows["Duplicate"].sum())
        rows = rows[~rows["Duplicate"]]
    else:
        skipped = 0
    rows = rows.drop(columns=["Duplicate"], errors="ignore")

    if rows.empty:
        return 0, skipped

    df = load_cohort(member_type)
    df = pd.concat([df, rows], ignore_index=True, sort=False)
    save_cohort(member_type, df)
    _sync_meeting_sheet_add(rows)
    return len(rows), skipped


# --------------------------------------------------------------------------
# Review workflow
# --------------------------------------------------------------------------

def load_reviews() -> pd.DataFrame:
    if os.path.exists(REVIEW_EXCEL_FILE):
        return pd.read_excel(REVIEW_EXCEL_FILE)
    return pd.DataFrame(columns=REVIEW_COLUMNS)


def add_review(
    member_type: str, username: str, name: str, registration_number: str,
    task: str, task_description: str, status: str, notes: str, outcome: str,
) -> None:
    reviews = load_reviews()
    new_row = {
        "Review Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Member Type": member_type,
        "Username": username,
        "Name": name,
        "Registration Number": registration_number,
        "Task": task,
        "Task Description": task_description,
        "Status": status,
        "Notes": notes,
        "Outcome": outcome,
    }
    reviews = pd.concat([reviews, pd.DataFrame([new_row])], ignore_index=True, sort=False)
    reviews.to_excel(REVIEW_EXCEL_FILE, index=False)


def review_eligible_members() -> pd.DataFrame:
    """Members are review-eligible if EITHER their meeting attendance OR
    their contest (Starters round) participation falls in the 0%-30%
    inclusive band. This generalizes the original rule (flag on exactly 0
    contests participated) into a percentage, and now also folds in
    meeting attendance rather than treating them as unrelated signals."""
    df = with_attendance(load_all_members())
    meeting_flag = df["Attendance %"].apply(is_review_eligible)
    contest_flag = df["Contest Attendance %"].apply(is_review_eligible)
    eligible = df[meeting_flag | contest_flag].copy()

    def reason(row):
        tags = []
        if is_review_eligible(row["Attendance %"]):
            tags.append("Meetings")
        if is_review_eligible(row["Contest Attendance %"]):
            tags.append("Contests")
        return " & ".join(tags)

    eligible["Review Reason"] = eligible.apply(reason, axis=1)
    return eligible


def delete_member(member_type: str, username: str) -> tuple[bool, str]:
    """Permanently removes a member from their cohort roster and the
    meeting sheet. No archive/trash is kept, per project requirements."""
    df = load_cohort(member_type)
    before = len(df)
    df = df[df["Username"].astype(str).str.strip() != str(username).strip()]
    if len(df) == before:
        return False, "Member not found."
    save_cohort(member_type, df)
    _sync_meeting_sheet_remove(username)
    return True, "Member permanently deleted."
