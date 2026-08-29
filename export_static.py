from __future__ import annotations

import json
import os
import re
from datetime import datetime

import pandas as pd

CORE_EXCEL = "CP_Members.xlsx"
FFCS_EXCEL = "FFCS_Members.xlsx"

FFCS_ROSTER_FALLBACK = "cccp202627.xlsx"  # raw Google Form export, no attendance yet
MEETING_EXCEL = "Meeting_Attendance.xlsx"

TEMPLATE_FILE = "index_template.html"
OUTPUT_HTML = "index.html"

STARTER_COL_RE = re.compile(r"^Starters\s+(\d+)$")


# --------------------------------------------------------------------------
# Loading & standardizing
# --------------------------------------------------------------------------

def _starter_columns(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if STARTER_COL_RE.match(str(c))]


def _pick(df_row: pd.Series, *candidates, default="N/A"):
    for c in candidates:
        if c in df_row.index and pd.notna(df_row[c]) and str(df_row[c]).strip():
            return df_row[c]
    return default


def load_cohort(excel_file: str, member_type: str, fallback_roster: str | None = None) -> pd.DataFrame:
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
    elif fallback_roster and os.path.exists(fallback_roster):
        raw = pd.read_excel(fallback_roster)
        cols = {str(c).lower(): c for c in raw.columns}
        name_col = next((cols[k] for k in cols if "name" in k), None)
        reg_col = next((cols[k] for k in cols if "reg" in k or "roll" in k), None)
        user_col = next(
            (cols[k] for k in cols if "codechef" in k or "username" in k or "handle" in k),
            None,
        )
        df = pd.DataFrame({
            "Name": raw[name_col] if name_col else "N/A",
            "Register number": raw[reg_col] if reg_col else "N/A",
            "Username": (raw[user_col].astype(str).str.strip().str.rstrip("/").str.split("/").str[-1]
                         if user_col else "N/A"),
        })
    else:
        return pd.DataFrame()

    if df.empty:
        return df

    starter_cols = _starter_columns(df)
    for c in starter_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    records = []
    for _, row in df.iterrows():
        name = _pick(row, "Name")
        reg_no = _pick(row, "Registration Number", "Register number")
        codechef_id = _pick(row, "CodeChef ID", "Username")
        total_solved = int(sum(row[c] for c in starter_cols)) if starter_cols else 0
        contests_participated = int(sum(1 for c in starter_cols if row[c] > 0)) if starter_cols else 0
        attendance_rate = (
            round(100 * contests_participated / len(starter_cols), 1) if starter_cols else 0.0
        )
        per_contest = {c: int(row[c]) for c in starter_cols}

        records.append({
            "name": str(name),
            "regNo": str(reg_no),
            "codechefId": str(codechef_id),
            "memberType": member_type,
            "attendanceStatus": "Present" if contests_participated > 0 else "Absent",
            "totalProblemsSolved": total_solved,
            "contestsParticipated": contests_participated,
            "contestsTracked": len(starter_cols),
            "attendanceRate": attendance_rate,
            "perContest": per_contest,
        })

    return pd.DataFrame(records)


def build_dataset() -> dict:
    core_df = load_cohort(CORE_EXCEL, "Core")
    ffcs_df = load_cohort(FFCS_EXCEL, "FFCS", fallback_roster=FFCS_ROSTER_FALLBACK)

    all_df = pd.concat([core_df, ffcs_df], ignore_index=True) if not (core_df.empty and ffcs_df.empty) else pd.DataFrame()

    # Union of every "Starters N" contest tracked across both cohorts, sorted.
    contest_numbers: set[int] = set()
    for df in (core_df, ffcs_df):
        if df.empty:
            continue
        for rec in df["perContest"]:
            for k in rec.keys():
                m = STARTER_COL_RE.match(k)
                if m:
                    contest_numbers.add(int(m.group(1)))
    contest_list = sorted(contest_numbers)
    contest_labels = [f"Starters {n}" for n in contest_list]

    # ---- KPIs -------------------------------------------------------
    total_members = len(all_df)
    total_active_solvers = int((all_df["totalProblemsSolved"] > 0).sum()) if not all_df.empty else 0

    ffcs_count = len(ffcs_df)
    ffcs_active = int((ffcs_df["attendanceStatus"] == "Present").sum()) if not ffcs_df.empty else 0
    ffcs_turnout_rate = round(100 * ffcs_active / ffcs_count, 1) if ffcs_count else 0.0

    core_count = len(core_df)
    core_active = int((core_df["attendanceStatus"] == "Present").sum()) if not core_df.empty else 0
    core_turnout_rate = round(100 * core_active / core_count, 1) if core_count else 0.0

    top_performers = {}
    for label in contest_labels:
        if all_df.empty:
            continue
        col_vals = all_df["perContest"].apply(lambda d: d.get(label, 0))
        max_val = col_vals.max() if len(col_vals) else 0
        if max_val > 0:
            best_idx = col_vals.idxmax()
            top_performers[label] = {
                "name": all_df.loc[best_idx, "name"],
                "memberType": all_df.loc[best_idx, "memberType"],
                "solved": int(col_vals.max()),
            }

    overall_top_performer = None
    if not all_df.empty and all_df["totalProblemsSolved"].max() > 0:
        idx = all_df["totalProblemsSolved"].idxmax()
        overall_top_performer = {
            "name": all_df.loc[idx, "name"],
            "memberType": all_df.loc[idx, "memberType"],
            "solved": int(all_df.loc[idx, "totalProblemsSolved"]),
        }

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
    }

    members = json.loads(all_df.to_json(orient="records")) if not all_df.empty else []

    # Meeting attendance (kept from the legacy dashboard, if present).
    meeting_records = []
    meeting_columns: list[str] = []
    if os.path.exists(MEETING_EXCEL):
        m_df = pd.read_excel(MEETING_EXCEL)
        if not m_df.empty:
            meeting_columns = list(m_df.columns)
            meeting_records = json.loads(m_df.to_json(orient="records"))

    return {
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "contestList": contest_list,
        "contestLabels": contest_labels,
        "kpis": kpis,
        "members": members,
        "meeting": {
            "columns": meeting_columns,
            "records": meeting_records,
        },
    }


# --------------------------------------------------------------------------
# HTML compilation
# --------------------------------------------------------------------------

def compile_html(dataset: dict) -> str:
    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError(
            f"{TEMPLATE_FILE} not found. Keep index_template.html next to export_static.py; "
            f"it is the static UI shell that this script injects data into."
        )

    with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
        template = f.read()

    payload_json = json.dumps(dataset, ensure_ascii=False)
    # Embed as a JSON literal inside a <script> tag so the page has zero
    # runtime network dependency for its data.
    injected = template.replace(
        "/*__DASHBOARD_DATA__*/",
        f"const DASHBOARD_DATA = {payload_json};",
    )
    return injected


def main():
    dataset = build_dataset()
    html = compile_html(dataset)

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ Generated static dashboard at '{OUTPUT_HTML}'")
    print(f"   Core members:  {dataset['kpis']['coreCount']}")
    print(f"   FFCS members:  {dataset['kpis']['ffcsCount']}")
    print(f"   Contests tracked: {dataset['kpis']['contestsTracked']}")
    print(f"   Total active solvers: {dataset['kpis']['totalActiveSolvers']}")


if __name__ == "__main__":
    main()
