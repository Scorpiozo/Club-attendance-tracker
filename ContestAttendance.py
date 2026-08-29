from __future__ import annotations

import random
import re
import sys
import time
import os
from pprint import pprint
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
from bs4 import BeautifulSoup
import pandas as pd

# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------
MEMBERS_TXT = "members.txt"

GFORM_FILE_XLSX = "cccp202627.xlsx"
GFORM_FILE_CSV = "cccp202627.csv"

CORE_FILENAME_XLSX = "CP_Members.xlsx"
FFCS_FILENAME_XLSX = "FFCS_Members.xlsx"

MEMBER_TYPE_CORE = "Core"
MEMBER_TYPE_FFCS = "FFCS"

CONTESTS_LABEL = "Contests"
CONTESTS_PARTICIPATED_LABEL = "Contests Participated"
TOTAL_PROBLEMS_LABEL = "Total Problems Solved"

# Canonical output columns shared by both cohorts.
STANDARD_COLUMNS = [
    "Name",
    "Registration Number",
    "CodeChef ID",
    "Member Type",
    "Attendance Status",
]

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 "
    "(KHTML, like Gecko) Version/17.0 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/123.0.0.0 Safari/537.36",
]

# Any of these headers can terminate the current contest's problem-list
# segment. The set intentionally covers CodeChef's other contest types so
# that a "Starters" block never accidentally swallows a following contest's
# entries.
CONTEST_HEADER_ALTERNATIVES = (
    r"Starters\s+\d+",
    r"Starters\s*Special\s*\d*",
    r"Monday\s+Munch\s*\d*",
    r"Lunchtime\s*\d*",
    r"Cook-?Off\s*\d*",
    r"Long\s+Challenge\s*\d*",
)
RE_ANY_HEADER = re.compile("|".join(f"(?:{p})" for p in CONTEST_HEADER_ALTERNATIVES), re.IGNORECASE)

RE_COUNT = re.compile(r"Contests\s*\((\d+)\)")

# Matches "Starters <num>" then swallows any division/rating noise that
# CodeChef renders next to it, e.g.:
#   Starters 253 (Rated for Div 3 & 4)
#   Starters 208 (Div 2)
#   Starters 190 (Division 4)
# The noise clause is intentionally generous: any run of parenthesised
# groups and/or bare words like "Rated", "for", "Div", "Division", "&",
# numerals and commas, so long as it doesn't itself start a new header.
RE_STARTERS_HEADER = re.compile(
    r"Starters\s+(\d+)\s*"
    r"((?:\((?:[^()]|\([^()]*\))*\)\s*)*)",  # zero or more (...) groups, allowing one level of nesting
    re.IGNORECASE,
)

# A "junk" line inside a problem segment that should never be treated as a
# problem name (stray labels CodeChef sometimes leaves in the flattened text).
RE_JUNK_LINE = re.compile(
    r"^\s*(rank|score|penalty|division|div\.?\s*\d|rated|view|result|"
    r"unrated|global\s+rank|country\s+rank)\b",
    re.IGNORECASE,
)

DEFAULT_RATE_LIMIT_SECONDS = 0.5  # min gap enforced inside fetch_html itself


# --------------------------------------------------------------------------
# Networking
# --------------------------------------------------------------------------

def build_user_url(handle: str) -> str:
    return f"https://www.codechef.com/users/{handle}"


def fetch_html(url: str, timeout: int = 15, max_retries: int = 4,
                rate_limit_seconds: float = DEFAULT_RATE_LIMIT_SECONDS) -> str:
    """Fetch a URL with retries, exponential backoff + jitter, and a
    randomised desktop User-Agent to reduce the odds of a 403 block.
    Returns "" (never raises) on unrecoverable failure."""

    last_error = None
    for attempt in range(max_retries):
        headers = {
            "User-Agent": random.choice(USER_AGENTS),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
            "Connection": "keep-alive",
        }
        try:
            time.sleep(rate_limit_seconds)  # simple per-request rate limiting
            req = Request(url, headers=headers)
            with urlopen(req, timeout=timeout) as resp:
                return resp.read().decode("utf-8", errors="replace")

        except HTTPError as e:
            last_error = e
            if e.code in (403, 429, 503):
                # Back off harder for explicit rate-limit / bot-block responses.
                wait_time = (2 ** attempt) + random.uniform(0, 1.0)
                time.sleep(wait_time)
                continue
            # Other HTTP errors (404, etc.) are not worth retrying.
            return ""

        except (URLError, TimeoutError) as e:
            last_error = e
            time.sleep((2 ** attempt) + random.uniform(0, 0.5))
            continue

        except Exception as e:  # noqa: BLE001 - fetch must never raise upstream
            last_error = e
            time.sleep(1.0)
            continue

    if last_error is not None:
        print(f"⚠️  Failed to fetch {url} after {max_retries} attempts: {last_error}",
              file=sys.stderr)
    return ""


# --------------------------------------------------------------------------
# Parsing
# --------------------------------------------------------------------------

def extract_contest_text(html: str) -> tuple[str, int]:
    """Isolate the raw "Contests" block from a CodeChef profile page."""
    if not html:
        return "", -1
    soup = BeautifulSoup(html, "html.parser")
    whole_text = soup.get_text(separator="\n", strip=False)

    participated_idx = whole_text.find(CONTESTS_PARTICIPATED_LABEL)
    start_idx = whole_text.find(CONTESTS_LABEL, participated_idx + 1 if participated_idx != -1 else 0)
    if start_idx == -1:
        return "", -1

    end_idx = whole_text.find(TOTAL_PROBLEMS_LABEL, start_idx)
    if end_idx == -1:
        end_idx = len(whole_text)

    section = whole_text[start_idx:end_idx].strip()
    m = RE_COUNT.search(section)
    expected = int(m.group(1)) if m else -1
    return section, expected


def _clean_problem_segment(segment: str) -> list[str]:
    """Turn the raw text between a contest header and the next header into
    a clean list of problem names, independent of how CodeChef happened to
    wrap whitespace/newlines around them."""
    if not segment:
        return []

    lines = [ln.strip() for ln in segment.replace("\r\n", "\n").splitlines()]
    lines = [ln for ln in lines if ln and not RE_JUNK_LINE.match(ln)]
    if not lines:
        return []

    joined = " ".join(lines)
    # Problems are typically comma separated in the flattened text; fall
    # back to treating the whole joined line as a single entry if there's
    # no comma at all (some layouts render one problem per line only).
    if "," in joined:
        parts = [p.strip(" .,-") for p in joined.split(",")]
    else:
        parts = [p.strip(" .,-") for p in lines]

    return [p for p in parts if p and not RE_JUNK_LINE.match(p)]


def parse_starters(contest_text: str) -> dict[int, list[str]]:
    """Parse every 'Starters N' block in the Contests section into
    {round_number: [problem names]}, tolerant of division-note noise like
    '(Div 3)', '(Division 4)', '(Rated for Div 3 & 4)', etc."""
    if not contest_text:
        return {}

    text = contest_text.replace("\r\n", "\n")
    contests: dict[int, list[str]] = {}

    # Collect every header in the section (any contest type) so we know
    # exactly where each Starters block ends, regardless of what comes next.
    all_headers = list(RE_ANY_HEADER.finditer(text))

    for m in RE_STARTERS_HEADER.finditer(text):
        round_no = int(m.group(1))
        seg_start = m.end()

        # Find the next header (of ANY contest type) that starts at/after
        # seg_start; that marks the end of this Starters block.
        seg_end = len(text)
        for h in all_headers:
            if h.start() >= seg_start:
                seg_end = h.start()
                break

        segment = text[seg_start:seg_end].strip()
        problems = _clean_problem_segment(segment)

        # A round can legitimately appear once; if CodeChef ever renders it
        # twice (pagination glitches) keep the richer (longer) list.
        if round_no not in contests or len(problems) > len(contests[round_no]):
            contests[round_no] = problems

    return contests


def get_user_contests(handle: str) -> dict[int, list[str]]:
    url = build_user_url(handle)
    html = fetch_html(url)
    contest_text, _ = extract_contest_text(html)
    return parse_starters(contest_text)


# --------------------------------------------------------------------------
# Data Initialization
# --------------------------------------------------------------------------

def _standardize_frame(df: pd.DataFrame, member_type: str) -> pd.DataFrame:
    """Ensure a member roster dataframe has all STANDARD_COLUMNS plus the
    legacy columns other tools in this project expect (Name, Register
    number, Phone number, Username), without clobbering any existing
    'Starters N' attendance columns."""
    out = df.copy()

    # Legacy columns used elsewhere in the pipeline (app.py, ContestAttendance)
    for col in ["Name", "Register number", "Phone number", "Username"]:
        if col not in out.columns:
            out[col] = "N/A"

    # New standardized columns requested for the FFCS/Core tracking feature.
    if "Registration Number" not in out.columns:
        out["Registration Number"] = out["Register number"]
    if "CodeChef ID" not in out.columns:
        out["CodeChef ID"] = out["Username"]
    if "Member Type" not in out.columns:
        out["Member Type"] = member_type
    else:
        out["Member Type"] = out["Member Type"].fillna(member_type)
    if "Attendance Status" not in out.columns:
        out["Attendance Status"] = "Unknown"

    return out


def _import_from_gform(path: str) -> pd.DataFrame | None:
    if path.lower().endswith(".csv"):
        raw_df = pd.read_csv(path)
    else:
        raw_df = pd.read_excel(path)

    df = pd.DataFrame()
    cols = {str(c).lower(): c for c in raw_df.columns}

    name_col = next((cols[k] for k in cols if "name" in k), None)
    reg_col = next((cols[k] for k in cols if "reg" in k or "roll" in k), None)
    phone_col = next((cols[k] for k in cols if "phone" in k or "mobile" in k or "contact" in k), None)
    user_col = next(
        (cols[k] for k in cols if "username" in k or "codechef" in k or "handle" in k),
        None,
    )

    df["Name"] = raw_df[name_col] if name_col else "N/A"
    df["Register number"] = raw_df[reg_col] if reg_col else "N/A"
    df["Phone number"] = raw_df[phone_col] if phone_col else "N/A"

    if user_col:
        raw_handle = raw_df[user_col].astype(str)
    else:
        raw_handle = raw_df.iloc[:, 0].astype(str)

    # CodeChef profile links are sometimes pasted instead of bare handles;
    # normalize "https://www.codechef.com/users/<handle>" -> "<handle>".
    df["Username"] = raw_handle.str.strip().str.rstrip("/").str.split("/").str[-1]

    return df


def initialize_member_dataframe(excel_file: str, member_type: str) -> pd.DataFrame:
    """Loads an existing roster Excel file, or seeds one from a Google Form
    export (xlsx/csv) or a plain members.txt handle list."""
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
        return _standardize_frame(df, member_type)

    gform_candidates = [GFORM_FILE_XLSX, GFORM_FILE_CSV]
    for gform_path in gform_candidates:
        if os.path.exists(gform_path):
            print(f"📥 Importing initial {member_type} roster from {gform_path}...")
            df = _import_from_gform(gform_path)
            return _standardize_frame(df, member_type)

    if os.path.exists(MEMBERS_TXT):
        print(f"ℹ️  No roster file found. Loading basic handles from {MEMBERS_TXT}...")
        with open(MEMBERS_TXT, "r") as f:
            handles = [line.strip() for line in f if line.strip()]
        df = pd.DataFrame({
            "Name": ["N/A"] * len(handles),
            "Register number": ["N/A"] * len(handles),
            "Phone number": ["N/A"] * len(handles),
            "Username": handles,
        })
        return _standardize_frame(df, member_type)

    print(f"❌ Error: No {excel_file}, Google Form export, or {MEMBERS_TXT} found.",
          file=sys.stderr)
    sys.exit(1)


# --------------------------------------------------------------------------
# Sequential Processing
# --------------------------------------------------------------------------
#
# CodeChef aggressively rate-limits concurrent requests from the same
# client (repeated HTTP 429s under ThreadPoolExecutor). Scraping is now
# strictly sequential, one profile at a time, with a polite delay between
# requests. This is slower but reliable - no more storms of 429s.

DEFAULT_INTER_REQUEST_DELAY_SECONDS = 1.5


def fetch_single_user(handle: str, starter_num: int) -> int:
    if not handle or handle.lower() == "nan":
        return 0
    try:
        contests = get_user_contests(handle)
        return len(contests.get(starter_num, []))
    except Exception as e:  # noqa: BLE001 - one bad profile must not kill the batch
        print(f"⚠️  Error processing @{handle}: {e}", file=sys.stderr)
        return 0


def process_attendance(
    excel_file: str,
    starter_num: int,
    member_type: str,
    inter_request_delay: float = DEFAULT_INTER_REQUEST_DELAY_SECONDS,
    on_progress=None,
) -> pd.DataFrame:
    """Sequentially scrapes every member's CodeChef profile for the given
    Starters round and writes the result back to `excel_file`.

    on_progress, if provided, is called after each member as:
        on_progress(completed: int, total: int, handle: str, count: int)
    so callers (e.g. a UI) can stream live progress instead of relying on
    stdout. If omitted, progress is printed to stdout (CLI behaviour).
    """
    members_table = initialize_member_dataframe(excel_file, member_type)
    col_name = f"Starters {starter_num}"

    if col_name not in members_table.columns:
        members_table[col_name] = 0

    handles = members_table["Username"].astype(str).str.strip().tolist()
    total = len(handles)

    print(f"\n🚀 Starting sequential scrape for {total} {member_type} members (Starters {starter_num})...")

    for completed, (idx, handle) in enumerate(enumerate(handles), start=1):
        count = fetch_single_user(handle, starter_num)
        members_table.loc[idx, col_name] = int(count)
        members_table.loc[idx, "Attendance Status"] = "Present" if count > 0 else "Absent"

        if on_progress:
            on_progress(completed, total, handle, count)
        else:
            status_icon = "✅" if count > 0 else "➖"
            print(f"[{completed}/{total}] {status_icon} @{handle}: {count} problem(s) solved in Starters {starter_num}")

        # Be polite between requests, but skip the delay after the last one.
        if completed < total and inter_request_delay > 0:
            time.sleep(inter_request_delay)

    # Recompute aggregate stats across every "Starters N" column tracked so far.
    starter_cols = [c for c in members_table.columns if re.match(r"^Starters\s+\d+$", str(c))]
    members_table["Total Problems Solved"] = members_table[starter_cols].apply(
        lambda row: sum(int(v) for v in row if pd.notna(v)), axis=1
    )
    members_table["Contests Participated"] = members_table[starter_cols].apply(
        lambda row: sum(1 for v in row if pd.notna(v) and int(v) > 0), axis=1
    )

    members_table.to_excel(excel_file, index=False)
    print(f"\n🎉 Success! Updated {excel_file} with column '{col_name}'.")
    return members_table


# --------------------------------------------------------------------------
# Execution
# --------------------------------------------------------------------------

if __name__ == "__main__":
    print("Choose operation mode:")
    print("1. Process attendance for Core CP Members")
    print("2. Process attendance for FFCS Course Students")
    print("3. Check contests for a specific user")

    try:
        mode = int(input("Enter choice (1, 2, or 3): ").strip())

        if mode == 3:
            handle = input("Enter CodeChef username: ").strip()
            pprint(get_user_contests(handle))
            sys.exit(0)

        if mode == 1:
            excel_file, member_type = CORE_FILENAME_XLSX, MEMBER_TYPE_CORE
        elif mode == 2:
            excel_file, member_type = FFCS_FILENAME_XLSX, MEMBER_TYPE_FFCS
        else:
            raise ValueError("mode must be 1, 2, or 3")

        starter_num = int(input("Enter Starters contest number (e.g., 253): ").strip())

        process_attendance(excel_file, starter_num, member_type)

    except ValueError:
        print("Invalid input.", file=sys.stderr)
        sys.exit(1)
