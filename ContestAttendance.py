import re
import sys
import time
import os
from pprint import pprint
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
from bs4 import BeautifulSoup
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- Constants ---
MEMBERS_TXT = "members.txt"
GFORM_FILE_XLSX = "cccp202627.xlsx"
GFORM_FILE_CSV = "cccp202627.csv"

FILENAME_XLSX = "CP_Members.xlsx"
FFCS_FILENAME_XLSX = "FFCS_Members.xlsx"

CONTESTS_LABEL = "Contests"
CONTESTS_PARTICIPATED_LABEL = "Contests Participated"
TOTAL_PROBLEMS_LABEL = "Total Problems Solved"

RE_COUNT = re.compile(r'Contests\s*\((\d+)\)')
RE_STARTERS = re.compile(
    r'Starters\s+(\d+)\b[^(]*\([^)]*\)\s*(.*?)(?=\s*(?:Monday\s+Munch|Starters\s+\d+)|$)',
    re.DOTALL | re.IGNORECASE
)

# --- Helper Functions ---

def build_user_url(handle: str) -> str:
    return f"https://www.codechef.com/users/{handle}"

def fetch_html(url: str, timeout: int = 15, max_retries: int = 3) -> str:
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    
    for attempt in range(max_retries):
        try:
            req = Request(url, headers=headers)
            with urlopen(req, timeout=timeout) as resp:
                return resp.read().decode("utf-8", errors="replace")
        
        except HTTPError as e:
            if e.code == 429:
                wait_time = 2 ** attempt
                time.sleep(wait_time)
                continue
            return ""
            
        except (URLError, Exception):
            time.sleep(2)
            continue

    return ""

def extract_contest_text(html: str) -> tuple[str, int]:
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

def parse_starters(contest_text: str) -> dict[int, list[str]]:
    if not contest_text:
        return {}
    text = contest_text.replace("\r\n", "\n")
    contests = {}

    for m in RE_STARTERS.finditer(text):
        round_no = int(m.group(1))
        segment = m.group(2).strip()
        lines = [ln.strip() for ln in segment.splitlines() if ln.strip()]
        all_text = " ".join(lines)
        problems = [p.strip() for p in all_text.split(",") if p.strip()]
        contests[round_no] = problems

    return contests

def get_user_contests(handle: str) -> dict[int, list[str]]:
    url = build_user_url(handle)
    html = fetch_html(url)
    contest_text, _ = extract_contest_text(html)
    return parse_starters(contest_text)

# --- Data Initialization & GForm Importer ---

def initialize_member_dataframe(excel_file: str) -> pd.DataFrame:
    """Loads existing Excel file or imports metadata from Google Form / members.txt."""
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
        # Ensure core columns exist
        for col in ["Name", "Register number", "Phone number", "Username"]:
            if col not in df.columns:
                df[col] = "N/A"
        return df

    # Check for Google Form export files
    gform_path = None
    if os.path.exists(GFORM_FILE_XLSX):
        gform_path = GFORM_FILE_XLSX
        raw_df = pd.read_excel(gform_path)
    elif os.path.exists(GFORM_FILE_CSV):
        gform_path = GFORM_FILE_CSV
        raw_df = pd.read_csv(gform_path)
    else:
        raw_df = None

    if raw_df is not None:
        print(f"📥 Importing initial member roster from {gform_path}...")
        df = pd.DataFrame()
        
        # Fuzzy column matching for Google Form export headers
        cols = {str(c).lower(): c for c in raw_df.columns}
        
        name_col = next((cols[k] for k in cols if "name" in k), None)
        reg_col = next((cols[k] for k in cols if "reg" in k or "roll" in k), None)
        phone_col = next((cols[k] for k in cols if "phone" in k or "mobile" in k or "contact" in k), None)
        user_col = next((cols[k] for k in cols if "username" in k or "codechef" in k or "handle" in k), None)

        df["Name"] = raw_df[name_col] if name_col else "N/A"
        df["Register number"] = raw_df[reg_col] if reg_col else "N/A"
        df["Phone number"] = raw_df[phone_col] if phone_col else "N/A"
        df["Username"] = raw_df[user_col] if user_col else raw_df.iloc[:, 0]
        
        df["Username"] = df["Username"].astype(str).str.strip()
        return df

    # Fallback to simple txt file
    if os.path.exists(MEMBERS_TXT):
        print(f"ℹ️ {GFORM_FILE_XLSX} not found. Loading basic handles from {MEMBERS_TXT}...")
        with open(MEMBERS_TXT, "r") as f:
            handles = [line.strip() for line in f if line.strip()]
        return pd.DataFrame({
            "Name": ["N/A"] * len(handles),
            "Register number": ["N/A"] * len(handles),
            "Phone number": ["N/A"] * len(handles),
            "Username": handles
        })

    print("❌ Error: No CP_Members.xlsx, Google Form export, or members.txt found.", file=sys.stderr)
    sys.exit(1)

# --- Multithreaded Processing ---

def fetch_single_user(index: int, handle: str, starter_num: int) -> tuple[int, str, int]:
    if not handle or handle.lower() == "nan":
        return index, handle, 0
    try:
        contests = get_user_contests(handle)
        count = len(contests.get(starter_num, []))
        return index, handle, count
    except Exception as e:
        return index, handle, 0

def process_attendance(excel_file: str, starter_num: int, max_workers: int = 5):
    members_table = initialize_member_dataframe(excel_file)
    col_name = f"Starters {starter_num}"
    
    if col_name not in members_table.columns:
        members_table[col_name] = 0

    handles = members_table["Username"].astype(str).str.strip().tolist()
    total = len(handles)
    
    print(f"\n🚀 Starting parallel scrape for {total} members (Starters {starter_num})...")

    # Multithreaded execution across handles
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [
            executor.submit(fetch_single_user, idx, handle, starter_num) 
            for idx, handle in enumerate(handles)
        ]
        
        completed = 0
        for future in as_completed(futures):
            idx, handle, count = future.result()
            members_table.loc[idx, col_name] = int(count)
            completed += 1
            print(f"[{completed}/{total}] ✅ @{handle}: {count} problem(s) solved in Starters {starter_num}")

    members_table.to_excel(excel_file, index=False)
    print(f"\n🎉 Success! Updated {excel_file} with column '{col_name}'.")

# --- Execution ---

if __name__ == "__main__":
    print("Choose operation mode:")
    print("1. Process attendance for CP Members")
    print("2. Process attendance for FFCS Members")
    print("3. Check contests for a specific user")

    try:
        mode = int(input("Enter choice (1, 2, or 3): ").strip())

        if mode == 3:
            handle = input("Enter CodeChef username: ").strip()
            pprint(get_user_contests(handle))
            sys.exit(0)

        excel_file = FILENAME_XLSX if mode == 1 else FFCS_FILENAME_XLSX
        starter_num = int(input("Enter Starters contest number (e.g., 208): ").strip())
        
        process_attendance(excel_file, starter_num, max_workers=5)

    except ValueError:
        print("Invalid input.", file=sys.stderr)
        sys.exit(1)