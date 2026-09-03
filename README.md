# CodeChef Contest Tracking Hub

Tracks CodeChef "Starters" contest attendance for two cohorts — **Core CP
Members** and **FFCS Course Students** — and publishes a static, filterable
dashboard.

## Files

| File | Purpose |
|---|---|
| `ContestAttendance.py` | Scrapes CodeChef profiles and updates per-contest attendance in `CP_Members.xlsx` (Core) or `FFCS_Members.xlsx` (FFCS). |
| `export_static.py` | Aggregates both cohorts' data + KPIs into JSON and compiles it into `index.html` via `index_template.html`. |
| `index_template.html` | Static dashboard shell (HTML/CSS/JS). Edit this, not `index.html`, for UI changes. |
| `index.html` | **Generated file** — self-contained dashboard, safe to publish on GitHub Pages. No runtime API calls. |
| `app.py` | Optional live NiceGUI app for browsing rosters and recording meeting attendance interactively. |
| `CP_Members.xlsx` | Core CP member roster + `Starters N` attendance columns. |
| `FFCS_Members.xlsx` | FFCS student roster + `Starters N` attendance columns (created on first run of mode 2). |
| `cccp202627.xlsx` / `.csv` | Raw Google Form export used to seed FFCS roster if `FFCS_Members.xlsx` doesn't exist yet. |
| `members.txt` | Fallback plain list of CodeChef handles if no Excel/Form roster is found. |

## Usage

1. **Scrape attendance for a contest**
   ```bash
   python ContestAttendance.py
   # Choose 1 (Core) or 2 (FFCS), then enter the Starters number, e.g. 253
   ```
   This updates the relevant `.xlsx` with a new `Starters N` column, plus
   recalculated `Total Problems Solved`, `Contests Participated`, and
   `Attendance Status`.

2. **Build the static dashboard**
   ```bash
   python export_static.py
   ```
   Produces `index.html` — commit/push this for GitHub Pages, or open it
   locally. It has category tabs (All / Core / FFCS), live search by name
   or registration number, and sortable columns, all client-side.

3. **(Optional) Run the live management app**
   ```bash
   python app.py
   ```
   Browse both rosters, view meeting logs, and record new meeting
   attendance sessions (`Meeting_Attendance.xlsx`).

## Notes on the parser fix

Earlier versions returned 0 problems for contests like Starters 253 because
the regex required an exact single parenthetical right after the round
number and a specific set of following headers. The parser now:
- Locates *any* contest-type header (Starters, Lunchtime, Cook-Off, Long
  Challenge, Monday Munch) to bound each block, instead of guessing what
  comes next.
- Strips arbitrary division/rating noise — `(Div 3)`, `(Division 4)`,
  `(Rated for Div 3 & 4)`, etc. — without needing to enumerate every phrasing.
- Does not replace an existing value when CodeChef is unavailable or the
   profile response cannot be verified; the member is reported as skipped.
- Uses CodeChef's embedded contest history to distinguish participation from
   solved-problem count, so a participant who solved zero problems is not
   incorrectly treated as absent.
- Falls back gracefully (empty list, not a crash) on any unexpected page
  layout or network failure.
