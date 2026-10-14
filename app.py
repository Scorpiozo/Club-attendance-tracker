from nicegui import ui
import pandas as pd
import os
import html
import threading
import queue
from datetime import date

import ContestAttendance as ca
import member_ops as mo

CORE_EXCEL_FILE = "CP_Members.xlsx"
FFCS_EXCEL_FILE = "FFCS_Members.xlsx"
MEETING_EXCEL_FILE = "Meeting_Attendance.xlsx"

COHORT_FILES = {"Core": CORE_EXCEL_FILE, "FFCS": FFCS_EXCEL_FILE}

BASE_COLS = ["Name", "Registration Number", "Phone number", "Username"]

# --- Data Loaders & Savers ---

def load_cohort(excel_file: str, member_type: str) -> pd.DataFrame:
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
    else:
        df = pd.DataFrame(columns=BASE_COLS)
    if "Member Type" not in df.columns:
        df["Member Type"] = member_type
    else:
        df["Member Type"] = df["Member Type"].fillna(member_type)
    return df


def load_all_members() -> pd.DataFrame:
    core = load_cohort(CORE_EXCEL_FILE, "Core")
    ffcs = load_cohort(FFCS_EXCEL_FILE, "FFCS")
    if core.empty and ffcs.empty:
        return pd.DataFrame(columns=BASE_COLS + ["Member Type"])
    return pd.concat([core, ffcs], ignore_index=True, sort=False)


def load_meeting_data() -> pd.DataFrame:
    if os.path.exists(MEETING_EXCEL_FILE):
        return pd.read_excel(MEETING_EXCEL_FILE)

    # Sync core member info from the combined roster if meeting excel doesn't exist yet
    df_all = load_all_members()
    if not df_all.empty:
        base_cols = [c for c in BASE_COLS + ["Member Type"] if c in df_all.columns]
        return df_all[base_cols].copy()

    return pd.DataFrame(columns=BASE_COLS + ["Member Type"])


def save_meeting_data(df: pd.DataFrame):
    df.to_excel(MEETING_EXCEL_FILE, index=False)


def starter_cols(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if str(c).startswith("Starters ")]


# --- App Layout ---

@ui.page('/')
def main_page():
    df_core = load_cohort(CORE_EXCEL_FILE, "Core")
    df_ffcs = load_cohort(FFCS_EXCEL_FILE, "FFCS")
    df_members = load_all_members()
    df_meetings = load_meeting_data()

    ui.colors(primary='#3b82f6', secondary='#22d3ee', dark=True)
    ui.dark_mode(True)
    ui.add_head_html(r'''
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        :root {
            --bg: #080b10;
            --bg-secondary: #0c1118;
            --surface: #101720;
            --surface-hover: #141d28;
            --surface-elevated: #182231;

            --border: #1b2735;
            --border-hover: #28558a;

            --text: #e8eef7;
            --text-secondary: #9aa8b8;
            --text-muted: #667487;

            --blue: #3b82f6;
            --blue-bright: #2b97fb;
            --cyan: #22d3ee;

            --green: #22c55e;
            --amber: #f59e0b;
            --red: #ef4444;
            --purple: #a78bfa;

            --radius: 10px;
            --transition: all 180ms ease;
        }

        :root, body, .q-field, .q-btn, .q-table, input, textarea, select {
            font-family: 'Inter', ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
        }
        .mono, .ag-theme-balham-dark, code, #log_area, .q-field.mono .q-field__native {
            font-family: 'JetBrains Mono', 'Fira Code', ui-monospace, monospace !important;
        }

        html, body { overflow-x: hidden; }
        body {
            background:
                radial-gradient(680px 320px at 18% -8%, rgba(59,130,246,0.10), transparent 60%),
                radial-gradient(520px 260px at 88% 0%, rgba(34,211,238,0.06), transparent 55%),
                var(--bg) !important;
        }
        h1, h2, h3, .text-lg, .text-xl, .font-bold, .font-extrabold, .font-semibold {
            letter-spacing: 0.01em;
        }

        ::selection { background: rgba(59,130,246,0.35); color: #fff; }

        ::-webkit-scrollbar { width: 9px; height: 9px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: #24354a; border-radius: 999px; border: 2px solid var(--bg); }
        ::-webkit-scrollbar-thumb:hover { background: #335c8f; }
        * { scrollbar-width: thin; scrollbar-color: #24354a transparent; }

        /* ---------------------------------------------------------------
           Surfaces & the signature border-lift interaction. Every card,
           panel and important container shares this one visual language
           instead of one-off styling per component.
        ------------------------------------------------------------------ */
        .bg-slate-900 { background-color: var(--bg) !important; }
        .bg-slate-800 {
            background-color: var(--surface) !important;
            border-radius: var(--radius) !important;
            box-shadow: 0 4px 18px rgba(0,0,0,0.18);
            transition: var(--transition);
        }
        .bg-slate-600, .bg-slate-500 { background-color: var(--surface-elevated) !important; }
        .border-slate-700, .border-rose-700 { border-color: var(--border) !important; }

        .text-slate-100, .text-slate-200 { color: var(--text) !important; }
        .text-slate-300 { color: var(--text-secondary) !important; }
        .text-slate-400, .text-slate-500 { color: var(--text-muted) !important; }

        .text-blue-400 { color: var(--blue-bright) !important; }
        .text-emerald-400 { color: var(--cyan) !important; }
        .text-purple-400 { color: var(--purple) !important; }
        .text-amber-400 { color: var(--amber) !important; }
        .text-rose-400 { color: var(--red) !important; }

        /* ---------------------------------------------------------------
           Buttons -- primary (electric blue), secondary (elevated
           charcoal), danger (restrained red). All share the lift effect.
        ------------------------------------------------------------------ */
        .bg-blue-600, .bg-emerald-600, .bg-purple-600 {
            background-color: rgba(59,130,246,0.10) !important;
            border: 1px solid rgba(59,130,246,0.35) !important;
            color: var(--blue-bright) !important;
            box-shadow: none !important;
            transition: var(--transition);
        }
        .hover\:bg-blue-500:hover, .hover\:bg-emerald-500:hover, .hover\:bg-purple-500:hover {
            background-color: rgba(59,130,246,0.18) !important;
            border-color: var(--blue-bright) !important;
            color: #fff !important;
            box-shadow: 0 8px 28px rgba(37,99,235,0.18);
            transform: translateY(-1px);
        }
        .bg-rose-600 {
            background-color: rgba(239,68,68,0.08) !important;
            border: 1px solid rgba(239,68,68,0.4) !important;
            color: #fca5a5 !important;
            transition: var(--transition);
        }
        .hover\:bg-rose-500:hover {
            background-color: rgba(239,68,68,0.16) !important;
            border-color: var(--red) !important;
            color: #fecaca !important;
            box-shadow: 0 8px 24px rgba(239,68,68,0.16);
            transform: translateY(-1px);
        }
        .bg-slate-600 {
            background-color: var(--surface-elevated) !important;
            border: 1px solid var(--border) !important;
            color: var(--text-secondary) !important;
            transition: var(--transition);
        }
        .hover\:bg-slate-500:hover {
            background-color: var(--surface-hover) !important;
            border-color: var(--border-hover) !important;
            transform: translateY(-1px);
        }

        .rounded-xl { border-radius: var(--radius) !important; }

        /* ---------------------------------------------------------------
           KPI cards
        ------------------------------------------------------------------ */
        .kpi-card {
            position: relative;
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 16px 18px;
            overflow: hidden;
            transition: var(--transition);
        }
        .kpi-card::before {
            content: "";
            position: absolute;
            inset: 0;
            background: radial-gradient(120px 60px at 15% 0%, var(--kpi-tint, rgba(59,130,246,0.14)), transparent 70%);
            pointer-events: none;
        }
        .kpi-card:hover {
            border-color: var(--border-hover);
            box-shadow: 0 8px 28px rgba(37,99,235,0.12);
            transform: translateY(-1px);
        }
        .kpi-card .kpi-label {
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            color: var(--text-muted);
            font-weight: 600;
        }
        .kpi-card .kpi-value {
            font-size: 1.9rem;
            font-weight: 700;
            margin-top: 6px;
            font-family: 'JetBrains Mono', ui-monospace, monospace;
            color: var(--text);
        }
        .kpi-card .kpi-sub {
            font-size: 0.72rem;
            color: var(--text-muted);
            margin-top: 4px;
        }
        .kpi-blue { --kpi-tint: rgba(59,130,246,0.16); }
        .kpi-blue .kpi-value { color: var(--blue-bright); }
        .kpi-cyan { --kpi-tint: rgba(34,211,238,0.16); }
        .kpi-cyan .kpi-value { color: var(--cyan); }
        .kpi-purple { --kpi-tint: rgba(167,139,250,0.16); }
        .kpi-purple .kpi-value { color: var(--purple); }
        .kpi-amber { --kpi-tint: rgba(245,158,11,0.16); }
        .kpi-amber .kpi-value { color: var(--amber); }

        /* ---------------------------------------------------------------
           Sidebar navigation
        ------------------------------------------------------------------ */
        .side-nav { padding: 8px 0; background: var(--bg-secondary) !important; }
        .side-nav .q-tabs { overflow: visible; }
        .side-nav .q-tabs__content { flex-direction: column; align-items: stretch; }
        .side-nav .q-tabs__arrow { display: none !important; }
        .side-nav .q-tab {
            justify-content: flex-start;
            text-transform: none;
            letter-spacing: normal;
            font-size: 0.85rem;
            font-weight: 500;
            min-height: 38px;
            height: 38px;
            padding: 0;
            border-radius: 7px;
            margin: 1px 8px;
            color: var(--text-muted);
            transition: var(--transition);
        }
        .side-nav .q-tab__content {
            flex-direction: row;
            justify-content: flex-start;
            align-items: center;
            gap: 10px;
            width: 100%;
            height: 100%;
            padding: 0 14px;
        }
        .side-nav .q-tab__icon { font-size: 18px; margin: 0; min-width: 18px; }
        .side-nav .q-tab__label { font-size: 0.85rem; line-height: 1; padding: 0; }
        .side-nav .q-tab--active {
            color: var(--blue-bright) !important;
            background: rgba(59,130,246,0.12);
            box-shadow: inset 0 0 0 1px rgba(59,130,246,0.35);
        }
        .side-nav .q-tab-indicator { display: none; }

        .side-nav .q-tab:not(.q-tab--active):hover {
            color: var(--text);
            background: var(--surface-hover);
        }
        .side-nav .q-tab .q-focus-helper { background: var(--blue) !important; }
        .side-nav .q-tab:not(.q-tab--active):hover .q-focus-helper { opacity: 0.06 !important; }
        .side-nav .q-tab--active .q-focus-helper { opacity: 0 !important; }
        .side-nav .q-ripple { color: var(--blue); }

        .side-nav .q-tab:nth-child(4),
        .side-nav .q-tab:nth-child(6),
        .side-nav .q-tab:nth-child(7) {
            margin-top: 14px;
            padding-top: 10px;
            border-top: 1px solid var(--border);
            border-radius: 0 0 7px 7px;
        }

        .nav-group-label {
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: var(--text-muted);
            padding: 14px 16px 4px;
        }
        .nav-group-label:first-child { padding-top: 6px; }

        .status-dot { width: 7px; height: 7px; border-radius: 999px; display: inline-block; }

        /* ---------------------------------------------------------------
           Forms
        ------------------------------------------------------------------ */
        .q-field--outlined .q-field__control {
            background: var(--surface-elevated);
            border-radius: 8px !important;
            transition: var(--transition);
        }
        .q-field--outlined .q-field__control:before { border-color: var(--border) !important; }
        .q-field--outlined.q-field--focused .q-field__control:before {
            border-color: var(--blue) !important;
            box-shadow: 0 0 0 3px rgba(59,130,246,0.18);
        }
        .q-field__label { color: var(--text-muted) !important; }
        .q-field--focused .q-field__label { color: var(--blue-bright) !important; }
        .q-field__native, .q-field__input { color: var(--text) !important; }

        /* ---------------------------------------------------------------
           ag-grid
        ------------------------------------------------------------------ */
        .ag-theme-balham-dark {
            --ag-background-color: var(--surface);
            --ag-odd-row-background-color: var(--bg-secondary);
            --ag-even-row-background-color: var(--surface);
            --ag-header-background-color: var(--surface-elevated);
            --ag-header-foreground-color: var(--text-secondary);
            --ag-foreground-color: var(--text);
            --ag-border-color: var(--border);
            --ag-row-hover-color: rgba(59,130,246,0.10);
            --ag-selected-row-background-color: rgba(59,130,246,0.16);
            --ag-font-size: 13px;
            --ag-cell-horizontal-padding: 12px;
        }
        .ag-theme-balham-dark .ag-header-cell-label { font-weight: 600; text-transform: uppercase; font-size: 0.68rem; letter-spacing: 0.04em; }
        .ag-theme-balham-dark .ag-root-wrapper { border-radius: var(--radius); overflow: hidden; }

        .desktop-table { display: block; }
        .mobile-cards { display: none; }
        .mobile-card {
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 12px 14px;
            background-color: var(--surface);
            transition: var(--transition);
        }
        .mobile-card:hover { border-color: var(--border-hover); transform: translateY(-1px); }
        .mobile-card .mc-row { display: flex; justify-content: space-between; gap: 12px; font-size: 0.8rem; padding: 3px 0; }
        .mobile-card .mc-label { color: var(--text-muted); }
        .mobile-card .mc-value { color: var(--text); text-align: right; }
        .mobile-card .mc-title { font-weight: 700; color: var(--text); font-size: 0.92rem; margin-bottom: 4px; }
        .mobile-card .mc-more { margin-top: 6px; border-top: 1px dashed var(--border); padding-top: 6px; }
        .mobile-card .mc-more summary { cursor: pointer; font-size: 0.74rem; color: var(--blue-bright); list-style: none; }
        .mobile-card .mc-more summary::-webkit-details-marker { display: none; }
        .mobile-card .mc-more summary::before { content: "+ "; }
        .mobile-card .mc-more[open] summary::before { content: "\2212 "; }

        /* ---------------------------------------------------------------
           Scrape console -- refined developer terminal
        ------------------------------------------------------------------ */
        .scrape-console {
            background: var(--bg-secondary);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 12px 14px;
        }
        .scrape-log {
            background: #05080c !important;
            border: 1px solid var(--border) !important;
            border-radius: 8px !important;
            color: var(--text-secondary) !important;
            padding: 10px 12px !important;
            line-height: 1.55;
            font-family: 'JetBrains Mono', ui-monospace, monospace !important;
        }
        .scrape-progress { border-radius: 999px; overflow: hidden; }

        .status-idle    { background: var(--text-muted); }
        .status-running { background: var(--blue); animation: pulse-dot 1.4s ease-in-out infinite; }
        .status-done    { background: var(--green); }
        .status-error   { background: var(--red); }
        @keyframes pulse-dot {
            0%, 100% { opacity: 1; box-shadow: 0 0 0 0 rgba(59,130,246,0.45); }
            50%      { opacity: 0.65; box-shadow: 0 0 0 4px rgba(59,130,246,0); }
        }

        /* ---------------------------------------------------------------
           Attendance roster rows + review rows
        ------------------------------------------------------------------ */
        .roster-row {
            border-bottom: 1px solid var(--border);
            padding: 7px 10px;
            border-radius: 6px;
            transition: var(--transition);
        }
        .roster-row:last-child { border-bottom: none; }
        .roster-row:hover { background: var(--surface-hover); }
        .roster-meta {
            font-size: 0.72rem;
            color: var(--text-muted);
            font-family: 'JetBrains Mono', ui-monospace, monospace;
        }

        .review-row {
            border: 1px solid var(--border);
            border-radius: var(--radius);
            background: var(--surface);
            transition: var(--transition);
        }
        .review-row:hover {
            border-color: var(--border-hover);
            box-shadow: 0 8px 28px rgba(37,99,235,0.12);
            transform: translateY(-1px);
        }
        .sev-red   { border-left: 3px solid var(--red); }
        .sev-amber { border-left: 3px solid var(--amber); }

        /* Dialogs: elevated surface, stronger border, subtle blue glow */
        .q-dialog .q-card {
            background: var(--surface-elevated) !important;
            border: 1px solid var(--border-hover) !important;
            border-radius: var(--radius) !important;
            box-shadow: 0 24px 60px rgba(0,0,0,0.55), 0 0 0 1px rgba(59,130,246,0.10);
        }

        .stat-divider { border-right: 1px solid var(--border); }

        .drawer-toggle-btn { display: none; }
        @media (max-width: 1024px) {
            .drawer-toggle-btn { display: inline-flex !important; }
        }

        /* ---------------------------------------------------------------
           Header
        ------------------------------------------------------------------ */
        .app-header {
            background: linear-gradient(to bottom, var(--bg-secondary), var(--bg)) !important;
            border-bottom: 1px solid var(--border) !important;
            box-shadow: 0 1px 0 rgba(59,130,246,0.10), 0 12px 24px -18px rgba(59,130,246,0.35);
        }
        .app-logo-mark {
            width: 30px; height: 30px;
            display: flex; align-items: center; justify-content: center;
            border-radius: 8px;
            background: radial-gradient(circle at 30% 20%, rgba(96,165,250,0.35), rgba(59,130,246,0.08));
            border: 1px solid rgba(59,130,246,0.35);
            color: var(--blue-bright);
        }

        /* ---------------------------------------------------------------
           Bulk import stepper
        ------------------------------------------------------------------ */
        .q-stepper { background: transparent !important; box-shadow: none !important; }
        .q-stepper__tab { color: var(--text-muted) !important; }
        .q-stepper__tab .q-stepper__dot { background: var(--surface-elevated); color: var(--text-muted); }
        .q-stepper__tab--active { color: var(--blue-bright) !important; }
        .q-stepper__tab--active .q-stepper__dot { background: var(--blue); color: #fff; }
        .q-stepper__tab--done { color: var(--green) !important; }
        .q-stepper__tab--done .q-stepper__dot { background: var(--green); color: #06210f; }
        .q-stepper__step-content { border-color: var(--border) !important; }

        /* ---------------------------------------------------------------
           Mobile responsiveness
        ------------------------------------------------------------------ */
        @media (max-width: 900px) {
            .kpi-grid-responsive { grid-template-columns: repeat(2, minmax(0, 1fr)) !important; }
            .mapping-grid-responsive { grid-template-columns: 1fr !important; }
        }
        @media (max-width: 720px) {
            .desktop-table { display: none; }
            .mobile-cards { display: flex; }
        }
        @media (max-width: 560px) {
            .kpi-grid-responsive { grid-template-columns: 1fr !important; }
            .page-container { padding: 10px !important; gap: 14px !important; }
            .q-header { flex-wrap: wrap; }
        }
    </style>
    ''')
    ui.query('body').classes('bg-slate-900 text-slate-100')

    left_drawer_ref = {}

    # Top Header
    with ui.header().classes('app-header items-center px-6 py-3 gap-3'):
        ui.button(icon='menu', on_click=lambda: left_drawer_ref['drawer'].toggle()).props('flat dense round color=blue-4').classes('drawer-toggle-btn')
        with ui.element('div').classes('app-logo-mark'):
            ui.icon('bolt', size='19px')
        with ui.column().classes('gap-0'):
            ui.label('Technical Department Hub').classes('text-base font-bold text-slate-100')
            ui.label('Core CP + FFCS \u00b7 attendance & contest manager').classes('text-xs text-slate-400')

    with ui.left_drawer(value=True).props('breakpoint=1024 bordered width=248').classes('bg-slate-800 border-r border-slate-700 side-nav') as drawer:
        left_drawer_ref['drawer'] = drawer

        with ui.tabs().props('vertical no-caps dense').classes('w-full') as tabs:
            tab_core = ui.tab('Core CP Roster', icon='groups')
            tab_ffcs = ui.tab('FFCS Roster', icon='school')
            tab_all = ui.tab('All Members', icon='filter_list')
            tab_meeting_grid = ui.tab('Meeting Sheet', icon='table_view')
            tab_mark_attendance = ui.tab('Mark Attendance', icon='fact_check')
            tab_scrape = ui.tab('Scrape Contest', icon='travel_explore')
            tab_add_import = ui.tab('Add / Import', icon='person_add')
            tab_review = ui.tab('Review', icon='rate_review')

    with ui.column().classes('page-container w-full max-w-6xl mx-auto p-6 gap-6'):

        # Compact stat bar (not heavy cards -- a single slim strip with dividers)
        total_members = len(df_members)
        core_count = len(df_core)
        ffcs_count = len(df_ffcs)
        contest_cols = [c for c in df_members.columns if 'Starters' in str(c)]
        meeting_cols = [c for c in df_meetings.columns if 'Meeting' in str(c)]
        active_solvers = 0
        if contest_cols:
            solved = df_members[contest_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)
            active_solvers = int((solved > 0).sum())

        stats = [
            ('groups', 'Total Members', str(total_members),
             f'{core_count} Core \u00b7 {ffcs_count} FFCS', 'kpi-blue', 'blue-4'),
            ('emoji_events', 'Starters Tracked', str(len(contest_cols)),
             'CodeChef rounds', 'kpi-cyan', 'cyan-4'),
            ('event_available', 'Meetings Held', str(len(meeting_cols)),
             'Recorded sessions', 'kpi-purple', 'purple-4'),
            ('bolt', 'Active Solvers', str(active_solvers),
             'Solved \u2265 1 problem', 'kpi-amber', 'amber-5'),
        ]
        with ui.grid(columns=4).classes('kpi-grid-responsive w-full gap-4'):
            for icon, label, value, sub, tint_cls, icon_color in stats:
                with ui.element('div').classes(f'kpi-card {tint_cls}'):
                    with ui.row().classes('items-center justify-between w-full no-wrap'):
                        ui.label(label).classes('kpi-label')
                        ui.icon(icon, size='17px').props(f'color={icon_color}')
                    ui.label(value).classes('kpi-value')
                    if sub:
                        ui.label(sub).classes('kpi-sub')

        def make_grid(df: pd.DataFrame, empty_msg: str, card_fields: list[str] | None = None):
            """Renders an ag-grid for desktop and a stacked card list for
            mobile from the same dataframe, toggled purely by CSS so wide
            rosters don't force horizontal scrolling on small screens.
            Nothing is dropped on mobile: primary fields show up front,
            everything else is one tap away behind "More details" so the
            mobile view stays convenient without losing data."""
            if df.empty:
                ui.label(empty_msg).classes('text-rose-400 text-sm')
                return

            cols = [{'headerName': col, 'field': col, 'sortable': True, 'filter': True} for col in df.columns]
            with ui.element('div').classes('desktop-table w-full'):
                ui.aggrid({
                    'columnDefs': cols,
                    'rowData': df.to_dict('records'),
                    'defaultColDef': {'flex': 1, 'minWidth': 130},
                    'pagination': True,
                    'paginationPageSize': 10,
                    'enableCellTextSelection': True,
                    'ensureDomOrder': True,
                }).classes('ag-theme-balham-dark h-96 w-full rounded-xl')

            preferred = card_fields or [
                'Name', 'Registration Number', 'Member Type', 'Batch',
                'Attendance %', 'Contest Attendance %', 'Attendance Status',
                'Total Problems Solved', 'Username',
            ]
            title_field = 'Name' if 'Name' in df.columns else df.columns[0]
            primary_fields = [f for f in preferred if f in df.columns and f != title_field][:6]
            remaining_fields = [c for c in df.columns if c != title_field and c not in primary_fields]

            def esc(v) -> str:
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    return ''
                return html.escape(str(v))

            with ui.column().classes('mobile-cards w-full gap-2'):
                for _, row in df.iterrows():
                    with ui.element('div').classes('mobile-card'):
                        ui.html(f'<div class="mc-title">{esc(row[title_field])}</div>')
                        for f in primary_fields:
                            ui.html(f'<div class="mc-row"><span class="mc-label">{esc(f)}</span><span class="mc-value mono">{esc(row[f])}</span></div>')
                        if remaining_fields:
                            rows_html = ''.join(
                                f'<div class="mc-row"><span class="mc-label">{esc(f)}</span><span class="mc-value mono">{esc(row[f])}</span></div>'
                                for f in remaining_fields
                            )
                            ui.html(f'<details class="mc-more"><summary>More details</summary>{rows_html}</details>')

        with ui.tab_panels(tabs, value=tab_core).classes('w-full bg-transparent mt-4'):


            # TAB 1: Core CP Roster
            with ui.tab_panel(tab_core):
                make_grid(df_core, 'No data found in CP_Members.xlsx. Run ContestAttendance.py (mode 1) first.')

            # TAB 2: FFCS Roster
            with ui.tab_panel(tab_ffcs):
                make_grid(df_ffcs, 'No data found in FFCS_Members.xlsx. Run ContestAttendance.py (mode 2) first.')

            # TAB 3: Combined Roster (Batch + attendance-range + type + search filters)
            with ui.tab_panel(tab_all):
                batch_options = ['All'] + mo.known_batches()

                with ui.card().classes('bg-slate-800 border border-slate-700 p-6 rounded-xl w-full gap-3'):
                    ui.label('Filters').classes('text-lg font-bold text-slate-200')
                    with ui.row().classes('gap-4 items-end flex-wrap'):
                        f_batch = ui.select(batch_options, value='All', label='Batch').props('dense outlined').classes('w-44')
                        f_type = ui.select(['All', 'Core', 'FFCS'], value='All', label='Member Type').props('dense outlined').classes('w-36')
                        f_att_min = ui.number('Min meeting attendance %', value=0, min=0, max=100).props('dense outlined').classes('w-40')
                        f_att_max = ui.number('Max meeting attendance %', value=100, min=0, max=100).props('dense outlined').classes('w-40')
                        f_catt_min = ui.number('Min contest attendance %', value=0, min=0, max=100).props('dense outlined').classes('w-40')
                        f_catt_max = ui.number('Max contest attendance %', value=100, min=0, max=100).props('dense outlined').classes('w-40')
                        f_search = ui.input('Search name / reg. no. / handle').props('dense outlined clearable').classes('w-64')
                    with ui.row().classes('gap-3 items-center'):
                        result_label = ui.label('').classes('text-xs text-slate-400')

                        def do_reset():
                            f_batch.value, f_type.value = 'All', 'All'
                            f_att_min.value, f_att_max.value = 0, 100
                            f_catt_min.value, f_catt_max.value = 0, 100
                            f_search.value = ''
                            refresh_all_grid()

                        ui.button('Apply Filters', icon='search', on_click=lambda: refresh_all_grid()).props('dense').classes(
                            'bg-blue-600 hover:bg-blue-500 text-white font-semibold px-4 rounded-lg'
                        )
                        ui.button('Reset', icon='refresh', on_click=do_reset).props('dense flat').classes('text-slate-300')

                all_grid_container = ui.column().classes('w-full mt-3')

                def refresh_all_grid():
                    all_grid_container.clear()
                    filtered = mo.filter_members(
                        df_members,
                        batch=f_batch.value,
                        member_type=f_type.value,
                        attendance_min=f_att_min.value,
                        attendance_max=f_att_max.value,
                        contest_attendance_min=f_catt_min.value,
                        contest_attendance_max=f_catt_max.value,
                        search=f_search.value,
                    )
                    result_label.set_text(f'{len(filtered)} of {len(df_members)} member(s) shown')
                    with all_grid_container:
                        make_grid(filtered, 'No members match the current filters — try Reset.')

                # Search box applies as you type (debounced); everything else needs Apply,
                # so dragging/typing numbers doesn't refetch the grid on every keystroke.
                f_search.on('keydown.enter', lambda e: refresh_all_grid())

                refresh_all_grid()

            # TAB 4: Meeting Excel Viewer
            with ui.tab_panel(tab_meeting_grid):
                if not df_meetings.empty:
                    make_grid(df_meetings, '')
                else:
                    ui.label('No meeting data created yet. Record a session to generate Meeting_Attendance.xlsx').classes('text-amber-400')

            # TAB 5: Mark Meeting Attendance
            with ui.tab_panel(tab_mark_attendance):
                with ui.card().classes('bg-slate-800 border border-slate-700 p-6 rounded-xl w-full gap-4'):
                    ui.label('Record Meeting Session').classes('text-lg font-bold text-slate-200')

                    meeting_date = ui.input('Meeting Date', value=str(date.today())).props('dense outlined').classes('w-64')
                    cohort_filter = ui.select(
                        {'All': 'All Members', 'Core': 'Core CP Only', 'FFCS': 'FFCS Only'},
                        value='All', label='Show'
                    ).props('dense outlined').classes('w-64')
                    selected_present = set()

                    ui.label('Select Present Members:').classes('font-semibold text-slate-300 mt-2')

                    checkbox_container = ui.column().classes('w-full')

                    def render_checkboxes():
                        checkbox_container.clear()
                        cohort = cohort_filter.value
                        rows = df_members if cohort == 'All' else df_members[df_members['Member Type'] == cohort]
                        with checkbox_container:
                            with ui.scroll_area().classes('h-72 border border-slate-700 rounded-lg px-2 py-1 bg-slate-900/50'):
                                for idx, row in rows.iterrows():
                                    name = str(row.get('Name', 'Unknown'))
                                    reg_no = str(row.get('Registration Number', 'N/A'))
                                    handle = str(row.get('Username', 'N/A'))
                                    m_type = str(row.get('Member Type', 'N/A'))

                                    def on_change(e, user=handle):
                                        if e.value:
                                            selected_present.add(user)
                                        else:
                                            selected_present.discard(user)

                                    with ui.row().classes('roster-row w-full items-center gap-3 no-wrap'):
                                        ui.checkbox(
                                            value=handle in selected_present,
                                            on_change=on_change,
                                        ).props('dense')
                                        with ui.column().classes('gap-0 min-w-0'):
                                            ui.label(name).classes('text-sm text-slate-200 leading-tight')
                                            ui.label(f'{m_type} \u00b7 {reg_no} \u00b7 @{handle}').classes('roster-meta leading-tight')

                    render_checkboxes()
                    cohort_filter.on('update:model-value', lambda e: render_checkboxes())

                    def save_meeting():
                        col_title = f"Meeting {meeting_date.value}"
                        m_df = load_meeting_data()

                        if m_df.empty:
                            ui.notify('Please run ContestAttendance.py first to build member data!', type='negative')
                            return

                        # Set 1 for present, 0 for absent
                        m_df[col_title] = m_df['Username'].apply(lambda u: 1 if str(u).strip() in selected_present else 0)
                        save_meeting_data(m_df)

                        ui.notify(f"Successfully saved {col_title} into {MEETING_EXCEL_FILE}!", type='positive')
                        ui.navigate.to('/')

                    ui.button('Save Attendance to Excel', on_click=save_meeting).classes('bg-blue-600 hover:bg-blue-500 text-white font-bold py-2 px-6 rounded-lg self-start mt-2')

            # TAB 6: Scrape a Contest (runs ContestAttendance.py from the dashboard)
            with ui.tab_panel(tab_scrape):
                with ui.card().classes('bg-slate-800 border border-slate-700 p-6 rounded-xl w-full gap-4'):
                    ui.label('Scrape CodeChef Attendance').classes('text-lg font-bold text-slate-200')
                    ui.label(
                        'Fetches each member\'s CodeChef profile one at a time (sequential, '
                        'rate-limited) to avoid HTTP 429s. This can take a while for large rosters.'
                    ).classes('text-xs text-slate-400')

                    with ui.row().classes('gap-4 items-end flex-wrap'):
                        scrape_cohort = ui.select(
                            {'Core': 'Core CP Members', 'FFCS': 'FFCS Members'},
                            value='Core', label='Cohort',
                        ).props('dense outlined').classes('w-56')
                        scrape_contest = ui.number(
                            'Starters Contest Number', value=None, format='%d', min=1,
                        ).props('dense outlined').classes('w-56')
                        scrape_delay = ui.number(
                            'Delay Between Requests (sec)', value=1.5, min=0.5, step=0.5,
                        ).props('dense outlined').classes('w-56')

                    with ui.element('div').classes('scrape-console w-full mt-2'):
                        with ui.row().classes('items-center justify-between w-full no-wrap mb-1'):
                            with ui.row().classes('items-center gap-2'):
                                scrape_status_dot = ui.element('div').classes('status-dot status-idle')
                                progress_label = ui.label('Idle.').classes('text-xs text-slate-300 mono')
                            ui.label('scrape console').classes('text-[10px] text-slate-500 uppercase tracking-widest')
                        progress_bar = ui.linear_progress(value=0).classes('w-full scrape-progress').props('instant-feedback size=4px')
                        log_area = ui.log(max_lines=500).classes('w-full h-64 scrape-log text-xs mt-2')

                    scrape_state = {'running': False}

                    def run_scrape():
                        if scrape_state['running']:
                            ui.notify('A scrape is already running.', type='warning')
                            return
                        if not scrape_contest.value:
                            ui.notify('Enter a Starters contest number first.', type='negative')
                            return

                        cohort = scrape_cohort.value
                        excel_file = COHORT_FILES[cohort]
                        starter_num = int(scrape_contest.value)
                        delay = float(scrape_delay.value or 1.5)

                        scrape_state['running'] = True
                        log_area.clear()
                        progress_bar.set_value(0)
                        scrape_status_dot.classes(replace='status-dot status-running')
                        progress_label.set_text(f'Starting scrape for {cohort} — Starters {starter_num}...')

                        progress_queue: queue.Queue = queue.Queue()

                        def worker():
                            def on_progress(completed, total, handle, count):
                                progress_queue.put(('progress', completed, total, handle, count))
                            try:
                                ca.process_attendance(
                                    excel_file, starter_num, cohort,
                                    inter_request_delay=delay,
                                    on_progress=on_progress,
                                )
                                progress_queue.put(('done', None, None, None, None))
                            except Exception as e:  # noqa: BLE001
                                progress_queue.put(('error', str(e), None, None, None))

                        threading.Thread(target=worker, daemon=True).start()

                        def poll():
                            drained_any = False
                            while True:
                                try:
                                    kind, a, b, c, d = progress_queue.get_nowait()
                                except queue.Empty:
                                    break
                                drained_any = True
                                if kind == 'progress':
                                    completed, total, handle, count = a, b, c, d
                                    progress_bar.set_value(completed / total if total else 0)
                                    if count is None:
                                        progress_label.set_text(f'{completed}/{total} — @{handle}: unavailable, unchanged')
                                        log_area.push(f'[{completed}/{total}] ⚠️ @{handle}: unavailable, unchanged')
                                    else:
                                        progress_label.set_text(f'{completed}/{total} — @{handle}: {count} solved')
                                        icon = '✅' if count > 0 else '➖'
                                        log_area.push(f'[{completed}/{total}] {icon} @{handle}: {count} problem(s) solved')
                                elif kind == 'done':
                                    scrape_status_dot.classes(replace='status-dot status-done')
                                    progress_label.set_text(f'✅ Done! Updated {excel_file}.')
                                    log_area.push(f'🎉 Finished scraping Starters {starter_num} for {cohort}.')
                                    ui.notify(f'Scrape complete for {cohort} — Starters {starter_num}.', type='positive')
                                    scrape_state['running'] = False
                                    poll_timer.deactivate()
                                elif kind == 'error':
                                    scrape_status_dot.classes(replace='status-dot status-error')
                                    progress_label.set_text('❌ Scrape failed.')
                                    log_area.push(f'❌ Error: {a}')
                                    ui.notify(f'Scrape failed: {a}', type='negative')
                                    scrape_state['running'] = False
                                    poll_timer.deactivate()
                            return drained_any

                        poll_timer = ui.timer(0.5, poll)

                    ui.button('Start Scrape', on_click=run_scrape).classes(
                        'bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-2 px-6 rounded-lg self-start mt-2'
                    )
                    ui.label(
                        'Tip: after a scrape finishes, refresh this page to see updated stats in the KPI '
                        'cards, rosters, and to regenerate the static dashboard with export_static.py.'
                    ).classes('text-xs text-slate-500 mt-2')

            # TAB 7: Add Member (single) + Bulk CSV/Excel Import
            with ui.tab_panel(tab_add_import):

                # --- Single Add ---
                with ui.card().classes('bg-slate-800 border border-slate-700 p-6 rounded-xl w-full gap-3'):
                    with ui.row().classes('items-center gap-2'):
                        ui.icon('person_add').classes('text-blue-400 text-2xl')
                        ui.label('Add One Member').classes('text-lg font-bold text-slate-200')
                    ui.label('For a single new recruit. Use Bulk Import below for a whole batch.').classes('text-xs text-slate-400 -mt-2')

                    with ui.grid(columns=3).classes('kpi-grid-responsive gap-4 w-full max-w-3xl mt-2'):
                        add_type = ui.select(['Core', 'FFCS'], value='Core', label='Member Type').props('dense outlined')
                        add_name = ui.input('Name *').props('dense outlined')
                        add_reg = ui.input('Registration Number *').props('dense outlined')
                        add_phone = ui.input('Phone number').props('dense outlined')
                        add_handle = ui.input('Username (CodeChef ID)').props('dense outlined')
                        add_batch = ui.select(
                            options=mo.known_batches(),
                            value=(mo.known_batches() or [ca.DEFAULT_BATCH_LABEL])[0],
                            label='Batch', new_value_mode='add-unique', with_input=True,
                        ).props('dense outlined').tooltip('Pick an existing batch or type a new one and press Enter')

                    add_status = ui.label('').classes('text-xs')

                    def do_add_member():
                        ok, msg = mo.add_single_member(
                            add_type.value, add_name.value or '', add_reg.value or '',
                            add_phone.value or '', add_handle.value or '', add_batch.value or '',
                        )
                        add_status.set_text(msg)
                        add_status.classes(replace='text-xs ' + ('text-emerald-400' if ok else 'text-rose-400'))
                        ui.notify(msg, type='positive' if ok else 'negative')
                        if ok:
                            ui.navigate.to('/')

                    ui.button('Add Member', icon='add', on_click=do_add_member).classes(
                        'bg-blue-600 hover:bg-blue-500 text-white font-bold py-2 px-6 rounded-lg self-start mt-2'
                    )

                # --- Bulk Import wizard ---
                with ui.card().classes('bg-slate-800 border border-slate-700 p-6 rounded-xl w-full gap-3 mt-4'):
                    with ui.row().classes('items-center gap-2'):
                        ui.icon('upload_file').classes('text-emerald-400 text-2xl')
                        ui.label('Bulk Import (CSV / Excel) — for recruitment drives').classes('text-lg font-bold text-slate-200')
                    ui.label(
                        'Follow the 3 steps below. Everything for a step is grouped together so you '
                        'always know what to do next.'
                    ).classes('text-xs text-slate-400 -mt-2 mb-2')

                    import_state = {'raw_df': None, 'preview_df': None}

                    with ui.stepper().props('vertical flat').classes('w-full bg-transparent') as stepper:

                        # Step 1: cohort, batch, upload
                        with ui.step('Choose cohort & upload file'):
                            with ui.row().classes('gap-4 items-end flex-wrap'):
                                import_cohort = ui.select(['Core', 'FFCS'], value='Core', label='Import into').props('dense outlined').classes('w-40')
                                import_batch = ui.select(
                                    options=mo.known_batches(),
                                    value=(mo.known_batches() or [ca.DEFAULT_BATCH_LABEL])[0],
                                    label='Batch for this import', new_value_mode='add-unique', with_input=True,
                                ).props('dense outlined').classes('w-56').tooltip('Pick an existing batch or type a new one and press Enter')
                            upload_status = ui.label('No file uploaded yet.').classes('text-xs text-slate-400 mt-2')

                            def handle_upload(e):
                                content = e.content.read()
                                try:
                                    raw_df = mo.parse_uploaded_table(e.name, content)
                                except Exception as ex:  # noqa: BLE001
                                    ui.notify(f'Could not read file: {ex}', type='negative')
                                    upload_status.set_text(f'❌ Failed to read {e.name}.')
                                    return
                                import_state['raw_df'] = raw_df
                                import_state['preview_df'] = None
                                upload_status.set_text(f'✅ Loaded {len(raw_df)} row(s) from {e.name}.')
                                render_mapping()
                                stepper.next()

                            ui.upload(on_upload=handle_upload, auto_upload=True).props('accept=".csv,.xlsx,.xls"').classes('w-full mt-2')
                            with ui.stepper_navigation():
                                ui.button('Next', icon='arrow_forward', on_click=stepper.next).props('dense')

                        # Step 2: mapping
                        with ui.step('Confirm column mapping'):
                            ui.label(
                                'We auto-detected which of your file\'s columns map to each field. '
                                'Fix any that show "(none)" before continuing.'
                            ).classes('text-xs text-slate-400 mb-1')
                            mapping_container = ui.column().classes('w-full gap-2')

                            def render_mapping():
                                mapping_container.clear()
                                raw_df = import_state['raw_df']
                                if raw_df is None:
                                    return
                                guess = mo.guess_column_mapping(raw_df)
                                options = ['(none)'] + [str(c) for c in raw_df.columns]
                                mapping_selects.clear()
                                with mapping_container:
                                    with ui.grid(columns=2).classes('mapping-grid-responsive gap-4 w-full max-w-2xl'):
                                        for field in mo.IMPORT_TARGET_FIELDS:
                                            default = guess.get(field) or '(none)'
                                            sel = ui.select(options, value=default, label=f'{field} ←').props('dense outlined')
                                            mapping_selects[field] = sel
                                    if any(guess.get(f) is None for f in mo.IMPORT_TARGET_FIELDS):
                                        ui.label(
                                            '⚠️ Some columns could not be auto-detected — please map them above.'
                                        ).classes('text-amber-400 text-xs mt-1')

                            mapping_selects = {}
                            render_mapping()

                            def do_preview():
                                raw_df = import_state['raw_df']
                                if raw_df is None:
                                    ui.notify('Upload a file first.', type='negative')
                                    return
                                mapping = {f: (None if s.value == '(none)' else s.value) for f, s in mapping_selects.items()}
                                mapped = mo.apply_mapping(raw_df, mapping, import_cohort.value, import_batch.value or ca.DEFAULT_BATCH_LABEL)
                                mapped = mo.detect_duplicates(mapped, import_cohort.value)
                                import_state['preview_df'] = mapped
                                render_preview()
                                stepper.next()

                            with ui.stepper_navigation():
                                ui.button('Preview →', icon='arrow_forward', on_click=do_preview).classes(
                                    'bg-slate-600 hover:bg-slate-500 text-white font-semibold py-2 px-6 rounded-lg'
                                )
                                ui.button('Back', on_click=stepper.previous).props('flat')

                        # Step 3: preview + confirm
                        with ui.step('Preview & confirm'):
                            preview_container = ui.column().classes('w-full gap-2')

                            def render_preview():
                                preview_container.clear()
                                mapped = import_state['preview_df']
                                if mapped is None:
                                    return
                                dup_count = int(mapped['Duplicate'].sum())
                                with preview_container:
                                    with ui.row().classes('gap-2 items-center'):
                                        ui.icon('info', color='amber' if dup_count else 'slate').classes('text-lg')
                                        ui.label(
                                            f'{len(mapped)} row(s) parsed — {dup_count} flagged as possible duplicate(s) '
                                            f'(matched by Registration Number) and will be skipped on import.'
                                        ).classes('text-sm text-slate-300')
                                    cols = [{'headerName': c, 'field': c, 'sortable': True, 'filter': True} for c in mapped.columns]
                                    ui.aggrid({
                                        'columnDefs': cols,
                                        'rowData': mapped.to_dict('records'),
                                        'defaultColDef': {'flex': 1, 'minWidth': 120},
                                        'pagination': True,
                                        'paginationPageSize': 10,
                                        'enableCellTextSelection': True,
                                        'ensureDomOrder': True,
                                    }).classes('ag-theme-balham-dark h-80 w-full rounded-xl')

                            render_preview()

                            def do_confirm():
                                mapped = import_state['preview_df']
                                if mapped is None:
                                    ui.notify('Nothing to import yet.', type='negative')
                                    return
                                imported, skipped = mo.commit_bulk_import(mapped, import_cohort.value, skip_duplicates=True)
                                ui.notify(f'Imported {imported} member(s), skipped {skipped} duplicate(s).', type='positive')
                                ui.navigate.to('/')

                            with ui.stepper_navigation():
                                ui.button('Confirm Import', icon='check', on_click=do_confirm).classes(
                                    'bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-2 px-6 rounded-lg'
                                )
                                ui.button('Back', on_click=stepper.previous).props('flat')

            # TAB 8: Review Workflow (0%-30% inclusive attendance)
            with ui.tab_panel(tab_review):
                eligible = mo.review_eligible_members()

                with ui.card().classes('bg-slate-800 border border-slate-700 p-6 rounded-xl w-full gap-3'):
                    ui.label(f'Review-Eligible Members ({mo.ATTENDANCE_LOW}%–{mo.ATTENDANCE_HIGH}% attendance, inclusive)').classes(
                        'text-lg font-bold text-slate-200'
                    )
                    ui.label(
                        'Flagged if EITHER meeting attendance or contest (Starters round) participation '
                        'falls in this band — enhanced from the original "0 contests" check into a percentage.'
                    ).classes('text-xs text-slate-400 -mt-2')
                    if eligible.empty:
                        ui.label(
                            'No members currently fall in the 0%-30% attendance band (or no attendance/contest data recorded yet).'
                        ).classes('text-slate-400 text-sm')
                    else:
                        eligible = eligible.sort_values('Attendance %', na_position='last')
                        for _, row in eligible.iterrows():
                            m_type = str(row.get('Member Type', 'N/A'))
                            username = str(row.get('Username', 'N/A'))
                            name = str(row.get('Name', 'Unknown'))
                            reg_no = str(row.get('Registration Number', 'N/A'))
                            batch = str(row.get('Batch', 'N/A'))
                            att = row.get('Attendance %')
                            catt = row.get('Contest Attendance %')
                            reason = str(row.get('Review Reason', ''))
                            att_display = f"{att:.0f}%" if att is not None and not pd.isna(att) else 'N/A'
                            catt_display = f"{catt:.0f}%" if catt is not None and not pd.isna(catt) else 'N/A'
                            worst = min([v for v in (att, catt) if v is not None and not pd.isna(v)], default=None)
                            is_critical = worst is not None and worst <= 15
                            badge_color = 'red' if is_critical else 'amber'
                            sev_cls = 'sev-red' if is_critical else 'sev-amber'

                            with ui.row().classes(
                                f'review-row {sev_cls} w-full items-center justify-between px-4 py-3 gap-3 flex-wrap'
                            ):
                                with ui.row().classes('items-center gap-3 min-w-0'):
                                    ui.badge(reason or '—', color=badge_color).classes('text-xs px-3 py-1')
                                    with ui.column().classes('gap-0 min-w-0'):
                                        ui.label(f'{name}  ·  {m_type}').classes('text-slate-100 font-semibold')
                                        ui.label(
                                            f'Reg. {reg_no}  ·  Batch: {batch}  ·  @{username}'
                                        ).classes('roster-meta')
                                        with ui.row().classes('items-center gap-3 mt-1'):
                                            ui.label(f'Meetings {att_display}').classes(
                                                'text-[11px] text-slate-400 mono'
                                            )
                                            ui.label(f'Contests {catt_display}').classes(
                                                'text-[11px] text-slate-400 mono'
                                            )

                                def open_review_dialog(m_type=m_type, username=username, name=name, reg_no=reg_no):
                                    with ui.dialog() as dialog, ui.card().classes('bg-slate-800 border border-slate-700 p-6 gap-3 w-96'):
                                        ui.label(f'Review: {name}').classes('text-lg font-bold text-slate-200')
                                        r_task = ui.input('Task').props('dense outlined').classes('w-full')
                                        r_desc = ui.textarea('Task Description').props('dense outlined').classes('w-full')
                                        r_status = ui.select(mo.REVIEW_STATUSES, value='Open', label='Status').props('dense outlined').classes('w-full')
                                        r_notes = ui.textarea('Notes').props('dense outlined').classes('w-full')
                                        r_outcome = ui.input('Outcome').props('dense outlined').classes('w-full')

                                        def save_review():
                                            mo.add_review(
                                                m_type, username, name, reg_no,
                                                r_task.value or '', r_desc.value or '',
                                                r_status.value or 'Open', r_notes.value or '', r_outcome.value or '',
                                            )
                                            ui.notify('Review saved.', type='positive')
                                            dialog.close()
                                            if r_status.value in ('Failed', 'Incomplete'):
                                                offer_removal(m_type, username, name)
                                            else:
                                                ui.navigate.to('/')

                                        with ui.row().classes('w-full justify-between mt-2'):
                                            ui.button('Save Review', on_click=save_review).classes(
                                                'bg-blue-600 hover:bg-blue-500 text-white font-semibold py-2 px-4 rounded-lg'
                                            )
                                            ui.button('Cancel', on_click=dialog.close).classes(
                                                'bg-slate-600 hover:bg-slate-500 text-white font-semibold py-2 px-4 rounded-lg'
                                            )
                                    dialog.open()

                                ui.button('Review', icon='rate_review', on_click=open_review_dialog).props('dense').classes(
                                    'bg-purple-600 hover:bg-purple-500 text-white text-xs font-semibold px-4 rounded-lg'
                                )

                def offer_removal(m_type, username, name):
                    with ui.dialog() as rdialog, ui.card().classes('bg-slate-800 border border-rose-700 p-6 gap-3 w-96'):
                        ui.label(f'Task failed/incomplete for {name}.').classes('text-slate-200 font-semibold')
                        ui.label('Remove this member from the roster? This is PERMANENT — there is no archive/trash.').classes(
                            'text-rose-400 text-sm'
                        )
                        with ui.row().classes('w-full justify-between mt-2'):
                            def confirm_delete():
                                ok, msg = mo.delete_member(m_type, username)
                                ui.notify(msg, type='positive' if ok else 'negative')
                                rdialog.close()
                                ui.navigate.to('/')
                            ui.button('Yes, Permanently Remove', on_click=confirm_delete).classes(
                                'bg-rose-600 hover:bg-rose-500 text-white font-semibold py-2 px-4 rounded-lg'
                            )
                            ui.button('Keep Member', on_click=rdialog.close).classes(
                                'bg-slate-600 hover:bg-slate-500 text-white font-semibold py-2 px-4 rounded-lg'
                            )
                    rdialog.open()

                with ui.card().classes('bg-slate-800 border border-slate-700 p-6 rounded-xl w-full gap-3 mt-4'):
                    ui.label('Review History').classes('text-lg font-bold text-slate-200')
                    reviews = mo.load_reviews()
                    if reviews.empty:
                        ui.label('No reviews recorded yet.').classes('text-slate-400 text-sm')
                    else:
                        cols = [{'headerName': c, 'field': c, 'sortable': True, 'filter': True} for c in reviews.columns]
                        ui.aggrid({
                            'columnDefs': cols,
                            'rowData': reviews.to_dict('records'),
                            'defaultColDef': {'flex': 1, 'minWidth': 120},
                            'pagination': True,
                            'paginationPageSize': 10,
                            'enableCellTextSelection': True,
                            'ensureDomOrder': True,
                        }).classes('ag-theme-balham-dark h-80 w-full rounded-xl')

ui.run(title='Department Attendance Hub', dark=False, port=8080)
