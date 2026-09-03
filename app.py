from nicegui import ui
import pandas as pd
import os
import threading
import queue
from datetime import date

import ContestAttendance as ca

CORE_EXCEL_FILE = "CP_Members.xlsx"
FFCS_EXCEL_FILE = "FFCS_Members.xlsx"
MEETING_EXCEL_FILE = "Meeting_Attendance.xlsx"

COHORT_FILES = {"Core": CORE_EXCEL_FILE, "FFCS": FFCS_EXCEL_FILE}

BASE_COLS = ["Name", "Register number", "Phone number", "Username"]

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

    ui.colors(primary='#3b82f6', secondary='#64748b', dark='#0f172a')
    ui.query('body').classes('bg-slate-900 text-slate-100')

    # Top Header
    with ui.header().classes('bg-slate-800 border-b border-slate-700 justify-between items-center px-6 py-4'):
        ui.label('⚡ Technical Department Hub').classes('text-xl font-bold text-blue-400')
        ui.label('Core CP + FFCS Attendance & Contest Manager').classes('text-sm text-slate-400')

    with ui.column().classes('w-full max-w-6xl mx-auto p-6 gap-6'):

        # Top Metrics
        with ui.row().classes('w-full gap-4'):
            total_members = len(df_members)
            core_count = len(df_core)
            ffcs_count = len(df_ffcs)
            contest_cols = [c for c in df_members.columns if 'Starters' in str(c)]
            meeting_cols = [c for c in df_meetings.columns if 'Meeting' in str(c)]

            with ui.card().classes('flex-1 bg-slate-800 border border-slate-700 p-4 rounded-xl'):
                ui.label('Total Members').classes('text-xs text-slate-400 uppercase font-semibold')
                ui.label(str(total_members)).classes('text-3xl font-extrabold text-blue-400')
                ui.label(f'{core_count} Core · {ffcs_count} FFCS').classes('text-xs text-slate-500')

            with ui.card().classes('flex-1 bg-slate-800 border border-slate-700 p-4 rounded-xl'):
                ui.label('Starters Tracked').classes('text-xs text-slate-400 uppercase font-semibold')
                ui.label(str(len(contest_cols))).classes('text-3xl font-extrabold text-emerald-400')

            with ui.card().classes('flex-1 bg-slate-800 border border-slate-700 p-4 rounded-xl'):
                ui.label('Meetings Held').classes('text-xs text-slate-400 uppercase font-semibold')
                ui.label(str(len(meeting_cols))).classes('text-3xl font-extrabold text-purple-400')

            with ui.card().classes('flex-1 bg-slate-800 border border-slate-700 p-4 rounded-xl'):
                active_solvers = 0
                if contest_cols:
                    solved = df_members[contest_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)
                    active_solvers = int((solved > 0).sum())
                ui.label('Active Solvers').classes('text-xs text-slate-400 uppercase font-semibold')
                ui.label(str(active_solvers)).classes('text-3xl font-extrabold text-amber-400')

        # Navigation Tabs
        with ui.tabs().classes('w-full border-b border-slate-700 text-slate-300') as tabs:
            tab_core = ui.tab('Core CP Roster')
            tab_ffcs = ui.tab('FFCS Roster')
            tab_all = ui.tab('All Members')
            tab_meeting_grid = ui.tab('Meeting Sheet')
            tab_mark_attendance = ui.tab('Mark Meeting Attendance')
            tab_scrape = ui.tab('Scrape Contest')

        def make_grid(df: pd.DataFrame, empty_msg: str):
            if not df.empty:
                cols = [{'headerName': col, 'field': col, 'sortable': True, 'filter': True} for col in df.columns]
                ui.aggrid({
                    'columnDefs': cols,
                    'rowData': df.to_dict('records'),
                    'defaultColDef': {'flex': 1, 'minWidth': 130},
                    'pagination': True,
                    'paginationPageSize': 10,
                }).classes('ag-theme-balham-dark h-96 w-full rounded-xl')
            else:
                ui.label(empty_msg).classes('text-rose-400')

        with ui.tab_panels(tabs, value=tab_core).classes('w-full bg-transparent mt-4'):

            # TAB 1: Core CP Roster
            with ui.tab_panel(tab_core):
                make_grid(df_core, 'No data found in CP_Members.xlsx. Run ContestAttendance.py (mode 1) first.')

            # TAB 2: FFCS Roster
            with ui.tab_panel(tab_ffcs):
                make_grid(df_ffcs, 'No data found in FFCS_Members.xlsx. Run ContestAttendance.py (mode 2) first.')

            # TAB 3: Combined Roster
            with ui.tab_panel(tab_all):
                make_grid(df_members, 'No member data found yet.')

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

                    meeting_date = ui.input('Meeting Date', value=str(date.today())).classes('w-64')
                    cohort_filter = ui.select(
                        {'All': 'All Members', 'Core': 'Core CP Only', 'FFCS': 'FFCS Only'},
                        value='All', label='Show'
                    ).classes('w-64')
                    selected_present = set()

                    ui.label('Select Present Members:').classes('font-semibold text-slate-300 mt-2')

                    checkbox_container = ui.column().classes('w-full')

                    def render_checkboxes():
                        checkbox_container.clear()
                        cohort = cohort_filter.value
                        rows = df_members if cohort == 'All' else df_members[df_members['Member Type'] == cohort]
                        with checkbox_container:
                            with ui.scroll_area().classes('h-64 border border-slate-700 rounded-lg p-4 bg-slate-900/50'):
                                for idx, row in rows.iterrows():
                                    name = str(row.get('Name', 'Unknown'))
                                    reg_no = str(row.get('Register number', 'N/A'))
                                    handle = str(row.get('Username', 'N/A'))
                                    m_type = str(row.get('Member Type', 'N/A'))

                                    display_label = f"[{m_type}] {name} | {reg_no} (@{handle})"

                                    def on_change(e, user=handle):
                                        if e.value:
                                            selected_present.add(user)
                                        else:
                                            selected_present.discard(user)

                                    ui.checkbox(
                                        display_label,
                                        value=handle in selected_present,
                                        on_change=on_change,
                                    ).classes('text-slate-200 py-1 font-mono text-sm')

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

                    with ui.row().classes('gap-4 items-end'):
                        scrape_cohort = ui.select(
                            {'Core': 'Core CP Members', 'FFCS': 'FFCS Members'},
                            value='Core', label='Cohort',
                        ).classes('w-56')
                        scrape_contest = ui.number(
                            'Starters Contest Number', value=None, format='%d', min=1,
                        ).classes('w-56')
                        scrape_delay = ui.number(
                            'Delay Between Requests (sec)', value=1.5, min=0.5, step=0.5,
                        ).classes('w-56')

                    progress_bar = ui.linear_progress(value=0).classes('w-full').props('instant-feedback')
                    progress_label = ui.label('Idle.').classes('text-xs text-slate-400')
                    log_area = ui.log(max_lines=500).classes('w-full h-64 bg-slate-900/60 border border-slate-700 rounded-lg text-xs')

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
                                    progress_label.set_text(f'✅ Done! Updated {excel_file}.')
                                    log_area.push(f'🎉 Finished scraping Starters {starter_num} for {cohort}.')
                                    ui.notify(f'Scrape complete for {cohort} — Starters {starter_num}.', type='positive')
                                    scrape_state['running'] = False
                                    poll_timer.deactivate()
                                elif kind == 'error':
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

ui.run(title='Department Attendance Hub', dark=True, port=8080)
