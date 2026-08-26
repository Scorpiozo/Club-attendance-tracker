from nicegui import ui
import pandas as pd
import os
from datetime import date

CP_EXCEL_FILE = "CP_Members.xlsx"
MEETING_EXCEL_FILE = "Meeting_Attendance.xlsx"

# --- Data Loaders & Savers ---

def load_cp_data() -> pd.DataFrame:
    if os.path.exists(CP_EXCEL_FILE):
        return pd.read_excel(CP_EXCEL_FILE)
    return pd.DataFrame(columns=["Name", "Register number", "Phone number", "Username"])

def load_meeting_data() -> pd.DataFrame:
    if os.path.exists(MEETING_EXCEL_FILE):
        return pd.read_excel(MEETING_EXCEL_FILE)
    
    # Sync core member info from main roster if meeting excel doesn't exist yet
    df_cp = load_cp_data()
    if not df_cp.empty:
        base_cols = [c for c in ["Name", "Register number", "Phone number", "Username"] if c in df_cp.columns]
        return df_cp[base_cols].copy()
        
    return pd.DataFrame(columns=["Name", "Register number", "Phone number", "Username"])

def save_meeting_data(df: pd.DataFrame):
    df.to_excel(MEETING_EXCEL_FILE, index=False)

# --- App Layout ---

@ui.page('/')
def main_page():
    df_members = load_cp_data()
    df_meetings = load_meeting_data()
    
    ui.colors(primary='#3b82f6', secondary='#64748b', dark='#0f172a')
    ui.query('body').classes('bg-slate-900 text-slate-100')

    # Top Header
    with ui.header().classes('bg-slate-800 border-b border-slate-700 justify-between items-center px-6 py-4'):
        ui.label('⚡ Technical Department Hub').classes('text-xl font-bold text-blue-400')
        ui.label('Attendance & Contest Manager').classes('text-sm text-slate-400')

    with ui.column().classes('w-full max-w-6xl mx-auto p-6 gap-6'):
        
        # Top Metrics
        with ui.row().classes('w-full gap-4'):
            total_members = len(df_members)
            contest_cols = [c for c in df_members.columns if 'Starters' in str(c)]
            meeting_cols = [c for c in df_meetings.columns if 'Meeting' in str(c)]
            
            with ui.card().classes('flex-1 bg-slate-800 border border-slate-700 p-4 rounded-xl'):
                ui.label('Total Members').classes('text-xs text-slate-400 uppercase font-semibold')
                ui.label(str(total_members)).classes('text-3xl font-extrabold text-blue-400')
                
            with ui.card().classes('flex-1 bg-slate-800 border border-slate-700 p-4 rounded-xl'):
                ui.label('Starters Tracked').classes('text-xs text-slate-400 uppercase font-semibold')
                ui.label(str(len(contest_cols))).classes('text-3xl font-extrabold text-emerald-400')
                
            with ui.card().classes('flex-1 bg-slate-800 border border-slate-700 p-4 rounded-xl'):
                ui.label('Meetings Held').classes('text-xs text-slate-400 uppercase font-semibold')
                ui.label(str(len(meeting_cols))).classes('text-3xl font-extrabold text-purple-400')

        # Navigation Tabs
        with ui.tabs().classes('w-full border-b border-slate-700 text-slate-300') as tabs:
            tab_members = ui.tab('Member Roster')
            tab_meeting_grid = ui.tab('Meeting Sheet')
            tab_mark_attendance = ui.tab('Mark Meeting Attendance')

        with ui.tab_panels(tabs, value=tab_members).classes('w-full bg-transparent mt-4'):
            
            # TAB 1: Main Member Roster
            with ui.tab_panel(tab_members):
                if not df_members.empty:
                    cols = [{'headerName': col, 'field': col, 'sortable': True, 'filter': True} for col in df_members.columns]
                    ui.aggrid({
                        'columnDefs': cols,
                        'rowData': df_members.to_dict('records'),
                        'defaultColDef': {'flex': 1, 'minWidth': 130},
                        'pagination': True,
                        'paginationPageSize': 10,
                    }).classes('ag-theme-balham-dark h-96 w-full rounded-xl')
                else:
                    ui.label('No data found in CP_Members.xlsx').classes('text-rose-400')

            # TAB 2: Meeting Excel Viewer
            with ui.tab_panel(tab_meeting_grid):
                if not df_meetings.empty:
                    cols = [{'headerName': col, 'field': col, 'sortable': True, 'filter': True} for col in df_meetings.columns]
                    ui.aggrid({
                        'columnDefs': cols,
                        'rowData': df_meetings.to_dict('records'),
                        'defaultColDef': {'flex': 1, 'minWidth': 130},
                        'pagination': True,
                        'paginationPageSize': 10,
                    }).classes('ag-theme-balham-dark h-96 w-full rounded-xl')
                else:
                    ui.label('No meeting data created yet. Record a session to generate Meeting_Attendance.xlsx').classes('text-amber-400')

            # TAB 3: Mark Meeting Attendance
            with ui.tab_panel(tab_mark_attendance):
                with ui.card().classes('bg-slate-800 border border-slate-700 p-6 rounded-xl w-full gap-4'):
                    ui.label('Record Meeting Session').classes('text-lg font-bold text-slate-200')
                    
                    meeting_date = ui.input('Meeting Date', value=str(date.today())).classes('w-64')
                    selected_present = set()

                    ui.label('Select Present Members:').classes('font-semibold text-slate-300 mt-2')
                    
                    with ui.scroll_area().classes('h-64 border border-slate-700 rounded-lg p-4 bg-slate-900/50'):
                        for idx, row in df_members.iterrows():
                            name = str(row.get('Name', 'Unknown'))
                            reg_no = str(row.get('Register number', 'N/A'))
                            handle = str(row.get('Username', 'N/A'))
                            
                            display_label = f"{name} | {reg_no} (@{handle})"
                            
                            def on_change(e, user=handle):
                                if e.value:
                                    selected_present.add(user)
                                else:
                                    selected_present.discard(user)
                                    
                            ui.checkbox(display_label, on_change=on_change).classes('text-slate-200 py-1 font-mono text-sm')

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

ui.run(title='Department Attendance Hub', dark=True, port=8080)