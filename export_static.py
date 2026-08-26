import pandas as pd
import os

CP_EXCEL = "CP_Members.xlsx"
MEETING_EXCEL = "Meeting_Attendance.xlsx"
OUTPUT_HTML = "index.html"

def load_table_html(file_path: str) -> str:
    if not os.path.exists(file_path):
        return "<p class='text-amber-400 p-4'>File not generated yet.</p>"
    
    df = pd.read_excel(file_path)
    if df.empty:
        return "<p class='text-slate-400 p-4'>No data recorded.</p>"
    
    # Convert DataFrame to clean HTML table
    return df.to_html(classes="display responsive nowrap w-full text-sm text-slate-200", index=False, border=0)

# Load data
cp_table_html = load_table_html(CP_EXCEL)
meeting_table_html = load_table_html(MEETING_EXCEL)

# Generate responsive HTML dashboard
html_content = f"""<!DOCTYPE html>
<html lang="en" class="dark">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Technical Department Dashboard</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- DataTables CSS for Search & Sorting -->
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
    <style>
        body {{ background-color: #0f172a; color: #f8fafc; font-family: sans-serif; }}
        .dataTables_wrapper {{ color: #94a3b8; font-size: 0.875rem; }}
        table.dataTable tbody tr {{ background-color: #1e293b; color: #f8fafc; border-bottom: 1px solid #334155; }}
        table.dataTable tbody tr:hover {{ background-color: #334155; }}
        table.dataTable header th {{ background-color: #0f172a; color: #60a5fa; }}
        .dataTables_wrapper .dataTables_length, .dataTables_wrapper .dataTables_filter, 
        .dataTables_wrapper .dataTables_info, .dataTables_wrapper .dataTables_paginate {{ color: #94a3b8 !important; margin-bottom: 10px; }}
        input, select {{ background-color: #1e293b !important; color: #f8fafc !important; border: 1px solid #475569 !important; border-radius: 6px; padding: 4px 8px; }}
    </style>
</head>
<body class="p-6">
    <div class="max-w-6xl mx-auto space-y-6">
        <!-- Header -->
        <header class="flex justify-between items-center bg-slate-800 p-6 rounded-xl border border-slate-700">
            <div>
                <h1 class="text-2xl font-bold text-blue-400">⚡ Technical Department Hub</h1>
                <p class="text-xs text-slate-400">Live Attendance & Contest Performance (Read-Only)</p>
            </div>
        </header>

        <!-- Navigation Tabs -->
        <div class="flex gap-4 border-b border-slate-700 pb-2">
            <button onclick="switchTab('roster')" id="tab-roster-btn" class="px-4 py-2 font-semibold text-blue-400 border-b-2 border-blue-400">Member Roster & Contests</button>
            <button onclick="switchTab('meeting')" id="tab-meeting-btn" class="px-4 py-2 font-semibold text-slate-400 hover:text-slate-200">Meeting Attendance</button>
        </div>

        <!-- Tab 1: CP Members -->
        <div id="tab-roster" class="bg-slate-800 p-6 rounded-xl border border-slate-700">
            <h2 class="text-lg font-bold text-slate-200 mb-4">CP Member & Contest Roster</h2>
            <div class="overflow-x-auto">
                {cp_table_html}
            </div>
        </div>

        <!-- Tab 2: Meetings -->
        <div id="tab-meeting" class="bg-slate-800 p-6 rounded-xl border border-slate-700 hidden">
            <h2 class="text-lg font-bold text-slate-200 mb-4">Meeting Attendance Logs</h2>
            <div class="overflow-x-auto">
                {meeting_table_html}
            </div>
        </div>
    </div>

    <!-- DataTables & Tab Switching JS -->
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script>
        $(document).ready(function() {{
            $('table').DataTable({{
                responsive: true,
                pageLength: 10,
                lengthMenu: [10, 25, 50, 100]
            }});
        }});

        function switchTab(tab) {{
            if (tab === 'roster') {{
                $('#tab-roster').removeClass('hidden');
                $('#tab-meeting').addClass('hidden');
                $('#tab-roster-btn').addClass('text-blue-400 border-b-2 border-blue-400').removeClass('text-slate-400');
                $('#tab-meeting-btn').removeClass('text-blue-400 border-b-2 border-blue-400').addClass('text-slate-400');
            }} else {{
                $('#tab-meeting').removeClass('hidden');
                $('#tab-roster').addClass('hidden');
                $('#tab-meeting-btn').addClass('text-blue-400 border-b-2 border-blue-400').removeClass('text-slate-400');
                $('#tab-roster-btn').removeClass('text-blue-400 border-b-2 border-blue-400').addClass('text-slate-400');
            }}
        }}
    </script>
</body>
</html>
"""

with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"✅ Generated static dashboard at '{OUTPUT_HTML}' successfully!")