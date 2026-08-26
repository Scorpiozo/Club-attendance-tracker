# ⚡ Technical Department - Club Attendance & Contest Tracker

A complete management suite and scraper system designed to track competitive programming contest performance on CodeChef and monitor meeting attendance for department members. 

It provides an interactive local management dashboard built with **NiceGUI** and **AgGrid**, updates Excel storage files automatically, and generates a responsive static HTML dashboard for remote viewing via **GitHub Pages**.

---

## 📸 Overview & Key Features

* **CodeChef Contest Scraper**: Automatically scrapes problem-solving data for CodeChef Starters contests using multithreading and Cloudflare-safe requests.
* **Google Form Integration**: Automatically imports initial member rosters (Name, Registration Number, Phone Number, Username) directly from Google Form CSV/Excel exports.
* **Dual Attendance Systems**:
  * **Contest Tracking**: Calculates problems solved per member per contest round in `CP_Members.xlsx`.
  * **Meeting Tracking**: Logs daily attendance sessions (Present: `1`, Absent: `0`) into a dedicated `Meeting_Attendance.xlsx`.
* **Interactive Local Management Hub (`app.py`)**: Web-based desktop GUI powered by NiceGUI to trigger live scraping, view member rosters, search records, and record meeting attendance.
* **Static HTML Exporter (`export_static.py`)**: Exports Excel data into a lightweight, dark-themed static webpage featuring **DataTables** (searchable, filterable, sortable) hostable on GitHub Pages.

---

## 📁 Repository Structure

```text
├── ContestAttendance.py  # Standalone multithreaded CodeChef web scraper & importer
├── app.py                # NiceGUI local interactive management dashboard
├── export_static.py      # Script to build index.html from Excel files for GitHub Pages
├── index.html            # Generated static dashboard published on GitHub Pages
├── CP_Members.xlsx       # Primary Excel sheet storing member details & contest records
├── Meeting_Attendance.xlsx # Excel sheet storing dated meeting attendance records
├── gform_responses.xlsx  # (Optional) Exported Google Form responses for auto-onboarding
└── members.txt           # (Optional) Fallback plain text list of CodeChef usernames
```

---

## 🛠️ System Architecture & How It Works

```
                     ┌────────────────────────┐
                     │  Google Form / CSV     │
                     └───────────┬────────────┘
                                 │ Initial Import
                                 ▼
   ┌──────────────────────────────────────────────────────────┐
   │                  ContestAttendance.py                    │
   │  - Scrapes CodeChef profile data concurrently            │
   │  - Parses problem count per Starters contest round      │
   └─────────────────────────────┬────────────────────────────┘
                                 │ Saves Data
                                 ▼
           ┌──────────────────────────────────────────┐
           │            Excel Storage                 │
           │  - CP_Members.xlsx                       │
           │  - Meeting_Attendance.xlsx               │
           └─────────────────────┬────────────────────┘
                                 │ Reads/Writes
                                 ▼
   ┌──────────────────────────────────────────────────────────┐
   │                        app.py                            │
   │  - Local NiceGUI web app (http://localhost:8080)        │
   │  - Interactive grid, meeting logger, live status         │
   └─────────────────────────────┬────────────────────────────┘
                                 │ Render Static Output
                                 ▼
   ┌──────────────────────────────────────────────────────────┐
   │                   export_static.py                       │
   │  - Compiles Excel files into index.html                  │
   └─────────────────────────────┬────────────────────────────┘
                                 │ Git Push
                                 ▼
   ┌──────────────────────────────────────────────────────────┐
   │                  GitHub Pages Deploy                     │
   │  - Live public dashboard viewable anywhere               │
   └─────────────────────────────┬────────────────────────────┘
```

---

## 🚀 Setup & Installation

### 1. Prerequisites
* Python 3.9 or higher installed on your system.
* Git installed and configured.

### 2. Install Required Dependencies
Clone this repository and install the Python packages:

```bash
git clone [https://github.com/Scorpiozo/Club-attendance-tracker.git](https://github.com/Scorpiozo/Club-attendance-tracker.git)
cd Club-attendance-tracker

pip install pandas openpyxl nicegui requests cloudscraper beautifulsoup4
```

---

## ⚙️ Initial Roster Onboarding

You can initialize member records using **one of two methods**:

1. **Google Form Export (Recommended)**: Place your exported Excel (`gform_responses.xlsx`) or CSV (`gform_responses.csv`) file in the project root directory. The script automatically extracts fields like `Name`, `Register number`, `Phone number`, and `Username`.
2. **Text File Fallback**: Create a `members.txt` file in the root folder containing one CodeChef username per line.

---

## 🖥️ Usage Guide

### Method A: Using the Interactive Local Dashboard (`app.py`)

Run the local UI dashboard to manage everything in a web interface:

```bash
python app.py
```

Open your browser and navigate to `http://localhost:8080`.

* **Member Roster Tab**: View full member details, registration numbers, and historic contest scores inside an interactive data grid.
* **Live CodeChef Scraper Tab**: Enter the Starters contest number (e.g., `208`) and click **Start Live Scraper** to update Excel files in real time.
* **Meeting Attendance Tab**: Select present members from a checklist to dynamically log a new dated meeting column into `Meeting_Attendance.xlsx`.

---

### Method B: Running Scraper via Command Line (`ContestAttendance.py`)

To run contest scraping directly from the terminal without the GUI:

```bash
python ContestAttendance.py
```

1. Select Option `1` (CP Members).
2. Enter the Starters contest number (e.g., `208`).
3. The multithreaded scraper will parse user profiles concurrently and write results directly to `CP_Members.xlsx`.

---

## 🌐 Publishing Live Updates to GitHub Pages

Because GitHub Pages only hosts static files, `export_static.py` compiles the local Excel files into a self-contained, responsive `index.html` dashboard powered by Tailwind CSS and DataTables.

### Step 1: Export Static Dashboard
After running the scraper or logging meeting attendance, generate the updated HTML:

```bash
python export_static.py
```

### Step 2: Push Updates to GitHub
Commit and push the updated `index.html` file:

```bash
git add index.html
git commit -m "Update dashboard data"
git push origin main
```

Your live public dashboard will automatically update at:
`https://scorpiozo.github.io/Club-attendance-tracker/`
