"""
Google Sheets database module for Pro Paint Teams.
Handles all read/write operations to a single Google Spreadsheet
with 5 worksheets: Clients, Jobs, Quote_Areas, Attendance, Bonus_Log.

Setup: the user creates a Google Sheet in their own Drive, shares it
with the service account email, and saves the spreadsheet ID to
sheet_config.json.  This avoids using the service account's own
(tiny) Drive storage quota.
"""

import os
import json
from datetime import date

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEET_CLIENTS = "Clients"
SHEET_JOBS = "Jobs"
SHEET_QUOTE_AREAS = "Quote_Areas"
SHEET_ATTENDANCE = "Attendance"
SHEET_BONUS_LOG = "Bonus_Log"

CLIENTS_HEADERS = ["Client Name", "Phone", "Email", "Address", "Notes", "Date Added"]
JOBS_HEADERS = [
    "Job No", "Job Name", "Client", "Area Manager", "Team Leader",
    "Start Date", "Status", "Total Labour Quoted", "Man Days Available", "Date Created",
]
QUOTE_AREAS_HEADERS = [
    "Job No", "Quote Area", "Unit", "Quantity", "Prod Rate", "Allowed Man-Days", "Description",
]
ATTENDANCE_HEADERS = [
    "Job No", "Painter Name", "Emp/ID", "Hourly Rate",
    "Day 1", "Day 2", "Day 3", "Day 4", "Day 5", "Day 6", "Day 7",
    "Day 8", "Day 9", "Day 10", "Day 11", "Day 12", "Day 13", "Day 14",
    "Total Hours", "Man-Days",
]
BONUS_LOG_HEADERS = [
    "Job No", "Man Days Available", "Actual Man-Days Used", "Days Saved",
    "Bonus Rate", "Total Bonus Pool", "Bonus Per Painter", "Date Completed",
]

_gc = None
_spreadsheet = None

_DIR = os.path.dirname(os.path.abspath(__file__))


def _creds_path():
    return os.path.join(_DIR, "credentials.json")


def _config_path():
    return os.path.join(_DIR, "sheet_config.json")


def _read_spreadsheet_id():
    """Read the spreadsheet ID from sheet_config.json."""
    path = _config_path()
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r") as f:
            data = json.load(f)
        return data.get("spreadsheet_id", "").strip() or None
    except Exception:
        return None


def save_spreadsheet_id(sheet_id):
    """Persist the spreadsheet ID to sheet_config.json."""
    with open(_config_path(), "w") as f:
        json.dump({"spreadsheet_id": sheet_id}, f, indent=2)


def get_service_account_email():
    """Return the service account email from credentials.json."""
    try:
        with open(_creds_path(), "r") as f:
            return json.load(f).get("client_email", "")
    except Exception:
        return ""


def _connect():
    """Authenticate and open the spreadsheet by ID (from sheet_config.json)."""
    global _gc, _spreadsheet
    if _gc is not None and _spreadsheet is not None:
        return _spreadsheet

    creds = Credentials.from_service_account_file(_creds_path(), scopes=SCOPES)
    _gc = gspread.authorize(creds)

    sheet_id = _read_spreadsheet_id()
    if not sheet_id:
        raise RuntimeError(
            "No spreadsheet ID configured. Run setup_sheets.py first "
            "(see README.md for instructions)."
        )

    _spreadsheet = _gc.open_by_key(sheet_id)
    _ensure_worksheets(_spreadsheet)
    return _spreadsheet


def _ensure_worksheets(sp):
    """Create any missing worksheets with headers."""
    existing = {ws.title for ws in sp.worksheets()}
    sheet_defs = [
        (SHEET_CLIENTS, CLIENTS_HEADERS),
        (SHEET_JOBS, JOBS_HEADERS),
        (SHEET_QUOTE_AREAS, QUOTE_AREAS_HEADERS),
        (SHEET_ATTENDANCE, ATTENDANCE_HEADERS),
        (SHEET_BONUS_LOG, BONUS_LOG_HEADERS),
    ]
    for title, headers in sheet_defs:
        if title not in existing:
            ws = sp.add_worksheet(title=title, rows=1000, cols=len(headers))
            ws.append_row(headers, value_input_option="RAW")

    if "Sheet1" in existing:
        try:
            default = sp.worksheet("Sheet1")
            if default.row_count <= 1:
                vals = default.get_all_values()
                if not vals or vals == [[]]:
                    sp.del_worksheet(default)
        except Exception:
            pass


def is_configured():
    """Return True if credentials.json + sheet_config.json both exist and are valid."""
    creds = _creds_path()
    if not os.path.isfile(creds):
        return False
    try:
        with open(creds, "r") as f:
            data = json.load(f)
        if "client_email" not in data:
            return False
    except Exception:
        return False

    return _read_spreadsheet_id() is not None


def connect():
    """Public entry point. Returns the spreadsheet or raises on failure."""
    return _connect()


# ---------------------------------------------------------------------------
# Clients
# ---------------------------------------------------------------------------

def save_client(name, phone="", email="", address="", notes=""):
    """Append client if not already present (matched by name). Returns True if new."""
    sp = _connect()
    ws = sp.worksheet(SHEET_CLIENTS)
    existing = ws.col_values(1)
    if name in existing:
        return False
    ws.append_row(
        [name, phone, email, address, notes, str(date.today())],
        value_input_option="USER_ENTERED",
    )
    return True


def get_all_clients():
    """Return all clients as a DataFrame."""
    sp = _connect()
    ws = sp.worksheet(SHEET_CLIENTS)
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=CLIENTS_HEADERS)
    return pd.DataFrame(data)


# ---------------------------------------------------------------------------
# Jobs
# ---------------------------------------------------------------------------

def save_job(job_no, job_name, client, area_manager, team_leader,
             start_date, total_labour, man_days_available, status="Open"):
    """Insert or update a job by Job No."""
    sp = _connect()
    ws = sp.worksheet(SHEET_JOBS)
    job_nos = ws.col_values(1)
    row_data = [
        job_no, job_name, client, area_manager, team_leader,
        str(start_date), status, float(total_labour),
        round(float(man_days_available), 2), str(date.today()),
    ]

    if job_no in job_nos:
        row_idx = job_nos.index(job_no) + 1
        ws.update(f"A{row_idx}:J{row_idx}", [row_data], value_input_option="USER_ENTERED")
    else:
        ws.append_row(row_data, value_input_option="USER_ENTERED")


def get_all_jobs():
    """Return all jobs as a DataFrame."""
    sp = _connect()
    ws = sp.worksheet(SHEET_JOBS)
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=JOBS_HEADERS)
    return pd.DataFrame(data)


def update_job_status(job_no, new_status):
    """Set job status to 'Open' or 'Complete'."""
    sp = _connect()
    ws = sp.worksheet(SHEET_JOBS)
    job_nos = ws.col_values(1)
    if job_no in job_nos:
        row_idx = job_nos.index(job_no) + 1
        ws.update_cell(row_idx, 7, new_status)


# ---------------------------------------------------------------------------
# Quote Areas
# ---------------------------------------------------------------------------

def save_quote_areas(job_no, df):
    """Replace all quote area rows for a job_no with the given DataFrame."""
    sp = _connect()
    ws = sp.worksheet(SHEET_QUOTE_AREAS)

    all_vals = ws.get_all_values()
    rows_to_keep = [all_vals[0]]
    for row in all_vals[1:]:
        if row and row[0] != job_no:
            rows_to_keep.append(row)

    for _, r in df.iterrows():
        rows_to_keep.append([
            job_no,
            str(r.get("Quote Area", "")),
            str(r.get("Unit", "")),
            float(r.get("Quantity", 0)),
            float(r.get("Prod Qty / Man-Day", 0)),
            round(float(r.get("Allowed Man-Days", 0)), 2),
            str(r.get("Description", "")),
        ])

    ws.clear()
    ws.update(f"A1:G{len(rows_to_keep)}", rows_to_keep, value_input_option="USER_ENTERED")


def get_quote_areas(job_no=None):
    """Return quote areas, optionally filtered by job_no."""
    sp = _connect()
    ws = sp.worksheet(SHEET_QUOTE_AREAS)
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=QUOTE_AREAS_HEADERS)
    df = pd.DataFrame(data)
    if job_no is not None:
        df = df[df["Job No"] == job_no]
    return df


# ---------------------------------------------------------------------------
# Attendance
# ---------------------------------------------------------------------------

def save_attendance(job_no, df):
    """Replace all attendance rows for a job_no."""
    sp = _connect()
    ws = sp.worksheet(SHEET_ATTENDANCE)

    all_vals = ws.get_all_values()
    rows_to_keep = [all_vals[0]] if all_vals else [ATTENDANCE_HEADERS]
    for row in all_vals[1:]:
        if row and row[0] != job_no:
            rows_to_keep.append(row)

    day_cols = [f"Day {i}" for i in range(1, 15)]
    for _, r in df.iterrows():
        row_data = [
            job_no,
            str(r.get("Painter Name", "")),
            str(r.get("Emp/ID", "")),
            float(r.get("Hourly Rate", 0)),
        ]
        for d in day_cols:
            row_data.append(float(r.get(d, 0)))
        row_data.append(float(r.get("Total Hours", 0)))
        row_data.append(round(float(r.get("Man-Days", 0)), 2))
        rows_to_keep.append(row_data)

    ws.clear()
    if rows_to_keep:
        ws.update(
            f"A1:{gspread.utils.rowcol_to_a1(len(rows_to_keep), len(ATTENDANCE_HEADERS))}",
            rows_to_keep,
            value_input_option="USER_ENTERED",
        )


def get_attendance(job_no=None):
    """Return attendance records, optionally filtered by job_no."""
    sp = _connect()
    ws = sp.worksheet(SHEET_ATTENDANCE)
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=ATTENDANCE_HEADERS)
    df = pd.DataFrame(data)
    if job_no is not None:
        df = df[df["Job No"] == job_no]
    return df


# ---------------------------------------------------------------------------
# Bonus Log
# ---------------------------------------------------------------------------

def save_bonus(job_no, man_days_available, actual_used, days_saved,
               bonus_rate, total_bonus_pool, bonus_per_painter):
    """Insert or update a bonus log entry for a job."""
    sp = _connect()
    ws = sp.worksheet(SHEET_BONUS_LOG)
    job_nos = ws.col_values(1)
    row_data = [
        job_no,
        round(float(man_days_available), 2),
        round(float(actual_used), 2),
        round(float(days_saved), 2),
        float(bonus_rate),
        round(float(total_bonus_pool), 2),
        round(float(bonus_per_painter), 2),
        str(date.today()),
    ]

    if job_no in job_nos:
        row_idx = job_nos.index(job_no) + 1
        ws.update(f"A{row_idx}:H{row_idx}", [row_data], value_input_option="USER_ENTERED")
    else:
        ws.append_row(row_data, value_input_option="USER_ENTERED")


def get_bonus_log():
    """Return all bonus log entries as a DataFrame."""
    sp = _connect()
    ws = sp.worksheet(SHEET_BONUS_LOG)
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=BONUS_LOG_HEADERS)
    return pd.DataFrame(data)


# ---------------------------------------------------------------------------
# Dashboard helpers
# ---------------------------------------------------------------------------

def get_job_history():
    """Return merged jobs + bonus data for dashboard charts."""
    jobs = get_all_jobs()
    bonus = get_bonus_log()
    if jobs.empty:
        return jobs
    if bonus.empty:
        return jobs
    merged = jobs.merge(bonus, on="Job No", how="left", suffixes=("", "_bonus"))
    return merged
