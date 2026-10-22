# COA Attendance Reconciliation System

## Prerequisites: seating report must exist on the server

The app assumes the seating report is **always available on the FTP server**. Before
running, make sure the report file is present at:

```
ftp_service/FTP_INBOX_PUBLIC/SeatingReport.xlsx
```

If it is missing, login will fail with "Could not read the seating report from the
server". To restore it, copy your source workbook into place, e.g.:

```bash
cp "CAT2_03-03-2019_AN-A1.xlsx" ftp_service/FTP_INBOX_PUBLIC/SeatingReport.xlsx
```

Note on folders: `ftp_service/FTP_INBOX_PUBLIC/` and `web_gateway/temp_uploads/` are
re-created automatically on startup if deleted, and the output subfolder
`FTP_INBOX_PUBLIC/absentee_reports/` is created on first submission. **However,
deleting `FTP_INBOX_PUBLIC/` also deletes `SeatingReport.xlsx`** — the folder comes
back empty, so you must place `SeatingReport.xlsx` back inside it before the app will work.

## Install dependencies

```bash
pip install -r requirements.txt
```

## Run (always use two separate terminals)

### for server

```bash
cd ftp_service
python ftp_server.py
```

### for client

```bash
cd web_gateway
python app.py
```

Then open http://127.0.0.1:5000 and log in (dummy credentials: `employee_123` / `password123`).
