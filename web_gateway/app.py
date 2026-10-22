import os
import re
import uuid
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, session
from ftplib import FTP
from werkzeug.utils import secure_filename
import openpyxl

app = Flask(__name__)
# Session is used to keep the authenticated user and their in-progress
# absentee list between the multiple form submissions of the workflow.
app.secret_key = "coa-attendance-prototype-secret"

UPLOAD_FOLDER = os.path.abspath(os.path.join(os.getcwd(), "temp_uploads"))
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# FTP target (same host/port the dummy server listens on)
FTP_HOST = "127.0.0.1"
FTP_PORT = 2123

# The seating report is assumed to always be present on the FTP server.
SEATING_REPORT_FILENAME = "SeatingReport.xlsx"
# Finished absentee workbooks are stored in a separate folder on the server.
ABSENTEE_FOLDER = "absentee_reports"

# Parsed seating data cached per login token so we do not re-download and
# re-parse the workbook on every single student lookup.
SEATING_CACHE = {}
# The absentee list being built, keyed by the same login token.
ABSENTEE_LISTS = {}

# Column headers we rely on inside the seating report worksheet.
COL_DATE = "EXAM DATE"
COL_SESSION = "SESSION NAME"
COL_VENUE = "VENUE"
COL_SEAT = "SEAT ROM COLUMN"
COL_REG = "REG NO"
COL_NAME = "STUDENT NAME"


def norm(value):
    """Normalise a cell / form value for tolerant matching."""
    return str(value).strip().upper() if value is not None else ""


def find_seating_sheet(wb):
    """Return the worksheet that actually holds the seating report rows.

    We do not hard-code the date-specific sheet name; instead we pick the
    first sheet whose header row contains the student columns we need.
    """
    for ws in wb.worksheets:
        headers = [norm(c.value) for c in ws[1]]
        if COL_REG in headers and COL_NAME in headers:
            return ws
    return None


def load_seating_rows(xlsx_path):
    """Parse the seating report into a list of {header: value} dicts."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = find_seating_sheet(wb)
    if ws is None:
        return []

    headers = [norm(c.value) for c in ws[1]]
    rows = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        if not any(raw):
            continue
        rows.append(dict(zip(headers, raw)))
    return rows


def fetch_seating_report(username, password):
    """Download the seating report from the FTP server over real FTP and parse it.

    Login (control channel) enforces the credentials; the file bytes are
    transferred with RETR over the FTP data channel.
    """
    local_path = os.path.join(app.config["UPLOAD_FOLDER"], f"seating_{uuid.uuid4().hex[:8]}.xlsx")

    ftp = FTP()
    ftp.connect(FTP_HOST, FTP_PORT)
    ftp.login(user=username, passwd=password)
    ftp.set_pasv(True)
    with open(local_path, "wb") as f:
        ftp.retrbinary(f"RETR {SEATING_REPORT_FILENAME}", f.write)
    ftp.quit()

    rows = load_seating_rows(local_path)
    try:
        os.remove(local_path)
    except OSError:
        pass
    return rows


def lookup_student(rows, date, sess, venue, seat_row, seat_col):
    """Find the student sitting at the given place/seat.

    The invigilator only knows where & when they are (date, session, venue)
    and the target student's seat row and column; from that we recover the
    student's registration number and name.
    """
    seat = f"R{seat_row}C{seat_col}"
    date_n, sess_n, venue_n, seat_n = norm(date), norm(sess), norm(venue), norm(seat)
    for r in rows:
        if (
            norm(r.get(COL_DATE)) == date_n
            and norm(r.get(COL_SESSION)) == sess_n
            and norm(r.get(COL_VENUE)) == venue_n
            and norm(r.get(COL_SEAT)) == seat_n
        ):
            return {
                "reg_no": r.get(COL_REG),
                "name": r.get(COL_NAME),
                "date": r.get(COL_DATE),
                "session": r.get(COL_SESSION),
                "venue": r.get(COL_VENUE),
                "seat": seat,
            }
    return None


def build_report_filename(students, employee_id=None, session_id=None):
    """date + session + venue + roomno + absentees_list
       + employee id + session id + timestamp.xlsx

    The data carries a single VENUE value (e.g. SJT315); we split it into the
    venue/building prefix (SJT) and the room number (315) to fill the format.
    We also tag the file with who sent it (employee id), their login session
    id, and the moment it was generated (timestamp) so submissions stay unique
    and traceable.
    """
    first = students[0]
    date = str(first.get("date") or "").replace("/", "-")
    sess = str(first.get("session") or "")
    venue_raw = str(first.get("venue") or "")

    m = re.match(r"^([A-Za-z]+)(\d+)$", venue_raw)
    if m:
        venue, roomno = m.group(1), m.group(2)
    else:
        venue, roomno = venue_raw, ""

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")

    parts = [p for p in [date, sess, venue, roomno] if p]
    name = "_".join(parts) + "_absentees_list"
    if employee_id:
        name += f"_{employee_id}"
    if session_id:
        name += f"_{session_id}"
    name += f"_{timestamp}.xlsx"
    return secure_filename(name)


def make_absentee_workbook(students, out_path):
    """Write the absentee list (reg no + name + seat details) to an xlsx."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Absentees"
    ws.append(["REG NO", "STUDENT NAME", "EXAM DATE", "SESSION", "VENUE", "SEAT"])
    for s in students:
        ws.append([s["reg_no"], s["name"], s["date"], s["session"], s["venue"], s["seat"]])
    wb.save(out_path)


def send_report_to_server(username, password, local_path, filename):
    """Upload the absentee workbook to a separate folder on the FTP server.

    Login (control channel) enforces the credentials; the file is stored into
    the ABSENTEE_FOLDER with STOR over the FTP data channel.
    """
    ftp = FTP()
    ftp.connect(FTP_HOST, FTP_PORT)
    ftp.login(user=username, passwd=password)
    ftp.set_pasv(True)

    # Ensure the separate folder exists on the server, then store into it.
    try:
        ftp.mkd(ABSENTEE_FOLDER)
    except Exception:
        pass  # already exists
    ftp.cwd(ABSENTEE_FOLDER)

    with open(local_path, "rb") as f:
        ftp.storbinary(f"STOR {filename}", f)
    ftp.quit()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/login", methods=["POST"])
def login():
    username = request.form.get("username")
    password = request.form.get("password")

    try:
        rows = fetch_seating_report(username, password)
    except Exception as e:
        return render_template("index.html", error=f"Authentication/data error: {e}")

    if not rows:
        return render_template("index.html", error="Could not read the seating report from the server.")

    token = uuid.uuid4().hex
    SEATING_CACHE[token] = rows
    ABSENTEE_LISTS[token] = []
    session["token"] = token
    session["username"] = username
    session["password"] = password
    # Context (where & when) is not set yet; invigilator fills it next.
    session.pop("ctx", None)
    return redirect(url_for("context_page"))


@app.route("/logout")
def logout():
    token = session.get("token")
    SEATING_CACHE.pop(token, None)
    ABSENTEE_LISTS.pop(token, None)
    session.clear()
    return redirect(url_for("index"))


@app.route("/context")
def context_page():
    token = current_token()
    if not token:
        return redirect(url_for("index"))
    return render_template("context.html")


@app.route("/set_context", methods=["POST"])
def set_context():
    token = current_token()
    if not token:
        return redirect(url_for("index"))

    # Fix the "where & when" for the whole login session. It stays put across
    # submissions and new lists; only logging out and back in can change it.
    session["ctx"] = {
        "date": request.form.get("date"),
        "session": request.form.get("session"),
        "venue": request.form.get("venue"),
    }
    return redirect(url_for("form"))


def current_token():
    token = session.get("token")
    if not token or token not in SEATING_CACHE:
        return None
    return token


@app.route("/form")
def form():
    token = current_token()
    if not token:
        return redirect(url_for("index"))
    ctx = session.get("ctx")
    if not ctx:
        return redirect(url_for("context_page"))
    return render_template(
        "form.html",
        ctx=ctx,
        students=ABSENTEE_LISTS[token],
        message=request.args.get("message"),
        error=request.args.get("error"),
    )


@app.route("/add", methods=["POST"])
def add_student():
    token = current_token()
    if not token:
        return redirect(url_for("index"))
    ctx = session.get("ctx")
    if not ctx:
        return redirect(url_for("context_page"))

    # Where & when are fixed for the session; only the seat varies per student.
    seat_row = request.form.get("seat_row")
    seat_col = request.form.get("seat_col")

    student = lookup_student(
        SEATING_CACHE[token], ctx["date"], ctx["session"], ctx["venue"], seat_row, seat_col
    )
    if student is None:
        return redirect(url_for("form", error="No student found for that place and seat. Check your values."))

    # Avoid adding the same student (same reg no) twice.
    if any(s["reg_no"] == student["reg_no"] for s in ABSENTEE_LISTS[token]):
        return redirect(url_for("form", message=f"{student['name']} ({student['reg_no']}) is already in the list."))

    ABSENTEE_LISTS[token].append(student)
    return redirect(url_for("form", message=f"Added {student['name']} ({student['reg_no']}) to the list."))


@app.route("/confirm")
def confirm():
    token = current_token()
    if not token:
        return redirect(url_for("index"))
    if not ABSENTEE_LISTS[token]:
        return redirect(url_for("form", error="The absentee list is empty. Add at least one student."))
    return render_template("confirm.html", students=ABSENTEE_LISTS[token])


@app.route("/finalize", methods=["POST"])
def finalize():
    token = current_token()
    if not token:
        return redirect(url_for("index"))

    students = ABSENTEE_LISTS[token]
    if not students:
        return redirect(url_for("form", error="The absentee list is empty."))

    filename = build_report_filename(
        students,
        employee_id=session.get("username"),
        session_id=token,
    )
    local_path = os.path.join(app.config["UPLOAD_FOLDER"], f"{uuid.uuid4().hex[:8]}_{filename}")
    make_absentee_workbook(students, local_path)

    send_report_to_server(
        session.get("username"), session.get("password"), local_path, filename
    )
    try:
        os.remove(local_path)
    except OSError:
        pass

    # Clear the list so a new session can start fresh.
    ABSENTEE_LISTS[token] = []

    return render_template("result.html", filename=filename, folder=ABSENTEE_FOLDER)


if __name__ == "__main__":
    app.run(port=5000, debug=True)
