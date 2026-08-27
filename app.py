from flask import (
    Flask,
    request,
    redirect,
    url_for,
    session,
    render_template_string,
    jsonify,
    send_from_directory
)

import sqlite3
import os
import secrets
import string
import base64
import hmac
from io import BytesIO
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import qrcode
import requests
from markupsafe import escape

GOOGLE_SHEET_URL = os.environ.get("GOOGLE_SHEET_URL", "")
# =========================================================
# BASIC SETTINGS
# =========================================================

app = Flask(__name__)

app.secret_key = os.environ.get("SECRET_KEY", "smartmess-local-development-key")
app.config.update(
    MAX_CONTENT_LENGTH=5 * 1024 * 1024,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("COOKIE_SECURE", "0") == "1",
)

ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "shahil123")
INDIA_TIMEZONE = ZoneInfo("Asia/Kolkata")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.environ.get("DATABASE_PATH", os.path.join(BASE_DIR, "smart_mess.db"))
PHOTO_DIR = os.environ.get("PHOTO_DIR", os.path.join(BASE_DIR, "photos"))

os.makedirs(PHOTO_DIR, exist_ok=True)


# =========================================================
# DATABASE
# =========================================================

def db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = db()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_uid TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            roll_number TEXT UNIQUE NOT NULL,
            branch TEXT NOT NULL,
            hostel_room TEXT NOT NULL,
            photo_filename TEXT,
            active INTEGER DEFAULT 1,
            created_at TEXT NOT NULL
        )
    """)

    # Phase 1 migration: preserve all existing students while adding hostel details.
    student_columns = {row[1] for row in conn.execute("PRAGMA table_info(students)")}
    if "gender" not in student_columns:
        conn.execute("ALTER TABLE students ADD COLUMN gender TEXT DEFAULT 'NOT SET'")
    if "hostel_name" not in student_columns:
        conn.execute("ALTER TABLE students ADD COLUMN hostel_name TEXT DEFAULT 'NOT SET'")
    if "hostel_block" not in student_columns:
        conn.execute("ALTER TABLE students ADD COLUMN hostel_block TEXT DEFAULT ''")

    conn.execute("""
        CREATE TABLE IF NOT EXISTS coupons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            token TEXT UNIQUE NOT NULL,
            student_uid TEXT NOT NULL,
            meal TEXT NOT NULL,
            generated_at TEXT NOT NULL,
            expires_at TEXT NOT NULL,
            used_at TEXT,
            status TEXT NOT NULL DEFAULT 'ACTIVE'
        )
    """)

    conn.commit()
    conn.close()


init_db()


# =========================================================
# HELPERS
# =========================================================

def current_time():
    # Store India-local timestamps so the meal date remains correct on Render.
    return datetime.now(INDIA_TIMEZONE).replace(tzinfo=None)


def make_student_uid():
    return "STU-" + secrets.token_hex(5).upper()


def make_coupon_token():
    chars = string.ascii_uppercase + string.digits
    return "CPN-" + "".join(
        secrets.choice(chars) for _ in range(20)
    )


def admin_required():
    return session.get("admin") is True


@app.after_request
def add_security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    return response


def image_data_uri(filename):
    if not filename:
        return ""

    path = os.path.join(PHOTO_DIR, filename)

    if not os.path.exists(path):
        return ""

    try:
        with open(path, "rb") as f:
            encoded = base64.b64encode(
                f.read()
            ).decode("utf-8")

        ext = os.path.splitext(filename)[1].lower()

        mime = "image/jpeg"

        if ext == ".png":
            mime = "image/png"
        elif ext == ".webp":
            mime = "image/webp"

        return f"data:{mime};base64,{encoded}"

    except Exception:
        return ""


def qr_data_uri(text):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=8,
        border=3
    )

    qr.add_data(text)
    qr.make(fit=True)

    image = qr.make_image(
        fill_color="black",
        back_color="white"
    )

    buffer = BytesIO()
    image.save(buffer, format="PNG")

    encoded = base64.b64encode(
        buffer.getvalue()
    ).decode("utf-8")

    return "data:image/png;base64," + encoded


# =========================================================
# COMMON CSS
# =========================================================

CSS = """
<style>
* {
    box-sizing: border-box;
}

body {
    margin: 0;
    background: #f3f4f6;
    font-family: Arial, Helvetica, sans-serif;
    color: #111827;
}

.container {
    width: 94%;
    max-width: 1150px;
    margin: 30px auto;
}

.card {
    background: #ffffff;
    padding: 24px;
    border-radius: 18px;
    margin-bottom: 20px;
    box-shadow: 0 8px 25px rgba(0,0,0,.07);
}

.nav {
    display: flex;
    gap: 10px;
    flex-wrap: wrap;
    margin-bottom: 20px;
}

.btn {
    display: inline-block;
    padding: 12px 17px;
    background: #111827;
    color: white;
    text-decoration: none;
    border-radius: 10px;
    border: 0;
    cursor: pointer;
}

.green {
    background: #16a34a;
}

.blue {
    background: #2563eb;
}

.red {
    background: #dc2626;
}

.gray {
    background: #6b7280;
}

input,
select {
    width: 100%;
    padding: 13px;
    margin: 7px 0 16px;
    border: 1px solid #d1d5db;
    border-radius: 10px;
    font-size: 15px;
}

.grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 16px;
}

.stat {
    background: white;
    padding: 20px;
    border-radius: 16px;
}

.stat h2 {
    margin: 5px 0 0;
    font-size: 30px;
}

.message {
    padding: 14px;
    border-radius: 10px;
    margin-bottom: 16px;
    background: #eff6ff;
}

.success {
    background: #dcfce7;
    color: #166534;
}

.error {
    background: #fee2e2;
    color: #991b1b;
}

.center {
    text-align: center;
}

.photo {
    width: 65px;
    height: 65px;
    border-radius: 50%;
    object-fit: cover;
}

.big-photo {
    width: 140px;
    height: 140px;
    object-fit: cover;
    border-radius: 20px;
}

.qr {
    width: 270px;
    max-width: 95%;
}

.countdown {
    font-size: 30px;
    font-weight: bold;
    margin: 15px 0;
}

table {
    width: 100%;
    border-collapse: collapse;
}

th,
td {
    padding: 11px;
    border-bottom: 1px solid #e5e7eb;
    text-align: left;
}

.badge {
    display: inline-block;
    padding: 5px 9px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: bold;
}

.active-badge {
    background: #dcfce7;
    color: #166534;
}

.used-badge {
    background: #dbeafe;
    color: #1d4ed8;
}

.expired-badge {
    background: #fee2e2;
    color: #991b1b;
}

.meal-buttons {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 12px;
}

.meal-buttons button {
    padding: 16px;
    border: 0;
    border-radius: 12px;
    cursor: pointer;
    font-size: 16px;
    font-weight: bold;
    background: #111827;
    color: white;
}

@media(max-width: 650px) {
    .meal-buttons {
        grid-template-columns: 1fr;
    }

    table {
        font-size: 12px;
    }
}

/* Premium SmartMess theme */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
:root { --navy:#081b33; --blue:#2563eb; --cyan:#06b6d4; --ink:#10233f; --muted:#64748b; }
body { min-height:100vh; color:var(--ink); font-family:Inter,system-ui,-apple-system,"Segoe UI",sans-serif;
 background:radial-gradient(circle at 8% 0%,rgba(37,99,235,.16),transparent 28rem),radial-gradient(circle at 95% 8%,rgba(6,182,212,.13),transparent 25rem),#f4f7fb; }
body:before { content:""; display:block; height:5px; background:linear-gradient(90deg,var(--blue),var(--cyan)); }
.container { width:min(94%,1180px); margin:32px auto; }
h1,h2 { color:var(--navy); letter-spacing:-.035em; }
p { line-height:1.65; color:var(--muted); }
.card,.stat { background:rgba(255,255,255,.94); border:1px solid rgba(148,163,184,.18); box-shadow:0 18px 55px rgba(15,35,65,.09); }
.card { padding:clamp(22px,4vw,34px); border-radius:24px; }
.nav { gap:10px; padding:10px; background:var(--navy); border-radius:18px; box-shadow:0 16px 35px rgba(8,27,51,.18); }
.btn,.meal-buttons button { display:inline-flex; align-items:center; justify-content:center; gap:7px; padding:12px 18px; font:600 14px Inter,sans-serif;
 background:linear-gradient(135deg,#172a46,#0b1d35); border-radius:12px; transition:transform .2s,box-shadow .2s,filter .2s; }
.btn:hover,.meal-buttons button:hover { transform:translateY(-2px); box-shadow:0 10px 24px rgba(15,35,65,.2); filter:brightness(1.08); }
.green { background:linear-gradient(135deg,#16a34a,#059669); }
.blue { background:linear-gradient(135deg,#2563eb,#0891b2); }
.red { background:linear-gradient(135deg,#ef4444,#be123c); }
.gray { background:linear-gradient(135deg,#64748b,#475569); }
label { display:block; margin-top:5px; font-weight:650; font-size:14px; }
input,select { padding:14px 15px; color:var(--ink); background:#f8fafc; border:1px solid #dbe3ee; border-radius:12px; outline:none; font:500 15px Inter,sans-serif; transition:.2s; }
input:focus,select:focus { background:#fff; border-color:#60a5fa; box-shadow:0 0 0 4px rgba(37,99,235,.1); }
.grid { gap:17px; margin-bottom:22px; }
.stat { position:relative; overflow:hidden; padding:23px; border-radius:20px; color:var(--muted); }
.stat:after { content:""; position:absolute; right:-22px; bottom:-28px; width:85px; height:85px; border-radius:50%; background:rgba(37,99,235,.10); }
.stat h2 { font-size:34px; font-weight:800; color:var(--navy); }
.message { border-left:4px solid #3b82f6; }
.success { background:#ecfdf5; border-color:#22c55e; }
.error { background:#fff1f2; border-color:#f43f5e; }
.photo,.big-photo { border:4px solid #fff; box-shadow:0 8px 24px rgba(15,35,65,.16); }
.photo { border-radius:17px; }
.big-photo { border-radius:28px; }
.qr { padding:12px; border-radius:22px; background:#fff; box-shadow:0 15px 40px rgba(15,35,65,.13); }
.countdown { display:inline-block; min-width:130px; padding:9px 16px; color:#1d4ed8; background:#eff6ff; border-radius:14px; }
table { border-collapse:separate; border-spacing:0; }
th { color:#475569; background:#f8fafc; font-size:12px; text-transform:uppercase; letter-spacing:.05em; }
tr:hover td { background:#f8fbff; }
.badge { font-weight:800; }
.meal-buttons form,.meal-buttons button { width:100%; }
.meal-buttons button { min-height:74px; background:linear-gradient(145deg,#102b4f,#2563eb); }
@media(max-width:650px) { .container{margin:18px auto}.card{border-radius:20px;padding:20px}.nav{position:sticky;top:8px;z-index:20;overflow-x:auto;flex-wrap:nowrap}.nav .btn{white-space:nowrap}h1{font-size:27px} }
</style>
"""


# =========================================================
# ROOT
# =========================================================

@app.route("/")
def root():
    return redirect(url_for("student_home"))


@app.route("/health")
def health():
    return jsonify({"status": "ok", "service": "SmartMess"})


# =========================================================
# ADMIN LOGIN
# =========================================================

@app.route("/admin", methods=["GET", "POST"])
def admin_login():

    if admin_required():
        return redirect(url_for("admin_dashboard"))

    error = ""

    if request.method == "POST":

        password = request.form.get("password", "")

        if hmac.compare_digest(password, ADMIN_PASSWORD):

            session["admin"] = True

            return redirect(url_for("admin_dashboard"))

        error = "Wrong admin password."

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Admin Login</title>
        {CSS}
    </head>

    <body>
    <div class="container">

        <div class="card" style="max-width:500px;margin:80px auto;">

            <h1>🔐 Admin Login</h1>

            {'<div class="message error">' + error + '</div>'
             if error else ''}

            <form method="POST">

                <label>Admin Password</label>

                <input
                    type="password"
                    name="password"
                    required
                >

                <button class="btn" type="submit">
                    Login
                </button>

            </form>

        </div>

    </div>
    </body>
    </html>
    """

    return html


@app.route("/admin/logout")
def admin_logout():

    session.pop("admin", None)

    return redirect(url_for("admin_login"))

# =========================================================
# ADMIN VERIFY COUPON
# =========================================================

@app.route("/admin/verify-coupon", methods=["POST"])
def verify_coupon():

    if not admin_required():
        return jsonify({
            "success": False,
            "message": "Admin login required."
        }), 401

    data = request.get_json(silent=True) or {}

    token = str(
        data.get("token", "")
    ).strip()

    if not token:
        return jsonify({
            "success": False,
            "message": "Invalid QR."
        }), 400

    conn = db()

    # ---------------------------------------------------------
    # FIND COUPON
    # ---------------------------------------------------------

    coupon = conn.execute("""
        SELECT *
        FROM coupons
        WHERE token = ?
    """, (token,)).fetchone()

    if not coupon:
        conn.close()

        return jsonify({
            "success": False,
            "message": "Coupon not found."
        }), 404

    # ---------------------------------------------------------
    # CHECK ALREADY USED
    # ---------------------------------------------------------

    if coupon["status"] == "USED":

        conn.close()

        return jsonify({
            "success": False,
            "message": "Coupon already used."
        })

    # ---------------------------------------------------------
    # CHECK EXPIRY
    # ---------------------------------------------------------

    expiry = datetime.strptime(
        coupon["expires_at"],
        "%Y-%m-%d %H:%M:%S"
    )

    if current_time() > expiry:

        conn.execute("""
            UPDATE coupons
            SET status = 'EXPIRED'
            WHERE id = ?
        """, (coupon["id"],))

        conn.commit()
        conn.close()

        return jsonify({
            "success": False,
            "message": "Coupon expired. 5-minute validity ended."
        })

    # ---------------------------------------------------------
    # FIND STUDENT
    # ---------------------------------------------------------

    student = conn.execute("""
        SELECT *
        FROM students
        WHERE student_uid = ?
        AND active = 1
    """, (
        coupon["student_uid"],
    )).fetchone()

    if not student:

        conn.close()

        return jsonify({
            "success": False,
            "message": "Student is not active."
        })

    # ---------------------------------------------------------
    # MARK COUPON AS USED
    # ---------------------------------------------------------

    used_time = current_time().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    cursor = conn.execute("""
        UPDATE coupons
        SET status = 'USED',
            used_at = ?
        WHERE id = ?
        AND status = 'ACTIVE'
    """, (
        used_time,
        coupon["id"]
    ))

    conn.commit()

    if cursor.rowcount != 1:

        conn.close()

        return jsonify({
            "success": False,
            "message": "Coupon was already used."
        })

    # ---------------------------------------------------------
    # SAVE SUCCESSFUL SCAN TO GOOGLE SHEET
    # ---------------------------------------------------------
    try:

        if not GOOGLE_SHEET_URL:
            raise ValueError("GOOGLE_SHEET_URL is not configured")

        now = current_time()

        response = requests.post(
            GOOGLE_SHEET_URL,
            json={
                "date": now.strftime("%d-%m-%Y"),
                "time": now.strftime("%I:%M:%S %p"),
                "name": student["name"],
                "roll": student["roll_number"],
                "branch": student["branch"],
                "room": student["hostel_room"],
                "gender": student["gender"],
                "hostel": student["hostel_name"],
                "block": student["hostel_block"],
                "meal": coupon["meal"],
                "status": "USED"
            },
            timeout=10
        )

        print(
            "GOOGLE SHEET:",
            response.status_code,
            response.text
        )

    except Exception as e:

        print(
            "Google Sheet Error:",
            e
        )
    # ---------------------------------------------------------
    # STUDENT PHOTO
    # ---------------------------------------------------------

    photo_url = ""

    if student["photo_filename"]:

        photo_url = (
            "/student-photo/"
            + str(student["id"])
        )

    conn.close()

    # ---------------------------------------------------------
    # SUCCESS RESPONSE
    # ---------------------------------------------------------

    return jsonify({
        "success": True,
        "name": student["name"],
        "roll": student["roll_number"],
        "branch": student["branch"],
        "room": student["hostel_room"],
        "gender": student["gender"],
        "hostel": student["hostel_name"],
        "block": student["hostel_block"],
        "meal": coupon["meal"],
        "photo": photo_url
    })# =========================================================
# ADMIN DASHBOARD
# =========================================================

@app.route("/admin/dashboard")
def admin_dashboard():

    if not admin_required():
        return redirect(url_for("admin_login"))

    today = current_time().strftime("%Y-%m-%d")

    conn = db()

    student_count = conn.execute("""
        SELECT COUNT(*) AS c
        FROM students
        WHERE active = 1
    """).fetchone()["c"]

    boys_count = conn.execute("SELECT COUNT(*) AS c FROM students WHERE active = 1 AND gender = 'BOY'").fetchone()["c"]
    girls_count = conn.execute("SELECT COUNT(*) AS c FROM students WHERE active = 1 AND gender = 'GIRL'").fetchone()["c"]

    breakfast = conn.execute("""
        SELECT COUNT(*) AS c
        FROM coupons
        WHERE meal = 'BREAKFAST'
        AND status = 'USED'
        AND substr(generated_at, 1, 10) = ?
    """, (today,)).fetchone()["c"]

    lunch = conn.execute("""
        SELECT COUNT(*) AS c
        FROM coupons
        WHERE meal = 'LUNCH'
        AND status = 'USED'
        AND substr(generated_at, 1, 10) = ?
    """, (today,)).fetchone()["c"]

    dinner = conn.execute("""
        SELECT COUNT(*) AS c
        FROM coupons
        WHERE meal = 'DINNER'
        AND status = 'USED'
        AND substr(generated_at, 1, 10) = ?
    """, (today,)).fetchone()["c"]

    conn.close()

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Admin Dashboard</title>
        {CSS}
    </head>

    <body>
    <div class="container">

        <h1>🍽️ Smart Mess Admin Panel</h1>

        <div class="nav">

            <a class="btn" href="/admin/dashboard">
                Dashboard
            </a>

            <a class="btn green" href="/admin/add-student">
                ➕ Register Student
            </a>

            <a class="btn blue" href="/admin/students">
                👨‍🎓 Students
            </a>

            <a class="btn" href="/admin/scanner">
                📷 Scanner
            </a>

            <a class="btn gray" href="/admin/records">
                📊 Records
            </a>

            <a class="btn red" href="/admin/logout">
                Logout
            </a>

        </div>

        <div class="grid">

            <div class="stat">
                <div>Hostel Students</div>
                <h2>{student_count}</h2>
            </div>

            <div class="stat">
                <div>👦 Boys</div>
                <h2>{boys_count}</h2>
            </div>

            <div class="stat">
                <div>👧 Girls</div>
                <h2>{girls_count}</h2>
            </div>

            <div class="stat">
                <div>Breakfast Today</div>
                <h2>{breakfast}</h2>
            </div>

            <div class="stat">
                <div>Lunch Today</div>
                <h2>{lunch}</h2>
            </div>

            <div class="stat">
                <div>Dinner Today</div>
                <h2>{dinner}</h2>
            </div>

        </div>

        <div class="card">
            <h2>How it works</h2>

            <p>
                Admin registers hostel students with photo.
            </p>

            <p>
                Student uses the Student Panel to generate
                a meal coupon.
            </p>

            <p>
                Coupon QR is valid for 5 minutes and can
                be used only once.
            </p>

        </div>

    </div>
    </body>
    </html>
    """

    return html


# =========================================================
# ADMIN - ADD STUDENT
# =========================================================


@app.route("/admin/add-student", methods=["GET", "POST"])
def admin_add_student():

    if not admin_required():
        return redirect(url_for("admin_login"))

    message = ""
    message_class = "message"

    if request.method == "POST":

        name = request.form.get("name", "").strip()
        roll = request.form.get("roll_number", "").strip()
        branch = request.form.get("branch", "").strip()
        room = request.form.get("hostel_room", "").strip()
        gender = request.form.get("gender", "").strip().upper()
        hostel_name = request.form.get("hostel_name", "").strip()
        hostel_block = request.form.get("hostel_block", "").strip()
        photo = request.files.get("photo")

        if (not name or not roll or not branch or not room or
                gender not in ["BOY", "GIRL"] or not hostel_name):

            message = "Please fill every field."
            message_class = "message error"

        elif not photo or not photo.filename:

            message = "Student photo is required."
            message_class = "message error"

        else:

            conn = db()

            exists = conn.execute("""
                SELECT id
                FROM students
                WHERE roll_number = ?
            """, (roll,)).fetchone()

            if exists:

                conn.close()

                message = "This Registration Number is already registered."
                message_class = "message error"

            else:

                uid = make_student_uid()

                ext = os.path.splitext(
                    photo.filename
                )[1].lower()

                if ext not in [".jpg", ".jpeg", ".png", ".webp"]:

                    conn.close()

                    message = "Use JPG, JPEG, PNG or WEBP."
                    message_class = "message error"

                else:

                    filename = uid + ext

                    photo.save(
                        os.path.join(
                            PHOTO_DIR,
                            filename
                        )
                    )

                    conn.execute("""
                        INSERT INTO students
                        (
                            student_uid,
                            name,
                            roll_number,
                            branch,
                            hostel_room,
                            gender,
                            hostel_name,
                            hostel_block,
                            photo_filename,
                            active,
                            created_at
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)
                    """, (
                        uid,
                        name,
                        roll,
                        branch,
                        room,
                        gender,
                        hostel_name,
                        hostel_block,
                        filename,
                        current_time().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )
                    ))

                    conn.commit()
                    conn.close()

                    message = (
                        "Student registered successfully. "
                        f"Student ID: {uid}"
                    )

                    message_class = "message success"

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Register Student</title>
        {CSS}
    </head>

    <body>
    <div class="container">

        <div class="nav">
            <a class="btn" href="/admin/dashboard">
                Dashboard
            </a>
            <a class="btn blue" href="/admin/students">
                Students
            </a>
        </div>

        <div class="card">

            <h1>👨‍🎓 Register Hostel Student</h1>

            {'<div class="' + message_class + '">' +
             message + '</div>' if message else ''}

            <form
                method="POST"
                enctype="multipart/form-data"
            >

                <label>Student Name</label>

                <input
                    type="text"
                    name="name"
                    required
                >

                <label>Registration Number</label>

                <input
                    type="text"
                    name="roll_number"
                    required
                >

                <label>Branch</label>

                <input
                    type="text"
                    name="branch"
                    placeholder="AI & ML / Electrical / etc."
                    required
                >

                <label>Hostel Room</label>

                <input
                    type="text"
                    name="hostel_room"
                    required
                >

                <label>Gender</label>
                <select name="gender" required>
                    <option value="">Select Boy/Girl</option>
                    <option value="BOY">Boy</option>
                    <option value="GIRL">Girl</option>
                </select>

                <label>Hostel</label>
                <select name="hostel_name" required>
                    <option value="">Select Hostel</option>
                    <option value="Boys Hostel">Boys Hostel</option>
                    <option value="Girls Hostel">Girls Hostel</option>
                </select>

                <label>Hostel Block</label>
                <input type="text" name="hostel_block" placeholder="Example: A Block">

                <label>Student Photo</label>

                <input
                    type="file"
                    name="photo"
                    accept="image/*"
                    required
                >

                <button class="btn green" type="submit">
                    Register Student
                </button>

            </form>

        </div>

    </div>
    </body>
    </html>
    """

    return html


# =========================================================
# ADMIN - STUDENT LIST
# =========================================================

@app.route("/admin/students")
def admin_students():

    if not admin_required():
        return redirect(url_for("admin_login"))

    search = request.args.get("q", "").strip()
    gender_filter = request.args.get("gender", "").strip().upper()
    hostel_filter = request.args.get("hostel", "").strip()
    where, params = ["1=1"], []
    if search:
        where.append("(name LIKE ? OR roll_number LIKE ? OR branch LIKE ?)")
        term = f"%{search}%"
        params.extend([term, term, term])
    if gender_filter in ["BOY", "GIRL"]:
        where.append("gender = ?")
        params.append(gender_filter)
    if hostel_filter:
        where.append("hostel_name = ?")
        params.append(hostel_filter)

    conn = db()
    students = conn.execute(
        "SELECT * FROM students WHERE " + " AND ".join(where) + " ORDER BY name",
        params
    ).fetchall()

    conn.close()

    rows = ""

    for student in students:

        photo = "No Photo"

        if student["photo_filename"]:

            photo = f"""
                <img
                    class="photo"
                    src="/student-photo/{student['id']}"
                >
            """

        status = (
            '<span class="badge active-badge">ACTIVE</span>'
            if student["active"]
            else
            '<span class="badge expired-badge">INACTIVE</span>'
        )

        rows += f"""
        <tr>

            <td>{photo}</td>
            <td>{escape(student["name"])}</td>
            <td>{escape(student["roll_number"])}</td>
            <td>{escape(student["gender"] or "NOT SET")}</td>
            <td>{escape(student["branch"])}</td>
            <td>{escape(student["hostel_name"] or "NOT SET")}<br><small>{escape(student["hostel_block"] or "")}</small></td>
            <td>{escape(student["hostel_room"])}</td>
            <td>{student["student_uid"]}</td>
            <td>{status}</td>
            <td><a class="btn blue" href="/admin/student/{student['id']}/edit">Edit</a>
                <form method="POST" action="/admin/student/{student['id']}/toggle" style="display:inline">
                  <button class="btn {'red' if student['active'] else 'green'}" type="submit">{'Inactive' if student['active'] else 'Active'}</button>
                </form></td>

        </tr>
        """

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Students</title>
        {CSS}
    </head>

    <body>

    <div class="container">

        <div class="nav">

            <a class="btn" href="/admin/dashboard">
                Dashboard
            </a>

            <a class="btn green" href="/admin/add-student">
                Register Student
            </a>

        </div>

        <div class="card">

            <h1>👨‍🎓 Hostel Students</h1>

            <form method="GET" class="grid" style="align-items:end">
              <div><label>Search</label><input name="q" value="{escape(search)}" placeholder="Name / Registration No. / Branch"></div>
              <div><label>Gender</label><select name="gender"><option value="">All</option><option value="BOY" {'selected' if gender_filter == 'BOY' else ''}>Boys</option><option value="GIRL" {'selected' if gender_filter == 'GIRL' else ''}>Girls</option></select></div>
              <div><label>Hostel</label><select name="hostel"><option value="">All</option><option value="Boys Hostel" {'selected' if hostel_filter == 'Boys Hostel' else ''}>Boys Hostel</option><option value="Girls Hostel" {'selected' if hostel_filter == 'Girls Hostel' else ''}>Girls Hostel</option></select></div>
              <div><button class="btn blue" type="submit">Apply Filter</button> <a class="btn gray" href="/admin/students">Clear</a></div>
            </form>

            <div style="overflow-x:auto;">

                <table>

                    <tr>
                        <th>Photo</th>
                        <th>Name</th>
                        <th>Registration No.</th>
                        <th>Gender</th>
                        <th>Branch</th>
                        <th>Hostel/Block</th>
                        <th>Room</th>
                        <th>Student ID</th>
                        <th>Status</th>
                        <th>Action</th>
                    </tr>

                    {rows}

                </table>

            </div>

        </div>

    </div>

    </body>
    </html>
    """

    return html


@app.route("/admin/student/<int:student_id>/edit", methods=["GET", "POST"])
def admin_edit_student(student_id):
    if not admin_required():
        return redirect(url_for("admin_login"))

    conn = db()
    student = conn.execute("SELECT * FROM students WHERE id = ?", (student_id,)).fetchone()
    if not student:
        conn.close()
        return "Student not found", 404

    message = ""
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        registration = request.form.get("registration_number", "").strip()
        branch = request.form.get("branch", "").strip()
        room = request.form.get("hostel_room", "").strip()
        gender = request.form.get("gender", "").strip().upper()
        hostel_name = request.form.get("hostel_name", "").strip()
        hostel_block = request.form.get("hostel_block", "").strip()
        photo = request.files.get("photo")
        if not name or not registration or not branch or not room or gender not in ["BOY", "GIRL"] or not hostel_name:
            message = "Please fill all required fields."
        else:
            duplicate = conn.execute("SELECT id FROM students WHERE roll_number = ? AND id != ?", (registration, student_id)).fetchone()
            if duplicate:
                message = "Registration Number already exists."
            else:
                photo_filename = student["photo_filename"]
                if photo and photo.filename:
                    ext = os.path.splitext(photo.filename)[1].lower()
                    if ext not in [".jpg", ".jpeg", ".png", ".webp"]:
                        message = "Use JPG, JPEG, PNG or WEBP photo."
                    else:
                        photo_filename = student["student_uid"] + ext
                        photo.save(os.path.join(PHOTO_DIR, photo_filename))
                if not message:
                    conn.execute("""UPDATE students SET name=?, roll_number=?, branch=?, hostel_room=?,
                                 gender=?, hostel_name=?, hostel_block=?, photo_filename=? WHERE id=?""",
                                 (name, registration, branch, room, gender, hostel_name, hostel_block, photo_filename, student_id))
                    conn.commit()
                    conn.close()
                    return redirect(url_for("admin_students"))

    student = conn.execute("SELECT * FROM students WHERE id = ?", (student_id,)).fetchone()
    conn.close()
    html = f"""<!DOCTYPE html><html><head><title>Edit Student</title>{CSS}</head><body><div class="container">
    <div class="nav"><a class="btn" href="/admin/students">Back to Students</a></div>
    <div class="card"><h1>✏️ Edit Student</h1>{f'<div class="message error">{escape(message)}</div>' if message else ''}
    <form method="POST" enctype="multipart/form-data">
      <label>Name</label><input name="name" value="{escape(student['name'])}" required>
      <label>Registration Number</label><input name="registration_number" value="{escape(student['roll_number'])}" required>
      <label>Branch</label><input name="branch" value="{escape(student['branch'])}" required>
      <label>Gender</label><select name="gender" required><option value="BOY" {'selected' if student['gender']=='BOY' else ''}>Boy</option><option value="GIRL" {'selected' if student['gender']=='GIRL' else ''}>Girl</option></select>
      <label>Hostel</label><select name="hostel_name" required><option value="Boys Hostel" {'selected' if student['hostel_name']=='Boys Hostel' else ''}>Boys Hostel</option><option value="Girls Hostel" {'selected' if student['hostel_name']=='Girls Hostel' else ''}>Girls Hostel</option></select>
      <label>Hostel Block</label><input name="hostel_block" value="{escape(student['hostel_block'] or '')}">
      <label>Room Number</label><input name="hostel_room" value="{escape(student['hostel_room'])}" required>
      <label>Change Photo (optional)</label><input type="file" name="photo" accept="image/*">
      <button class="btn green" type="submit">Save Changes</button>
    </form></div></div></body></html>"""
    return html


@app.route("/admin/student/<int:student_id>/toggle", methods=["POST"])
def admin_toggle_student(student_id):
    if not admin_required():
        return redirect(url_for("admin_login"))
    conn = db()
    conn.execute("UPDATE students SET active = CASE WHEN active = 1 THEN 0 ELSE 1 END WHERE id = ?", (student_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("admin_students"))


# =========================================================
# STUDENT PHOTO
# =========================================================

@app.route("/student-photo/<int:student_id>")
def student_photo(student_id):

    conn = db()

    student = conn.execute("""
        SELECT photo_filename
        FROM students
        WHERE id = ?
    """, (student_id,)).fetchone()

    conn.close()

    if not student or not student["photo_filename"]:
        return "Photo not found", 404

    return send_from_directory(
        PHOTO_DIR,
        student["photo_filename"]
    )


# =========================================================
# STUDENT PANEL
# =========================================================

@app.route("/student")
def student_home():

    # Student session is not required permanently.
    # Registration number verification is done before coupon generation.

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>

        <title>Student Panel</title>

        {CSS}

    </head>

    <body>

    <div class="container">

        <div class="card" style="max-width:650px;margin:40px auto;">

            <div class="center">

                <h1>🍽️ Smart Mess Student Panel</h1>

                <p>
                    Hostel student apna Registration Number enter kare.
                </p>

            </div>

            <form method="POST" action="/student/verify">

                <label>Registration Number</label>

                <input
                    type="text"
                    name="roll_number"
                    placeholder="Enter your Registration Number"
                    required
                >

                <button class="btn blue" type="submit">
                    Verify Student
                </button>

            </form>

            <hr style="margin:25px 0;">

            <p class="center">
                Outside / unregistered student ko coupon
                generate nahi hoga.
            </p>

        </div>

    </div>

    </body>
    </html>
    """

    return html


@app.route("/student/verify", methods=["POST"])
def student_verify():

    roll = request.form.get(
        "roll_number",
        ""
    ).strip()

    conn = db()

    student = conn.execute("""
        SELECT *
        FROM students
        WHERE roll_number = ?
        AND active = 1
    """, (roll,)).fetchone()

    conn.close()

    if not student:

        return f"""
        <!DOCTYPE html>
        <html>
        <head>{CSS}</head>
        <body>

        <div class="container">

            <div class="card center">

                <h1>❌ Student Not Found</h1>

                <div class="message error">
                    This student is not registered as an active hostel student.
                </div>

                <a class="btn" href="/student">
                    Try Again
                </a>

                </div>

        </div>

        </body>
        </html>
        """

    session["student_uid"] = student["student_uid"]

    return redirect(
        url_for(
            "student_meals"
        )
    )


# =========================================================
# STUDENT MEAL PAGE
# =========================================================

@app.route("/student/meals")
def student_meals():

    student_uid = session.get("student_uid")

    if not student_uid:
        return redirect(url_for("student_home"))

    conn = db()

    student = conn.execute("""
        SELECT *
        FROM students
        WHERE student_uid = ?
        AND active = 1
    """, (student_uid,)).fetchone()

    conn.close()

    if not student:

        session.pop("student_uid", None)

        return redirect(
            url_for("student_home")
        )

    photo = ""

    if student["photo_filename"]:

        photo = image_data_uri(
            student["photo_filename"]
        )

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Student Meals</title>
        {CSS}
    </head>

    <body>

    <div class="container">

        <div class="card center">

            {(
                '<img class="big-photo" src="' +
                photo +
                '">'
            ) if photo else ''}

            <h1>
                Hello, {student["name"]} 👋
            </h1>

            <p>
                Registration No.: <b>{student["roll_number"]}</b>
            </p>

            <p>
                Branch: <b>{student["branch"]}</b>
            </p>

            <p>
                Room: <b>{student["hostel_room"]}</b>
            </p>

            <p>Gender: <b>{student["gender"] or "NOT SET"}</b></p>
            <p>Hostel: <b>{student["hostel_name"] or "NOT SET"} {student["hostel_block"] or ""}</b></p>

        </div>


        <div class="card">

            <h2 class="center">
                Select Meal
            </h2>

            <div class="meal-buttons">

                <form
                    method="POST"
                    action="/student/generate"
                >

                    <input
                        type="hidden"
                        name="meal"
                        value="BREAKFAST"
                    >

                    <button type="submit">
                        🌅 Breakfast
                    </button>

                </form>


                <form
                    method="POST"
                    action="/student/generate"
                >

                    <input
                        type="hidden"
                        name="meal"
                        value="LUNCH"
                    >

                    <button type="submit">
                        ☀️ Lunch
                    </button>

                </form>


                <form
                    method="POST"
                    action="/student/generate"
                >

                    <input
                        type="hidden"
                        name="meal"
                        value="DINNER"
                    >

                    <button type="submit">
                        🌙 Dinner
                    </button>

                </form>

            </div>

        </div>


        <div class="card center">

            <a
                class="btn red"
                href="/student/logout"
            >
                Student Logout
            </a>

        </div>

    </div>

    </body>
    </html>
    """

    return html


# =========================================================
# STUDENT GENERATE COUPON
# =========================================================

@app.route("/student/generate", methods=["POST"])
def student_generate():

    student_uid = session.get("student_uid")

    if not student_uid:

        return redirect(
            url_for("student_home")
        )

    meal = request.form.get(
        "meal",
        ""
    ).upper()

    if meal not in [
        "BREAKFAST",
        "LUNCH",
        "DINNER"
    ]:

        return "Invalid meal.", 400

    conn = db()

    student = conn.execute("""
        SELECT *
        FROM students
        WHERE student_uid = ?
        AND active = 1
    """, (student_uid,)).fetchone()

    if not student:

        conn.close()

        return redirect(
            url_for("student_home")
        )

    # -----------------------------------------------------
    # Check today's same meal
    # -----------------------------------------------------

    today = current_time().strftime("%Y-%m-%d")

    existing = conn.execute("""
        SELECT *
        FROM coupons
        WHERE student_uid = ?
        AND meal = ?
        AND substr(generated_at, 1, 10) = ?
        ORDER BY id DESC
        LIMIT 1
    """, (
        student_uid,
        meal,
        today
    )).fetchone()

    if existing:

        if existing["status"] == "USED":

            conn.close()

            return f"""
            <!DOCTYPE html>
            <html>
            <head>{CSS}</head>
            <body>

            <div class="container">

                <div class="card center">

                    <h1>❌ Already Used</h1>

                    <div class="message error">

                        Your {meal.title()} coupon
                        has already been used today.

                    </div>

                    <a
                        class="btn"
                        href="/student/meals"
                    >
                        Back
                    </a>

                </div>

            </div>

            </body>
            </html>
            """

        if existing["status"] == "ACTIVE":

            expiry = datetime.strptime(
                existing["expires_at"],
                "%Y-%m-%d %H:%M:%S"
            )

            if current_time() < expiry:

                conn.close()

                # Show existing active coupon.
                return render_coupon(
                    student,
                    existing
                )

            else:

                conn.execute("""
                    UPDATE coupons
                    SET status = 'EXPIRED'
                    WHERE id = ?
                """, (existing["id"],))

                conn.commit()

    # -----------------------------------------------------
    # Create new coupon
    # -----------------------------------------------------

    generated = current_time()

    expires = generated + timedelta(minutes=5)

    token = make_coupon_token()

    conn.execute("""
        INSERT INTO coupons
        (
            token,
            student_uid,
            meal,
            generated_at,
            expires_at,
            status
        )
        VALUES (?, ?, ?, ?, ?, 'ACTIVE')
    """, (
        token,
        student_uid,
        meal,
        generated.strftime(
            "%Y-%m-%d %H:%M:%S"
        ),
        expires.strftime(
            "%Y-%m-%d %H:%M:%S"
        ),
    ))

    conn.commit()

    coupon = conn.execute("""
        SELECT *
        FROM coupons
        WHERE token = ?
    """, (token,)).fetchone()

    conn.close()

    return render_coupon(
        student,
        coupon
    )


# =========================================================
# COUPON HTML
# =========================================================

def render_coupon(student, coupon):

    photo = ""

    if student["photo_filename"]:

        photo = image_data_uri(
            student["photo_filename"]
        )

    qr = qr_data_uri(
        coupon["token"]
    )

    expires = datetime.strptime(
        coupon["expires_at"],
        "%Y-%m-%d %H:%M:%S"
    )

    expires_iso = expires.isoformat()

    meal_name = coupon["meal"].title()

    html = f"""
    <!DOCTYPE html>
    <html>

    <head>

        <title>Mess Coupon</title>

        {CSS}

    </head>

    <body>

    <div class="container">

        <div class="card center">

            <h1>🎫 Meal Coupon</h1>

            {(
                '<img class="big-photo" src="' +
                photo +
                '">'
            ) if photo else ''}

            <h2>
                {student["name"]}
            </h2>

            <p>
                Registration Number:
                <b>{student["roll_number"]}</b>
            </p>

            <p>
                Branch:
                <b>{student["branch"]}</b>
            </p>

            <p>
                Hostel Room:
                <b>{student["hostel_room"]}</b>
            </p>

            <p>Gender: <b>{student["gender"] or "NOT SET"}</b></p>
            <p>Hostel: <b>{student["hostel_name"] or "NOT SET"} {student["hostel_block"] or ""}</b></p>

            <h2>
                🍽️ {meal_name}
            </h2>

            <img
                class="qr"
                src="{qr}"
                alt="Secure Coupon QR"
            >

            <p>
                Show this QR at the mess scanner.
            </p>

            <div
                id="timer"
                class="countdown"
            >
                05:00
            </div>

            <p>
                Valid for 5 minutes only.
            </p>

            <p>
                After successful scan this coupon
                becomes USED.
            </p>

        </div>

    </div>


    <script>

    const expiry =
        new Date(
            "{expires_iso}"
        ).getTime();

    const timerElement =
        document.getElementById("timer");

    const interval =
        setInterval(function() {{

            const now =
                new Date().getTime();

            const distance =
                expiry - now;

            if (distance <= 0) {{

                clearInterval(interval);

                timerElement.innerText =
                    "EXPIRED";

                timerElement.style.color =
                    "red";

                return;

            }}

            const minutes =
                Math.floor(
                    distance / 60000
                );

            const seconds =
                Math.floor(
                    (distance % 60000) / 1000
                );

            timerElement.innerText =
                String(minutes).padStart(
                    2, "0"
                )
                + ":"
                +
                String(seconds).padStart(
                    2, "0"
                );

        }}, 1000);

    </script>

    </body>
    </html>
    """

    return render_template_string(html)


# =========================================================
# STUDENT LOGOUT
# =========================================================

@app.route("/student/logout")
def student_logout():

    session.pop("student_uid", None)

    return redirect(
        url_for("student_home")
    )


# =========================================================
# ADMIN SCANNER
# =========================================================

# =========================================================
# ADMIN QR SCANNER
# =========================================================

@app.route("/admin/scanner")
def admin_scanner():

    if not admin_required():
        return redirect(url_for("admin_login"))

    html = f"""
    <!DOCTYPE html>
    <html>

    <head>

        <title>Mess QR Scanner</title>

        {CSS}

        <script src="https://unpkg.com/html5-qrcode"></script>

        <style>
            #reader {{
                width: 100%;
                max-width: 600px;
                margin: 20px auto;
            }}

            #result {{
                margin-top: 20px;
            }}

            .big-photo {{
                width: 150px;
                height: 150px;
                object-fit: cover;
                border-radius: 15px;
                display: block;
                margin: 15px auto;
            }}
        </style>

    </head>

    <body>

    <div class="container">

        <div class="nav">

            <a class="btn"
               href="/admin/dashboard">
                Dashboard
            </a>

            <a class="btn gray"
               href="/admin/records">
                📊 Records
            </a>

            <a class="btn red"
               href="/admin/logout">
                Logout
            </a>

        </div>

        <div class="card">

            <h1>📷 Mess QR Scanner</h1>

            <p>
                Allow camera permission and scan the student's coupon QR.
            </p>

            <div id="reader"></div>

            <div id="result"></div>

        </div>

    </div>

    <script>

    let processing = false;

    async function verifyCoupon(token) {{

        if (processing) {{
            return;
        }}

        processing = true;

        try {{

            const response = await fetch(
                "/admin/verify-coupon",
                {{
                    method: "POST",
                    headers: {{
                        "Content-Type": "application/json"
                    }},
                    body: JSON.stringify({{
                        token: token
                    }})
                }}
            );

            const data = await response.json();

            if (data.success) {{

                const photo =
                    data.photo
                    ?
                    "<img class='big-photo' src='"
                    + data.photo
                    + "'>"
                    :
                    "";

                document.getElementById(
                    "result"
                ).innerHTML =

                    "<div class='message success'>"

                    + "<h2>✅ COUPON VALID</h2>"

                    + photo

                    + "<h2>"
                    + data.name
                    + "</h2>"

                    + "<p>Registration No.: "
                    + data.roll
                    + "</p>"

                    + "<p>Branch: "
                    + data.branch
                    + "</p>"

                    + "<p>Room: "
                    + data.room
                    + "</p>"

                    + "<p>Gender: " + (data.gender || "NOT SET") + "</p>"
                    + "<p>Hostel: " + (data.hostel || "NOT SET") + " " + (data.block || "") + "</p>"

                    + "<p>Meal: <b>"
                    + data.meal
                    + "</b></p>"

                    + "<p>Coupon is now USED.</p>"

                    + "</div>";

            }} else {{

                document.getElementById(
                    "result"
                ).innerHTML =

                    "<div class='message error'>"

                    + "<h2>❌ COUPON REJECTED</h2>"

                    + "<p>"
                    + data.message
                    + "</p>"

                    + "</div>";
            }}

        }} catch (error) {{

            document.getElementById(
                "result"
            ).innerHTML =

                "<div class='message error'>"

                + "<h2>❌ Scanner Error</h2>"

                + "<p>"
                + error
                + "</p>"

                + "</div>";
        }}

        setTimeout(
            function() {{
                processing = false;
            }},
            2000
        );

    }}

    function scanSuccess(
        decodedText,
        decodedResult
    ) {{

        verifyCoupon(decodedText);

    }}

    function scanFailure(error) {{

        // Ignore scan errors

    }}

    const scanner =
        new Html5QrcodeScanner(
            "reader",
            {{
                fps: 10,
                qrbox: {{
                    width: 250,
                    height: 250
                }},
                rememberLastUsedCamera: true
            }},
            false
        );

    scanner.render(
        scanSuccess,
        scanFailure
    );

    </script>

    </body>
    </html>
    """

    return html
    # =====================================================
    # SAVE SUCCESSFUL SCAN TO GOOGLE SHEET
    # =====================================================

    try:

        GOOGLE_SHEET_URL = (
            "https://script.google.com/macros/s/"
            "AKfycbzRjb3Xh_O95l9N18VJKhUVs6-4qb99Ybr60zmyxIp-amMOtBPnFzArpyJo2tlyl1zE/exec"
        )

        sheet_data = {
            "date": current_time().strftime("%d-%m-%Y"),
            "time": current_time().strftime("%H:%M:%S"),
            "name": student["name"],
            "roll": student["roll_number"],
            "branch": student["branch"],
            "room": student["hostel_room"],
            "meal": coupon["meal"],
            "status": "USED"
        }

        response = requests.post(
            GOOGLE_SHEET_URL,
            json=sheet_data,
            timeout=10
        )

        print(
            "Google Sheet:",
            response.text
        )

    except Exception as e:

        print(
            "Google Sheet Error:",
            e
        )

    conn.close()

    photo_url = ""

    if student["photo_filename"]:
        photo_url = (
            "/student-photo/"
            + str(student["id"])
        )

    return jsonify({
        "success": True,
        "name": student["name"],
        "roll": student["roll_number"],
        "branch": student["branch"],
        "room": student["hostel_room"],
        "meal": coupon["meal"],
        "photo": photo_url
    })

# =========================================================
# ADMIN RECORDS
# =========================================================

@app.route("/admin/records")
def admin_records():

    if not admin_required():
        return redirect(
            url_for("admin_login")
        )

    conn = db()

    records = conn.execute("""
        SELECT
            coupons.*,
            students.name,
            students.roll_number,
            students.branch,
            students.hostel_room
        FROM coupons
        LEFT JOIN students
        ON coupons.student_uid =
           students.student_uid
        ORDER BY coupons.id DESC
    """).fetchall()

    conn.close()

    rows = ""

    for record in records:

        if record["status"] == "USED":

            badge = (
                '<span class="badge used-badge">'
                'USED</span>'
            )

        elif record["status"] == "EXPIRED":

            badge = (
                '<span class="badge expired-badge">'
                'EXPIRED</span>'
            )

        else:

            badge = (
                '<span class="badge active-badge">'
                'ACTIVE</span>'
            )

        rows += f"""
        <tr>

            <td>{record["name"] or "-"}</td>

            <td>{record["roll_number"] or "-"}</td>

            <td>{record["meal"]}</td>

            <td>{record["generated_at"]}</td>

            <td>{record["expires_at"]}</td>

            <td>{record["used_at"] or "-"}</td>

            <td>{badge}</td>

        </tr>
        """

    html = f"""
    <!DOCTYPE html>
    <html>

    <head>

        <title>Records</title>

        {CSS}

    </head>

    <body>

    <div class="container">

        <div class="nav">

            <a
                class="btn"
                href="/admin/dashboard"
            >
                Dashboard
            </a>

        </div>

        <div class="card">

            <h1>📊 Mess Records</h1>

            <div style="overflow-x:auto;">

                <table>

                    <tr>
                        <th>Name</th>
                        <th>Registration No.</th>
                        <th>Meal</th>
                        <th>Generated</th>
                        <th>Expires</th>
                        <th>Used At</th>
                        <th>Status</th>
                    </tr>

                    {rows}

                </table>

            </div>

        </div>

    </div>

    </body>

    </html>
    """

    return html


# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )
