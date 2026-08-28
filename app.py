from flask import (
    Flask,
    request,
    redirect,
    url_for,
    session,
    render_template_string,
    jsonify,
    send_from_directory,
    send_file,
    Response
)

import sqlite3
import os
import secrets
import string
import base64
import hmac
import re
import json
from io import BytesIO
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import qrcode
import requests
from markupsafe import escape
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

try:
    from pymongo import MongoClient, ASCENDING, ReturnDocument
    from pymongo.errors import DuplicateKeyError
except ImportError:  # Local SQLite mode remains available.
    MongoClient = None
    ASCENDING = 1
    ReturnDocument = None
    DuplicateKeyError = Exception

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
BRANCHES = [
    "AI & ML",
    "Civil (Construction Technology)",
    "Electronics (Robotics)",
    "Mechanical (CAD/CAM)",
]
HOSTEL_BLOCKS = ["BH-1", "BH-2"]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.environ.get("DATABASE_PATH", os.path.join(BASE_DIR, "smart_mess.db"))
MONGODB_URI = os.environ.get("MONGODB_URI", "").strip()
MONGODB_DB = os.environ.get("MONGODB_DB", "smartmess").strip() or "smartmess"
USE_MONGO = MONGODB_URI.startswith("mongodb")
PHOTO_DIR = os.environ.get("PHOTO_DIR", os.path.join(BASE_DIR, "photos"))

os.makedirs(PHOTO_DIR, exist_ok=True)


# =========================================================
# DATABASE
# =========================================================

def db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


mongo_client = None
mongo_database = None
if USE_MONGO:
    if MongoClient is None:
        raise RuntimeError("pymongo is not installed")
    mongo_client = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=10000)
    mongo_database = mongo_client[MONGODB_DB]


def next_mongo_id(sequence_name):
    result = mongo_database.counters.find_one_and_update(
        {"_id": sequence_name}, {"$inc": {"value": 1}}, upsert=True,
        return_document=ReturnDocument.AFTER
    )
    return result["value"]


def init_db():
    defaults = {
        "breakfast_open": "1", "breakfast_start": "07:00", "breakfast_end": "09:00",
        "lunch_open": "1", "lunch_start": "12:00", "lunch_end": "14:30",
        "dinner_open": "1", "dinner_start": "19:00", "dinner_end": "22:00",
        "breakfast_menu": "", "lunch_menu": "", "dinner_menu": "",
        "student_notice": "Welcome to SmartMess.",
    }

    if USE_MONGO:
        mongo_client.admin.command("ping")
        mongo_database.students.create_index([("student_uid", ASCENDING)], unique=True)
        mongo_database.students.create_index([("roll_number", ASCENDING)], unique=True)
        mongo_database.coupons.create_index([("token", ASCENDING)], unique=True)
        mongo_database.coupons.create_index([("student_uid", ASCENDING), ("meal", ASCENDING), ("generated_at", ASCENDING)])
        mongo_database.skipped_meals.create_index([("student_uid", ASCENDING), ("meal", ASCENDING), ("skip_date", ASCENDING)], unique=True)
        mongo_database.admins.create_index([("username", ASCENDING)], unique=True)
        for key, value in defaults.items():
            mongo_database.settings.update_one({"_id": key}, {"$setOnInsert": {"value": value}}, upsert=True)
        for student in mongo_database.students.find({"$or": [{"pin_hash": {"$exists": False}}, {"pin_hash": ""}]}):
            temporary_pin = str(student["roll_number"])[-4:].zfill(4)
            mongo_database.students.update_one({"_id": student["_id"]}, {"$set": {"pin_hash": generate_password_hash(temporary_pin)}})
        return

    conn = db()
    conn.execute("""CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT, student_uid TEXT UNIQUE NOT NULL,
        name TEXT NOT NULL, roll_number TEXT UNIQUE NOT NULL, branch TEXT NOT NULL,
        hostel_room TEXT NOT NULL, photo_filename TEXT, photo_data BLOB, photo_mime TEXT,
        pin_hash TEXT, gender TEXT DEFAULT 'NOT SET', hostel_name TEXT DEFAULT 'NOT SET',
        hostel_block TEXT DEFAULT '', active INTEGER DEFAULT 1, created_at TEXT NOT NULL
    )""")
    columns = {row[1] for row in conn.execute("PRAGMA table_info(students)")}
    for column, definition in {
        "gender": "TEXT DEFAULT 'NOT SET'", "hostel_name": "TEXT DEFAULT 'NOT SET'",
        "hostel_block": "TEXT DEFAULT ''", "photo_data": "BLOB",
        "photo_mime": "TEXT", "pin_hash": "TEXT",
    }.items():
        if column not in columns:
            conn.execute(f"ALTER TABLE students ADD COLUMN {column} {definition}")
    conn.execute("""CREATE TABLE IF NOT EXISTS coupons (
        id INTEGER PRIMARY KEY AUTOINCREMENT, token TEXT UNIQUE NOT NULL,
        student_uid TEXT NOT NULL, meal TEXT NOT NULL, generated_at TEXT NOT NULL,
        expires_at TEXT NOT NULL, used_at TEXT, status TEXT NOT NULL DEFAULT 'ACTIVE'
    )""")
    conn.execute("CREATE TABLE IF NOT EXISTS settings (setting_key TEXT PRIMARY KEY, setting_value TEXT NOT NULL)")
    conn.execute("""CREATE TABLE IF NOT EXISTS skipped_meals (
        id INTEGER PRIMARY KEY AUTOINCREMENT, student_uid TEXT NOT NULL, meal TEXT NOT NULL,
        skip_date TEXT NOT NULL, created_at TEXT NOT NULL, UNIQUE(student_uid, meal, skip_date)
    )""")
    conn.execute("""CREATE TABLE IF NOT EXISTS admins (
        id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL, role TEXT NOT NULL, active INTEGER DEFAULT 1,
        created_at TEXT NOT NULL
    )""")
    for key, value in defaults.items():
        conn.execute("INSERT OR IGNORE INTO settings (setting_key, setting_value) VALUES (?, ?)", (key, value))
    for student in conn.execute("SELECT id, roll_number FROM students WHERE pin_hash IS NULL OR pin_hash = ''"):
        temporary_pin = str(student["roll_number"])[-4:].zfill(4)
        conn.execute("UPDATE students SET pin_hash = ? WHERE id = ?", (generate_password_hash(temporary_pin), student["id"]))
    for student in conn.execute("SELECT id, photo_filename FROM students WHERE photo_data IS NULL AND photo_filename IS NOT NULL"):
        path = os.path.join(PHOTO_DIR, student["photo_filename"])
        if os.path.exists(path):
            ext = os.path.splitext(path)[1].lower()
            mime = {".png": "image/png", ".webp": "image/webp"}.get(ext, "image/jpeg")
            with open(path, "rb") as photo_file:
                conn.execute("UPDATE students SET photo_data = ?, photo_mime = ? WHERE id = ?", (photo_file.read(), mime, student["id"]))

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


def admin_required(roles=None):
    if session.get("admin") is not True:
        return False
    if roles is None:
        return True
    return session.get("admin_role", "MAIN") in roles


def admin_scope_gender():
    role = session.get("admin_role", "MAIN")
    return "BOY" if role == "BOYS" else "GIRL" if role == "GIRLS" else ""


def row_dict(row):
    return dict(row) if row is not None else None


def get_setting(key, default=""):
    if USE_MONGO:
        item = mongo_database.settings.find_one({"_id": key})
        return item.get("value", default) if item else default
    conn = db()
    row = conn.execute("SELECT setting_value FROM settings WHERE setting_key = ?", (key,)).fetchone()
    conn.close()
    return row["setting_value"] if row else default


def save_settings(values):
    if USE_MONGO:
        for key, value in values.items():
            mongo_database.settings.update_one({"_id": key}, {"$set": {"value": value}}, upsert=True)
        return
    conn = db()
    for key, value in values.items():
        conn.execute("INSERT OR REPLACE INTO settings (setting_key, setting_value) VALUES (?, ?)", (key, value))
    conn.commit()
    conn.close()


def meal_is_available(meal):
    prefix = meal.lower()
    if get_setting(prefix + "_open", "1") != "1":
        return False, f"{meal.title()} is closed by admin."
    now = current_time().strftime("%H:%M")
    start = get_setting(prefix + "_start")
    end = get_setting(prefix + "_end")
    if start and end and not (start <= now <= end):
        return False, f"{meal.title()} coupon time is {start} to {end}."
    return True, ""


def student_by_id(student_id):
    if USE_MONGO:
        return mongo_database.students.find_one({"id": student_id})
    conn = db(); row = conn.execute("SELECT * FROM students WHERE id = ?", (student_id,)).fetchone(); conn.close()
    return row_dict(row)


def student_by_roll(registration, active_only=False):
    if USE_MONGO:
        query = {"roll_number": registration}
        if active_only: query["active"] = 1
        return mongo_database.students.find_one(query)
    conn = db(); sql = "SELECT * FROM students WHERE roll_number = ?" + (" AND active = 1" if active_only else "")
    row = conn.execute(sql, (registration,)).fetchone(); conn.close(); return row_dict(row)


def student_by_uid(uid, active_only=False):
    if USE_MONGO:
        query = {"student_uid": uid}
        if active_only: query["active"] = 1
        return mongo_database.students.find_one(query)
    conn = db(); sql = "SELECT * FROM students WHERE student_uid = ?" + (" AND active = 1" if active_only else "")
    row = conn.execute(sql, (uid,)).fetchone(); conn.close(); return row_dict(row)


def count_students(extra=None):
    extra = extra or {}
    if USE_MONGO:
        return mongo_database.students.count_documents({"active": 1, **extra})
    where, params = ["active = 1"], []
    for key, value in extra.items(): where.append(f"{key} = ?"); params.append(value)
    conn = db(); count = conn.execute("SELECT COUNT(*) AS c FROM students WHERE " + " AND ".join(where), params).fetchone()["c"]; conn.close()
    return count


def add_student_record(data):
    if USE_MONGO:
        data = dict(data); data["id"] = next_mongo_id("students")
        mongo_database.students.insert_one(data); return data["id"]
    columns = list(data); placeholders = ",".join("?" for _ in columns)
    conn = db(); cursor = conn.execute(f"INSERT INTO students ({','.join(columns)}) VALUES ({placeholders})", [data[c] for c in columns]); conn.commit(); new_id = cursor.lastrowid; conn.close(); return new_id


def update_student_record(student_id, changes):
    if USE_MONGO:
        return mongo_database.students.update_one({"id": student_id}, {"$set": changes}).modified_count
    assignments = ", ".join(f"{key} = ?" for key in changes)
    conn = db(); cur = conn.execute(f"UPDATE students SET {assignments} WHERE id = ?", [*changes.values(), student_id]); conn.commit(); count = cur.rowcount; conn.close(); return count


def toggle_student_record(student_id):
    student = student_by_id(student_id)
    if student: update_student_record(student_id, {"active": 0 if student["active"] else 1})


def list_student_records(search="", gender="", hostel="", branch="", block=""):
    filters = {k: v for k, v in {"gender": gender, "hostel_name": hostel, "branch": branch, "hostel_block": block}.items() if v}
    if USE_MONGO:
        query = dict(filters)
        if search:
            safe = re.escape(search); query["$or"] = [{field: {"$regex": safe, "$options": "i"}} for field in ["name", "roll_number", "branch"]]
        return list(mongo_database.students.find(query).sort("name", ASCENDING))
    where, params = ["1=1"], []
    if search:
        where.append("(name LIKE ? OR roll_number LIKE ? OR branch LIKE ?)"); term = f"%{search}%"; params.extend([term, term, term])
    for key, value in filters.items(): where.append(f"{key} = ?"); params.append(value)
    conn = db(); rows = conn.execute("SELECT * FROM students WHERE " + " AND ".join(where) + " ORDER BY name", params).fetchall(); conn.close()
    return [dict(row) for row in rows]


def coupon_by_token(token):
    if USE_MONGO: return mongo_database.coupons.find_one({"token": token})
    conn = db(); row = conn.execute("SELECT * FROM coupons WHERE token = ?", (token,)).fetchone(); conn.close(); return row_dict(row)


def update_coupon(coupon_id, changes, required_status=None):
    if USE_MONGO:
        query = {"id": coupon_id}
        if required_status: query["status"] = required_status
        return mongo_database.coupons.update_one(query, {"$set": changes}).modified_count
    assignments = ", ".join(f"{key} = ?" for key in changes); params = list(changes.values()) + [coupon_id]
    sql = f"UPDATE coupons SET {assignments} WHERE id = ?"
    if required_status: sql += " AND status = ?"; params.append(required_status)
    conn = db(); cur = conn.execute(sql, params); conn.commit(); count = cur.rowcount; conn.close(); return count


def coupon_count(meal, date_prefix):
    if USE_MONGO:
        return mongo_database.coupons.count_documents({"meal": meal, "status": "USED", "generated_at": {"$regex": "^" + re.escape(date_prefix)}})
    conn = db(); count = conn.execute("SELECT COUNT(*) AS c FROM coupons WHERE meal=? AND status='USED' AND substr(generated_at,1,10)=?", (meal, date_prefix)).fetchone()["c"]; conn.close(); return count


def latest_coupon(uid, meal, date_prefix):
    if USE_MONGO:
        return mongo_database.coupons.find_one({"student_uid": uid, "meal": meal, "generated_at": {"$regex": "^" + re.escape(date_prefix)}}, sort=[("id", -1)])
    conn = db(); row = conn.execute("SELECT * FROM coupons WHERE student_uid=? AND meal=? AND substr(generated_at,1,10)=? ORDER BY id DESC LIMIT 1", (uid, meal, date_prefix)).fetchone(); conn.close(); return row_dict(row)


def add_coupon_record(data):
    if USE_MONGO:
        data = dict(data); data["id"] = next_mongo_id("coupons"); mongo_database.coupons.insert_one(data); return data
    columns = list(data); conn = db(); conn.execute(f"INSERT INTO coupons ({','.join(columns)}) VALUES ({','.join('?' for _ in columns)})", [data[c] for c in columns]); conn.commit(); row = conn.execute("SELECT * FROM coupons WHERE token=?", (data["token"],)).fetchone(); conn.close(); return dict(row)


def all_records():
    if USE_MONGO:
        records = list(mongo_database.coupons.find().sort("id", -1)); students = {s["student_uid"]: s for s in mongo_database.students.find()}
        for record in records:
            student = students.get(record["student_uid"], {})
            record.update({key: student.get(key) for key in ["name", "roll_number", "branch", "hostel_room"]})
        return records
    conn = db(); rows = conn.execute("""SELECT coupons.*,students.name,students.roll_number,students.branch,students.hostel_room FROM coupons LEFT JOIN students ON coupons.student_uid=students.student_uid ORDER BY coupons.id DESC""").fetchall(); conn.close(); return [dict(row) for row in rows]


def list_used_records(start_date="", end_date="", gender="", meal="", registration=""):
    records = [r for r in all_records() if r.get("status") == "USED"]
    students = {s["student_uid"]: s for s in list_student_records()}
    result = []
    for record in records:
        student = students.get(record.get("student_uid"), {})
        item = dict(record)
        item.update({key: student.get(key, item.get(key)) for key in [
            "name", "roll_number", "branch", "hostel_room", "gender", "hostel_name", "hostel_block"
        ]})
        used_date = (item.get("used_at") or item.get("generated_at") or "")[:10]
        if start_date and used_date < start_date: continue
        if end_date and used_date > end_date: continue
        if gender and item.get("gender") != gender: continue
        if meal and item.get("meal") != meal: continue
        if registration and item.get("roll_number") != registration: continue
        result.append(item)
    return result


def daily_meal_series(days, gender=""):
    end = current_time().date()
    labels, breakfast, lunch, dinner = [], [], [], []
    records = list_used_records((end - timedelta(days=days - 1)).isoformat(), end.isoformat(), gender=gender)
    for offset in range(days - 1, -1, -1):
        day = end - timedelta(days=offset)
        date_text = day.isoformat()
        labels.append(day.strftime("%d %b"))
        for meal, target in [("BREAKFAST", breakfast), ("LUNCH", lunch), ("DINNER", dinner)]:
            target.append(sum(1 for r in records if (r.get("used_at") or "")[:10] == date_text and r.get("meal") == meal))
    return {"labels": labels, "breakfast": breakfast, "lunch": lunch, "dinner": dinner}


def add_skip_record(uid, meal, skip_date):
    data = {"student_uid": uid, "meal": meal, "skip_date": skip_date,
            "created_at": current_time().strftime("%Y-%m-%d %H:%M:%S")}
    try:
        if USE_MONGO:
            data["id"] = next_mongo_id("skipped_meals"); mongo_database.skipped_meals.insert_one(data)
        else:
            conn = db(); conn.execute("INSERT INTO skipped_meals (student_uid,meal,skip_date,created_at) VALUES (?,?,?,?)",
                                      (uid, meal, skip_date, data["created_at"])); conn.commit(); conn.close()
        return True
    except (DuplicateKeyError, sqlite3.IntegrityError):
        return False


def list_skips(uid="", start_date="", end_date=""):
    if USE_MONGO:
        query = {}
        if uid: query["student_uid"] = uid
        if start_date or end_date:
            query["skip_date"] = {}
            if start_date: query["skip_date"]["$gte"] = start_date
            if end_date: query["skip_date"]["$lte"] = end_date
        return list(mongo_database.skipped_meals.find(query).sort("skip_date", -1))
    where, params = ["1=1"], []
    if uid: where.append("student_uid=?"); params.append(uid)
    if start_date: where.append("skip_date>=?"); params.append(start_date)
    if end_date: where.append("skip_date<=?"); params.append(end_date)
    conn = db(); rows = conn.execute("SELECT * FROM skipped_meals WHERE " + " AND ".join(where) + " ORDER BY skip_date DESC", params).fetchall(); conn.close()
    return [dict(row) for row in rows]


def find_admin(username):
    if USE_MONGO: return mongo_database.admins.find_one({"username": username.lower()})
    conn = db(); row = conn.execute("SELECT * FROM admins WHERE username=?", (username.lower(),)).fetchone(); conn.close(); return row_dict(row)


def list_admins():
    if USE_MONGO: return list(mongo_database.admins.find().sort("username", ASCENDING))
    conn = db(); rows = conn.execute("SELECT * FROM admins ORDER BY username").fetchall(); conn.close(); return [dict(r) for r in rows]


def add_admin_account(username, password, role):
    data = {"username": username.lower(), "password_hash": generate_password_hash(password), "role": role,
            "active": 1, "created_at": current_time().strftime("%Y-%m-%d %H:%M:%S")}
    try:
        if USE_MONGO:
            data["id"] = next_mongo_id("admins"); mongo_database.admins.insert_one(data)
        else:
            conn = db(); conn.execute("INSERT INTO admins (username,password_hash,role,active,created_at) VALUES (?,?,?,?,?)",
                                      (data["username"], data["password_hash"], role, 1, data["created_at"])); conn.commit(); conn.close()
        return True
    except (DuplicateKeyError, sqlite3.IntegrityError): return False


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
    return jsonify({
        "status": "ok",
        "service": "SmartMess",
        "database": "mongodb" if USE_MONGO else "sqlite",
    })


# =========================================================
# ADMIN LOGIN
# =========================================================

@app.route("/admin", methods=["GET", "POST"])
def admin_login():

    if admin_required():
        return redirect(url_for("admin_dashboard"))

    error = ""

    if request.method == "POST":

        username = request.form.get("username", "main").strip().lower() or "main"
        password = request.form.get("password", "")

        if username in ["main", "admin"] and hmac.compare_digest(password, ADMIN_PASSWORD):

            session["admin"] = True
            session["admin_role"] = "MAIN"
            session["admin_username"] = "main"

            return redirect(url_for("admin_dashboard"))

        account = find_admin(username)
        if account and account.get("active", 1) and check_password_hash(account["password_hash"], password):
            session["admin"] = True
            session["admin_role"] = account["role"]
            session["admin_username"] = account["username"]
            return redirect(url_for("admin_scanner" if account["role"] == "SCANNER" else "admin_dashboard"))

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

                <label>Username</label>
                <input name="username" value="main" required>

                <label>Password</label>

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
    session.clear()

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

    coupon = coupon_by_token(token)

    if not coupon:
        return jsonify({
            "success": False,
            "message": "Coupon not found."
        }), 404

    # ---------------------------------------------------------
    # CHECK ALREADY USED
    # ---------------------------------------------------------

    if coupon["status"] == "USED":

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

        update_coupon(coupon["id"], {"status": "EXPIRED"})

        return jsonify({
            "success": False,
            "message": "Coupon expired. 5-minute validity ended."
        })

    # ---------------------------------------------------------
    # FIND STUDENT
    # ---------------------------------------------------------

    student = student_by_uid(coupon["student_uid"], active_only=True)

    if not student:

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

    if update_coupon(coupon["id"], {"status": "USED", "used_at": used_time}, required_status="ACTIVE") != 1:

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

    if student.get("photo_filename") or student.get("photo_data"):

        photo_url = (
            "/student-photo/"
            + str(student["id"])
        )

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
    if session.get("admin_role") == "SCANNER":
        return redirect(url_for("admin_scanner"))

    today = current_time().strftime("%Y-%m-%d")

    scope_gender = admin_scope_gender()
    scope = {"gender": scope_gender} if scope_gender else {}
    student_count = count_students(scope)
    boys_count = count_students({"gender": "BOY"}) if not scope_gender or scope_gender == "BOY" else 0
    girls_count = count_students({"gender": "GIRL"}) if not scope_gender or scope_gender == "GIRL" else 0
    bh1_count = count_students({**scope, "hostel_block": "BH-1"})
    bh2_count = count_students({**scope, "hostel_block": "BH-2"})
    today_records = list_used_records(today, today, gender=scope_gender)
    breakfast = sum(r["meal"] == "BREAKFAST" for r in today_records)
    lunch = sum(r["meal"] == "LUNCH" for r in today_records)
    dinner = sum(r["meal"] == "DINNER" for r in today_records)
    present_uids = {r["student_uid"] for r in today_records}
    absent = max(student_count - len(present_uids), 0)
    skipped = len(list_skips(start_date=today, end_date=today))
    chart7 = json.dumps(daily_meal_series(7, scope_gender))
    chart30 = json.dumps(daily_meal_series(30, scope_gender))
    database_label = "MongoDB Atlas (Permanent)" if USE_MONGO else "SQLite (Local)"

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Admin Dashboard</title>
        {CSS}
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
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

            <a class="btn blue" href="/admin/reports">📥 Reports</a>
            <a class="btn green" href="/admin/menu-notice">📋 Menu & Notice</a>
            {'<a class="btn gray" href="/admin/roles">🔐 Admin Roles</a>' if admin_required(['MAIN']) else ''}

            <a class="btn green" href="/admin/meal-settings">
                ⏰ Meal Settings
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

            <div class="stat"><div>🏢 BH-1 Students</div><h2>{bh1_count}</h2></div>
            <div class="stat"><div>🏢 BH-2 Students</div><h2>{bh2_count}</h2></div>

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

            <div class="stat"><div>Absent Today</div><h2>{absent}</h2></div>
            <div class="stat"><div>Skipped Today</div><h2>{skipped}</h2></div>

            <div class="stat">
                <div>Database</div>
                <h2 style="font-size:20px;">{database_label}</h2>
            </div>

        </div>

        <div class="grid">
          <div class="card"><h2>Last 7 Days</h2><canvas id="chart7"></canvas></div>
          <div class="card"><h2>Last 30 Days</h2><canvas id="chart30"></canvas></div>
        </div>

    </div>
    <script>
    function draw(id, data) {{ new Chart(document.getElementById(id), {{type:'line',data:{{labels:data.labels,datasets:[
      {{label:'Breakfast',data:data.breakfast,borderColor:'#f59e0b',tension:.35}},
      {{label:'Lunch',data:data.lunch,borderColor:'#0ea5e9',tension:.35}},
      {{label:'Dinner',data:data.dinner,borderColor:'#8b5cf6',tension:.35}}
    ]}},options:{{responsive:true,plugins:{{legend:{{position:'bottom'}}}},scales:{{y:{{beginAtZero:true,ticks:{{precision:0}}}}}}}}}}); }}
    draw('chart7', {chart7}); draw('chart30', {chart30});
    </script>
    </body>
    </html>
    """

    return html


# =========================================================
# ADMIN - MEAL SETTINGS
# =========================================================

@app.route("/admin/meal-settings", methods=["GET", "POST"])
def admin_meal_settings():
    if not admin_required():
        return redirect(url_for("admin_login"))

    message = ""
    message_class = "message success"
    if request.method == "POST":
        values = {}
        valid = True
        for meal in ["breakfast", "lunch", "dinner"]:
            start = request.form.get(meal + "_start", "").strip()
            end = request.form.get(meal + "_end", "").strip()
            try:
                start_time = datetime.strptime(start, "%H:%M")
                end_time = datetime.strptime(end, "%H:%M")
                if start_time >= end_time:
                    valid = False
            except ValueError:
                valid = False
            values[meal + "_open"] = "1" if request.form.get(meal + "_open") == "1" else "0"
            values[meal + "_start"] = start
            values[meal + "_end"] = end
        if valid:
            save_settings(values)
            message = "Meal settings saved successfully."
        else:
            message = "Please select valid start and end times. End time must be later than start time."
            message_class = "message error"

    settings = {
        meal: {
            "open": get_setting(meal + "_open", "1") == "1",
            "start": get_setting(meal + "_start"),
            "end": get_setting(meal + "_end"),
        }
        for meal in ["breakfast", "lunch", "dinner"]
    }

    cards = ""
    for meal, icon in [("breakfast", "🌅"), ("lunch", "☀️"), ("dinner", "🌙")]:
        item = settings[meal]
        cards += f"""
        <div class="card">
          <h2>{icon} {meal.title()}</h2>
          <label style="display:flex;gap:10px;align-items:center;margin-bottom:18px">
            <input type="checkbox" name="{meal}_open" value="1" {'checked' if item['open'] else ''} style="width:auto;margin:0">
            Open for students
          </label>
          <label>Start Time</label><input type="time" name="{meal}_start" value="{item['start']}" required>
          <label>End Time</label><input type="time" name="{meal}_end" value="{item['end']}" required>
        </div>
        """

    return f"""<!DOCTYPE html><html><head><title>Meal Settings</title>{CSS}</head><body>
    <div class="container">
      <div class="nav"><a class="btn" href="/admin/dashboard">Dashboard</a><a class="btn blue" href="/admin/scanner">Scanner</a></div>
      <h1>⏰ Meal Open/Close & Timings</h1>
      {f'<div class="{message_class}">{escape(message)}</div>' if message else ''}
      <form method="POST"><div class="grid">{cards}</div>
        <div class="card center"><button class="btn green" type="submit">Save Meal Settings</button></div>
      </form>
    </div></body></html>"""


# =========================================================
# ADMIN - ADD STUDENT
# =========================================================


@app.route("/admin/add-student", methods=["GET", "POST"])
def admin_add_student():

    if not admin_required(["MAIN", "BOYS", "GIRLS"]):
        return redirect(url_for("admin_login"))

    message = ""
    message_class = "message"

    if request.method == "POST":

        name = request.form.get("name", "").strip()
        roll = request.form.get("roll_number", "").strip()
        branch = request.form.get("branch", "").strip()
        room = request.form.get("hostel_room", "").strip()
        gender = request.form.get("gender", "").strip().upper()
        if admin_scope_gender(): gender = admin_scope_gender()
        hostel_name = request.form.get("hostel_name", "").strip()
        hostel_block = request.form.get("hostel_block", "").strip()
        pin = request.form.get("pin", "").strip()
        photo = request.files.get("photo")

        if (not name or not roll or branch not in BRANCHES or not room or
                gender not in ["BOY", "GIRL"] or hostel_name not in ["Boys Hostel", "Girls Hostel"] or
                hostel_block not in HOSTEL_BLOCKS or len(pin) != 4 or not pin.isdigit()):

            message = "Please fill every field."
            message_class = "message error"

        elif not photo or not photo.filename:

            message = "Student photo is required."
            message_class = "message error"

        else:

            exists = student_by_roll(roll)

            if exists:

                message = "This Registration Number is already registered."
                message_class = "message error"

            else:

                uid = make_student_uid()

                ext = os.path.splitext(
                    photo.filename
                )[1].lower()

                if ext not in [".jpg", ".jpeg", ".png", ".webp"]:

                    message = "Use JPG, JPEG, PNG or WEBP."
                    message_class = "message error"

                else:

                    filename = uid + ext
                    photo_bytes = photo.read()
                    mime = {".png": "image/png", ".webp": "image/webp"}.get(ext, "image/jpeg")
                    try:
                        add_student_record({
                            "student_uid": uid, "name": name, "roll_number": roll,
                            "branch": branch, "hostel_room": room, "gender": gender,
                            "hostel_name": hostel_name, "hostel_block": hostel_block,
                            "photo_filename": filename, "photo_data": photo_bytes,
                            "photo_mime": mime, "pin_hash": generate_password_hash(pin),
                            "active": 1, "created_at": current_time().strftime("%Y-%m-%d %H:%M:%S")
                        })
                    except (DuplicateKeyError, sqlite3.IntegrityError):
                        message = "This Registration Number is already registered."
                        message_class = "message error"
                        return redirect(url_for("admin_add_student"))

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
                <select name="branch" required>
                    <option value="">Select Branch</option>
                    <option value="AI &amp; ML">AI &amp; ML</option>
                    <option value="Civil (Construction Technology)">Civil (Construction Technology)</option>
                    <option value="Electronics (Robotics)">Electronics (Robotics)</option>
                    <option value="Mechanical (CAD/CAM)">Mechanical (CAD/CAM)</option>
                </select>

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
                <select name="hostel_block" required>
                    <option value="">Select Hostel Block</option>
                    <option value="BH-1">BH-1</option>
                    <option value="BH-2">BH-2</option>
                </select>

                <label>4-digit Student PIN</label>
                <input type="password" name="pin" inputmode="numeric" minlength="4" maxlength="4" pattern="[0-9]{{4}}" required>

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

    if not admin_required(["MAIN", "BOYS", "GIRLS"]):
        return redirect(url_for("admin_login"))

    search = request.args.get("q", "").strip()
    gender_filter = request.args.get("gender", "").strip().upper()
    if admin_scope_gender(): gender_filter = admin_scope_gender()
    hostel_filter = request.args.get("hostel", "").strip()
    branch_filter = request.args.get("branch", "").strip()
    block_filter = request.args.get("block", "").strip()
    students = list_student_records(
        search=search,
        gender=gender_filter if gender_filter in ["BOY", "GIRL"] else "",
        hostel=hostel_filter if hostel_filter in ["Boys Hostel", "Girls Hostel"] else "",
        branch=branch_filter if branch_filter in BRANCHES else "",
        block=block_filter if block_filter in HOSTEL_BLOCKS else "",
    )

    rows = ""

    for student in students:

        photo = "No Photo"

        if student.get("photo_filename") or student.get("photo_data"):

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
              <div><label>Branch</label><select name="branch"><option value="">All Branches</option>{''.join(f'<option value="{escape(branch)}" {"selected" if branch_filter == branch else ""}>{escape(branch)}</option>' for branch in BRANCHES)}</select></div>
              <div><label>Block</label><select name="block"><option value="">All Blocks</option>{''.join(f'<option value="{block}" {"selected" if block_filter == block else ""}>{block}</option>' for block in HOSTEL_BLOCKS)}</select></div>
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
    if not admin_required(["MAIN", "BOYS", "GIRLS"]):
        return redirect(url_for("admin_login"))

    student = student_by_id(student_id)
    if not student:
        return "Student not found", 404
    if admin_scope_gender() and student.get("gender") != admin_scope_gender():
        return "Not allowed for this hostel admin.", 403

    message = ""
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        registration = request.form.get("registration_number", "").strip()
        branch = request.form.get("branch", "").strip()
        room = request.form.get("hostel_room", "").strip()
        gender = request.form.get("gender", "").strip().upper()
        if admin_scope_gender(): gender = admin_scope_gender()
        hostel_name = request.form.get("hostel_name", "").strip()
        hostel_block = request.form.get("hostel_block", "").strip()
        pin = request.form.get("pin", "").strip()
        photo = request.files.get("photo")
        if (not name or not registration or branch not in BRANCHES or not room or
                gender not in ["BOY", "GIRL"] or hostel_name not in ["Boys Hostel", "Girls Hostel"] or
                hostel_block not in HOSTEL_BLOCKS):
            message = "Please fill all required fields."
        elif pin and (len(pin) != 4 or not pin.isdigit()):
            message = "New PIN must contain exactly 4 numbers."
        else:
            duplicate = student_by_roll(registration)
            if duplicate and duplicate["id"] != student_id:
                message = "Registration Number already exists."
            else:
                changes = {
                    "name": name,
                    "roll_number": registration,
                    "branch": branch,
                    "hostel_room": room,
                    "gender": gender,
                    "hostel_name": hostel_name,
                    "hostel_block": hostel_block,
                }
                if photo and photo.filename:
                    ext = os.path.splitext(photo.filename)[1].lower()
                    if ext not in [".jpg", ".jpeg", ".png", ".webp"]:
                        message = "Use JPG, JPEG, PNG or WEBP photo."
                    else:
                        changes.update({
                            "photo_filename": student["student_uid"] + ext,
                            "photo_data": photo.read(),
                            "photo_mime": {".png": "image/png", ".webp": "image/webp"}.get(ext, "image/jpeg"),
                        })
                if pin:
                    changes["pin_hash"] = generate_password_hash(pin)
                if not message:
                    update_student_record(student_id, changes)
                    return redirect(url_for("admin_students"))

    student = student_by_id(student_id)
    html = f"""<!DOCTYPE html><html><head><title>Edit Student</title>{CSS}</head><body><div class="container">
    <div class="nav"><a class="btn" href="/admin/students">Back to Students</a></div>
    <div class="card"><h1>✏️ Edit Student</h1>{f'<div class="message error">{escape(message)}</div>' if message else ''}
    <form method="POST" enctype="multipart/form-data">
      <label>Name</label><input name="name" value="{escape(student['name'])}" required>
      <label>Registration Number</label><input name="registration_number" value="{escape(student['roll_number'])}" required>
      <label>Branch</label><select name="branch" required>
        {''.join(f'<option value="{escape(branch)}" {"selected" if student["branch"] == branch else ""}>{escape(branch)}</option>' for branch in BRANCHES)}
      </select>
      <label>Gender</label><select name="gender" required><option value="BOY" {'selected' if student['gender']=='BOY' else ''}>Boy</option><option value="GIRL" {'selected' if student['gender']=='GIRL' else ''}>Girl</option></select>
      <label>Hostel</label><select name="hostel_name" required><option value="Boys Hostel" {'selected' if student['hostel_name']=='Boys Hostel' else ''}>Boys Hostel</option><option value="Girls Hostel" {'selected' if student['hostel_name']=='Girls Hostel' else ''}>Girls Hostel</option></select>
      <label>Hostel Block</label><select name="hostel_block" required>
        {''.join(f'<option value="{block}" {"selected" if student["hostel_block"] == block else ""}>{block}</option>' for block in HOSTEL_BLOCKS)}
      </select>
      <label>Room Number</label><input name="hostel_room" value="{escape(student['hostel_room'])}" required>
      <label>Reset 4-digit PIN (optional)</label><input type="password" name="pin" inputmode="numeric" minlength="4" maxlength="4" pattern="[0-9]{{4}}" placeholder="Leave blank to keep current PIN">
      <label>Change Photo (optional)</label><input type="file" name="photo" accept="image/*">
      <button class="btn green" type="submit">Save Changes</button>
    </form></div></div></body></html>"""
    return html


@app.route("/admin/student/<int:student_id>/toggle", methods=["POST"])
def admin_toggle_student(student_id):
    if not admin_required(["MAIN", "BOYS", "GIRLS"]):
        return redirect(url_for("admin_login"))
    student = student_by_id(student_id)
    if admin_scope_gender() and student and student.get("gender") != admin_scope_gender():
        return "Not allowed for this hostel admin.", 403
    toggle_student_record(student_id)
    return redirect(url_for("admin_students"))


# =========================================================
# STUDENT PHOTO
# =========================================================

@app.route("/student-photo/<int:student_id>")
def student_photo(student_id):
    student = student_by_id(student_id)
    if not student:
        return "Photo not found", 404
    if student.get("photo_data"):
        return Response(bytes(student["photo_data"]), mimetype=student.get("photo_mime") or "image/jpeg")
    filename = student.get("photo_filename")
    if filename and os.path.exists(os.path.join(PHOTO_DIR, filename)):
        return send_from_directory(PHOTO_DIR, filename)
    return "Photo not found", 404


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
                    Registration Number aur 4-digit PIN se login kare.
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

                <label>4-digit PIN</label>
                <input
                    type="password"
                    name="pin"
                    inputmode="numeric"
                    minlength="4"
                    maxlength="4"
                    pattern="[0-9]{{4}}"
                    placeholder="Enter your 4-digit PIN"
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
    pin = request.form.get("pin", "").strip()
    student = student_by_roll(roll, active_only=True)

    if not student or not student.get("pin_hash") or not check_password_hash(student["pin_hash"], pin):

        return f"""
        <!DOCTYPE html>
        <html>
        <head>{CSS}</head>
        <body>

        <div class="container">

            <div class="card center">

                <h1>❌ Login Failed</h1>

                <div class="message error">
                    Registration Number or PIN is incorrect, or this student is inactive.
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

    student = student_by_uid(student_uid, active_only=True)

    if not student:

        session.pop("student_uid", None)

        return redirect(
            url_for("student_home")
        )

    photo = f"/student-photo/{student['id']}" if student.get("photo_filename") or student.get("photo_data") else ""

    meal_status_cards = ""
    for meal, icon in [("BREAKFAST", "🌅"), ("LUNCH", "☀️"), ("DINNER", "🌙")]:
        available, status_message = meal_is_available(meal)
        if available:
            status_html = '<span class="badge used-badge">OPEN NOW</span>'
        else:
            status_html = '<span class="badge expired-badge">CLOSED</span>'
        meal_status_cards += (
            f'<div class="stat"><div>{icon} {meal.title()}</div>'
            f'<p>{status_html}</p><small>{escape(status_message) if status_message else "Coupon available"}</small></div>'
        )

    today = current_time().strftime("%Y-%m-%d")
    month_start = current_time().strftime("%Y-%m-01")
    history = list_used_records(month_start, today, registration=student["roll_number"])
    history_rows = "".join(f"<tr><td>{(r.get('used_at') or '')[:10]}</td><td>{r.get('meal','-')}</td><td>{(r.get('used_at') or '')[11:19]}</td></tr>" for r in history[:31])
    menus = {meal: get_setting(meal.lower() + "_menu", "Menu not updated") or "Menu not updated" for meal in ["BREAKFAST", "LUNCH", "DINNER"]}
    notice = get_setting("student_notice", "")
    tomorrow = (current_time().date() + timedelta(days=1)).isoformat()

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

            {f'<div class="message"><b>📢 Notice:</b> {escape(notice)}</div>' if notice else ''}
            <h2 class="center">🍽️ Today's Menu</h2>
            <div class="grid">
              <div class="stat"><b>Breakfast</b><p>{escape(menus['BREAKFAST'])}</p></div>
              <div class="stat"><b>Lunch</b><p>{escape(menus['LUNCH'])}</p></div>
              <div class="stat"><b>Dinner</b><p>{escape(menus['DINNER'])}</p></div>
            </div>

            <h2 class="center">
                Select Meal
            </h2>

            <div class="grid" style="margin-bottom:20px;">{meal_status_cards}</div>

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

        <div class="grid">
          <div class="card"><h2>⏭️ Skip a Meal</h2><p>Tell the mess in advance to reduce food waste.</p>
            <form method="post" action="/student/skip-meal"><label>Meal</label><select name="meal"><option>BREAKFAST</option><option>LUNCH</option><option>DINNER</option></select>
            <label>Date</label><select name="skip_date"><option value="{today}">Today</option><option value="{tomorrow}">Tomorrow</option></select><button class="btn red">Skip Meal</button></form>
          </div>
          <div class="card"><h2>📅 This Month History</h2><p>Total meals: <b>{len(history)}</b></p><div style="overflow:auto"><table><tr><th>Date</th><th>Meal</th><th>Time</th></tr>{history_rows or '<tr><td colspan="3">No meals yet.</td></tr>'}</table></div></div>
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

    student = student_by_uid(student_uid, active_only=True)

    if not student:
        return redirect(
            url_for("student_home")
        )

    available, unavailable_message = meal_is_available(meal)
    if not available:
        return f"""
        <!DOCTYPE html><html><head><title>Meal Closed</title>{CSS}</head><body>
        <div class="container"><div class="card center">
          <h1>⏰ {meal.title()} Closed</h1>
          <div class="message error">{escape(unavailable_message)}</div>
          <a class="btn" href="/student/meals">Back to Meals</a>
        </div></div></body></html>
        """

    # -----------------------------------------------------
    # Check today's same meal
    # -----------------------------------------------------

    today = current_time().strftime("%Y-%m-%d")

    if any(item["meal"] == meal for item in list_skips(student_uid, today, today)):
        return f"""<!DOCTYPE html><html><head><title>Meal Skipped</title>{CSS}</head><body><div class="container"><div class="card center">
        <h1>⏭️ Meal Skipped</h1><div class="message error">You marked today's {meal.title()} as skipped.</div><a class="btn" href="/student/meals">Back</a>
        </div></div></body></html>"""

    existing = latest_coupon(student_uid, meal, today)

    if existing:

        if existing["status"] == "USED":

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

                # Show existing active coupon.
                return render_coupon(
                    student,
                    existing
                )

            else:

                update_coupon(existing["id"], {"status": "EXPIRED"})

    # -----------------------------------------------------
    # Create new coupon
    # -----------------------------------------------------

    generated = current_time()

    expires = generated + timedelta(minutes=5)

    token = make_coupon_token()

    coupon = add_coupon_record({
        "token": token,
        "student_uid": student_uid,
        "meal": meal,
        "generated_at": generated.strftime("%Y-%m-%d %H:%M:%S"),
        "expires_at": expires.strftime("%Y-%m-%d %H:%M:%S"),
        "used_at": None,
        "status": "ACTIVE",
    })

    return render_coupon(
        student,
        coupon
    )


@app.route("/student/skip-meal", methods=["POST"])
def student_skip_meal():
    uid = session.get("student_uid")
    if not uid: return redirect(url_for("student_home"))
    meal, skip_date = request.form.get("meal", "").upper(), request.form.get("skip_date", "")
    allowed_dates = {current_time().date().isoformat(), (current_time().date() + timedelta(days=1)).isoformat()}
    if meal not in ["BREAKFAST", "LUNCH", "DINNER"] or skip_date not in allowed_dates: return "Invalid meal or date.", 400
    success = add_skip_record(uid, meal, skip_date)
    text = "Skip meal saved. Thank you for helping reduce food waste." if success else "This meal is already marked as skipped."
    return f"""<!DOCTYPE html><html><head><title>Skip Meal</title>{CSS}</head><body><div class="container"><div class="card center"><h1>⏭️ Skip Meal</h1><div class="message {'success' if success else 'error'}">{text}</div><a class="btn" href="/student/meals">Back to Meals</a></div></div></body></html>"""


# =========================================================
# COUPON HTML
# =========================================================

def render_coupon(student, coupon):
    photo = f"/student-photo/{student['id']}" if student.get("photo_filename") or student.get("photo_data") else ""

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

@app.route("/admin/reports")
def admin_reports():
    if not admin_required(["MAIN", "BOYS", "GIRLS"]): return redirect(url_for("admin_login"))
    today = current_time().strftime("%Y-%m-%d")
    month = current_time().strftime("%Y-%m")
    students = list_student_records(gender=admin_scope_gender())
    options = "".join(f'<option value="{escape(s["roll_number"])}">{escape(s["name"])} - {escape(s["roll_number"])}</option>' for s in students)
    return f"""<!DOCTYPE html><html><head><title>SmartMess Reports</title>{CSS}</head><body><div class="container">
    <div class="nav"><a class="btn" href="/admin/dashboard">Dashboard</a><a class="btn gray" href="/admin/records">Records</a></div>
    <div class="card"><h1>📥 Excel & PDF Reports</h1><p>Today, monthly, meal-wise, gender-wise and individual student reports.</p>
    <form action="/admin/report/download" method="get">
      <div class="grid">
       <div><label>From Date</label><input type="date" name="start" value="{today}"></div>
       <div><label>To Date</label><input type="date" name="end" value="{today}"></div>
       <div><label>Gender</label><select name="gender"><option value="">All</option><option value="BOY">Boys</option><option value="GIRL">Girls</option></select></div>
       <div><label>Meal</label><select name="meal"><option value="">All Meals</option><option>BREAKFAST</option><option>LUNCH</option><option>DINNER</option></select></div>
       <div><label>Individual Student</label><select name="registration"><option value="">All Students</option>{options}</select></div>
       <div><label>Format</label><select name="format"><option value="xlsx">Excel (.xlsx)</option><option value="pdf">PDF</option></select></div>
      </div><button class="btn green" type="submit">Download Report</button>
      <a class="btn blue" href="/admin/report/download?start={month}-01&end={today}&format=xlsx">This Month Excel</a>
    </form></div></div></body></html>"""


@app.route("/admin/report/download")
def admin_report_download():
    if not admin_required(["MAIN", "BOYS", "GIRLS"]): return redirect(url_for("admin_login"))
    today = current_time().strftime("%Y-%m-%d")
    start, end = request.args.get("start", today), request.args.get("end", today)
    gender = admin_scope_gender() or request.args.get("gender", "")
    meal = request.args.get("meal", "")
    registration = request.args.get("registration", "")
    records = list_used_records(start, end, gender, meal, registration)
    title = f"SmartMess Attendance Report ({start} to {end})"
    headers = ["Date", "Time", "Name", "Registration No.", "Gender", "Branch", "Hostel", "Block", "Room", "Meal"]
    rows = []
    for r in records:
        used = r.get("used_at") or r.get("generated_at") or ""
        rows.append([used[:10], used[11:19], r.get("name", "-"), r.get("roll_number", "-"), r.get("gender", "-"),
                     r.get("branch", "-"), r.get("hostel_name", "-"), r.get("hostel_block", "-"), r.get("hostel_room", "-"), r.get("meal", "-")])
    fmt = request.args.get("format", "xlsx")
    if fmt == "pdf":
        buffer = BytesIO(); doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=10*mm, leftMargin=10*mm, topMargin=12*mm, bottomMargin=12*mm)
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle("SmartMessTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=colors.HexColor("#102b4f"), spaceAfter=8)
        body_style = ParagraphStyle("SmartMessBody", parent=styles["Normal"], fontName="Helvetica", fontSize=9, leading=12)
        story = [Paragraph(title, title_style), Paragraph(f"Total attendance records: {len(rows)}", body_style), Spacer(1, 10)]
        table = Table([headers] + rows, repeatRows=1, colWidths=[23*mm,19*mm,33*mm,31*mm,18*mm,29*mm,25*mm,16*mm,16*mm,22*mm])
        table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#102b4f")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),7),("GRID",(0,0),(-1,-1),.35,colors.HexColor("#cbd5e1")),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#f8fafc")]),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("PADDING",(0,0),(-1,-1),4)]))
        story.append(table); doc.build(story); buffer.seek(0)
        return send_file(buffer, mimetype="application/pdf", as_attachment=True, download_name=f"SmartMess_Report_{start}_{end}.pdf")
    wb = Workbook(); ws = wb.active; ws.title = "Attendance"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers)); ws["A1"] = title
    ws["A1"].font = Font(size=18, bold=True, color="FFFFFF"); ws["A1"].fill = PatternFill("solid", fgColor="102B4F"); ws["A1"].alignment = Alignment(horizontal="center")
    ws.append([f"Total records: {len(rows)}"]); ws.append([]); ws.append(headers)
    for cell in ws[4]: cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor="2563EB")
    for row in rows: ws.append(row)
    thin = Side(style="thin", color="D9E2F0")
    for row in ws.iter_rows(min_row=4):
        for cell in row: cell.border = Border(bottom=thin); cell.alignment = Alignment(vertical="center")
    widths = [13,11,24,21,12,25,20,12,12,15]
    for i, width in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A5"; ws.auto_filter.ref = f"A4:J{max(4, ws.max_row)}"
    buffer = BytesIO(); wb.save(buffer); buffer.seek(0)
    return send_file(buffer, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name=f"SmartMess_Report_{start}_{end}.xlsx")


@app.route("/admin/menu-notice", methods=["GET", "POST"])
def admin_menu_notice():
    if not admin_required(["MAIN", "BOYS", "GIRLS"]): return redirect(url_for("admin_login"))
    message = ""
    if request.method == "POST":
        save_settings({key: request.form.get(key, "").strip() for key in ["breakfast_menu", "lunch_menu", "dinner_menu", "student_notice"]}); message = "Menu and notice saved."
    values = {key: get_setting(key, "") for key in ["breakfast_menu", "lunch_menu", "dinner_menu", "student_notice"]}
    return f"""<!DOCTYPE html><html><head><title>Menu & Notice</title>{CSS}</head><body><div class="container">
    <div class="nav"><a class="btn" href="/admin/dashboard">Dashboard</a></div><div class="card"><h1>📋 Menu & Notice</h1>
    {f'<div class="message success">{message}</div>' if message else ''}<form method="post">
    <label>Breakfast Menu</label><input name="breakfast_menu" value="{escape(values['breakfast_menu'])}">
    <label>Lunch Menu</label><input name="lunch_menu" value="{escape(values['lunch_menu'])}">
    <label>Dinner Menu</label><input name="dinner_menu" value="{escape(values['dinner_menu'])}">
    <label>Student Notice</label><input name="student_notice" value="{escape(values['student_notice'])}">
    <button class="btn green">Save Menu & Notice</button></form></div></div></body></html>"""


@app.route("/admin/roles", methods=["GET", "POST"])
def admin_roles():
    if not admin_required(["MAIN"]): return redirect(url_for("admin_dashboard"))
    message = ""
    if request.method == "POST":
        username, password, role = request.form.get("username", "").strip(), request.form.get("password", ""), request.form.get("role", "")
        if len(username) < 3 or len(password) < 8 or role not in ["BOYS", "GIRLS", "SCANNER"]: message = "Use 3+ character username and 8+ character password."
        elif add_admin_account(username, password, role): message = "Admin account created successfully."
        else: message = "Username already exists."
    rows = "".join(f"<tr><td>{escape(a['username'])}</td><td>{a['role']}</td><td>{'Active' if a.get('active',1) else 'Inactive'}</td></tr>" for a in list_admins())
    return f"""<!DOCTYPE html><html><head><title>Admin Roles</title>{CSS}</head><body><div class="container"><div class="nav"><a class="btn" href="/admin/dashboard">Dashboard</a></div>
    <div class="card"><h1>🔐 Admin Roles</h1>{f'<div class="message">{escape(message)}</div>' if message else ''}<form method="post"><div class="grid">
    <div><label>Username</label><input name="username" required></div><div><label>Password</label><input type="password" name="password" minlength="8" required></div>
    <div><label>Role</label><select name="role"><option value="BOYS">Boys Hostel Admin</option><option value="GIRLS">Girls Hostel Admin</option><option value="SCANNER">Scanner Operator</option></select></div></div>
    <button class="btn green">Create Account</button></form></div><div class="card"><table><tr><th>Username</th><th>Role</th><th>Status</th></tr>{rows}</table></div></div></body></html>"""

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
                width: 230px;
                height: 230px;
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

    function sound(ok) {{
        const audio = new (window.AudioContext || window.webkitAudioContext)();
        const oscillator = audio.createOscillator(); const gain = audio.createGain();
        oscillator.connect(gain); gain.connect(audio.destination);
        oscillator.frequency.value = ok ? 880 : 190; gain.gain.value = .22;
        oscillator.start(); oscillator.stop(audio.currentTime + (ok ? .22 : .55));
    }}

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

                sound(true);

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

                    + "<h1 style='font-size:42px'>✅ MEAL APPROVED</h1>"

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

                sound(false);
                const alreadyUsed = (data.message || '').toLowerCase().includes('already used');

                document.getElementById(
                    "result"
                ).innerHTML =

                    "<div class='message error'>"

                    + "<h1 style='font-size:42px'>❌ " + (alreadyUsed ? "ALREADY USED" : "COUPON REJECTED") + "</h1>"

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

# =========================================================
# ADMIN RECORDS
# =========================================================

@app.route("/admin/records")
def admin_records():

    if not admin_required():
        return redirect(
            url_for("admin_login")
        )

    records = all_records()
    if admin_scope_gender():
        allowed = {s["student_uid"] for s in list_student_records(gender=admin_scope_gender())}
        records = [r for r in records if r.get("student_uid") in allowed]

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
