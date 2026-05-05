from flask import Flask, render_template, request, redirect, url_for, session, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from flask import jsonify
from werkzeug.utils import secure_filename
from PIL import Image,ImageDraw,ImageFont
from datetime import datetime, timedelta
from flask import send_from_directory
import pandas as pd
import sqlite3
import tempfile
import math
import matplotlib.pyplot as plt
import threading
import time
import io
import os
import random
import base64
from datetime import datetime
from datetime import date
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

TARGET_PER_DAY = 20
print(os.listdir("templates"))

app = Flask(__name__)
app.secret_key = "cpv_secret_key"


db_path = "cpv_database.db"

sqlite3.connect(db_path, timeout=20, check_same_thread=False)

def init_db():
    with sqlite3.connect(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("PRAGMA journal_mode=WAL;")
        conn.commit()

init_db()

def create_agents_table():
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS agents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT,
        role TEXT,
        latitude REAL,
        longitude REAL
    )
    """)

    conn.commit()
    conn.close()

create_agents_table()

agents = ["Kumar","Ravi","Arjun","Siva"]

COLUMN_MAPPING = {
    "client_name": ["client_name", "client name", "customer name", "name"],
    "ref_no": ["ref_no", "ref no", "reference number"],
    "cpv_type": ["cpv_type", "cpv type"],
    "bank_name": ["bank_name", "bank"],
    "priority": ["priority"],
    "initiation_date": ["initiation_date", "date", "initiation date"],
    "agency": ["agency"],
    "assigned_agent": ["assigned_agent", "agent"],
    "pincode": ["pincode", "pin", "zip"],
    "status": ["status"]
}

def seed_agents():
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM agents")
    count = cursor.fetchone()[0]

    if count == 0:
        cursor.execute("""
        INSERT INTO agents (username, role, latitude, longitude)
        VALUES
        ('agent1','agent',13.0827,80.2707),
        ('agent2','agent',13.0500,80.2500)
        """)

    conn.commit()
    conn.close()

seed_agents()

UPLOAD_FOLDER = "uploads"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# 🔥 FE PINCODE MAPPING
FE_PINCODE_MAP = {

    # 🔵 VELLORE
    "Jainmose": ["631501","631502","631561"],
    "Jay Ranipet": ["632001","632009","632010","632014","632055","632403","632513"],
    "Jaikum Gudiyatham": ["632602","632601"],
    "Irfan": ["631209","631001","631204","631205","632003","632502"],
    "Aravind": ["635802","635810"],
    "Suresh": ["635651","635601","635653","635751"],
    "Anand": ["606709","606755","606804"],

    # 🟢 TIRUNELVELI
    "Arun Kumar": ["627151","627401","627414","627417","627425","627751","627804","627805"],
    "Arun Valiyan": ["627117","627103"],
    "Bala": ["626125","626117","626142"],
    "Dhanvesh": ["629001","629002","629204","629401","629301","629851","629801"],
    "Iyappan": ["626204","626303","626128","626123"],
    "Jebastin": ["626613"],

    # 🔵 TUTICORIN / NAGERCOIL
    "Joseph": ["628001","628002","628003","628004","628005","628006","628008","628101","628102"],
    "Marimuthu": ["627106","629169","629403","629003","627116","627164"],
    "Muneeswaran": ["626189","626203","623123","626131","626140"],
    "Muthu Krishnan": ["626121","626142"],
    "Prabhu": ["629101","629152"],
    "Raja": ["626125"],
    "Rajaguru": ["628501","628502","628503"],
    "Ram": ["627713"],
    "Ramesh": ["628502","628712","628728","628903","628904"],
    "Selvakumar": ["628581","627109","628704","627657"],
    "Riyas": ["627357","627001","627002"],
    "Sithamala": ["626103"],
    "Sudalai": ["627719","627751","627756","627757","627758","627855","628721"],
    "Thalapathy": ["626118"]
}

def is_agent_present(agent_name):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    today = datetime.now().strftime("%Y-%m-%d")

    cursor.execute("""
        SELECT status FROM attendance 
        WHERE agent_name=? AND date=?
    """, (agent_name, today))

    row = cursor.fetchone()
    conn.close()

    return row and row[0] == "Present"


def auto_assign_agent(pincode):
    pincode = str(pincode)

    for agent, pincodes in FE_PINCODE_MAP.items():
        if pincode in pincodes:
            
            # 🔥 CHECK ATTENDANCE
            if is_agent_present(agent):
                return agent
            else:
                return "Not Available"

    return "Unassigned"

@app.route("/mark_attendance", methods=["POST"])
def mark_attendance():
    agent = request.form.get("agent")
    status = request.form.get("status")

    today = datetime.now().strftime("%Y-%m-%d")

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO attendance (agent_name, status, date)
        VALUES (?, ?, ?)
    """, (agent, status, today))

    conn.commit()
    conn.close()

    return "Attendance Updated"

def init_db():

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # =========================================
    # 🔥 MAIN CASES TABLE
    # =========================================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS cases (
        id INTEGER PRIMARY KEY AUTOINCREMENT,

        case_number TEXT,
        ref_no TEXT,
        application_no TEXT,

        client_name TEXT,
        cpv_type TEXT,
        bank_name TEXT,
        portal_type TEXT,

        address TEXT,
        landmark TEXT,
        city TEXT,
        state TEXT,
        pincode TEXT,

        assigned_agent TEXT,
        status TEXT,
        priority TEXT,

        initiation_date TEXT,
        created_time TEXT,
        assigned_time TEXT,

        agency TEXT,
        feedback TEXT,

        latitude TEXT,
        longitude TEXT,
        map_link TEXT
    )
    """)

    # =========================================
    # 🔍 EXISTING COLUMNS
    # =========================================
    cursor.execute("PRAGMA table_info(cases)")
    columns = [col[1] for col in cursor.fetchall()]

    # =========================================
    # 🔥 EXTRA COMMON FIELDS
    # =========================================
    extra_columns = {
        "photo1": "TEXT",
        "photo2": "TEXT",
        "photo3": "TEXT",
        "photo4": "TEXT",
        "photo5": "TEXT"
    }

    # =========================================
    # 🔥 MUTHOOT FULL FIELDS (UPDATED 🔥)
    # =========================================
    muthoot_columns = {
        "residence_type": "TEXT",
        "locality": "TEXT",
        "accessibility": "TEXT",
        "person_met": "TEXT",
        "years_at_residence": "TEXT",
        "ownership": "TEXT",
        "family_status": "TEXT",
        "vehicle": "TEXT",
        "living_standard": "TEXT",
        "verification_status": "TEXT",
        "final_status": "TEXT",
        "visit_date": "TEXT",
        "document_photo": "TEXT"
    }

    # =========================================
    # ✅ SAFE COLUMN ADD FUNCTION
    # =========================================
    def add_column(name, col_type):
        if name not in columns:
            print(f"➕ Adding column: {name}")
            cursor.execute(f"ALTER TABLE cases ADD COLUMN {name} {col_type}")

    # =========================================
    # 🔥 ADD ALL EXTRA + MUTHOOT COLUMNS
    # =========================================
    for col, typ in extra_columns.items():
        add_column(col, typ)

    for col, typ in muthoot_columns.items():
        add_column(col, typ)

    # =========================================
    # 🔹 USERS TABLE
    # =========================================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        password TEXT,
        role TEXT
    )
    """)

    # =========================================
    # 🔹 CASE NOTES
    # =========================================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS case_notes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        case_id INTEGER,
        note TEXT,
        created_at TEXT
    )
    """)

    # =========================================
    # 🔔 NOTIFICATIONS
    # =========================================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS notifications (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        message TEXT,
        created_at TEXT
    )
    """)

    # =========================================
    # 📍 LIVE LOCATION
    # =========================================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS locations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        case_id INTEGER,
        lat TEXT,
        lng TEXT,
        updated_at TEXT
    )
    """)

    # =========================================
    # 📜 CASE LOGS
    # =========================================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS case_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        case_id INTEGER,
        message TEXT,
        created_at TEXT
    )
    """)

    # =========================================
    # 🔥 AGENTS TABLE
    # =========================================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS agents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT,
        role TEXT,
        latitude REAL,
        longitude REAL
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        agent TEXT,
        status TEXT,
        date TEXT
    )
    """)

    print("✅ DB READY (HDFC + MUTHOOT FULL SUPPORT)")

    # =========================================
    # 👤 DEFAULT USERS
    # =========================================
    try:
        cursor.execute(
            "INSERT INTO users (username,password,role) VALUES (?,?,?)",
            ("admin", generate_password_hash("admin123"), "admin")
        )

        cursor.execute(
            "INSERT INTO users (username,password,role) VALUES (?,?,?)",
            ("agent1", generate_password_hash("agent123"), "agent")
        )
    except:
        pass

    conn.commit()

    # =========================================
    # 🔍 FINAL DEBUG
    # =========================================
    cursor.execute("PRAGMA table_info(cases)")
    print("📊 CASE TABLE STRUCTURE:")
    for col in cursor.fetchall():
        print(col)

    conn.close()

@app.route("/clear_muthoot")
def clear_muthoot():
    with sqlite3.connect(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM cases WHERE portal_type='MUTHOOT'")
        conn.commit()

    return "Muthoot data cleared"

def auto_assign_agent(lat=None, lng=None, priority="Normal", portal_type="HDFC"):

    try:
        with sqlite3.connect(db_path, timeout=20) as conn:
            cursor = conn.cursor()

            cursor.execute("SELECT username FROM agents WHERE role='agent'")
            agents = cursor.fetchall()

            print("DEBUG AGENTS:", agents)  # ✅ DEBUG

            if not agents:
                return "unassigned"

            # ✅ FIX (tuple index error avoid)
            return agents[0][0]

    except Exception as e:
        print("Auto Reassign Error:", e)
        return "unassigned"

def generate_summary(client_name, bank_name, cpv_type, status):

    if status == "Completed":
        result = "Verification completed successfully. Recommended: Approved."
    elif status == "Pending":
        result = "Verification is still pending. Awaiting further action."
    else:
        result = "Verification in progress. Field check ongoing."

    summary = f"""
    Customer {client_name} applied for {cpv_type} verification from {bank_name}.
    {result}
    """

    return summary

def generate_case_number(client):

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 🔍 Get last ID
    cursor.execute("SELECT MAX(id) FROM cases")
    last_id = cursor.fetchone()[0]

    conn.close()

    if last_id is None:
        last_id = 0

    new_id = last_id + 1

    # 🔥 Prefix based on client
    if client.upper() == "HDFC":
        prefix = "HDFC"
    else:
        prefix = "MUT"

    # 🔥 Random number
    rand = random.randint(100, 999)

    # 🔥 Final Case Number (Professional format)
    return f"{prefix}-{new_id:05d}-{rand}"

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route("/download_pdf/<int:id>")
def download_pdf(id):

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM cases WHERE id=?", (id,))
    case = cursor.fetchone()

    conn.close()

    #  Safety check
    if not case:
        return "Case not found"

    # 📄 PDF content
    styles = getSampleStyleSheet()
    content = []

    content.append(Paragraph(f"Case Number: {case[1]}", styles['Normal']))
    content.append(Paragraph(f"Client Name: {case[3]}", styles['Normal']))
    content.append(Paragraph(f"Bank: {case[5]}", styles['Normal']))
    content.append(Paragraph(f"Status: {case[11]}", styles['Normal']))  # ✅ FIXED INDEX

    # 🔥 Memory buffer (NO FILE SAVE)
    buffer = io.BytesIO()

    doc = SimpleDocTemplate(buffer)
    doc.build(content)

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"case_{id}.pdf",
        mimetype="application/pdf"
    )

@app.route("/", methods=["GET","POST"])
def login():

    if request.method == "POST":

        username = request.form["username"]
        password = request.form["password"]

        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute("SELECT * FROM users WHERE username=?", (username,))
        user = cursor.fetchone()

        conn.close()

        if user and len(user) > 3 and check_password_hash(user[2], password):

            session["user"] = username
            session["role"] = user[3]

            return redirect(url_for("dashboard"))

        else:
            # 🔥 UI ERROR SHOW
            return render_template("login.html", error="Invalid Username or Password")

    return render_template("login.html")

@app.route("/register", methods=["GET","POST"])
def register():

    if request.method == "POST":

        username = request.form["username"]
        password = generate_password_hash(request.form["password"])
        role = request.form["role"]

        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        try:
            cursor.execute("INSERT INTO users (username,password,role) VALUES (?,?,?)",
                           (username,password,role))
            conn.commit()
        except:
            return "User already exists"

        conn.close()
        return redirect(url_for("login"))

    return render_template("register.html")

def calculate_distance(lat1, lon1, lat2, lon2):
    R = 6371

    d_lat = math.radians(lat2 - lat1)
    d_lon = math.radians(lon2 - lon1)

    a = math.sin(d_lat/2)*2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(d_lon/2)*2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))

    return R * c

# 🚀 AUTO ASSIGN (FINAL)
def auto_assign_agent(lat=None, lng=None, priority="Normal", portal_type="HDFC"):

    try:
        with sqlite3.connect(db_path, timeout=20) as conn:
            conn.row_factory = sqlite3.Row   # ✅ safe access
            cursor = conn.cursor()

            cursor.execute("""
            SELECT username 
            FROM agents 
            WHERE role='agent'
            """)

            agents = cursor.fetchall()

            print("DEBUG AGENTS:", agents)  # 🔍 debug

            # ❌ no agents
            if not agents or len(agents) == 0:
                return "unassigned"

            # ✅ FIX: avoid tuple index error
            first_agent = agents[0]

            if isinstance(first_agent, sqlite3.Row):
                return first_agent["username"]
            else:
                return first_agent[0]

    except Exception as e:
        print("Auto Reassign Error:", e)
        return "unassigned"

@app.route("/get_locations")
def get_locations():

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("""
    SELECT id, latitude, longitude 
    FROM cases
    WHERE latitude IS NOT NULL AND longitude IS NOT NULL
    """)

    rows = cursor.fetchall()
    conn.close()

    data = []

    for r in rows:
        data.append({
            "case_id": r[0],
            "lat": r[1],
            "lng": r[2]
        })

    return jsonify(data)

@app.route("/live_map")
def live_map():

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("""
    SELECT client_name, latitude, longitude 
    FROM cases 
    WHERE latitude IS NOT NULL AND longitude IS NOT NULL
    """)

    locations = cursor.fetchall()
    conn.close()

    return render_template("live_map.html", locations=locations)

@app.route("/update_case/<int:case_id>", methods=["GET","POST"])
def update_case(case_id):

    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    if request.method == "POST":

        feedback = request.form.get("feedback","").strip()
        status = request.form.get("status","Pending").strip()

        # 📷 photos update
        photos = []
        for i in range(1,6):
            file = request.files.get(f"photo{i}")

            if file and file.filename != "":
                filename = f"{int(time.time())}_{secure_filename(file.filename)}"
                filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(filepath)
                photos.append(filepath)
            else:
                photos.append(None)

        # 🔥 UPDATE QUERY
        cursor.execute("""
        UPDATE cases SET
            feedback=?,
            status=?,
            photo1=COALESCE(?, photo1),
            photo2=COALESCE(?, photo2),
            photo3=COALESCE(?, photo3),
            photo4=COALESCE(?, photo4),
            photo5=COALESCE(?, photo5)
        WHERE id=?
        """, (
            feedback, status,
            photos[0], photos[1], photos[2], photos[3], photos[4],
            case_id
        ))

        # 🔔 log
        cursor.execute("""
        INSERT INTO case_logs (case_id, message, created_at)
        VALUES (?, ?, datetime('now'))
        """, (case_id, f"Updated by {session['user']}"))

        conn.commit()
        conn.close()

        return redirect("/dashboard")

    # 🔍 GET CASE
    cursor.execute("SELECT * FROM cases WHERE id=?", (case_id,))
    case = cursor.fetchone()

    conn.close()

    return render_template("update_case.html", case=case)

@app.route("/attendance")
def attendance_page():
    return render_template("attendance.html")

@app.route("/map")
def map_view():
    return render_template("map.html")

@app.route("/get_agent_location")
def get_agent_location():

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("""
    SELECT latitude, longitude 
    FROM agents
    WHERE latitude IS NOT NULL
    LIMIT 1
    """)

    row = cursor.fetchone()
    conn.close()

    if row:
        return jsonify({"lat": row[0], "lng": row[1]})
    else:
        return jsonify({})

@app.route("/forgot_password", methods=["GET","POST"])
def forgot_password():

    if request.method == "POST":

        username = request.form["username"]
        new_password = generate_password_hash(request.form["password"])

        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute("SELECT * FROM users WHERE username=?", (username,))
        user = cursor.fetchone()

        if user:
            cursor.execute("UPDATE users SET password=? WHERE username=?",
                           (new_password, username))
            conn.commit()
            conn.close()

            return redirect(url_for("login"))

        else:
            conn.close()
            return render_template("forgot_password.html", error="User not found")

    return render_template("forgot_password.html")

@app.route("/profile", methods=["GET","POST"])
def profile():

    if "user" not in session:
        return redirect(url_for("login"))

    username = session["user"]

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("SELECT username, role FROM users WHERE username=?", (username,))
    user = cursor.fetchone()

    if request.method == "POST":

        new_password = generate_password_hash(request.form["password"])

        cursor.execute("UPDATE users SET password=? WHERE username=?",
                       (new_password, username))
        conn.commit()

        conn.close()
        return render_template("profile.html", user=user, success="Password Updated!")

    conn.close()

    return render_template("profile.html", user=user)

@app.route("/dashboard")
def dashboard():

    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    role = session.get("role")
    username = session.get("user")
    today = str(date.today())

    portal_filter = request.args.get("portal", "").upper()
    status_filter = request.args.get("status", "").upper()

    # =========================================
    # 📊 OVERALL COUNTS
    # =========================================
    cursor.execute("SELECT COUNT(*) FROM cases")
    total_cases = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM cases WHERE status='Completed'")
    completed_cases = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM cases WHERE status='Pending'")
    pending_cases = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM cases WHERE status='In Progress'")
    progress_cases = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM cases WHERE UPPER(TRIM(portal_type))='HDFC'")
    hdfc_count = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM cases WHERE UPPER(TRIM(portal_type))='MUTHOOT'")
    muthoot_count = cursor.fetchone()[0] or 0

    # =========================================
    # 📅 TODAY DATA
    # =========================================
    cursor.execute("SELECT COUNT(*) FROM cases WHERE initiation_date=?", (today,))
    today_total = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM cases WHERE status='Completed' AND initiation_date=?", (today,))
    today_completed = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM cases WHERE status='Pending' AND initiation_date=?", (today,))
    today_pending = cursor.fetchone()[0] or 0

    # =========================================
    # 🔍 FILTER
    # =========================================
    query = "SELECT * FROM cases WHERE 1=1"
    params = []

    if portal_filter:
        query += " AND UPPER(TRIM(portal_type))=?"
        params.append(portal_filter)

    if status_filter:
        query += " AND UPPER(TRIM(status))=?"
        params.append(status_filter)

    query += " ORDER BY id DESC LIMIT 20"

    cursor.execute(query, params)
    filtered_cases = cursor.fetchall()

    # =========================================
    # 🏦 RECENT CASES
    # =========================================
    cursor.execute("""
    SELECT * FROM cases
    WHERE UPPER(TRIM(portal_type))='HDFC'
    ORDER BY id DESC LIMIT 5
    """)
    recent_hdfc = cursor.fetchall()

    cursor.execute("""
    SELECT * FROM cases
    WHERE UPPER(TRIM(portal_type))='MUTHOOT'
    ORDER BY id DESC LIMIT 5
    """)
    recent_muthoot = cursor.fetchall()

    # =========================================
    # 👤 AGENT / ADMIN
    # =========================================
    leaderboard = []

    if role == "agent":

        cursor.execute("SELECT COUNT(*) FROM cases WHERE assigned_agent=?", (username,))
        total_cases = cursor.fetchone()[0] or 0

        cursor.execute("SELECT COUNT(*) FROM cases WHERE status='Completed' AND assigned_agent=?", (username,))
        completed_cases = cursor.fetchone()[0] or 0

        cursor.execute("SELECT COUNT(*) FROM cases WHERE status='Pending' AND assigned_agent=?", (username,))
        pending_cases = cursor.fetchone()[0] or 0

        cursor.execute("SELECT COUNT(*) FROM cases WHERE status='In Progress' AND assigned_agent=?", (username,))
        progress_cases = cursor.fetchone()[0] or 0

        cursor.execute("""
        SELECT strftime('%m', initiation_date), COUNT(*)
        FROM cases
        WHERE assigned_agent=?
        GROUP BY strftime('%m', initiation_date)
        """, (username,))
        monthly_raw = cursor.fetchall()

        agent_data = [(username, total_cases)]

    else:

        cursor.execute("""
        SELECT assigned_agent, COUNT(*)
        FROM cases
        WHERE status='Completed' AND initiation_date=?
        GROUP BY assigned_agent
        ORDER BY COUNT(*) DESC
        LIMIT 5
        """, (today,))
        leaderboard = cursor.fetchall()

        cursor.execute("""
        SELECT assigned_agent, COUNT(*)
        FROM cases
        GROUP BY assigned_agent
        """)
        agent_data = cursor.fetchall()

        cursor.execute("""
        SELECT strftime('%m', initiation_date), COUNT(*)
        FROM cases
        GROUP BY strftime('%m', initiation_date)
        """)
        monthly_raw = cursor.fetchall()

    # =========================================
    # 📊 MONTHLY CONVERT (🔥 IMPORTANT)
    # =========================================
    month_map = {
        "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
        "05": "May", "06": "Jun", "07": "Jul", "08": "Aug",
        "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"
    }

    months = []
    monthly_data = []

    for m, count in monthly_raw:
        months.append(month_map.get(m, m))
        monthly_data.append(count)

    # =========================================
    # 🔔 NOTIFICATIONS
    # =========================================
    cursor.execute("""
    SELECT message FROM notifications
    ORDER BY id DESC LIMIT 5
    """)
    notifications = cursor.fetchall()

    conn.close()

    # =========================================
    # 🎯 TARGET
    # =========================================
    target = TARGET_PER_DAY
    achieved = today_completed
    achievement_percent = int((achieved / target) * 100) if target > 0 else 0

    # =========================================
    # 🚀 FINAL RETURN
    # =========================================
    return render_template(
        "dashboard.html",

        total_cases=total_cases,
        completed_cases=completed_cases,
        pending_cases=pending_cases,
        progress_cases=progress_cases,

        today_total=today_total,
        today_completed=today_completed,
        today_pending=today_pending,

        target=target,
        achieved=achieved,
        achievement_percent=achievement_percent,

        recent_hdfc=recent_hdfc,
        recent_muthoot=recent_muthoot,

        filtered_cases=filtered_cases,
        agent_data=agent_data,

        months=months,
        monthly_data=monthly_data,

        role=role,
        notifications=notifications,
        leaderboard=leaderboard,

        hdfc_count=hdfc_count,
        muthoot_count=muthoot_count,

        portal_filter=portal_filter,
        status_filter=status_filter
    )

@app.route("/add_case", methods=["GET","POST"])
def add_case():

    if "user" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        try:

            # 🔹 FORM DATA
            ref_no = request.form.get("ref_no","").strip()
            client_name = request.form.get("client_name","").strip()
            cpv_type = request.form.get("cpv_type","").strip()

            # 🔥 PORTAL TYPE
            portal_type = request.form.get("portal_type", "HDFC")

            if portal_type == "MUTHOOT":
                bank_name = "MUTHOOT"
            else:
                bank_name = "HDFC"

            priority = request.form.get("priority","Normal").strip()
            agency = request.form.get("agency","").strip()
            pincode = request.form.get("pincode","").strip()
            address = request.form.get("address","").strip()

            status = request.form.get("status","Pending").strip().capitalize()
            if status not in ["Pending","Completed","In progress"]:
                status = "Pending"

            # 🕒 TIME
            now = datetime.now()
            initiation_date = now.strftime("%Y-%m-%d")
            created_time = now.strftime("%Y-%m-%d %H:%M:%S")
            assigned_time = created_time

            # 📍 LOCATION
            lat = request.form.get("lat")
            lng = request.form.get("lng")

            lat = float(lat) if lat else None
            lng = float(lng) if lng else None

            map_link = f"https://www.google.com/maps?q={lat},{lng}" if lat and lng else ""

            # 📷 PHOTOS
            photos = []
            for i in range(1,6):
                file = request.files.get(f"photo{i}")
                if file and file.filename != "":
                    filename = f"{int(time.time())}_{secure_filename(file.filename)}"
                    filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                    file.save(filepath)
                    photos.append(filepath)
                else:
                    photos.append("")

            # 👤 AUTO ASSIGN
            manual_agent = request.form.get("assigned_agent","").strip()
            if manual_agent:
                assigned_agent = manual_agent
            else:
                assigned_agent = auto_assign_agent(
                    lat, lng, priority, portal_type, pincode
                )

            # 🔢 CASE NUMBER
            case_number = generate_case_number(portal_type)

            # =========================================
            # ✅ DATABASE INSERT (NO DUPLICATE BLOCK ❌)
            # =========================================
            with sqlite3.connect(db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()

                cursor.execute("""
                INSERT INTO cases (
                    case_number, ref_no, application_no,
                    client_name, cpv_type, bank_name, portal_type,

                    address, landmark, city, state, pincode,

                    residence_type, family_status,
                    verification_status, final_status,

                    assigned_agent, status, priority,

                    initiation_date, created_time, assigned_time,

                    agency,

                    photo1, photo2, photo3, photo4, photo5,

                    latitude, longitude, map_link
                )
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    case_number, ref_no, "",

                    client_name, cpv_type, bank_name, portal_type,

                    address, "", "", "", pincode,

                    "", "", "", "",

                    assigned_agent, status, priority,

                    initiation_date, created_time, assigned_time,

                    agency,

                    photos[0], photos[1], photos[2], photos[3], photos[4],

                    lat, lng, map_link
                ))

                case_id = cursor.lastrowid

                # 🔥 LOG
                cursor.execute("""
                INSERT INTO case_logs (case_id, message, created_at)
                VALUES (?, ?, datetime('now'))
                """, (case_id, f"{portal_type} Case Created"))

                # 🔔 NOTIFICATION
                cursor.execute("""
                INSERT INTO notifications (message, created_at)
                VALUES (?, datetime('now'))
                """, (
                    f"New {portal_type} case added for {client_name} (Agent: {assigned_agent})",
                ))

                conn.commit()

            
            # 🔁 REDIRECT
            if portal_type == "MUTHOOT":
                return redirect(url_for("muthoot_cases"))
            else:
                return redirect(url_for("hdfc_cases"))

        except Exception as e:
            print("❌ ERROR:", e)
            return f"❌ ERROR: {e}"

    return render_template("add_case.html")

def auto_reassign_cases():

    try:
        with sqlite3.connect(db_path, timeout=20) as conn:
            cursor = conn.cursor()

            cursor.execute("""
                SELECT id, assigned_agent, assigned_time, pincode
                FROM cases
                WHERE status='Pending'
            """)

            cases = cursor.fetchall()

            for case in cases:

                case_id = case[0]
                old_agent = case[1]
                assigned_time = case[2]
                pincode = case[3]

                if not assigned_time:
                    continue

                try:
                    assigned_dt = datetime.strptime(assigned_time, "%Y-%m-%d %H:%M:%S")
                except:
                    continue

                # ⏰ 30 mins rule
                if datetime.now() - assigned_dt > timedelta(minutes=30):

                    # ✅ SIMPLE SAFE ASSIGN (NO LOCK ISSUE)
                    new_agent = auto_assign_agent()

                    if new_agent != old_agent:

                        cursor.execute("""
                            UPDATE cases
                            SET assigned_agent=?, assigned_time=?
                            WHERE id=?
                        """, (
                            new_agent,
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            case_id
                        ))

                        print(f"🔥 Reassigned Case {case_id} → {new_agent}")

    except Exception as e:
        print("Auto Reassign Error:", e)

@app.route("/upload_excel", methods=["GET", "POST"])
def upload_excel():

    if "user" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":

        file = request.files.get("file")

        if not file or file.filename == "":
            return "❌ No file uploaded"

        import pandas as pd
        from datetime import datetime

        try:
            # ✅ READ EXCEL
            df = pd.read_excel(file)

            # ✅ CLEAN HEADERS
            df.columns = df.columns.str.strip().str.lower()

            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()

            count = 0

            for _, row in df.iterrows():

                try:
                    # 🔥 SAFE GET (avoid None / NaN issues)
                    case_number = str(row.get("application number", "")).strip()
                    client_name = str(row.get("applicant name", "")).strip()
                    cpv_type = str(row.get("cpv type", "Residence")).strip()
                    address = str(row.get("address", "")).strip()
                    city = str(row.get("city", "")).strip()
                    pincode = str(row.get("pincode", "")).strip()
                    status = str(row.get("status", "Pending")).strip()

                    # ❌ SKIP EMPTY ROW
                    if case_number == "" or case_number.lower() == "nan":
                        continue

                    # 🔥 PORTAL AUTO DETECT
                    filename = file.filename.lower()
                    if "hdfc" in filename:
                        portal_type = "HDFC"
                    elif "muthoot" in filename:
                        portal_type = "MUTHOOT"
                    else:
                        portal_type = "HDFC"  # default

                    # 🕒 DATE
                    now = datetime.now()
                    initiation_date = now.strftime("%Y-%m-%d")
                    created_time = now.strftime("%Y-%m-%d %H:%M:%S")

                    # 🔥 AUTO AGENT (safe fallback)
                    try:
                        assigned_agent = auto_assign_agent(None, None, "Normal", portal_type)
                    except:
                        assigned_agent = "agent1"

                    # 🚫 DUPLICATE CHECK
                    cursor.execute(
                        "SELECT id FROM cases WHERE case_number=?",
                        (case_number,)
                    )
                    exists = cursor.fetchone()

                    if exists:
                        continue  # skip duplicate

                    # ✅ INSERT
                    cursor.execute("""
                    INSERT INTO cases (
                        case_number, client_name, cpv_type,
                        address, city, pincode,
                        status, portal_type,
                        initiation_date, created_time,
                        assigned_agent
                    )
                    VALUES (?,?,?,?,?,?,?,?,?,?,?)
                    """, (
                        case_number,
                        client_name,
                        cpv_type,
                        address,
                        city,
                        pincode,
                        status,
                        portal_type,
                        initiation_date,
                        created_time,
                        assigned_agent
                    ))

                    count += 1

                except Exception as row_error:
                    print("❌ Row Error:", row_error)

            conn.commit()
            conn.close()

            return f"✅ {count} cases imported successfully!"

        except Exception as e:
            return f"❌ Error: {str(e)}"

    return render_template("upload_excel.html")

@app.route("/search_hdfc", methods=["POST"])
def search_hdfc():

    ref_no = request.form.get("ref_no", "").strip()

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id FROM cases
        WHERE ref_no=? AND UPPER(TRIM(portal_type))='HDFC'
    """, (ref_no,))

    case = cursor.fetchone()
    conn.close()

    if case:
        return redirect(f"/hdfc_form/{case['id']}")
    else:
        return "<h4 style='color:red'>❌ Case not found</h4>"

@app.route("/mark_completed/<int:id>")
def mark_completed(id):

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("""
    UPDATE cases
    SET status='Completed',
        final_status='Completed'
    WHERE id=?
    """, (id,))

    conn.commit()
    conn.close()

    return redirect(url_for("dashboard"))

@app.route("/hdfc_dashboard")
def hdfc_dashboard():

    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # =========================================
    # 🔥 HDFC CASES
    # =========================================
    cursor.execute("""
    SELECT * FROM cases
    WHERE UPPER(TRIM(portal_type))='HDFC'
    ORDER BY id DESC
    """)
    cases = cursor.fetchall()

    cursor.execute("""
    SELECT COUNT(*) FROM cases
    WHERE UPPER(TRIM(portal_type))='HDFC'
    """)
    total = cursor.fetchone()[0] or 0

    cursor.execute("""
    SELECT COUNT(*) FROM cases
    WHERE UPPER(TRIM(portal_type))='HDFC'
    AND UPPER(TRIM(status))='COMPLETED'
    """)
    completed = cursor.fetchone()[0] or 0

    cursor.execute("""
    SELECT COUNT(*) FROM cases
    WHERE UPPER(TRIM(portal_type))='HDFC'
    AND UPPER(TRIM(status))='PENDING'
    """)
    pending = cursor.fetchone()[0] or 0

    conn.close()

    return render_template(
        "hdfc_dashboard.html",
        cases=cases,
        total=total,
        completed=completed,
        pending=pending
    )

@app.route("/muthoot_dashboard")
def muthoot_dashboard():

    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # 🔥 ONLY MUTHOOT DATA
    cursor.execute("""
    SELECT * FROM cases
    WHERE UPPER(TRIM(portal_type))='MUTHOOT'
    ORDER BY id DESC
    """)
    cases = cursor.fetchall()

    # 🔥 COUNTS
    cursor.execute("""
    SELECT COUNT(*) FROM cases
    WHERE portal_type='MUTHOOT'
    """)
    total = cursor.fetchone()[0]

    cursor.execute("""
    SELECT COUNT(*) FROM cases
    WHERE portal_type='MUTHOOT' AND status='Completed'
    """)
    completed = cursor.fetchone()[0]

    cursor.execute("""
    SELECT COUNT(*) FROM cases
    WHERE portal_type='MUTHOOT' AND status='Pending'
    """)
    pending = cursor.fetchone()[0]

    conn.close()

    return render_template(
        "muthoot_dashboard.html",
        cases=cases,
        total=total,
        completed=completed,
        pending=pending
    )

@app.route("/case_summary/<int:id>")
def case_summary(id):

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("SELECT client_name, bank_name, cpv_type, status FROM cases WHERE id=?", (id,))
    case = cursor.fetchone()

    conn.close()

    summary = generate_summary(case[0], case[1], case[2], case[3])

    return render_template("summary.html", summary=summary)

@app.route("/case/<int:case_id>")
def case_detail(case_id):

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM cases WHERE id=?", (case_id,))
    case = cursor.fetchone()

    conn.close()

    return render_template("case_detail.html", case=case)

@app.route("/search_case")
def search_case():

    ref_no = request.args.get("ref_no")

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
    SELECT * FROM cases WHERE ref_no=?
    """, (ref_no,))

    case = cursor.fetchone()

    conn.close()

    if case:
        return redirect(url_for("case_form", id=case["id"]))
    else:
        return "❌ Case not found"

@app.route("/view_case")
def view_case():

    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    role = session.get("role")
    username = session.get("user")

    # 🔍 GET VALUES
    search = request.args.get("search")
    status_filter = request.args.get("status")

    # ✅ BASE QUERY
    if role == "agent":
        base_query = "SELECT * FROM cases WHERE assigned_agent=?"
        params = [username]
    else:
        base_query = "SELECT * FROM cases WHERE 1=1"
        params = []

    # 🔥 SEARCH (UPDATED)
    if search:
        base_query += """
        AND (
            client_name LIKE ? OR
            ref_no LIKE ? OR
            application_no LIKE ?
        )
        """
        params.extend([
            '%' + search + '%',
            '%' + search + '%',
            '%' + search + '%'
        ])

    # ✅ STATUS FILTER
    if status_filter:
        base_query += " AND status=?"
        params.append(status_filter)

    # ✅ ORDER
    base_query += " ORDER BY id DESC"

    # ✅ EXECUTE
    cursor.execute(base_query, tuple(params))
    cases = cursor.fetchall()

    conn.close()

    return render_template(
        "view_case.html",
        cases=cases,
        search=search,
        status_filter=status_filter,
        role=role
    )

@app.route("/view_case/<int:case_id>")
def view_case_detail(case_id):

    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM cases WHERE id=?", (case_id,))
    case = cursor.fetchone()

    conn.close()

    if not case:
        return "❌ Case Not Found"

    return render_template("view_case_detail.html", case=case)

@app.route("/muthoot_form/<int:case_id>")
def muthoot_form(case_id):

    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
    SELECT * FROM cases 
    WHERE id=? AND portal_type='MUTHOOT'
    """, (case_id,))

    case = cursor.fetchone()
    conn.close()

    if not case:
        return "❌ Muthoot case not found"

    return render_template("muthoot_form.html", case=case)

def get_column(df, possible_names):
    for col in df.columns:
        if col.lower().strip() in possible_names:
            return col
    return None

@app.route("/muthoot_cases")
def muthoot_cases():

    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    status = request.args.get("status", "").strip()
    search = request.args.get("search", "").strip()

    # 🔥 BASE QUERY
    query = "SELECT * FROM cases WHERE UPPER(TRIM(portal_type))='MUTHOOT'"
    params = []

    # 🔍 STATUS FILTER
    if status:
        query += " AND UPPER(TRIM(status))=?"
        params.append(status.upper())

    # 🔍 SEARCH FILTER (case-insensitive 🔥)
    if search:
        query += " AND LOWER(client_name) LIKE ?"
        params.append(f"%{search.lower()}%")

    query += " ORDER BY id DESC"

    cursor.execute(query, params)
    cases = cursor.fetchall()

    conn.close()

    return render_template("muthoot_cases.html", cases=cases)

@app.route("/edit_muthoot/<int:id>", methods=["GET", "POST"])
def edit_muthoot(id):

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    if request.method == "POST":

        customer_name = request.form["customer_name"]
        address = request.form["address"]
        city = request.form["city"]
        status = request.form["status"]
        final_status = request.form["final_status"]

        cursor.execute("""
        UPDATE cases
        SET client_name=?,
            address=?,
            city=?,
            status=?,
            final_status=?
        WHERE id=? AND portal_type='MUTHOOT'
        """, (customer_name, address, city, status, final_status, id))

        conn.commit()
        conn.close()

        return redirect(url_for("view_case"))

    cursor.execute("SELECT * FROM cases WHERE id=? AND portal_type='MUTHOOT'", (id,))
    case = cursor.fetchone()
    conn.close()

    return render_template("edit_muthoot.html", case=case)

@app.route("/hdfc_form/<int:case_id>")
def hdfc_form(case_id):

    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # ✅ FIXED LINE
    cursor.execute("SELECT * FROM cases WHERE id=?", (case_id,))
    case = cursor.fetchone()

    conn.close()

    if not case:
        return "❌ Case not found"

    return render_template("hdfc_form.html", case=case)

@app.route("/hdfc_cases")
def hdfc_cases():

    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    status = request.args.get("status", "").strip()
    search = request.args.get("search", "").strip()

    # 🔥 BASE QUERY
    query = "SELECT * FROM cases WHERE UPPER(TRIM(portal_type))='HDFC'"
    params = []

    # 🔍 STATUS FILTER
    if status:
        query += " AND UPPER(TRIM(status))=?"
        params.append(status.upper())

    # 🔍 SEARCH FILTER
    if search:
        query += " AND LOWER(client_name) LIKE ?"
        params.append(f"%{search.lower()}%")

    query += " ORDER BY id DESC"

    cursor.execute(query, params)
    cases = cursor.fetchall()

    conn.close()

    return render_template("hdfc_cases.html", cases=cases)

@app.route("/upload_cases", methods=["GET","POST"])
def upload_cases():

    if "user" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":

        file = request.files.get("file")

        if not file:
            return "No file uploaded"

        # ✅ Read Excel
        df = pd.read_excel(file)

        # ✅ Clean columns
        df.columns = df.columns.str.strip().str.lower()

        print("📊 Columns Found:", df.columns)

        if df.empty:
            return "Excel file is empty"

        # ✅ Replace NaN
        df = df.fillna("")

        # 🔥 Portal detect (FIXED LOWERCASE)
        if "application_no" in df.columns or "application number" in df.columns:
            portal_type = "muthoot"
        else:
            portal_type = "hdfc"

        print("🔥 Portal Type:", portal_type)

        # 🔥 SMART COLUMN DETECT
        client_col = get_column(df, COLUMN_MAPPING["client_name"])
        ref_col = get_column(df, COLUMN_MAPPING["ref_no"])
        cpv_col = get_column(df, COLUMN_MAPPING["cpv_type"])
        bank_col = get_column(df, COLUMN_MAPPING["bank_name"])
        priority_col = get_column(df, COLUMN_MAPPING["priority"])
        date_col = get_column(df, COLUMN_MAPPING["initiation_date"])
        agency_col = get_column(df, COLUMN_MAPPING["agency"])
        agent_col = get_column(df, COLUMN_MAPPING["assigned_agent"])
        pin_col = get_column(df, COLUMN_MAPPING["pincode"])
        status_col = get_column(df, COLUMN_MAPPING["status"])

        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # 🔁 LOOP ROWS
        for index, row in df.iterrows():

            try:
                # ✅ SAFE AGENT ASSIGN
                assigned_agent = (
                    str(row.get(agent_col)).strip()
                    if agent_col and str(row.get(agent_col)).strip() != ""
                    else auto_assign_agent()
                )

                cursor.execute("""
                INSERT INTO cases
                (case_number, ref_no, client_name, cpv_type, bank_name,
                 priority, initiation_date, agency, assigned_agent,
                 pincode, status, portal_type)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
                """, (

                generate_case_number(),

                str(row.get(ref_col, "")),
                str(row.get(client_col, "")),
                str(row.get(cpv_col, "")),
                str(row.get(bank_col, "")),
                str(row.get(priority_col, "")),
                str(row.get(date_col, "")),
                str(row.get(agency_col, "")),
                assigned_agent,
                str(row.get(pin_col, "")),
                str(row.get(status_col, "")),
                portal_type   # ✅ FINAL FIX
                ))

            except Exception as e:
                print(f"❌ Row {index} Error:", e)

        conn.commit()
        conn.close()

        print("✅ Upload Completed Successfully")

        return redirect(url_for("dashboard"))

    return render_template("upload_cases.html")

@app.route("/case_details/<int:id>")
def case_details(id):

    conn = sqlite3.connect(db_path, timeout=10)
    cursor = conn.cursor()

    # 🔹 CASE DATA
    cursor.execute("SELECT * FROM cases WHERE id=?", (id,))
    case = cursor.fetchone()

    # 🔥 CASE LOGS (IMPORTANT FIX)
    cursor.execute("""
    SELECT * FROM case_logs
    WHERE case_id=?
    ORDER BY created_at DESC
    """, (id,))
    logs = cursor.fetchall()

    conn.close()

    # 🔥 SEND BOTH case + logs
    return render_template(
        "case_detail.html",
        case=case,
        logs=logs
    )

@app.route("/submit_hdfc/<int:case_id>", methods=["POST"])
def submit_hdfc(case_id):

    if "user" not in session:
        return redirect(url_for("login"))

    try:
        mobile = request.form.get("mobile")
        contact_person = request.form.get("contact_person")
        remarks = request.form.get("remarks")

        status = request.form.get("status", "Completed")
        feedback = request.form.get("final_feedback", "")

        # 📸 photo
        photo1 = request.files.get("photo1")
        photo_path = ""

        if photo1 and photo1.filename:
            filename = secure_filename(photo1.filename)
            path = os.path.join("static/uploads", filename)
            photo1.save(path)
            photo_path = path

        with sqlite3.connect(db_path) as conn:
            cursor = conn.cursor()

            cursor.execute("""
            UPDATE cases SET
                status=?,
                feedback=?,
                photo1=?
            WHERE id=?
            """, (status, feedback, photo_path, case_id))

            conn.commit()

        return redirect(url_for("hdfc_cases"))

    except Exception as e:
        print("❌ ERROR:", e)
        return str(e)

@app.route("/submit_muthoot/<int:case_id>", methods=["POST"])
def submit_muthoot(case_id):

    if "user" not in session:
        return redirect(url_for("login"))

    try:
        # 🔹 FORM DATA
        residence_type = request.form.get("residence_type", "").strip()
        locality = request.form.get("locality", "").strip()
        accessibility = request.form.get("accessibility", "").strip()
        person_met = request.form.get("person_met", "").strip()
        years_at_residence = request.form.get("years_at_residence", "").strip()
        ownership = request.form.get("ownership", "").strip()

        verification_status = request.form.get("verification_status", "").strip()
        final_status = request.form.get("final_status", "").strip()
        visit_date = request.form.get("visit_date")

        # 📍 LOCATION
        lat = request.form.get("lat")
        lng = request.form.get("lng")

        lat = float(lat) if lat else None
        lng = float(lng) if lng else None

        # 📸 DOCUMENT PHOTO (SAFE NAME 🔥)
        doc = request.files.get("document_photo")
        doc_path = ""

        if doc and doc.filename != "":
            filename = f"{int(time.time())}_{secure_filename(doc.filename)}"
            path = os.path.join("static/uploads", filename)
            doc.save(path)
            doc_path = path

        # 🤖 AUTO
        status = "Completed"

        # 🧠 FEEDBACK
        if verification_status == "Positive" and final_status == "Eligible":
            feedback = "Customer verified successfully"
        else:
            feedback = "Verification failed"

        with sqlite3.connect(db_path) as conn:
            cursor = conn.cursor()

            # 🔥 CHECK CASE EXIST
            cursor.execute("SELECT id FROM cases WHERE id=?", (case_id,))
            if not cursor.fetchone():
                return "❌ Case not found"

            # ✅ UPDATE (ONLY EXISTING COLUMNS 🔥)
            cursor.execute("""
            UPDATE cases SET
                residence_type=?,
                verification_status=?,
                final_status=?,
                latitude=?,
                longitude=?,
                status=?,
                feedback=?
            WHERE id=?
            """, (
                residence_type,
                verification_status,
                final_status,
                lat,
                lng,
                status,
                feedback,
                case_id
            ))

            # 🔥 LOG
            cursor.execute("""
            INSERT INTO case_logs (case_id, message, created_at)
            VALUES (?, ?, datetime('now'))
            """, (case_id, "Muthoot Verification Completed"))

            conn.commit()

        return redirect(url_for("muthoot_cases"))

    except Exception as e:
        print("❌ ERROR:", e)
        return f"Error: {e}"

@app.route("/edit_case/<int:id>", methods=["GET","POST"])
def edit_case(id):

    if "user" not in session:
        return redirect(url_for("login"))

    if session.get("role") != "admin":
        return "Access Denied"

    with sqlite3.connect(db_path, timeout=10) as conn:
     cursor = conn.cursor()

    if request.method == "POST":

        status = request.form["status"]
        priority = request.form["priority"]
        assigned_agent = request.form["assigned_agent"]

        # 🔥 UPDATE CASE
        cursor.execute("""
        UPDATE cases
        SET status=?, priority=?, assigned_agent=?
        WHERE id=?
        """,(status,priority,assigned_agent,id))

        # 🔥 LOG INSERT (ADD HERE ✅)
        cursor.execute("""
        INSERT INTO case_logs (case_id, message, created_at)
        VALUES (?, ?, datetime('now'))
        """, (
            id,   # ✅ case_id
            f"Status changed to {status}"
        ))

        conn.commit()
        conn.close()

        return redirect(url_for("view_case"))

    cursor.execute("SELECT * FROM cases WHERE id=?", (id,))
    case = cursor.fetchone()

    conn.close()

    return render_template("edit_case.html", case=case)

@app.route("/delete_case/<int:id>")
def delete_case(id):

    if "user" not in session:
        return redirect(url_for("login"))

    try:
        # ✅ ONLY THIS BLOCK (no extra conn outside)
        with sqlite3.connect(db_path, timeout=20) as conn:
            cursor = conn.cursor()

            cursor.execute("DELETE FROM cases WHERE id=?", (id,))

            conn.commit()   # 🔥 IMPORTANT

        return redirect(url_for("dashboard"))

    except Exception as e:
        print("❌ DELETE ERROR:", e)
        return "Error deleting case"

@app.route("/export_excel")
def export_excel():

    if "user" not in session:
        return redirect(url_for("login"))

    portal = request.args.get("portal", "").upper()  # 🔥 filter optional

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 🔥 DYNAMIC QUERY
    query = """
    SELECT case_number, ref_no, client_name, cpv_type,
           bank_name, priority, initiation_date,
           agency, assigned_agent, pincode, status
    FROM cases
    WHERE 1=1
    """

    params = []

    if portal:
        query += " AND UPPER(TRIM(portal_type))=?"
        params.append(portal)

    query += " ORDER BY id DESC"

    cursor.execute(query, params)
    data = cursor.fetchall()

    conn.close()

    # 📄 EXCEL
    wb = Workbook()
    ws = wb.active

    ws.title = "CPV Report"

    # ✅ HEADERS
    ws.append([
        "Case No", "Ref No", "Client Name", "CPV Type",
        "Bank", "Priority", "Date",
        "Agency", "Agent", "Pincode", "Status"
    ])

    # ✅ DATA
    for row in data:
        ws.append(row)

    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(temp.name)
    temp.close()

    return send_file(
        temp.name,
        as_attachment=True,
        download_name=f"{portal if portal else 'ALL'}_CPV_Report.xlsx"
    )

@app.route("/download_hdfc_report")
def download_hdfc_report():

    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(db_path)

    # 🔥 SAFE FILTER (important)
    query = """
    SELECT case_number, ref_no, client_name, cpv_type, status, priority,
           initiation_date, agency, assigned_agent, pincode
    FROM cases
    WHERE UPPER(TRIM(portal_type))='HDFC'
    ORDER BY id DESC
    """

    df = pd.read_sql_query(query, conn)
    conn.close()

    # ✅ RENAME COLUMNS (clean header)
    df.columns = [
        "Case No", "Ref No", "Customer Name", "CPV Type",
        "Status", "Priority", "Initiation Date",
        "Agency", "Assigned Agent", "Pincode"
    ]

    file_path = "hdfc_report.xlsx"
    df.to_excel(file_path, index=False)

    return send_file(file_path, as_attachment=True)

@app.route("/download_muthoot_report")
def download_muthoot_report():

    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(db_path)

    query = """
    SELECT case_number, ref_no, client_name, cpv_type, status, priority,
           initiation_date, agency, assigned_agent, pincode,
           verification_status, final_status
    FROM cases
    WHERE UPPER(TRIM(portal_type))='MUTHOOT'
    ORDER BY id DESC
    """

    df = pd.read_sql_query(query, conn)
    conn.close()

    df.columns = [
        "Case No", "Ref No", "Customer Name", "CPV Type",
        "Status", "Priority", "Initiation Date",
        "Agency", "Assigned Agent", "Pincode",
        "Verification", "Final Status"
    ]

    file_path = "muthoot_report.xlsx"
    df.to_excel(file_path, index=False)

    return send_file(file_path, as_attachment=True)

@app.route("/download_report")
def download_report():

    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(db_path)

    role = session.get("role")
    username = session.get("user")

    if role == "agent":
        query = "SELECT * FROM cases WHERE assigned_agent=?"
        df = pd.read_sql_query(query, conn, params=(username,))
    else:
        query = "SELECT * FROM cases"
        df = pd.read_sql_query(query, conn)

    conn.close()

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Report')

    output.seek(0)

    return send_file(
        output,
        download_name="CPV_Report.xlsx",
        as_attachment=True
    )

@app.route("/logout")
def logout():

    session.pop("user", None)
    return redirect(url_for("login"))


if __name__ == "__main__":

    # ✅ INIT DB
    init_db()

    # ✅ AUTO REASSIGN THREAD
    def run_auto_reassign():
        import time
        time.sleep(5)

        while True:
            try:
                auto_reassign_cases()
            except Exception as e:
                print("Auto Reassign Error:", e)

            time.sleep(60)

    # ✅ IMPORTANT CHANGE
    app.run(host="0.0.0.0", port=5000)

   
   
       
    