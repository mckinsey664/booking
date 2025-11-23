from flask import Flask, render_template, request, redirect, url_for, session, flash
import sqlite3, random, string
from functools import wraps
from flask import abort
from flask import jsonify
from flask import request
import gspread
from google.oauth2.service_account import Credentials
import smtplib
from email.message import EmailMessage
from datetime import datetime, timedelta


SMTP_EMAIL = "lynn.m@mckinsey-electronics.com"
SMTP_PASSWORD = "kpuf hgzt yycc tcan"   # Gmail app password

app = Flask(__name__)
app.secret_key = "yoursecretkey"

ADMIN_EMAIL = "admin@mckinsey-electronics.com"  # admin account that skips verification

######################## GOOGLE SHEETS INTEGRATION #########################
# ---- Google Sheets setup ----

import json, os

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds_json = json.loads(os.environ["GOOGLE_CREDS"])
CREDS = Credentials.from_service_account_info(creds_json, scopes=SCOPES)

gc = gspread.authorize(CREDS)


SHEET_ID = "16gIa4CoAbnNhlQrqL1xu21VwmAnrqFaQ7KThpxmz8xs"  # part after /d/ and before /edit in the sheet URL
SHEET_NAME = "DSS Form Test 2"            # change if your tab is named differently


def get_users_from_sheets():
    """Fetch only approved users (column R == 'yes') and keep selected columns"""
    sh = gc.open_by_key(SHEET_ID)
    worksheet = sh.worksheet(SHEET_NAME)

    # Get all rows including header
    rows = worksheet.get_all_values()

    if not rows:
        return []

    headers = rows[0]
    # Map header names to their column index
    header_map = {h.strip(): i for i, h in enumerate(headers)}

    desired_keys = ["First name", "Last name", "Email", "Company name", "Position"]

    # Column R is the 18th column (0-based index 17) — adjust if needed
    # Or safer: just use absolute index if you know it's column R
    APPROVE_COL_INDEX = 17  # column R = index 17 (A=0,B=1,...,R=17)

    filtered = []
    for row in rows[1:]:  # skip header
        #if len(row) > APPROVE_COL_INDEX and str(row[APPROVE_COL_INDEX]).strip().lower() == "yes":
            # Only approved rows
            filtered_row = {}
            for key in desired_keys:
                if key in header_map and header_map[key] < len(row):
                    filtered_row[key] = row[header_map[key]]
                else:
                    filtered_row[key] = ""
            filtered.append(filtered_row)

    return filtered

def send_plain_email(to_email, subject, message):
    try:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = SMTP_EMAIL
        msg["To"] = to_email
        msg.set_content(message)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(SMTP_EMAIL, SMTP_PASSWORD)
            smtp.send_message(msg)

        print(f"📧 Email sent to {to_email}")
        return True
    except Exception as e:
        print("❌ Email failed:", e)
        return False

def get_users_from_sheets2():
    """
    Fetch ALL columns for approved users (column R == 'yes').
    Returns a list of dicts where keys are the header names from the sheet.
    """
    sh = gc.open_by_key(SHEET_ID)
    worksheet = sh.worksheet(SHEET_NAME)

    rows = worksheet.get_all_values()
    if not rows:
        return []

    headers = [h.strip() for h in rows[0]]  # use full header row
    APPROVE_COL_INDEX = 17  # column R = index 17 (A=0,B=1,...,R=17)

    approved = []
    for row in rows[1:]:  # skip header
        if len(row) > APPROVE_COL_INDEX and str(row[APPROVE_COL_INDEX]).strip().lower() == "yes":
            # Build a dict for this row using all headers
            data = {}
            for idx, header in enumerate(headers):
                if idx < len(row):
                    data[header] = row[idx]
                else:
                    data[header] = ""
            approved.append(data)

    return approved

# FOR LOGIN
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "email" not in session:
            # user is not logged in
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function

def verified_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "email" not in session:
            return redirect(url_for("login"))
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT verified FROM users WHERE email=?", (session["email"],))
        user = c.fetchone()
        if not user or user["verified"] != 1:
            flash("You must verify your email first.")
            return redirect(url_for("verify"))
        return f(*args, **kwargs)
    return decorated_function

# def get_db():
#     conn = sqlite3.connect("rooms.db")
#     conn.row_factory = sqlite3.Row
#     return conn

def get_db():
    DB_PATH = "/var/data/rooms.db"
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

@app.route("/debug/schema")
def debug_schema():
    conn = get_db()
    c = conn.cursor()

    tables = c.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()

    output = ""
    for t in tables:
        table_name = t[0]
        output += f"\n===== {table_name} =====\n"
        rows = c.execute(f"PRAGMA table_info({table_name})").fetchall()
        for r in rows:
            output += f"{r}\n"

    return f"<pre>{output}</pre>"


def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        # Check if the user is logged in and is the admin
        if session.get("email") != ADMIN_EMAIL:
            flash("⛔ Access denied. Admins only.", "danger")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function

@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form["email"].strip().lower()

        # ✅ 1️⃣ Admin bypass: check FIRST before any database logic
        if email == "admin@mckinsey-electronics.com":
            session.clear()
            session["email"] = email
            flash("✅ Logged in as Administrator!", "success")
            return redirect(url_for("admin_dashboard"))

        # ✅ 2️⃣ Continue normal login for everyone else
        conn = get_db()
        conn.row_factory = sqlite3.Row
        c = conn.cursor()

        # ✅ Check if email exists in approved_users
        c.execute("SELECT * FROM approved_users WHERE lower(email)=?", (email,))
        approved_user = c.fetchone()

        # ✅ Check if email exists in company_contacts
        c.execute("SELECT company_id FROM company_contacts WHERE lower(email)=?", (email,))
        contact_user = c.fetchone()

        # ❌ Reject if not in either table
        if not approved_user and not contact_user:
            flash("❌ This email is not approved. Please contact the administrator.", "danger")
            conn.close()
            return redirect(url_for("login"))

        # ✅ Insert in local users table if not already there
        c.execute("INSERT OR IGNORE INTO users(email) VALUES(?)", (email,))
        conn.commit()

        # Generate 4-digit code
        import random
        code = f"{random.randint(1000,9999)}"

        # Save temporary session info
        session["pending_email"] = email
        session["code"] = code

        # Send verification email using Gmail SMTP
        sender_email = "lynn.m@mckinsey-electronics.com"
        sender_password = "kpuf hgzt yycc tcan"  # Gmail App Password ONLY

        msg = EmailMessage()
        msg["Subject"] = "Your verification code"
        msg["From"] = sender_email
        msg["To"] = email
        msg.set_content(f"Your verification code is: {code}")

        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(sender_email, sender_password)
                smtp.send_message(msg)
        except Exception as e:
            print("❌ Email failed:", e)

        conn.close()

        flash("✅ Verification code sent to your email.", "success")
        return redirect(url_for("verify"))

    return render_template("login.html")

@app.route("/verify", methods=["GET", "POST"])
def verify():
    # If no pending_email (user not through login), redirect to login
    if "pending_email" not in session:
        return redirect(url_for("login"))

    # If the pending email is the admin, just log in immediately
    if session.get("pending_email", "").lower() == ADMIN_EMAIL:
        session["email"] = ADMIN_EMAIL
        session.pop("pending_email", None)
        session.pop("code", None)
        session["company_id"] = None
        flash("✅ Logged in as admin.", "success")
        return redirect(url_for("admin_dashboard"))

    if request.method == "POST":
        code_entered = request.form.get("code", "").strip()
        if code_entered == session.get("code"):
            email = session.pop("pending_email")
            session["email"] = email
            session.pop("code", None)

            # ✅ Mark the user as verified
            conn = get_db()
            conn.row_factory = sqlite3.Row
            c = conn.cursor()
            c.execute("UPDATE users SET verified = 1 WHERE email = ?", (email,))
            conn.commit()

            # 🔍 Check if this user is a company contact
            c.execute("SELECT company_id FROM company_contacts WHERE lower(email)=?", (email.lower(),))
            contact = c.fetchone()
            if contact:
                session["company_id"] = contact["company_id"]
                flash("✅ Logged in as company contact!", "success")
                return redirect(url_for("company_slots_user"))
            else:
                session["company_id"] = None
                flash("✅ Login successful!", "success")
                return redirect(url_for("reserve"))
        else:
            flash("❌ Invalid verification code. Please try again.", "danger")

    return render_template("verify.html")

@app.route("/reserve", methods=["GET", "POST"])
@verified_required
def reserve():
    email = session.get("email")
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Ensure user is approved
    c.execute("SELECT id, first_name, last_name FROM approved_users WHERE lower(email)=?", (email,))
    user = c.fetchone()
    if not user:
        return redirect(url_for("login"))
    user_id = user["id"]
    requester_full_name = f"{user['first_name']} {user['last_name']}"

    # # Load companies
    # companies = c.execute("""
    #     SELECT DISTINCT company_name 
    #     FROM approved_users
    #     WHERE company_name IS NOT NULL AND company_name <> ''
    #     ORDER BY company_name
    # """).fetchall()

    # Allowed exhibitor list
    allowed_companies = [
    "Amphenol Corporation",
    "ams-OSRAM AG",
    "Avnet Silica",
    "Diotec Semiconductor AG",
    "Energizer Holdings",
    "Epson Europe Electronics GmbH",
    "FoxEMS",
    "Holtek Semiconductor Inc.",
    "Keysight Technologies",
    "Marvell Technology, Inc.",
    "Nuvoton Technology Corporation",
    "Quectel Wireless Solutions",
    "Renesas Electronics Europe GmbH",
    "Samtec Europe Ltd",
    "SIMCom Wireless Solutions",
    "STMicroelectronics",
    "Swissbit AG",
    "Viking Tech Europe",
    "WAGO Middle East FZC",
    "Wurth Elektronik eiSos GmbH & Co. KG",
    "BPM"
    ]

    # Load only allowed companies
    companies = c.execute("""
    SELECT DISTINCT company_name
    FROM approved_users
    WHERE company_name IN ({placeholders})
    ORDER BY company_name
    """.format(placeholders=",".join("?" * len(allowed_companies))),
    allowed_companies).fetchall()


    selected_company = request.args.get("company_name")
    entity_id = request.args.get("entity_id")
    selected_date = request.args.get("date")

    # Load people
    people = []
    if selected_company:
        people = c.execute("""
            SELECT id, first_name, last_name, email 
            FROM approved_users
            WHERE company_name = ?
            ORDER BY first_name, last_name
        """, (selected_company,)).fetchall()

    # POST – Reserve
    if request.method == "POST":
        chosen_time = request.form.get("time")
        slot_id = request.form.get("slot_id")
        entity_id = request.form.get("entity_id")
        selected_date = request.form.get("date")

        if not entity_id:
            flash("❌ Please select a person.", "danger")
            return redirect(url_for("reserve"))

        # Check double booking
        c.execute("""
            SELECT id FROM reservations
            WHERE entity_type='person' AND entity_id=? 
              AND date=? AND start_time=? 
              AND status IN ('Pending', 'Approved')
        """, (entity_id, selected_date, chosen_time))
        if c.fetchone():
            flash("❌ This slot is already booked.", "danger")
            return redirect(url_for("reserve",
                                    company_name=selected_company,
                                    entity_id=entity_id,
                                    date=selected_date))

        # Assign room
        rooms = [r["name"] for r in c.execute("SELECT name FROM rooms").fetchall()]
        c.execute("""
            SELECT room_name FROM reservations 
            WHERE date=? AND start_time=? AND status IN ('Pending','Approved')
        """, (selected_date, chosen_time))
        taken = [row["room_name"] for row in c.fetchall()]
        free_room = next((r for r in rooms if r not in taken), None)

        if not free_room:
            flash("❌ No rooms left at this time.", "danger")
            return redirect(url_for("reserve",
                                    company_name=selected_company,
                                    entity_id=entity_id,
                                    date=selected_date))

        # Target person (guest)
        person = c.execute("""
            SELECT first_name, last_name, email 
            FROM approved_users WHERE id=?
        """, (entity_id,)).fetchone()

        full_name = f"{person['first_name']} {person['last_name']}"
        target_email = person["email"]

        # Save reservation
        c.execute("""
            INSERT INTO reservations 
            (user_id, entity_type, entity_id, date, start_time, room_name, invites, status, slot_id)
            VALUES (?,?,?,?,?,?,?,?,?)
        """, (
            user_id, "person", entity_id, selected_date,
            chosen_time, free_room, f"{email},{target_email}", "Pending", slot_id
        ))
        conn.commit()

        reservation_id = c.lastrowid

        # Links for YES / NO
        approve_link = f"https://mckinsey-booking-system.onrender.com/respond_meeting/{reservation_id}?decision=approve"
        reject_link = f"https://mckinsey-booking-system.onrender.com/respond_meeting/{reservation_id}?decision=reject"

        # Format date
        pretty_date = datetime.strptime(selected_date, "%Y-%m-%d").strftime("%-d %B %Y")


        start_dt = datetime.strptime(chosen_time, "%H:%M")
        end_dt = start_dt + timedelta(minutes=20)
        pretty_time = start_dt.strftime("%I:%M") + " – " + end_dt.strftime("%I:%M %p")

        # Build Google-Calendar style subject
        subject_guest = build_invitation_subject(
            selected_date,
            start_dt.strftime("%H:%M"),
            end_dt.strftime("%H:%M")
        )

# requester gets the same subject
        subject_requester = subject_guest

        # -----------------------------
        # EMAIL TO REQUESTER
        # -----------------------------
        # subject_requester = "Your Meeting Request Has Been Submitted"

        body_requester = (
            f"Dear {user['first_name']},\n\n"
            f"Your meeting request with {full_name} of {selected_company} has been successfully scheduled.\n\n"
            f"Date: {pretty_date}\n"
            f"Time: {pretty_time}\n"
            f"Meeting Room: {free_room}\n\n"
            f"Please wait for the guest's confirmation.\n"
        )

        send_plain_email(email, subject_requester, body_requester)

        # -----------------------------
        # EMAIL TO GUEST (HTML)
        # -----------------------------

        # subject_guest = "New Meeting Request – Action Required"

        html_guest = f"""
<div style='font-family:Arial,sans-serif;font-size:15px;color:#202124'>

  <p>Dear {full_name},</p>

  <p>You have received a new meeting request from <b>{requester_full_name}</b> 
     of <b>{selected_company}</b>.</p>

  <p><b>Date:</b> {pretty_date}<br>
     <b>Time:</b> {pretty_time}<br>
     <b>Meeting Room:</b> {free_room}<br>
     <b>Requested by:</b> {email}</p>

  <p>Please select one of the options below:</p>

  <div style="margin-top:15px;">
    <a href="{approve_link}" 
       style="padding:10px 18px;border:1px solid #ccc;
              border-radius:4px;text-decoration:none;
              color:#000;margin-right:10px;">
      Confirm
    </a>

    <a href="{reject_link}" 
       style="padding:10px 18px;border:1px solid #ccc;
              border-radius:4px;text-decoration:none;
              color:#000;">
      Reject
    </a>
  </div>

  <br><br>
  <p style="color:#5f6368;font-size:12px;">
    
  </p>

</div>
"""


        send_html_email(target_email, subject_guest, html_guest)

        flash("Meeting request submitted!", "success")
        return redirect(url_for("my_meetings"))

    # -----------------------------
    # BUILD 20-MIN SLOTS
    # -----------------------------
    available_times = []
    if entity_id and selected_date:
        start_minutes = 9 * 60
        end_minutes = 17 * 60
        current = start_minutes
        slot_counter = 0

        while current < end_minutes:
            sh, sm = divmod(current, 60)
            eh, em = divmod(current + 20, 60)

            start_str = f"{sh:02d}:{sm:02d}"
            end_str = f"{eh:02d}:{em:02d}"

            taken = c.execute("""
                SELECT 1 FROM reservations
                WHERE entity_type='person' AND entity_id=? 
                  AND date=? AND start_time=?
                  AND status IN ('Pending','Approved')
            """, (entity_id, selected_date, start_str)).fetchone()

            if not taken:
                slot_counter += 1
                available_times.append({
                    "start": start_str,
                    "end": end_str,
                    "slot_id": f"p{slot_counter}"
                })

            current += 20

    return render_template(
        "reserve.html",
        companies=companies,
        people=people,
        selected_company=selected_company,
        selected_date=selected_date,
        entity_id=entity_id,
        available_times=available_times
    )

def send_html_email(to_email, subject, html_body):
    try:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = SMTP_EMAIL
        msg["To"] = to_email
        msg.add_alternative(html_body, subtype="html")

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(SMTP_EMAIL, SMTP_PASSWORD)
            smtp.send_message(msg)

        print(f"📧 HTML email sent to {to_email}")
        return True
    except Exception as e:
        print("❌ HTML Email failed:", e)
        return False

def build_invitation_subject(date_str, start_time, end_time):
    dt = datetime.strptime(f"{date_str} {start_time}", "%Y-%m-%d %H:%M")
    end_dt = datetime.strptime(f"{date_str} {end_time}", "%Y-%m-%d %H:%M")

    weekday = dt.strftime("%a")  # Mon
    month = dt.strftime("%B")    # December
    day = dt.strftime("%-d")     # 9
    year = dt.strftime("%Y")     # 2025

    pretty_start = dt.strftime("%-I:%M%p").lower()   # 1:30pm
    pretty_end   = end_dt.strftime("%-I:%M%p").lower()

    return (
        f"Invitation: Semicon Summit Dubai 2025 @ "
        f"{weekday} {month} {day}, {year} {pretty_start} - {pretty_end} (GMT+4)"
    )

################################################################################################

############################################## ADMIN ###########################################

################################################################################################

# DISPLAY THE MENU FOR ADMIN

@app.route("/admin", methods=["GET", "POST"])
@login_required
@admin_required
def admin_dashboard():
    email = session.get("email")

    # ✅ Ensure only admin can access this
    if email != ADMIN_EMAIL:
        flash("⚠️ You don’t have access to the admin dashboard.", "danger")
        return redirect(url_for("login"))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Handle adding new room
    if request.method == "POST":
        if "room_name" in request.form:
            room_name = request.form["room_name"]
            c.execute("INSERT INTO rooms(name) VALUES(?)", (room_name,))
            conn.commit()
            flash("✅ Room added successfully!", "success")
        elif "slot_room" in request.form:
            room_id = request.form["slot_room"]
            date = request.form["date"]
            start = request.form["start_time"]
            end = request.form["end_time"]
            c.execute("""
                INSERT INTO time_slots (room_id, date, start_time, end_time)
                VALUES (?, ?, ?, ?)
            """, (room_id, date, start, end))
            conn.commit()
            flash("✅ Slot added successfully!", "success")

    # Fetch rooms & slots to show them
    rooms = c.execute("SELECT * FROM rooms ORDER BY name").fetchall()
    slots = c.execute("""
        SELECT time_slots.*, rooms.name AS room_name, users.email AS reserver_email
        FROM time_slots
        JOIN rooms ON time_slots.room_id = rooms.id
        LEFT JOIN users ON time_slots.reserved_by = users.id
        ORDER BY date, start_time
    """).fetchall()

    conn.close()

    # ✅ Force-render the admin layout only
    return render_template("admin_dashboard.html", rooms=rooms, slots=slots)


# 1- GET ALL USERS FROM GOOGLE SHEETS

@app.route("/admin/users")
@admin_required
def admin_users():
    users = get_users_from_sheets()
    return render_template("admin_users.html", users=users)


# 2- GET ONLY APPROVED USERS FROM SHEETS AND STORE IN SQLITE

@app.route("/admin/approved_users", methods=["GET", "POST"])
@admin_required
def admin_approved_users():
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    if request.method == "POST" and request.form.get("action") == "import":
        # ---- get approved users from Google Sheet ----
        sheet_users = get_users_from_sheets2()

        # Insert only if not exists
        added = 0
        for u in sheet_users:
            try:
                c.execute("""
        INSERT INTO approved_users
        (first_name,last_name,email,company_name,position,phone,passport_number,
         passport_place,passport_expiry,birth_date)
        VALUES (?,?,?,?,?,?,?,?,?,?)
    """, (
        u.get("First name"), 
        u.get("Last name"), 
        u.get("Email"),
        u.get("Company name"),
        u.get("Position"),
        u.get("Phone"), 
        u.get("Passport Number"),
        u.get("Passport Place of Issuance"),
        u.get("Passport Expiry Date"),
        u.get("Birth Date")
    ))
                added += 1
            except sqlite3.IntegrityError:
                # skip if email already exists
                pass
        conn.commit()
        flash(f"✅ Imported {added} new approved user(s).", "success")

    users = c.execute("SELECT * FROM approved_users ORDER BY created_date DESC").fetchall()
    return render_template("admin_approved_users.html", users=users)


# 3- DELETE APPROVED USER

@app.route("/admin/delete_approved_user/<int:user_id>", methods=["POST"])
@admin_required
def delete_approved_user(user_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM approved_users WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()
    flash("🗑️ User deleted successfully.", "success")
    return redirect(url_for("admin_approved_users"))


# 4- CHECK ALL ROOMS + ADD NEW ONE

@app.route("/admin/rooms", methods=["GET", "POST"])
@admin_required
def admin_rooms():

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    if request.method == "POST":
        # Add new room
        name = request.form.get("name")
        if name:
            c.execute("INSERT INTO rooms (name) VALUES (?)", (name,))
            conn.commit()
            flash("Room added successfully!", "success")
        return redirect(url_for("admin_rooms"))

    rooms = c.execute("SELECT * FROM rooms ORDER BY name").fetchall()
    return render_template("admin_rooms.html", rooms=rooms)

#5- DELETE ROOM

@app.route("/admin/rooms/delete/<int:room_id>", methods=["POST"])
def delete_room(room_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM rooms WHERE id=?", (room_id,))
    conn.commit()
    return redirect(url_for("admin_rooms"))

#6- MANAGE COMPANIES + CONTACTS

@app.route("/admin/companies", methods=["GET", "POST"])
@admin_required
def admin_companies():

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    if request.method == "POST":
        action = request.form.get("action")

        if action == "add_company":
            name = request.form["name"]
            description = request.form.get("description")
            c.execute("INSERT INTO companies (name, description) VALUES (?,?)", (name, description))
            conn.commit()
            flash("✅ Company added successfully!", "success")
            return redirect(url_for("admin_companies"))

        elif action == "add_contact":
            company_id = request.form.get("company_id")
            email = request.form.get("email")
            c.execute("INSERT INTO company_contacts (company_id, email) VALUES (?,?)", (company_id, email))
            conn.commit()
            flash("✅ Contact added!", "success")
            return redirect(url_for("admin_companies"))

    # Fetch companies with their contacts
    companies = c.execute("SELECT * FROM companies ORDER BY name").fetchall()
    companies_list = []
    for comp in companies:
        contacts = c.execute("SELECT * FROM company_contacts WHERE company_id=?", (comp["id"],)).fetchall()
        companies_list.append({
            **dict(comp),
            "contacts": contacts
        })

    return render_template("admin_companies.html", companies=companies_list)

#7- DELETE CONTACT

@app.route("/admin/companies/delete_contact/<int:id>")
@admin_required
def delete_contact(id):
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM company_contacts WHERE id=?", (id,))
    conn.commit()
    flash("🗑️ Contact deleted.", "info")
    return redirect(url_for("admin_companies"))


#8- DELETE COMPANY (and its contacts)

@app.route("/admin/companies/delete/<int:id>")
@admin_required
def delete_company(id):
    conn = get_db()
    c = conn.cursor()
    # Delete related contacts first (to keep DB clean)
    c.execute("DELETE FROM company_contacts WHERE company_id=?", (id,))
    c.execute("DELETE FROM companies WHERE id=?", (id,))
    conn.commit()

    flash("🗑️ Company deleted.", "info")
    return redirect(url_for("admin_companies"))


#9- EDIT COMPANY

@app.route("/admin/companies/edit/<int:id>", methods=["GET","POST"])
@admin_required
def edit_company(id):
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    if request.method == "POST":
        name = request.form["name"]
        description = request.form.get("description")
        c.execute("UPDATE companies SET name=?, description=? WHERE id=?", (name, description, id))
        conn.commit()
        flash("✅ Company updated!", "success")
        return redirect(url_for("admin_companies"))

    company = c.execute("SELECT * FROM companies WHERE id=?", (id,)).fetchone()
    if not company:
        flash("Company not found.", "danger")
        return redirect(url_for("admin_companies"))
    return render_template("edit_company.html", company=company)


#10- MANAGE COMPANY-SPECIFIC TIME SLOTS

@app.route("/admin/company_slots", methods=["GET", "POST"])
@admin_required
def admin_company_slots():
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # ✅ Handle Add Slot (POST)
    if request.method == "POST":
        company_id = request.form.get("company_id")
        date = request.form.get("date")
        start_time = request.form.get("start_time")
        end_time = request.form.get("end_time")

        if not (company_id and date and start_time and end_time):
            flash("⚠️ Please fill all fields.", "warning")
        else:
            # 🧠 Prevent duplicate same company/date/start_time
            existing = c.execute("""
                SELECT 1 FROM company_slots
                WHERE company_id = ? AND date = ? AND start_time = ?
            """, (company_id, date, start_time)).fetchone()

            if existing:
                flash("⚠️ This slot already exists for the selected company and time.", "warning")
            else:
                c.execute("""
                    INSERT INTO company_slots (company_id, date, start_time, end_time)
                    VALUES (?, ?, ?, ?)
                """, (company_id, date, start_time, end_time))
                conn.commit()
                flash("✅ Time slot added successfully!", "success")

        return redirect(url_for("admin_company_slots"))

    # ---- Handle optional filter (GET)
    company_filter = request.args.get("company_id")

    if company_filter:
        c.execute("""
            SELECT cs.*, co.name AS company_name 
            FROM company_slots cs
            JOIN companies co ON cs.company_id = co.id
            WHERE cs.company_id = ?
            ORDER BY cs.date, cs.start_time
        """, (company_filter,))
    else:
        c.execute("""
            SELECT cs.*, co.name AS company_name 
            FROM company_slots cs
            JOIN companies co ON cs.company_id = co.id
            ORDER BY cs.date, cs.start_time
        """)

    slots = c.fetchall()
    companies = c.execute("SELECT id, name FROM companies ORDER BY name").fetchall()

    return render_template(
        "admin_company_slots.html",
        slots=slots,
        companies=companies,
        company_filter=company_filter
    )


#11- DELETE COMPANY SLOT

@app.route("/admin/delete_company_slot/<int:slot_id>", methods=["POST"])
@admin_required
def delete_company_slot(slot_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM company_slots WHERE id=?", (slot_id,))
    conn.commit()
    conn.close()
    flash("🗑️ Slot deleted successfully.", "success")
    return redirect(url_for("admin_company_slots"))


#12- VIEW ALL RESERVATIONS

@app.route("/admin/reservations")
@admin_required
def admin_reservations():
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    reservations = c.execute("""
        SELECT r.*,
               au.first_name || ' ' || au.last_name AS booker_name,
               au.email AS booker_email,
               CASE
                 WHEN r.entity_type='company' THEN (SELECT name FROM companies WHERE id = r.entity_id)
                 WHEN r.entity_type='person'  THEN (SELECT first_name || ' ' || last_name FROM approved_users WHERE id = r.entity_id)
                 ELSE 'Unknown'
               END AS entity_name
        FROM reservations r
        LEFT JOIN approved_users au ON r.user_id = au.id
        ORDER BY r.date, r.start_time
    """).fetchall()

    return render_template("admin_reservations.html", reservations=reservations)

#13- ADD RESERVATION MANUALLY

@app.route("/admin/add_reservation", methods=["GET", "POST"])
@admin_required
def admin_add_reservation():
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Load companies and rooms for dropdowns
    companies = c.execute("SELECT id, name FROM companies ORDER BY name").fetchall()
    rooms = c.execute("SELECT name FROM rooms ORDER BY name").fetchall()

    if request.method == "POST":
        company_id = request.form.get("company_id")
        date = request.form.get("date")
        time = request.form.get("time")
        room_name = request.form.get("room_name")
        invites = request.form.get("invites")

        # Validate inputs
        if not company_id or not date or not time or not room_name:
            flash("⚠️ Please fill all required fields.", "warning")
            return redirect(url_for("admin_add_reservation"))

        # Prevent double-booking same room and time
        existing = c.execute("""
            SELECT id FROM reservations WHERE date=? AND start_time=? AND room_name=?
        """, (date, time, room_name)).fetchone()
        if existing:
            flash("❌ This room is already booked for that time.", "danger")
            return redirect(url_for("admin_add_reservation"))

        # Insert into database
        c.execute("""
            INSERT INTO reservations (user_id, entity_type, entity_id, date, start_time, room_name, invites)
            VALUES (NULL, 'company', ?, ?, ?, ?, ?)
        """, (company_id, date, time, room_name, invites))
        conn.commit()

        flash("✅ Reservation added successfully!", "success")
        return redirect(url_for("admin_reservations"))

    return render_template("admin_add_reservation.html", companies=companies, rooms=rooms)

#14- DELETE RESERVATION

@app.route("/admin/delete_reservation/<int:reservation_id>", methods=["POST"])
@admin_required
def delete_reservation(reservation_id):
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    c.execute("DELETE FROM reservations WHERE id=?", (reservation_id,))
    conn.commit()
    conn.close()

    flash("🗑️ Reservation deleted successfully!", "success")
    return redirect(url_for("admin_reservations"))




################################################################################################

############################################## COMPANY ###########################################

################################################################################################

# 1- COMPANY CONTACT VIEW + MANAGE THEIR COMPANY'S SLOTS
@app.route("/company_slots", methods=["GET", "POST"])
@verified_required
def company_slots_user():
    company_id = session.get("company_id")
    if not company_id:
        flash("❌ You are not linked to any company.", "danger")
        return redirect(url_for("reserve"))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    if request.method == "POST":
        date = request.form.get("date")
        start_time = request.form.get("start_time")
        end_time = request.form.get("end_time")

        if not (date and start_time and end_time):
            flash("⚠️ Please fill all fields.", "warning")
            return redirect(url_for("company_slots_user"))

        c.execute("INSERT INTO company_slots (company_id, date, start_time, end_time) VALUES (?, ?, ?, ?)",
                  (company_id, date, start_time, end_time))
        conn.commit()
        flash("✅ Time slot added successfully!", "success")
        return redirect(url_for("company_slots_user"))

    c.execute("""
        SELECT cs.*, co.name AS company_name 
        FROM company_slots cs
        JOIN companies co ON cs.company_id = co.id
        WHERE cs.company_id = ?
        ORDER BY cs.date, cs.start_time
    """, (company_id,))
    slots = c.fetchall()

    company = c.execute("SELECT name FROM companies WHERE id=?", (company_id,)).fetchone()
    company_name = company["name"] if company else "Your Company"

    return render_template("company_slots_user.html", slots=slots, company_name=company_name)

# 2- VIEW ALL MEETING REQUESTS
@app.route("/company_requests")
@verified_required
def company_requests():
    company_id = session.get("company_id")
    if not company_id:
        flash("❌ You are not linked to a company.", "danger")
        return redirect(url_for("reserve"))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    requests = c.execute("""
        SELECT r.*, u.email AS booker_email
        FROM reservations r
        LEFT JOIN approved_users u ON r.user_id = u.id
        WHERE r.entity_type='company' AND r.entity_id=? AND r.status='Pending'
        ORDER BY r.date, r.start_time
    """, (company_id,)).fetchall()

    return render_template("company_requests.html", requests=requests)

# 3- APPROVE REQUESTS
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import os



import json
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request


CALENDAR_SCOPES = ["https://www.googleapis.com/auth/calendar"]


def create_calendar_event(summary, description, start_datetime, end_datetime, attendees, timezone="Asia/Beirut"):
    # Load creds from Render environment variables
    creds_json = os.environ.get("CALENDAR_CREDS_JSON")
    token_json = os.environ.get("CALENDAR_TOKEN_JSON")

    if not creds_json or not token_json:
        print("❌ Missing Google Calendar environment variables")
        return None

    creds_raw = json.loads(creds_json)["installed"]
    token_raw = json.loads(token_json)

    creds = Credentials(
        token=token_raw.get("token"),
        refresh_token=token_raw.get("refresh_token"),
        token_uri=creds_raw["token_uri"],
        client_id=creds_raw["client_id"],
        client_secret=creds_raw["client_secret"],
        scopes=CALENDAR_SCOPES,
    )

    # Refresh token if expired
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())

    # Build service
    service = build("calendar", "v3", credentials=creds)

    event = {
        "summary": summary,
        "description": description,
        "start": {"dateTime": start_datetime.isoformat(), "timeZone": timezone},
        "end": {"dateTime": end_datetime.isoformat(), "timeZone": timezone},
        "attendees": [{"email": e} for e in sorted(set(attendees))],
        "reminders": {
            "useDefault": False,
            "overrides": [
                {"method": "email", "minutes": 60},
                {"method": "popup", "minutes": 10},
            ],
        },
    }

    created = service.events().insert(
        calendarId="primary",
        body=event,
        sendUpdates="all"
    ).execute()

    print("✅ Calendar event created:", created.get("htmlLink"))
    return created

# ✅ APPROVE REQUEST ENDPOINT
@app.route("/approve_request/<int:reservation_id>", methods=["POST"])
@verified_required
def approve_request(reservation_id):
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # --- Fetch reservation details ---
    reservation = c.execute("""
        SELECT r.*, au.email AS requester_email,
               co.name AS company_name
        FROM reservations r
        LEFT JOIN approved_users au ON r.user_id = au.id
        LEFT JOIN companies co ON r.entity_id = co.id
        WHERE r.id=?
    """, (reservation_id,)).fetchone()

    if not reservation:
        flash("❌ Reservation not found.", "danger")
        return redirect(url_for("company_all_requests"))

    # --- Fetch company contacts ---
    contacts = c.execute(
        "SELECT email FROM company_contacts WHERE company_id=?",
        (reservation["entity_id"],)
    ).fetchall()
    company_emails = [row["email"] for row in contacts]

    # --- Split invites (comma-separated) and combine all recipients ---
    invites_list = [e.strip() for e in (reservation["invites"] or "").split(",") if e.strip()]
    recipients = list(dict.fromkeys(invites_list + company_emails))  # unique + ordered

    # --- Update reservation status ---
    c.execute("UPDATE reservations SET status='Approved' WHERE id=?", (reservation_id,))
    conn.commit()

    # --- Send approval email ---
    try:
        subject = f"Meeting Request Approved - {reservation['company_name']}"
        body = (
            f"Hello,\n\n"
            f"Your meeting request with {reservation['company_name']} has been approved.\n\n"
            f"📅 Date: {reservation['date']}\n"
            f"⏰ Time: {reservation['start_time']}\n"
            f"🏠 Room: {reservation['room_name'] or 'TBD'}\n\n"
            f"This meeting has also been added to your Google Calendar.\n\n"
            f"Thank you,\nMcKinsey Electronics Team"
        )

        to_field = ", ".join(recipients)
        send_plain_email(to_field, subject, body)
        flash("✅ Email sent successfully!", "success")

        # --- Create Google Calendar event ---
        start_str = f"{reservation['date']} {reservation['start_time']}"
        start_dt = datetime.strptime(start_str, "%Y-%m-%d %H:%M")
        end_dt = start_dt + timedelta(minutes=30)

        summary = f"Meeting with {reservation['company_name']}"
        description = (
            f"Approved meeting between {reservation['invites']} and {reservation['company_name']}.\n"
            f"Room: {reservation['room_name'] or 'TBD'}"
        )

        # Call helper to create event
        create_calendar_event(summary, description, start_dt, end_dt, recipients)
        flash("📅 Meeting added to Google Calendar for all attendees.", "success")

    except Exception as e:
        flash(f"⚠️ Meeting approved but email or calendar invite failed: {e}", "warning")

    # --- Redirect back to company requests page ---
    return redirect(url_for("company_all_requests"))


# 4-REJECT REQUESTS
@app.route("/reject_request/<int:reservation_id>", methods=["POST"])
@verified_required
def reject_request(reservation_id):
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    # Fetch reservation details
    reservation = c.execute("""
        SELECT r.*, au.email AS requester_email,
               co.name AS company_name
        FROM reservations r
        LEFT JOIN approved_users au ON r.user_id = au.id
        LEFT JOIN companies co ON r.entity_id = co.id
        WHERE r.id=?
    """, (reservation_id,)).fetchone()

    if not reservation:
        flash("❌ Reservation not found.", "danger")
        return redirect(url_for("company_requests"))
    # ✅ Fetch company contacts
    contacts = c.execute(
        "SELECT email FROM company_contacts WHERE company_id=?",
        (reservation["entity_id"],)
    ).fetchall()
    company_emails = [c["email"] for c in contacts]

    # ✅ Combine requester + company contacts
    recipients = list(set([reservation["invites"]] + company_emails))
    c.execute("UPDATE reservations SET status='Rejected' WHERE id=?", (reservation_id,))
    conn.commit()

    # --- Send email notification
    try:
        subject = f"Meeting Request Rejected - {reservation['company_name']}"
        body = (
            f"Hello,\n\n"
            f"Unfortunately, your meeting request with {reservation['company_name']} "
            f"has been rejected.\n\n"
            f"📅 Date: {reservation['date']}\n"
            f"⏰ Time: {reservation['start_time']}\n\n"
            f"Please contact the company for further details.\n\n"
            f"Thank you,\n"
        )

        to_field = ", ".join(recipients)
        send_plain_email(to_field, subject, body)

        flash("❌ Meeting rejected and email sent successfully.", "danger")
    except Exception as e:
        flash(f"⚠️ Meeting rejected but email failed to send: {e}", "warning")

    return redirect(url_for("company_requests"))

# 5- VIEW APPROVED MEETINGS (with FullCalendar integration)
@app.route("/company_approved_meetings")
@verified_required
def company_approved_meetings():
    company_id = session.get("company_id")
    if not company_id:
        flash("❌ You are not linked to a company.", "danger")
        return redirect(url_for("reserve"))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # ✅ Fetch only approved reservations for this company
    approved = c.execute("""
        SELECT r.*, u.email AS booker_email
        FROM reservations r
        LEFT JOIN approved_users u ON r.user_id = u.id
        WHERE r.entity_type='company' AND r.entity_id=? AND r.status='Approved'
        ORDER BY r.date, r.start_time
    """, (company_id,)).fetchall()

    # ✅ Return events as JSON for FullCalendar
    if request.args.get("format") == "json":
        from datetime import datetime, timedelta
        events = []
        for m in approved:
            start_str = m["start_time"]
            try:
                start_dt = datetime.strptime(start_str, "%H:%M")
                end_dt = start_dt + timedelta(minutes=30)
                end_str = end_dt.strftime("%H:%M")
            except Exception:
                end_str = start_str

            events.append({
                "title": f"Meeting in {m['room_name']} ({m['booker_email']})",
                "start": f"{m['date']}T{start_str}",
                "end": f"{m['date']}T{end_str}",
                "backgroundColor": "#198754",
                "borderColor": "#198754",
                "textColor": "white"
            })
        return jsonify(events)

    return render_template("company_approved_meetings.html")

# 6- VIEW ALL REQUESTS (with filters and search)
@app.route("/company_all_requests")
@verified_required
def company_all_requests():
    company_id = session.get("company_id")
    if not company_id:
        flash("❌ You are not linked to a company.", "danger")
        return redirect(url_for("reserve"))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # --- Filters ---
    status_filter = request.args.get("status")
    search_query = request.args.get("q", "").strip().lower()

    query = """
        SELECT r.*, u.email AS booker_email
        FROM reservations r
        LEFT JOIN approved_users u ON r.user_id = u.id
        WHERE r.entity_type='company' AND r.entity_id=?
    """
    params = [company_id]

    if status_filter:
        query += " AND r.status=?"
        params.append(status_filter)

    if search_query:
        query += " AND (LOWER(u.email) LIKE ? OR LOWER(r.date) LIKE ?)"
        params += [f"%{search_query}%", f"%{search_query}%"]

    query += " ORDER BY r.date DESC, r.start_time ASC"
    requests = c.execute(query, params).fetchall()

    return render_template(
        "company_all_requests.html",
        requests=requests,
        status_filter=status_filter,
        search_query=search_query
    )




######################################################################################################################
####################################################################################################

# === PERSON MEETING REQUESTS (you are the invited person) ===

@app.route("/person_requests")
@verified_required
def person_requests():
    """
    Show all reservations where the logged-in user is the *target person*
    (entity_type='person' and entity_id = this user's approved_users.id)
    """
    email = session.get("email")
    if not email:
        return redirect(url_for("login"))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Find this user in approved_users
    c.execute("SELECT id, first_name, last_name FROM approved_users WHERE lower(email)=?", (email.lower(),))
    me = c.fetchone()
    if not me:
        flash("Your email is not in approved users.", "danger")
        return redirect(url_for("reserve"))

    # Fetch all reservations where this user is the invited *person*
    requests_rows = c.execute("""
        SELECT r.*,
               req.first_name || ' ' || req.last_name AS requester_name,
               req.email AS requester_email,
               req.company_name AS requester_company
        FROM reservations r
        JOIN approved_users req ON r.user_id = req.id             -- who booked
        WHERE r.entity_type = 'person'
          AND r.entity_id = ?
        ORDER BY r.date, r.start_time
    """, (me["id"],)).fetchall()

    conn.close()
    return render_template("person_requests.html", requests=requests_rows, me=me)


@app.route("/person_requests/approve/<int:reservation_id>", methods=["POST"])
@verified_required
def approve_person_request(reservation_id):
    """
    Logged-in invited person approves the meeting.
    Uses same email structure as respond_meeting.
    """
    email = session.get("email")
    if not email:
        return redirect(url_for("login"))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Fetch reservation
    reservation = c.execute("""
        SELECT r.*,
               req.first_name || ' ' || req.last_name AS requester_name,
               req.first_name AS requester_first,
               req.email  AS requester_email,
               target.first_name || ' ' || target.last_name AS target_name,
               target.company_name AS target_company,
               target.email AS target_email
        FROM reservations r
        JOIN approved_users req    ON r.user_id   = req.id
        JOIN approved_users target ON r.entity_id = target.id
        WHERE r.id = ? AND r.entity_type = 'person'
    """, (reservation_id,)).fetchone()

    if not reservation:
        flash("❌ Reservation not found.", "danger")
        return redirect(url_for("person_requests"))

    if reservation["target_email"].lower() != email.lower():
        abort(403)

    # Date formatting
    start_dt = datetime.strptime(
        f"{reservation['date']} {reservation['start_time']}",
        "%Y-%m-%d %H:%M"
    )
    end_dt = start_dt + timedelta(minutes=20)

    pretty_date = start_dt.strftime("%-d %B %Y")
    pretty_time = start_dt.strftime("%I:%M") + " – " + end_dt.strftime("%I:%M %p")

    weekday = start_dt.strftime("%a")
    month = start_dt.strftime("%B")
    day = start_dt.strftime("%-d")
    year = start_dt.strftime("%Y")
    pretty_start = start_dt.strftime("%-I:%M%p").lower()
    pretty_end = end_dt.strftime("%-I:%M%p").lower()

    # EXACT SAME subject as respond_meeting
    subject_line = (
        f"Invitation: Semicon Summit Dubai 2025 @ "
        f"{weekday} {month} {day}, {year} {pretty_start} - {pretty_end} (GMT+4)"
    )

    # Recipients list
    invites_list = [e.strip() for e in (reservation["invites"] or "").split(",") if e.strip()]
    recipients = list(dict.fromkeys(invites_list + [
        reservation["requester_email"],
        reservation["target_email"]
    ]))

    # Update DB
    c.execute("UPDATE reservations SET status='Approved' WHERE id=?", (reservation_id,))
    conn.commit()

    # Email to requester (same as respond_meeting)
    body_requester = (
        f"Dear {reservation['requester_first']},\n\n"
        f"Your meeting request with {reservation['target_name']} of {reservation['target_company']} "
        f"has been accepted.\n\n"
        f"Below are the meeting details:\n\n"
        f"Date: {pretty_date}\n"
        f"Time: {pretty_time}\n"
        f"Meeting Room: {reservation['room_name'] or 'Room TBD'}\n\n"
        f"Thank you,\n"
    )

    send_plain_email(reservation["requester_email"], subject_line, body_requester)

    # Calendar event
    try:
        summary = f"Meeting with {reservation['target_name']}"
        description = (
            f"Approved meeting between {reservation['requester_name']} "
            f"and {reservation['target_name']}.\n"
            f"Room: {reservation['room_name'] or 'TBD'}"
        )
        create_calendar_event(summary, description, start_dt, end_dt, recipients)
    except Exception as e:
        print("Calendar Error:", e)

    flash("Meeting approved!", "success")
    return redirect(url_for("person_requests"))


@app.route("/person_requests/reject/<int:reservation_id>", methods=["POST"])
@verified_required
def reject_person_request(reservation_id):
    """
    Logged-in invited person rejects the meeting.
    Uses same email structure as respond_meeting.
    """
    email = session.get("email")
    if not email:
        return redirect(url_for("login"))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Fetch reservation
    reservation = c.execute("""
        SELECT r.*,
               req.first_name || ' ' || req.last_name AS requester_name,
               req.first_name AS requester_first,
               req.email  AS requester_email,
               target.first_name || ' ' || target.last_name AS target_name,
               target.company_name AS target_company,
               target.email AS target_email
        FROM reservations r
        JOIN approved_users req    ON r.user_id   = req.id
        JOIN approved_users target ON r.entity_id = target.id
        WHERE r.id = ? AND r.entity_type = 'person'
    """, (reservation_id,)).fetchone()

    if not reservation:
        flash("❌ Reservation not found.", "danger")
        return redirect(url_for("person_requests"))

    if reservation["target_email"].lower() != email.lower():
        abort(403)

    # Date formatting
    start_dt = datetime.strptime(
        f"{reservation['date']} {reservation['start_time']}",
        "%Y-%m-%d %H:%M"
    )
    end_dt = start_dt + timedelta(minutes=20)

    pretty_date = start_dt.strftime("%-d %B %Y")
    pretty_time = start_dt.strftime("%I:%M") + " – " + end_dt.strftime("%I:%M %p")

    weekday = start_dt.strftime("%a")
    month = start_dt.strftime("%B")
    day = start_dt.strftime("%-d")
    year = start_dt.strftime("%Y")
    pretty_start = start_dt.strftime("%-I:%M%p").lower()
    pretty_end = end_dt.strftime("%-I:%M%p").lower()

    # SAME SUBJECT
    subject_line = (
        f"Invitation: Semicon Summit Dubai 2025 @ "
        f"{weekday} {month} {day}, {year} {pretty_start} - {pretty_end} (GMT+4)"
    )

    # Update DB
    c.execute("UPDATE reservations SET status='Rejected' WHERE id=?", (reservation_id,))
    conn.commit()

    # Email to requester (same as respond_meeting)
    body_requester = (
        f"Dear {reservation['requester_first']},\n\n"
        f"Your meeting request with {reservation['target_name']} of {reservation['target_company']} "
        f"has been rejected.\n\n"
        f"Date: {pretty_date}\n"
        f"Time: {pretty_time}\n\n"
        f"Thank you,\n"
    )

    send_plain_email(reservation["requester_email"], subject_line, body_requester)

    flash("Meeting rejected.", "danger")
    return redirect(url_for("person_requests"))


######################################## USER DASHBOARD ###########################################
######################################################################################################################

@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.")
    return redirect(url_for("login"))


@app.route("/search_entities")
@verified_required
def search_entities():
    q = request.args.get("q", "").strip().lower()
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    results = []

    # 🔎 Search PEOPLE (by name or email, including domain/company part)
    people = c.execute("""
        SELECT id, first_name, last_name, email, company_name
        FROM approved_users
        WHERE lower(first_name) LIKE ? 
           OR lower(last_name) LIKE ? 
           OR lower(email) LIKE ?
        LIMIT 15
    """, (f"%{q}%", f"%{q}%", f"%{q}%")).fetchall()

    for p in people:
        company_part = p["email"].split("@")[-1] if p["email"] else ""
        results.append({
            "type": "person",
            "id": p["id"],
            "label": f"{p['first_name']} {p['last_name']} ({p['email']})"
        })

    # 🔎 Search COMPANIES (by name)
    companies = c.execute("""
        SELECT id, name
        FROM companies
        WHERE lower(name) LIKE ?
        LIMIT 10
    """, (f"%{q}%",)).fetchall()

    for row in companies:
        results.append({
            "type": "company",
            "id": row["id"],
            "label": row["name"]
        })

    return jsonify(results)

@app.route("/my_meetings")
@verified_required
def my_meetings():
    email = session.get("email")
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    c.execute("SELECT id FROM approved_users WHERE lower(email)=?", (email,))
    user = c.fetchone()
    if not user:
        return redirect(url_for("login"))
    
    user_id = user["id"]



    # Fetch meetings:
    # 1) Meetings I booked (r.user_id = me)
    # 2) Meetings where I am the invited PERSON (entity_type='person' AND entity_id = me)
    meetings = c.execute("""
        SELECT r.*, 
               COALESCE(c.name, au.first_name || ' ' || au.last_name) AS entity_name
        FROM reservations r
        LEFT JOIN companies c ON (r.entity_type='company' AND r.entity_id = c.id)
        LEFT JOIN approved_users au ON (r.entity_type='person' AND r.entity_id = au.id)
        WHERE r.user_id = ?
           OR (r.entity_type = 'person' AND r.entity_id = ?)
        ORDER BY r.date, r.start_time
    """, (user_id, user_id)).fetchall()


    return render_template("my_meetings.html", meetings=meetings)

@app.route("/cancel_meeting/<int:meeting_id>", methods=["POST"])
@verified_required
def cancel_meeting(meeting_id):
    email = session.get("email")
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Fetch meeting
    meeting = c.execute("SELECT * FROM reservations WHERE id=?", (meeting_id,)).fetchone()
    if not meeting:
        flash("Meeting not found.", "danger")
        return redirect(url_for("my_meetings"))

    # Ensure user owns it
    c.execute("SELECT id FROM approved_users WHERE lower(email)=?", (email,))
    user = c.fetchone()
    if not user or meeting["user_id"] != user["id"]:
        flash("You are not authorized to cancel this meeting.", "danger")
        return redirect(url_for("my_meetings"))

    # Determine meeting status safely
    status = meeting["status"].lower() if "status" in meeting.keys() and meeting["status"] else "pending"

    if status == "approved":
        subject = "Meeting Cancellation Notice"
        body = (
            f"Dear attendee,\n\n"
            f"The meeting scheduled for {meeting['date']} at {meeting['start_time']} "
            f"in room {meeting['room_name']} has been CANCELLED.\n\n"
            f"Cancelled by: {email}\n\n"
            f"Regards,\nRoom Reservation System"
        )

        # Send email if invites exist
        attendees = meeting["invites"].split(",") if meeting["invites"] else []
        if attendees:
            to_field = ", ".join(attendees)
            send_plain_email(to_field, subject, body)
            flash("📧 Cancellation email sent to all attendees.", "info")

    # Delete record
    c.execute("DELETE FROM reservations WHERE id=?", (meeting_id,))
    conn.commit()
    conn.close()

    flash("🗑️ Meeting cancelled successfully.", "success")
    return redirect(url_for("my_meetings"))


@app.route("/admin/clear_approved_users", methods=["POST"])
def clear_approved_users():
    conn = get_db()
    c = conn.cursor()

    # Delete all rows
    c.execute("DELETE FROM approved_users")

    # Reset auto-increment counter (optional but recommended)
    c.execute("DELETE FROM sqlite_sequence WHERE name='approved_users'")

    conn.commit()
    conn.close()

    flash("✅ All approved users have been deleted successfully.", "success")
    return redirect(url_for("admin_users"))

@app.route("/respond_meeting/<int:reservation_id>")
def respond_meeting(reservation_id):
    """
    Handles approval/rejection from email links.
    No login required.
    """
    decision = request.args.get("decision")

    if decision not in ("approve", "reject"):
        return "Invalid action.", 400

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Pull reservation, requester info, target info
    reservation = c.execute("""
        SELECT r.*,
               req.first_name || ' ' || req.last_name AS requester_name,
               req.first_name AS requester_first,
               req.email  AS requester_email,
               target.first_name || ' ' || target.last_name AS target_name,
               target.company_name AS target_company,
               target.email AS target_email
        FROM reservations r
        JOIN approved_users req    ON r.user_id   = req.id
        JOIN approved_users target ON r.entity_id = target.id
        WHERE r.id = ? AND r.entity_type = 'person'
    """, (reservation_id,)).fetchone()

    if not reservation:
        return "Reservation not found.", 404

    # Prepare times
    start_dt = datetime.strptime(f"{reservation['date']} {reservation['start_time']}", "%Y-%m-%d %H:%M")
    end_dt = start_dt + timedelta(minutes=20)

    pretty_date = start_dt.strftime("%-d %B %Y")                      # e.g. 9 December 2025
    pretty_time = start_dt.strftime("%I:%M") + " – " + end_dt.strftime("%I:%M %p")

    # Build Google-calendar-like subject
    weekday = start_dt.strftime("%a")    # Tue
    month = start_dt.strftime("%B")      # December
    day = start_dt.strftime("%-d")       # 9
    year = start_dt.strftime("%Y")       # 2025
    pretty_start = start_dt.strftime("%-I:%M%p").lower()   # 11:40am
    pretty_end = end_dt.strftime("%-I:%M%p").lower()       # 12:00pm

    subject_line = (
        f"Invitation: Semicon Summit Dubai 2025 @ "
        f"{weekday} {month} {day}, {year} {pretty_start} - {pretty_end} (GMT+4)"
    )

    # Build recipients (requester + target + invites)
    invites_list = [e.strip() for e in (reservation["invites"] or "").split(",") if e.strip()]
    recipients = list(dict.fromkeys(invites_list + [
        reservation["requester_email"],
        reservation["target_email"]
    ]))

    # ------------------------------------------------
    # APPROVE
    # ------------------------------------------------
    if decision == "approve":

        c.execute("UPDATE reservations SET status='Approved' WHERE id=?", (reservation_id,))
        conn.commit()

        # Email to requester — meeting accepted
        body_requester = (
            f"Dear {reservation['requester_first']},\n\n"
            f"Your meeting request with {reservation['target_name']} of {reservation['target_company']} "
            f"has been accepted.\n\n"
            f"Below are the meeting details:\n\n"
            f"Date: {pretty_date}\n"
            f"Time: {pretty_time}\n"
            f"Meeting Room: {reservation['room_name'] or 'Room TBD'}\n\n"
            f"Thank you,\n"
            f""
        )

        send_plain_email(reservation["requester_email"], subject_line, body_requester)

        # Calendar event (optional)
        try:
            summary = f"Meeting with {reservation['target_name']}"
            description = (
                f"Approved meeting between {reservation['requester_name']} "
                f"and {reservation['target_name']}.\n"
                f"Room: {reservation['room_name'] or 'TBD'}"
            )
            create_calendar_event(summary, description, start_dt, end_dt, recipients)
        except Exception as e:
            print("Calendar error:", e)

        return """
        <h2>Meeting Approved</h2>
        <p>Thank you. The meeting has been approved and both parties were notified.</p>
        """

    # ------------------------------------------------
    # REJECT
    # ------------------------------------------------
    if decision == "reject":

        c.execute("UPDATE reservations SET status='Rejected' WHERE id=?", (reservation_id,))
        conn.commit()

        # Email to requester — meeting rejected
        body_requester = (
            f"Dear {reservation['requester_first']},\n\n"
            f"Your meeting request with {reservation['target_name']} of {reservation['target_company']} "
            f"has been rejected.\n\n"
            f"Date: {pretty_date}\n"
            f"Time: {pretty_time}\n\n"
            f"Thank you,\n"
            f""
        )

        send_plain_email(reservation["requester_email"], subject_line, body_requester)

        return """
        <h2>Meeting Rejected</h2>
        <p>The meeting has been declined and the requester has been notified.</p>
        """


# if __name__ == "__main__":
#     app.run(host="0.0.0.0", port=5000, debug=True)

if __name__ == "__main__":
    app.run()
