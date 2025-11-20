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

# @app.route("/reserve", methods=["GET", "POST"])
# @verified_required
# def reserve():
#     email = session.get("email")
#     conn = get_db()
#     conn.row_factory = sqlite3.Row
#     c = conn.cursor()

#     # ✅ Only logged-in approved users can reserve
#     c.execute("SELECT id FROM approved_users WHERE lower(email)=?", (email,))
#     user = c.fetchone()
#     if not user:
#         return redirect(url_for("login"))

#     # ---- Get URL params
#     entity_type = request.args.get("entity_type") or "company"
#     entity_id = request.args.get("entity_id")
#     selected_date = request.args.get("date")

#     # ---- POST: User picked a slot to reserve
#     if request.method == "POST":
#         chosen_time = request.form.get("time")
#         slot_id = request.form.get("slot_id")
#         entity_id = request.form.get("entity_id")
#         selected_date = request.form.get("date")
#         entity_type = request.form.get("entity_type") or "company"

#         # ✅ Check if this specific slot_id or time is already booked
#         if entity_type == "company":
#             c.execute("""
#                 SELECT id FROM reservations
#                 WHERE slot_id=? AND status IN ('Pending','Approved')
#             """, (slot_id,))
#         else:
#             c.execute("""
#                 SELECT id FROM reservations
#                 WHERE entity_type='person' AND entity_id=? AND date=? AND start_time=? 
#                       AND status IN ('Pending','Approved')
#             """, (entity_id, selected_date, chosen_time))
#         if c.fetchone():
#             flash("❌ This specific slot is already booked.", "danger")
#             return redirect(url_for("reserve", entity_type=entity_type, entity_id=entity_id, date=selected_date))

#         # ✅ Assign first available room
#         rooms = [r["name"] for r in c.execute("SELECT name FROM rooms").fetchall()]
#         c.execute("""
#             SELECT room_name FROM reservations 
#             WHERE date=? AND start_time=? AND status IN ('Pending','Approved')
#         """, (selected_date, chosen_time))
#         taken = [row["room_name"] for row in c.fetchall()]
#         free_room = next((r for r in rooms if r not in taken), None)

#         if not free_room:
#             flash("❌ No rooms left at this time. Please choose another time.", "danger")
#             return redirect(url_for("reserve", entity_type=entity_type, entity_id=entity_id, date=selected_date))

#         # ✅ Build meeting details depending on type
#         if entity_type == "company":
#             company = c.execute("SELECT name FROM companies WHERE id=?", (entity_id,)).fetchone()
#             if not company:
#                 flash("❌ Company not found.", "danger")
#                 return redirect(url_for("reserve"))
#             company_name = company["name"]

#             # ✅ Fetch company contacts
#             contacts = c.execute(
#                 "SELECT email FROM company_contacts WHERE company_id=?",
#                 (entity_id,)
#             ).fetchall()
#             company_emails = [row["email"] for row in contacts]

#             recipients = list(set([email] + company_emails))
#             invites_str = ",".join(recipients)

#             # ✅ Save reservation
#             c.execute("""
#                 INSERT INTO reservations (user_id, entity_type, entity_id, date, start_time, room_name, invites, status, slot_id)
#                 VALUES (?,?,?,?,?,?,?,?,?)
#             """, (
#                 user["id"], "company", entity_id, selected_date,
#                 chosen_time, free_room, invites_str, "Pending", slot_id
#             ))
#             conn.commit()

            
#             subject = f"Meeting Request with {company_name}"
#             body = (
#                 f"Hello,\n\nA meeting has been requested with {company_name}.\n\n"
#                 f"📅 Date: {selected_date}\n"
#                 f"⏰ Time: {chosen_time}\n"
#                 f"🏢 Company: {company_name}\n"
#                 f"🏠 Room: {free_room}\n\n"
#                 f"Requested by: {email}\n"
#             )

#         elif entity_type == "person":
#             person = c.execute("""
#                 SELECT first_name, last_name, email FROM approved_users WHERE id=?
#             """, (entity_id,)).fetchone()
#             if not person:
#                 flash("❌ Person not found.", "danger")
#                 return redirect(url_for("reserve"))
#             full_name = f"{person['first_name']} {person['last_name']}"
#             person_email = person["email"]

#             # ✅ Save reservation
#             c.execute("""
#                 INSERT INTO reservations (user_id, entity_type, entity_id, date, start_time, room_name, invites, status, slot_id)
#                 VALUES (?,?,?,?,?,?,?,?,?)
#             """, (
#                 user["id"], "person", entity_id, selected_date,
#                 chosen_time, free_room, f"{email},{person_email}", "Pending", slot_id
#             ))
#             conn.commit()

#             recipients = list(set([email, person_email]))
#             subject = f"Meeting Request with {full_name}"
#             body = (
#                 f"Hello,\n\nA meeting has been requested with {full_name}.\n\n"
#                 f"📅 Date: {selected_date}\n"
#                 f"⏰ Time: {chosen_time}\n"
#                 f"👤 Person: {full_name}\n"
#                 f"🏠 Room: {free_room}\n\n"
#                 f"Requested by: {email}\n"
#             )

#         # ✅ Send confirmation email
#         to_field = ", ".join(recipients)
#         send_plain_email(to_field, subject, body)
#         flash(f"✅ Meeting request sent successfully!", "success")
#         return redirect(url_for("my_meetings"))


#     # ---- GET: Build available times ----
#     available_times = []
#     if entity_id and selected_date:
#         if entity_type == "company":
#             # Company slots from admin-defined table
#             c.execute("""
#                 SELECT id, start_time FROM company_slots
#                 WHERE company_id=? AND date=? ORDER BY start_time
#             """, (entity_id, selected_date))
#             slots = c.fetchall()

#             for s in slots:
#                 reserved = c.execute("""
#                     SELECT 1 FROM reservations
#                     WHERE slot_id=? AND status IN ('Pending','Approved')
#                 """, (s["id"],)).fetchone()
#                 if not reserved:
#                     available_times.append({"time": s["start_time"], "slot_id": s["id"]})

#         elif entity_type == "person":
#             # Everyone has fixed default 30-min slots 10:00 → 14:30
#             start_hour = 10
#             end_hour = 15  # 14:30 is the last
#             slot_counter = 0
#             for hour in range(start_hour, end_hour):
#                 for minute in (0, 30):
#                     if hour == 14 and minute == 30:
#                         continue
#                     t = f"{hour:02d}:{minute:02d}"
#                     taken = c.execute("""
#                         SELECT 1 FROM reservations
#                         WHERE entity_type='person' AND entity_id=? AND date=? AND start_time=? 
#                           AND status IN ('Pending','Approved')
#                     """, (entity_id, selected_date, t)).fetchone()
#                     if not taken:
#                         slot_counter += 1
#                         available_times.append({"time": t, "slot_id": f"p{slot_counter}"})

#     # ---- Display label for search box ----
#     display_label = None
#     if entity_id:
#         if entity_type == "company":
#             row = c.execute("SELECT name FROM companies WHERE id=?", (entity_id,)).fetchone()
#             if row:
#                 display_label = row["name"]
#         else:
#             row = c.execute("SELECT first_name, last_name FROM approved_users WHERE id=?", (entity_id,)).fetchone()
#             if row:
#                 display_label = f"{row['first_name']} {row['last_name']}"

                

#     return render_template(
#         "reserve.html",
#         entity_type=entity_type,
#         entity_id=entity_id,
#         selected_date=selected_date,
#         available_times=available_times,
#         display_label=display_label
#     )


# @app.route("/reserve", methods=["GET", "POST"])
# @verified_required
# def reserve():
#     email = session.get("email")
#     conn = get_db()
#     conn.row_factory = sqlite3.Row
#     c = conn.cursor()

#     # Ensure user is an approved user
#     c.execute("SELECT id FROM approved_users WHERE lower(email)=?", (email,))
#     user = c.fetchone()
#     if not user:
#         return redirect(url_for("login"))
#     user_id = user["id"]

#     # -----------------------------
#     # 1) Load companies dynamically
#     # -----------------------------
#     companies = c.execute("""
#         SELECT DISTINCT company_name 
#         FROM approved_users
#         WHERE company_name IS NOT NULL AND company_name <> ''
#         ORDER BY company_name
#     """).fetchall()

#     # GET parameters
#     selected_company = request.args.get("company_name")
#     entity_id = request.args.get("entity_id")
#     selected_date = request.args.get("date")

#     # ----------------------------------------
#     # 2) Load people belonging to a company
#     # ----------------------------------------
#     people = []
#     if selected_company:
#         people = c.execute("""
#             SELECT id, first_name, last_name, email 
#             FROM approved_users
#             WHERE company_name = ?
#             ORDER BY first_name, last_name
#         """, (selected_company,)).fetchall()

#     # --------------------------
#     # 3) POST → Reserve a slot
#     # --------------------------
#     if request.method == "POST":
#         chosen_time = request.form.get("time")
#         slot_id = request.form.get("slot_id")
#         entity_id = request.form.get("entity_id")
#         selected_date = request.form.get("date")
#         entity_type = "person"   # always person now

#         if not entity_id:
#             flash("❌ Please select a person.", "danger")
#             return redirect(url_for("reserve"))

#         # Already booked?
#         c.execute("""
#             SELECT id FROM reservations
#             WHERE entity_type='person' AND entity_id=? 
#               AND date=? AND start_time=? 
#               AND status IN ('Pending', 'Approved')
#         """, (entity_id, selected_date, chosen_time))
#         if c.fetchone():
#             flash("❌ This slot is already booked.", "danger")
#             return redirect(url_for("reserve", company_name=selected_company, entity_id=entity_id, date=selected_date))

#         # Assign room
#         rooms = [r["name"] for r in c.execute("SELECT name FROM rooms").fetchall()]
#         c.execute("""
#             SELECT room_name FROM reservations 
#             WHERE date=? AND start_time=? AND status IN ('Pending','Approved')
#         """, (selected_date, chosen_time))
#         taken = [row["room_name"] for row in c.fetchall()]
#         free_room = next((r for r in rooms if r not in taken), None)

#         if not free_room:
#             flash("❌ No rooms left at this time.", "danger")
#             return redirect(url_for("reserve", company_name=selected_company, entity_id=entity_id, date=selected_date))

#         # Get invited person info
#         person = c.execute("""
#             SELECT first_name, last_name, email 
#             FROM approved_users WHERE id=?
#         """, (entity_id,)).fetchone()
#         full_name = f"{person['first_name']} {person['last_name']}"
#         target_email = person["email"]

#         invites_str = f"{email},{target_email}"

#         # Save reservation
#         c.execute("""
#             INSERT INTO reservations 
#             (user_id, entity_type, entity_id, date, start_time, room_name, invites, status, slot_id)
#             VALUES (?,?,?,?,?,?,?,?,?)
#         """, (
#             user_id, "person", entity_id, selected_date,
#             chosen_time, free_room, invites_str, "Pending", slot_id
#         ))
#         conn.commit()

#         # Email
#         subject = f"Meeting Request with {full_name}"
#         body = (
#             f"Hello,\n\nA meeting has been requested with {full_name}.\n\n"
#             f"📅 Date: {selected_date}\n"
#             f"⏰ Time: {chosen_time}\n"
#             f"👤 Person: {full_name}\n"
#             f"🏠 Room: {free_room}\n\n"
#             f"Requested by: {email}\n"
#         )
#         send_plain_email(f"{email},{target_email}", subject, body)

#         flash("✅ Meeting request submitted!", "success")
#         return redirect(url_for("my_meetings"))

#     # -----------------------------
#     # 4) Build available times list
#     # -----------------------------
#     available_times = []
#     if entity_id and selected_date:
#         # Fixed 30-min slots from 10:00 → 14:30
#         start_hour = 10
#         end_hour = 15
#         slot_counter = 0
#         for hour in range(start_hour, end_hour):
#             for minute in (0, 30):
#                 if hour == 14 and minute == 30:
#                     continue
#                 t = f"{hour:02d}:{minute:02d}"

#                 # Check booking
#                 taken = c.execute("""
#                     SELECT 1 FROM reservations
#                     WHERE entity_type='person' AND entity_id=? 
#                       AND date=? AND start_time=?
#                       AND status IN ('Pending','Approved')
#                 """, (entity_id, selected_date, t)).fetchone()

#                 if not taken:
#                     slot_counter += 1
#                     available_times.append({
#                         "time": t,
#                         "slot_id": f"p{slot_counter}"
#                     })

#     return render_template(
#         "reserve.html",
#         companies=companies,
#         people=people,
#         selected_company=selected_company,
#         selected_date=selected_date,
#         entity_id=entity_id,
#         available_times=available_times
#     )

@app.route("/reserve", methods=["GET", "POST"])
@verified_required
def reserve():
    email = session.get("email")
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Ensure user is an approved user
    c.execute("SELECT id FROM approved_users WHERE lower(email)=?", (email,))
    user = c.fetchone()
    if not user:
        return redirect(url_for("login"))
    user_id = user["id"]

    # -----------------------------
    # 1) Load companies dynamically
    # -----------------------------
    companies = c.execute("""
        SELECT DISTINCT company_name 
        FROM approved_users
        WHERE company_name IS NOT NULL AND company_name <> ''
        ORDER BY company_name
    """).fetchall()

    # GET parameters
    selected_company = request.args.get("company_name")
    entity_id = request.args.get("entity_id")
    selected_date = request.args.get("date")

    # ----------------------------------------
    # 2) Load people belonging to a company
    # ----------------------------------------
    people = []
    if selected_company:
        people = c.execute("""
            SELECT id, first_name, last_name, email 
            FROM approved_users
            WHERE company_name = ?
            ORDER BY first_name, last_name
        """, (selected_company,)).fetchall()

    # --------------------------
    # 3) POST → Reserve a slot
    # --------------------------
    if request.method == "POST":
        chosen_time = request.form.get("time")      # start time (HH:MM)
        slot_id = request.form.get("slot_id")
        entity_id = request.form.get("entity_id")
        selected_date = request.form.get("date")
        entity_type = "person"   # always person now

        if not entity_id:
            flash("❌ Please select a person.", "danger")
            return redirect(url_for("reserve"))

        # Already booked?
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

        # Get invited person info
        person = c.execute("""
            SELECT first_name, last_name, email 
            FROM approved_users WHERE id=?
        """, (entity_id,)).fetchone()
        full_name = f"{person['first_name']} {person['last_name']}"
        target_email = person["email"]

        invites_str = f"{email},{target_email}"

        # Save reservation
        c.execute("""
            INSERT INTO reservations 
            (user_id, entity_type, entity_id, date, start_time, room_name, invites, status, slot_id)
            VALUES (?,?,?,?,?,?,?,?,?)
        """, (
            user_id, "person", entity_id, selected_date,
            chosen_time, free_room, invites_str, "Pending", slot_id
        ))
        conn.commit()

        # Email
        subject = f"Meeting Request with {full_name}"
        body = (
            f"Hello,\n\nA meeting has been requested with {full_name}.\n\n"
            f"📅 Date: {selected_date}\n"
            f"⏰ Time: {chosen_time}\n"
            f"👤 Person: {full_name}\n"
            f"🏠 Room: {free_room}\n\n"
            f"Requested by: {email}\n"
        )
        send_plain_email(f"{email},{target_email}", subject, body)

        flash("✅ Meeting request submitted!", "success")
        return redirect(url_for("my_meetings"))

    # -----------------------------
    # 4) Build available times list
    # -----------------------------
    available_times = []
    if entity_id and selected_date:
        # ✅ NEW: 20-min slots from 09:00 → 17:00
        start_minutes = 9 * 60      # 09:00
        end_minutes = 17 * 60       # 17:00
        slot_counter = 0
        current = start_minutes

        while current < end_minutes:
            slot_start = current
            slot_end = current + 20   # 20 minutes

            if slot_end > end_minutes:
                break

            sh = slot_start // 60
            sm = slot_start % 60
            eh = slot_end // 60
            em = slot_end % 60

            start_str = f"{sh:02d}:{sm:02d}"
            end_str = f"{eh:02d}:{em:02d}"

            # Check booking for that start time
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

# CALENDAR_SCOPES = ['https://www.googleapis.com/auth/calendar']

# def create_calendar_event(summary, description, start_datetime, end_datetime, attendees, timezone='Asia/Beirut'):
#     creds = None
#     # token_path = 'token_calendar.json'
#     # creds_path = 'credentials_calendar.json'

#     creds_json = os.environ.get("CALENDAR_CREDS_JSON")
#     token_json = os.environ.get("CALENDAR_TOKEN_JSON")




#     if os.path.exists(token_path):
#         creds = Credentials.from_authorized_user_file(token_path, CALENDAR_SCOPES)

#     if not creds or not creds.valid:
#         if creds and creds.expired and creds.refresh_token:
#             creds.refresh(Request())
#         else:
#             flow = InstalledAppFlow.from_client_secrets_file(creds_path, CALENDAR_SCOPES)
#             creds = flow.run_local_server(port=0, prompt='consent')
#         with open(token_path, 'w') as token:
#             token.write(creds.to_json())

#     service = build('calendar', 'v3', credentials=creds)

#     event = {
#         'summary': summary,
#         'description': description,
#         'start': {'dateTime': start_datetime.isoformat(), 'timeZone': timezone},
#         'end': {'dateTime': end_datetime.isoformat(), 'timeZone': timezone},
#         'attendees': [{'email': e} for e in sorted(set(attendees))],
#         'reminders': {
#             'useDefault': False,
#             'overrides': [
#                 {'method': 'email', 'minutes': 60},
#                 {'method': 'popup', 'minutes': 10},
#             ],
#         },
#     }

#     created = service.events().insert(
#         calendarId='primary',
#         body=event,
#         sendUpdates='all'
#     ).execute()

#     print("✅ Calendar event:", created.get('htmlLink'))
#     return created


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
    Approve a meeting where the logged-in user is the invited person.
    """
    email = session.get("email")
    if not email:
        return redirect(url_for("login"))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Fetch reservation and ensure this user is the target person
    reservation = c.execute("""
        SELECT r.*,
               req.first_name || ' ' || req.last_name AS requester_name,
               req.email  AS requester_email,
               req.company_name AS requester_company,
               target.id  AS target_id,
               target.first_name || ' ' || target.last_name AS target_name,
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
        # Someone else trying to approve
        abort(403)

    # Build recipients list: requester + all invites + target
    invites_list = [e.strip() for e in (reservation["invites"] or "").split(",") if e.strip()]
    recipients = invites_list + [reservation["requester_email"], reservation["target_email"]]
    # make unique in order
    recipients = list(dict.fromkeys(recipients))

    # Update status
    c.execute("UPDATE reservations SET status = 'Approved' WHERE id = ?", (reservation_id,))
    conn.commit()

    # Send approval email + calendar event
    try:
        subject = f"Meeting Request Approved - {reservation['target_name']}"
        body = (
            f"Hello,\n\n"
            f"Your meeting request with {reservation['target_name']} has been approved.\n\n"
            f"📅 Date: {reservation['date']}\n"
            f"⏰ Time: {reservation['start_time']}\n"
            f"🏠 Room: {reservation['room_name'] or 'TBD'}\n\n"
            f"Thank you,\nMcKinsey Electronics Team"
        )

        to_field = ", ".join(recipients)
        send_plain_email(to_field, subject, body)
        flash("✅ Meeting approved and email sent.", "success")

        # Calendar event
        start_str = f"{reservation['date']} {reservation['start_time']}"
        start_dt = datetime.strptime(start_str, "%Y-%m-%d %H:%M")
        end_dt = start_dt + timedelta(minutes=30)

        summary = f"Meeting with {reservation['target_name']}"
        description = (
            f"Approved meeting between {reservation['requester_name']} "
            f"and {reservation['target_name']}.\n"
            f"Room: {reservation['room_name'] or 'TBD'}"
        )

        create_calendar_event(summary, description, start_dt, end_dt, recipients)
        flash("📅 Meeting added to Google Calendar for all attendees.", "success")

    except Exception as e:
        flash(f"⚠️ Meeting approved but email or calendar invite failed: {e}", "warning")

    conn.close()
    return redirect(url_for("person_requests"))

@app.route("/person_requests/reject/<int:reservation_id>", methods=["POST"])
@verified_required
def reject_person_request(reservation_id):
    """
    Reject a meeting where the logged-in user is the invited person.
    """
    email = session.get("email")
    if not email:
        return redirect(url_for("login"))

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    reservation = c.execute("""
        SELECT r.*,
               req.first_name || ' ' || req.last_name AS requester_name,
               req.email  AS requester_email,
               target.first_name || ' ' || target.last_name AS target_name,
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

    # Recipients: requester + all invites + target
    invites_list = [e.strip() for e in (reservation["invites"] or "").split(",") if e.strip()]
    recipients = invites_list + [reservation["requester_email"], reservation["target_email"]]
    recipients = list(dict.fromkeys(recipients))

    # Update status
    c.execute("UPDATE reservations SET status = 'Rejected' WHERE id = ?", (reservation_id,))
    conn.commit()

    # Send rejection email
    try:
        subject = f"Meeting Request Rejected - {reservation['target_name']}"
        body = (
            f"Hello,\n\n"
            f"Unfortunately, your meeting request with {reservation['target_name']} "
            f"has been rejected.\n\n"
            f"📅 Date: {reservation['date']}\n"
            f"⏰ Time: {reservation['start_time']}\n\n"
            f"Thank you,\nMcKinsey Electronics Team"
        )

        to_field = ", ".join(recipients)
        send_plain_email(to_field, subject, body)

        flash("❌ Meeting rejected and email sent successfully.", "danger")
    except Exception as e:
        flash(f"⚠️ Meeting rejected but email failed to send: {e}", "warning")

    conn.close()
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



def import_companies_and_contacts():
    sh = gc.open_by_key(SHEET_ID)
    
    # Load companies sheet
    ws_companies = sh.worksheet("DSS Company Names")
    rows_companies = ws_companies.get_all_records()
    
    # Load contacts sheet
    ws_contacts = sh.worksheet("DSS Contacts linked to Companies")
    rows_contacts = ws_contacts.get_all_records()

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # 1️⃣ Import Companies
    company_map = {}  # company_code → id in DB
    
    for row in rows_companies:
        code = row["Company Code"]
        name = row["Company Name"].strip()

        # Insert OR IGNORE (avoid duplicates)
        c.execute("""
            INSERT OR IGNORE INTO companies(id, name, description)
            VALUES (?, ?, '')
        """, (code, name))
        
        company_map[code] = code  # DB id = code (we keep same ID)

    # 2️⃣ Import Contacts
    for row in rows_contacts:
        code = row["Company Code"]
        email = row["Contacts Email"].strip().lower()

        if code not in company_map:
            continue  # skip unknown company

        c.execute("""
            INSERT OR IGNORE INTO company_contacts(company_id, email)
            VALUES (?, ?)
        """, (code, email))

    conn.commit()
    conn.close()
    print("✅ Companies & Contacts Imported Successfully")

@app.route("/admin/import_companies", methods=["POST"])
@admin_required
def admin_import_companies():
    import_companies_and_contacts()
    flash("✅ Companies & contacts imported successfully!", "success")
    return redirect(url_for("admin_companies"))

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




# if __name__ == "__main__":
#     app.run(host="0.0.0.0", port=5000, debug=True)

if __name__ == "__main__":
    app.run()
