"""
NTT DATA Room Booking System — Production (Azure App Service)
Changes from local version:
  - win32com (Outlook) replaced with smtplib (works on Linux/Azure)
  - Excel files stored in Azure Blob Storage (not local disk)
  - Secret key and credentials loaded from environment variables
  - debug=False, production WSGI via gunicorn
"""

import os
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

import pandas as pd
from flask import Flask, request, jsonify, render_template, session, redirect, url_for
from azure.storage.blob import BlobServiceClient

# ── APP SETUP ──────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-me-in-production")

# ── AZURE BLOB STORAGE CONFIG ─────────────────────────────────────────
# Set these in Azure App Service → Configuration → Application Settings
AZURE_STORAGE_CONNECTION_STRING = os.environ.get("AZURE_STORAGE_CONNECTION_STRING", "DefaultEndpointsProtocol=https;AccountName=adminbooking;AccountKey=YiFCuEJAeDJ3IUsSHrp/Xh0DvOqGhBwL/4RK0aU3gxoinzHmMQ7UO9i5ogieVq6PuxJczc3gWReo+ASt9/cyUw==;EndpointSuffix=core.windows.net")
BLOB_CONTAINER = os.environ.get("BLOB_CONTAINER", "bookingbot")

ROOM_BLOB     = "RoomMaster.xlsx"
BOOKING_BLOB  = "Bookings.xlsx"
EMPLOYEE_BLOB = "login.xlsx"

# ── EMAIL CONFIG (SMTP) ───────────────────────────────────────────────
# Use your corporate SMTP or any relay (e.g. Office 365, SendGrid, Gmail)
SMTP_HOST     = os.environ.get("SMTP_HOST", "smtp.office365.com")
SMTP_PORT     = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER     = os.environ.get("SMTP_USER", "")      # sender email address
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")  # app password / SMTP password

# ── LOGO & ADMIN ───────────────────────────────────────────────────────
LOGO_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAU4AAABcCAYAAAABOlxNAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAFxEAABcRAcom8z8AABwCSURBVHhe7Z0JlFxVmcc/Rh1xG0V0BjfE5biN2wgzqON2cANHx3FBBwVRlChZut69r6o7QbE5bgNHcdRxVBQVdUSJihhCd9e791X1ko0krCLIMrIMm7KEBBJCSNJz/t+tV6m6773qqu7q7iT1/c75Tifd7771vv+797vf/S6RIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAiCIAjCfktp1ZOoFL+YdPx6Wlw+igKYfSMV7Stpmf07f/M6CzY+hgaqh1FQPYIK5i3UN/RW/hlER1AYP5cGJ//KLyIIgrDvokeeQ3rkJFLR2RTaYdLmClLmVgrMn0mZv5Ayt5My15KyYxSYX5Cyn6ewciSVomdSIXo/KfMN0vZCUnY9KXsjqeguV45/3kiBuYSU/Q2p6EwK7TE0uOLx/ikIgiDsGwT2NaTjsymsXE3aPkTL1k6yDayapP7xSSqN1Wx8kvonJmnpmkk69RL3OwirtleRNpvc79dN0sBqt1293Fit3GpXDvsN47tJl9dSUOmnvhYtWEEQhL0KNfJCUubbpO09LG4QtuLoJGnbnoUVJ4gQSpQP4/Q2WYbtIMI4XljZScpeQ4H9DJUufJJ/ioIgCHsH8EMqezzp+HoWsE5Er9tWrDrxxfFD+zsKhl/jn64gCML8suiCg0mXv0ZhZUddsHwxmw+DgKLlGo7eSDo61j9tQRCE+WFR9ExS5fPrfktfvBqNW4AVJ2gzFdf6fiqTpKfYl3MXbKLALKHJyQP8SxAEQZg7EEIUmAt4YAYi5gtWo8hBVPf4OreRNttT27VtZjcp+yBps8P5NmuDRblibFzXvVjdSioK/MsQBEGYG8LyE0hHP+GWZthCNJO/B/ZWCqKfU1/0EVLV15KOhnNbniy0ENmMvxUrKAPRPZVC84+kTIm0qZCy97tR9xYDUfi7irdSYE70L0cQBGH2UeXTXXc5R6ggiOxfjG/jGMul5vn1bnJgFBUrD7AI+uWS/aFFmdeKLI3tpmJ8LRWrr+b9LV/+KNLRW0lFF1JY2U4DE+kyiXHYUuUOKo2+0bsiQRCEWSSw76Gwcj+3Jn1hgnFI0fgkBVGVZwY1Uqy+jcLKvdz688uhtRhWd5KO/5O0+Sgps"

ADMIN_CREDENTIALS = {
    "email":    "admin",
    "password": os.environ.get("ADMIN_PASSWORD", "Admin@123"),
    "name":     "Admin",
    "role":     "admin"
}
ADMIN_EMAIL = os.environ.get("ADMIN_EMAIL", "admin@nttdata.com")


# ── BLOB STORAGE HELPERS ──────────────────────────────────────────────

def get_blob_client(blob_name):
    service = BlobServiceClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING)
    return service.get_blob_client(container=BLOB_CONTAINER, blob=blob_name)


def read_excel_from_blob(blob_name, sheet_name=0):
    """Download an Excel file from Blob Storage and return a DataFrame."""
    client = get_blob_client(blob_name)
    data = client.download_blob().readall()
    return pd.read_excel(io.BytesIO(data), sheet_name=sheet_name)


def write_excel_to_blob(df, blob_name, sheet_name="Sheet1"):
    """Write a DataFrame back to Blob Storage as Excel."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    buffer.seek(0)
    client = get_blob_client(blob_name)
    client.upload_blob(buffer.read(), overwrite=True)


# ── HELPERS ───────────────────────────────────────────────────────────

def time_to_minutes(t):
    t = str(t).strip()[:5]
    h, m = t.split(":")
    return int(h) * 60 + int(m)


def normalize_date(d):
    return pd.to_datetime(d).strftime("%Y-%m-%d")


def load_employees():
    try:
        df = read_excel_from_blob(EMPLOYEE_BLOB)
        df.columns = [c.strip() for c in df.columns]
        return df
    except Exception as e:
        print(f"Employee file error: {e}")
        return pd.DataFrame(columns=["Employee Name", "Emp_ID", "Email", "Password"])


def validate_employee(email_input, password_input):
    df = load_employees()
    if df.empty:
        return None
    match = df[
        (df["Email"].str.strip().str.lower() == email_input.lower()) &
        (df["Password"].str.strip() == password_input)
    ]
    if match.empty:
        return None
    row = match.iloc[0]
    return {
        "name":   str(row.get("Employee Name", "Employee")).strip(),
        "email":  str(row["Email"]).strip(),
        "emp_id": str(row.get("Emp_ID", "")).strip(),
        "role":   "employee"
    }


def load_rooms():
    all_sheets = read_excel_from_blob(ROOM_BLOB, sheet_name=None)
    dfs = [df for name, df in all_sheets.items() if name != "Bookingsdummy"]
    return pd.concat(dfs, ignore_index=True)


def load_bookings():
    try:
        df = read_excel_from_blob(BOOKING_BLOB, sheet_name="Sheet1")
        if not df.empty:
            df["Date"] = df["Date"].apply(normalize_date)
        return df
    except Exception:
        return pd.DataFrame(columns=[
            "Booking_ID", "Name", "Room_ID", "Location", "Floor",
            "No. of people", "Date", "Start_Time", "End_Time",
            "Employee_Name", "Email", "Purpose",
            "Booking date", "Booking time", "Status", "Admin_Comment"
        ])


def save_booking(row):
    try:
        df = read_excel_from_blob(BOOKING_BLOB, sheet_name="Sheet1")
    except Exception:
        df = pd.DataFrame()
    updated = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    write_excel_to_blob(updated, BOOKING_BLOB, sheet_name="Sheet1")


def update_booking_status(booking_id, status, comment=""):
    df = read_excel_from_blob(BOOKING_BLOB, sheet_name="Sheet1")
    if "Admin_Comment" in df.columns:
        df["Admin_Comment"] = df["Admin_Comment"].astype(str)
    if "Status" in df.columns:
        df["Status"] = df["Status"].astype(str)
    df.loc[df["Booking_ID"] == booking_id, "Status"] = status
    df.loc[df["Booking_ID"] == booking_id, "Admin_Comment"] = comment
    write_excel_to_blob(df, BOOKING_BLOB, sheet_name="Sheet1")
    return df[df["Booking_ID"] == booking_id].iloc[0]


# ── EMAIL (SMTP — works on Linux/Azure) ───────────────────────────────

def send_email_smtp(to_email, subject, html_body):
    """Send email via SMTP (replaces Outlook/win32com)."""
    try:
        if not to_email or "@" not in str(to_email):
            print(f"Email skipped – invalid address: {to_email!r}")
            return False
        if not SMTP_USER or not SMTP_PASSWORD:
            print("Email skipped – SMTP credentials not configured.")
            return False

        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = SMTP_USER
        msg["To"]      = str(to_email).strip()
        msg.attach(MIMEText(html_body, "html"))

        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, str(to_email).strip(), msg.as_string())

        print(f"Email sent → {to_email}")
        return True
    except Exception as e:
        print(f"Email error: {e}")
        return False


def build_email_html(title, body_content, color="#7b52a3"):
    return f"""
    <html><body style="font-family:Arial,sans-serif;background:#f4f0fb;padding:20px">
    <table width="600" style="background:#fff;border-radius:8px;overflow:hidden;margin:auto">
      <tr><td style="background:{color};padding:20px 30px;text-align:left">
        <strong style="color:#fff;font-size:20px">NTT DATA</strong>
      </td></tr>
      <tr><td style="padding:30px">
        <h2 style="color:{color};margin-bottom:16px">{title}</h2>
        {body_content}
        <hr style="margin-top:28px;border:none;border-top:1px solid #ddd"/>
        <p style="color:#888;font-size:11px;margin-top:10px">NTT DATA — Room Booking System &nbsp;|&nbsp; This is an automated email, please do not reply.</p>
      </td></tr>
    </table></body></html>
    """


# ── AUTH ──────────────────────────────────────────────────────────────

@app.route("/")
def root():
    if "user" in session:
        return redirect(url_for("admin_dashboard") if session["role"] == "admin" else url_for("booking"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        data     = request.get_json(force=True)
        email    = data.get("email", "").strip()
        password = data.get("password", "").strip()

        if email.lower() == "admin" and password == ADMIN_CREDENTIALS["password"]:
            session["user"]  = "admin"
            session["role"]  = "admin"
            session["name"]  = "Admin"
            session["email"] = ADMIN_EMAIL
            return jsonify({"status": "ok", "role": "admin"})

        emp = validate_employee(email, password)
        if emp:
            session["user"]  = emp["email"]
            session["role"]  = "employee"
            session["name"]  = emp["name"]
            session["email"] = emp["email"]
            return jsonify({"status": "ok", "role": "employee"})

        return jsonify({"status": "error", "message": "Invalid email or password."}), 401

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ── EMPLOYEE ──────────────────────────────────────────────────────────

@app.route("/booking")
def booking():
    if "user" not in session or session["role"] != "employee":
        return redirect(url_for("login"))
    return render_template("booking.html", name=session["name"], email=session.get("email", ""))


@app.route("/api/filters")
def get_filters():
    rooms_df  = load_rooms()
    locations = rooms_df["location"].unique().tolist()
    return jsonify({"locations": locations})


@app.route("/api/floors")
def get_floors():
    location = request.args.get("location")
    rooms_df = load_rooms()
    floors   = rooms_df[rooms_df["location"] == location]["floor"].unique().tolist()
    return jsonify({"floors": floors})


@app.route("/api/check", methods=["POST"])
def check_availability():
    if "user" not in session:
        return jsonify({"error": "Unauthorized"}), 401

    data       = request.get_json(force=True)
    location   = data.get("location")
    floor      = data.get("floor")
    date       = data.get("date")
    start_time = data.get("start_time")
    end_time   = data.get("end_time")
    people     = int(data.get("people", 0))

    if not all([location, floor, date, start_time, end_time]):
        return jsonify({"error": "Missing fields"}), 400

    req_start = time_to_minutes(start_time)
    req_end   = time_to_minutes(end_time)
    if req_end <= req_start:
        return jsonify({"error": "End time must be after start time"}), 400

    rooms_df    = load_rooms()
    bookings_df = load_bookings()

    filtered = rooms_df[
        (rooms_df["location"] == location) &
        (rooms_df["floor"]    == floor)    &
        (rooms_df["capacity"] >= people)
    ]

    available = []
    for _, room in filtered.iterrows():
        room_id   = room["id"]
        conflicts = bookings_df[
            (bookings_df["Room_ID"] == room_id) &
            (bookings_df["Date"]    == date)    &
            (bookings_df["Status"].isin(["Pending", "Approved"]))
        ] if not bookings_df.empty else pd.DataFrame()

        overlap = False
        for _, b in conflicts.iterrows():
            if req_start < time_to_minutes(b["End_Time"]) and req_end > time_to_minutes(b["Start_Time"]):
                overlap = True
                break

        if not overlap:
            available.append({
                "room_id":    room_id,
                "room_name":  room["name"],
                "capacity":   int(room["capacity"]),
                "type":       room["type"],
                "floor":      room["floor"],
                "facilities": str(room.get("facilities", "")).split(",") if room.get("facilities") else [],
            })

    suggest = data.get("suggest", False)
    if not available and suggest:
        all_rooms = rooms_df[rooms_df["capacity"] >= people].copy()
        all_rooms["_same_loc"] = (all_rooms["location"] == location).astype(int)
        all_rooms = all_rooms.sort_values(["_same_loc", "capacity"], ascending=[False, True])

        seen_ids = set()
        for _, room in all_rooms.iterrows():
            if len(available) >= 6:
                break
            room_id = room["id"]
            if room_id in seen_ids:
                continue
            seen_ids.add(room_id)
            if room["location"] == location and room["floor"] == floor:
                continue
            conflicts = bookings_df[
                (bookings_df["Room_ID"] == room_id) &
                (bookings_df["Date"] == date) &
                (bookings_df["Status"].isin(["Pending", "Approved"]))
            ] if not bookings_df.empty else pd.DataFrame()
            overlap = False
            for _, b in conflicts.iterrows():
                if req_start < time_to_minutes(b["End_Time"]) and req_end > time_to_minutes(b["Start_Time"]):
                    overlap = True
                    break
            if not overlap:
                available.append({
                    "room_id":    room_id,
                    "room_name":  room["name"],
                    "capacity":   int(room["capacity"]),
                    "type":       room["type"],
                    "floor":      room["floor"],
                    "location":   room["location"],
                    "facilities": str(room.get("facilities", "")).split(",") if room.get("facilities") else [],
                })

    return jsonify({"rooms": available})


@app.route("/api/book", methods=["POST"])
def book_room():
    if "user" not in session or session["role"] != "employee":
        return jsonify({"error": "Unauthorized"}), 401

    data       = request.get_json(force=True)
    booking_id = f"BK{datetime.now().strftime('%Y%m%d%H%M%S')}"

    rooms_df  = load_rooms()
    room      = rooms_df[rooms_df["id"] == data["room_id"]].iloc[0]
    emp_email = session.get("email", "")

    row = {
        "Booking_ID":    booking_id,
        "Name":          room["name"],
        "Room_ID":       data["room_id"],
        "Location":      data["location"],
        "Floor":         data["floor"],
        "No. of people": data["people"],
        "Date":          data["date"],
        "Start_Time":    data["start_time"],
        "End_Time":      data["end_time"],
        "Employee_Name": session["name"],
        "Email":         emp_email,
        "Purpose":       data.get("purpose", ""),
        "Facilities":    ", ".join(data.get("facilities", [])),
        "Booking date":  datetime.now().strftime("%Y-%m-%d"),
        "Booking time":  datetime.now().strftime("%H:%M:%S"),
        "Status":        "Pending",
        "Admin_Comment": ""
    }
    save_booking(row)

    body = f"""
    <p>Dear {session['name']},</p>
    <p>Your room booking request has been submitted and is <strong>awaiting admin approval</strong>.</p>
    <table style="border-collapse:collapse;width:100%;margin-top:12px">
      <tr><td style="padding:9px 12px;background:#f0ebfa;font-weight:600;width:140px">Booking ID</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{booking_id}</td></tr>
      <tr><td style="padding:9px 12px;background:#f0ebfa;font-weight:600">Room</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{room['name']}</td></tr>
      <tr><td style="padding:9px 12px;background:#f0ebfa;font-weight:600">Location</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{data['location']} — {data['floor']}</td></tr>
      <tr><td style="padding:9px 12px;background:#f0ebfa;font-weight:600">Date</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{data['date']}</td></tr>
      <tr><td style="padding:9px 12px;background:#f0ebfa;font-weight:600">Time</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{data['start_time']} – {data['end_time']}</td></tr>
      <tr><td style="padding:9px 12px;background:#f0ebfa;font-weight:600">People</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{data['people']}</td></tr>
      <tr><td style="padding:9px 12px;background:#f0ebfa;font-weight:600">Purpose</td><td style="padding:9px 12px">{data.get('purpose','–')}</td></tr>
    </table>
    <p style="margin-top:18px;color:#555">You will receive another email once the admin processes your request.</p>
    """
    if emp_email:
        send_email_smtp(emp_email,
                        f"[NTT DATA] Booking Request Submitted — {booking_id}",
                        build_email_html("Booking Request Submitted", body))

    abody = f"""
    <p>A new room booking request is awaiting your approval.</p>
    <table style="border-collapse:collapse;width:100%;margin-top:12px">
      <tr><td style="padding:9px 12px;background:#fff3e0;font-weight:600;width:140px">Booking ID</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{booking_id}</td></tr>
      <tr><td style="padding:9px 12px;background:#fff3e0;font-weight:600">Employee</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{session['name']} ({emp_email})</td></tr>
      <tr><td style="padding:9px 12px;background:#fff3e0;font-weight:600">Room</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{room['name']}</td></tr>
      <tr><td style="padding:9px 12px;background:#fff3e0;font-weight:600">Location</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{data['location']} — {data['floor']}</td></tr>
      <tr><td style="padding:9px 12px;background:#fff3e0;font-weight:600">Date</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{data['date']}</td></tr>
      <tr><td style="padding:9px 12px;background:#fff3e0;font-weight:600">Time</td><td style="padding:9px 12px">{data['start_time']} – {data['end_time']}</td></tr>
    </table>
    <p style="margin-top:18px;color:#555">Please log in to the admin panel to approve or deny this request.</p>
    """
    send_email_smtp(ADMIN_EMAIL,
                    f"[NTT DATA] New Booking Request — {booking_id}",
                    build_email_html("New Booking Request", abody, "#e65c00"))

    return jsonify({"status": "ok", "booking_id": booking_id})


# ── ADMIN ─────────────────────────────────────────────────────────────

@app.route("/admin")
def admin_dashboard():
    if "user" not in session or session["role"] != "admin":
        return redirect(url_for("login"))
    return render_template("admin.html", name=session["name"])


@app.route("/api/admin/bookings")
def get_all_bookings():
    if "user" not in session or session["role"] != "admin":
        return jsonify({"error": "Unauthorized"}), 401
    df = load_bookings()
    if df.empty:
        return jsonify({"bookings": []})
    return jsonify({"bookings": df.fillna("").to_dict(orient="records")})


@app.route("/api/admin/action", methods=["POST"])
def admin_action():
    if "user" not in session or session["role"] != "admin":
        return jsonify({"error": "Unauthorized"}), 401

    data       = request.get_json(force=True)
    booking_id = data.get("booking_id")
    action     = data.get("action")
    comment    = data.get("comment", "")

    status  = "Approved" if action == "approve" else "Denied"
    booking = update_booking_status(booking_id, status, comment)

    emp_email = str(booking.get("Email", "")).strip()
    emp_name  = str(booking.get("Employee_Name", "Employee")).strip()
    room_name = str(booking.get("Name", "")).strip()
    date      = str(booking.get("Date", "")).strip()
    start     = str(booking.get("Start_Time", "")).strip()
    end       = str(booking.get("End_Time", "")).strip()
    location  = str(booking.get("Location", "")).strip()
    floor     = str(booking.get("Floor", "")).strip()
    people    = str(booking.get("No. of people", "")).strip()
    purpose   = str(booking.get("Purpose", "–")).strip()

    if action == "approve":
        body = f"""
        <p>Dear {emp_name},</p>
        <p>&#127881; Your room booking has been <strong style="color:#1a8a3d">APPROVED</strong>!</p>
        <table style="border-collapse:collapse;width:100%;margin-top:12px">
          <tr><td style="padding:9px 12px;background:#e8f8ee;font-weight:600;width:140px">Booking ID</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{booking_id}</td></tr>
          <tr><td style="padding:9px 12px;background:#e8f8ee;font-weight:600">Room</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{room_name}</td></tr>
          <tr><td style="padding:9px 12px;background:#e8f8ee;font-weight:600">Location</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{location} — {floor}</td></tr>
          <tr><td style="padding:9px 12px;background:#e8f8ee;font-weight:600">Date</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{date}</td></tr>
          <tr><td style="padding:9px 12px;background:#e8f8ee;font-weight:600">Time</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{start} – {end}</td></tr>
          <tr><td style="padding:9px 12px;background:#e8f8ee;font-weight:600">People</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{people}</td></tr>
          <tr><td style="padding:9px 12px;background:#e8f8ee;font-weight:600">Purpose</td><td style="padding:9px 12px">{purpose}</td></tr>
        </table>
        <p style="margin-top:18px;color:#555">Please arrive on time. Contact facilities if you need any assistance.</p>
        """
        if emp_email and "@" in emp_email:
            send_email_smtp(emp_email,
                            f"[NTT DATA] ✅ Booking Approved — {booking_id}",
                            build_email_html("Booking Approved!", body, "#1a8a3d"))
    else:
        body = f"""
        <p>Dear {emp_name},</p>
        <p>We regret to inform you that your room booking has been <strong style="color:#c0392b">DENIED</strong>.</p>
        <table style="border-collapse:collapse;width:100%;margin-top:12px">
          <tr><td style="padding:9px 12px;background:#fde8e8;font-weight:600;width:140px">Booking ID</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{booking_id}</td></tr>
          <tr><td style="padding:9px 12px;background:#fde8e8;font-weight:600">Room</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{room_name}</td></tr>
          <tr><td style="padding:9px 12px;background:#fde8e8;font-weight:600">Date</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{date}</td></tr>
          <tr><td style="padding:9px 12px;background:#fde8e8;font-weight:600">Time</td><td style="padding:9px 12px;border-bottom:1px solid #eee">{start} – {end}</td></tr>
          <tr><td style="padding:9px 12px;background:#fde8e8;font-weight:600">Reason</td><td style="padding:9px 12px">{comment}</td></tr>
        </table>
        <p style="margin-top:18px;color:#555">Please try booking a different time slot or contact the admin for assistance.</p>
        """
        if emp_email and "@" in emp_email:
            send_email_smtp(emp_email,
                            f"[NTT DATA] ❌ Booking Denied — {booking_id}",
                            build_email_html("Booking Request Denied", body, "#c0392b"))

    return jsonify({"status": "ok"})


# ── ENTRYPOINT ────────────────────────────────────────────────────────
if __name__ == "__main__":
    # For local testing only; Azure uses gunicorn
    app.run(debug=False, host="0.0.0.0", port=5000)