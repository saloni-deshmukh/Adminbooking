from flask import Flask, request, jsonify, render_template, session, redirect, url_for
import pandas as pd
from datetime import datetime
import win32com.client as win32

app = Flask(__name__)
app.secret_key = "nttdata_secret_key_2024"

ROOM_FILE      = "RoomMaster.xlsx"
BOOKING_FILE   = "Bookings.xlsx"
EMPLOYEE_FILE  = "login.xlsx"   # Excel with Employee Name, Emp_ID, Email, Password

LOGO_BASE64 = "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"

ADMIN_CREDENTIALS = {
    "email":    "admin",          # admin logs in with literal "admin" as email
    "password": "Admin@123",
    "name":     "Admin",
    "role":     "admin"
}
ADMIN_EMAIL = "Abhigna.Javvaji@bs.nttdata.com"


# ── HELPERS ───────────────────────────────────────────────────────────

def time_to_minutes(t):
    t = str(t).strip()[:5]
    h, m = t.split(":")
    return int(h) * 60 + int(m)


def normalize_date(d):
    return pd.to_datetime(d).strftime("%Y-%m-%d")


def load_employees():
    """Load employee records from Employees.xlsx.
    Expected columns: Employee Name, Emp_ID, Email, Password
    """
    try:
        df = pd.read_excel(EMPLOYEE_FILE)
        df.columns = [c.strip() for c in df.columns]
        return df
    except Exception as e:
        print(f"Employee file error: {e}")
        return pd.DataFrame(columns=["Employee Name", "Emp_ID", "Email", "Password"])


def validate_employee(email_input, password_input):
    """Return employee dict or None."""
    df = load_employees()
    if df.empty:
        return None
    # Match by Email (case-insensitive) and Password
    match = df[
        (df["Email"].str.strip().str.lower() == email_input.lower()) &
        (df["Password"].str.strip() == password_input)
    ]
    if match.empty:
        return None
    row = match.iloc[0]
    return {
        "name":  str(row.get("Employee Name", "Employee")).strip(),
        "email": str(row["Email"]).strip(),
        "emp_id": str(row.get("Emp_ID", "")).strip(),
        "role":  "employee"
    }


def load_rooms():
    all_sheets = pd.read_excel(ROOM_FILE, sheet_name=None)
    dfs = [df for name, df in all_sheets.items() if name != "Bookingsdummy"]
    return pd.concat(dfs, ignore_index=True)


def load_bookings():
    try:
        df = pd.read_excel(BOOKING_FILE, sheet_name="Sheet1")
        if not df.empty:
            df["Date"] = df["Date"].apply(normalize_date)
        return df
    except:
        return pd.DataFrame(columns=[
            "Booking_ID", "Name", "Room_ID", "Location", "Floor",
            "No. of people", "Date", "Start_Time", "End_Time",
            "Employee_Name", "Email", "Purpose",
            "Booking date", "Booking time", "Status", "Admin_Comment"
        ])


def save_booking(row):
    try:
        df = pd.read_excel(BOOKING_FILE, sheet_name="Sheet1")
    except:
        df = pd.DataFrame()
    updated = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    with pd.ExcelWriter(BOOKING_FILE, engine="openpyxl", mode="w") as w:
        updated.to_excel(w, sheet_name="Sheet1", index=False)


def update_booking_status(booking_id, status, comment=""):
    df = pd.read_excel(BOOKING_FILE, sheet_name="Sheet1")
    if "Admin_Comment" in df.columns:
        df["Admin_Comment"] = df["Admin_Comment"].astype(str)
    if "Status" in df.columns:
        df["Status"] = df["Status"].astype(str)
    df.loc[df["Booking_ID"] == booking_id, "Status"] = status
    df.loc[df["Booking_ID"] == booking_id, "Admin_Comment"] = comment
    with pd.ExcelWriter(BOOKING_FILE, engine="openpyxl", mode="w") as w:
        df.to_excel(w, sheet_name="Sheet1", index=False)
    return df[df["Booking_ID"] == booking_id].iloc[0]


def send_email_outlook(to_email, subject, html_body):
    """Send via Outlook desktop client (Windows only)."""
    try:
        if not to_email or "@" not in str(to_email):
            print(f"Email skipped – invalid address: {to_email!r}")
            return False
        import pythoncom
        pythoncom.CoInitialize()
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.To = str(to_email).strip()
        mail.Subject = subject
        mail.HTMLBody = html_body
        mail.Send()
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
        <img src="data:image/png;base64,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";
" height="44" alt="NTT DATA"/>
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
        email    = data.get("email", "").strip()      # employees use their Excel email
        password = data.get("password", "").strip()

        # ── Admin check (email field contains literal "admin") ──
        if email.lower() == "admin" and password == ADMIN_CREDENTIALS["password"]:
            session["user"]  = "admin"
            session["role"]  = "admin"
            session["name"]  = "Admin"
            session["email"] = ADMIN_EMAIL
            return jsonify({"status": "ok", "role": "admin"})

        # ── Employee check against Excel ──
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
    facilities = data.get("facilities", [])  # list of required facilities

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
                "room_id":   room_id,
                "room_name": room["name"],
                "capacity":  int(room["capacity"]),
                "type":      room["type"],
                "floor":     room["floor"],
                "facilities": str(room.get("facilities", "")).split(",") if room.get("facilities") else [],
            })

    suggest = data.get("suggest", False)
    if not available and suggest:
        # Smart suggestions: same location other floors first, then other locations
        # Sort by: same location first, then by capacity (closest match first)
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
            # Skip same floor (already checked above)
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
                    "room_id":   room_id,
                    "room_name": room["name"],
                    "capacity":  int(room["capacity"]),
                    "type":      room["type"],
                    "floor":     room["floor"],
                    "location":  room["location"],
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
    emp_email = session.get("email", "")   # always from session (Excel)

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

    # ── Confirmation email → employee ──
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
        send_email_outlook(emp_email,
                           f"[NTT DATA] Booking Request Submitted — {booking_id}",
                           build_email_html("Booking Request Submitted", body))

    # ── Notification email → admin ──
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
    send_email_outlook(ADMIN_EMAIL,
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
    action     = data.get("action")   # "approve" or "deny"
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

    print(f"[admin_action] {action.upper()} booking={booking_id}  emp_email={emp_email!r}")

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
            send_email_outlook(emp_email,
                               f"[NTT DATA] ✅ Booking Approved — {booking_id}",
                               build_email_html("Booking Approved!", body, "#1a8a3d"))
        else:
            print(f"[admin_action] SKIP approval email – no valid email for booking {booking_id}")

    else:  # deny
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
            send_email_outlook(emp_email,
                               f"[NTT DATA] ❌ Booking Denied — {booking_id}",
                               build_email_html("Booking Request Denied", body, "#c0392b"))
        else:
            print(f"[admin_action] SKIP denial email – no valid email for booking {booking_id}")

    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(debug=True)
