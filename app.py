
from flask import Flask, render_template, request, jsonify, send_from_directory, send_file, session, redirect, url_for
import pandas as pd
from pathlib import Path
from datetime import datetime
import uuid, os, shutil, io, re
from functools import wraps
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from openpyxl import Workbook

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = BASE_DIR / "static" / "uploads"
BACKUP_DIR = BASE_DIR / "backups"
for d in (DATA_DIR, UPLOAD_DIR, BACKUP_DIR):
    d.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.secret_key = "change-this-secret-key-in-production"
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024

TABLES = {
    "users": ["id","username","password","role","status","created_at"],
    "students": ["id","adm_no","form_no","full_name","father_name","mother_name","dob","gender","blood_group",
                 "aadhaar_no","pen_no","class_name","section","status","admission_date","category","religion",
                 "nationality","mobile","alt_mobile","email","address","pin_code","district","state",
                 "transport_fee","photo","remarks","created_at"],
    "enquiries": ["id","enquiry_no","student_name","parent_name","mobile","email","class_name","source",
                  "status","follow_up_date","remarks","created_at"],
    "fees": ["id","receipt_no","adm_no","student_name","fee_type","amount","payment_mode","payment_date","remarks"],
    "attendance": ["id","date","adm_no","student_name","class_name","status","remarks"],
    "staff": ["id","staff_id","name","role","mobile","email","joining_date","salary","status"],
    "marks": ["id","exam_name","exam_date","adm_no","student_name","class_name","subject","max_marks","marks","grade","remarks"],
}

def ensure_files():
    for table, cols in TABLES.items():
        path = DATA_DIR / f"{table}.csv"
        if not path.exists():
            pd.DataFrame(columns=cols).to_csv(path, index=False)
    users = read_table("users")
    if users.empty:
        row = {"id": uuid.uuid4().hex[:12], "username":"admin", "password":"admin123",
               "role":"Super Admin", "status":"Active", "created_at":datetime.now().isoformat(timespec="seconds")}
        write_table("users", pd.DataFrame([row]))

def read_table(table):
    path = DATA_DIR / f"{table}.csv"
    if not path.exists():
        pd.DataFrame(columns=TABLES[table]).to_csv(path, index=False)
    return pd.read_csv(path, dtype=str).fillna("")

def write_table(table, df):
    df.to_csv(DATA_DIR / f"{table}.csv", index=False)

def next_number(prefix, table, field):
    df = read_table(table)
    if df.empty:
        n = 1001
    else:
        nums = pd.to_numeric(df[field].astype(str).str.extract(r"(\d+)$")[0], errors="coerce").dropna()
        n = int(nums.max()) + 1 if not nums.empty else 1001
    return f"{prefix}{n}"

def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("user"):
            if request.path.startswith("/api/"):
                return jsonify({"error":"Authentication required"}), 401
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper

def grade_for(m, max_marks):
    try:
        pct = float(m) / float(max_marks) * 100
        if pct >= 90: return "A+"
        if pct >= 80: return "A"
        if pct >= 70: return "B+"
        if pct >= 60: return "B"
        if pct >= 50: return "C"
        if pct >= 40: return "D"
        return "F"
    except Exception:
        return ""

@app.route("/login", methods=["GET","POST"])
def login():
    ensure_files()
    if request.method == "POST":
        username = request.form.get("username","").strip()
        password = request.form.get("password","")
        users = read_table("users")
        match = users[(users.username == username) & (users.password == password) & (users.status == "Active")]
        if not match.empty:
            session["user"] = username
            session["role"] = match.iloc[0]["role"]
            return redirect(url_for("dashboard"))
        return render_template("login.html", error="Invalid username or password")
    return render_template("login.html")

@app.get("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/")
@login_required
def dashboard(): return render_template("dashboard.html")

@app.route("/new-admission")
@login_required
def new_admission(): return render_template("new_admission.html", edit_id="")

@app.route("/students")
@login_required
def students_page(): return render_template("students.html")

@app.route("/students/edit/<student_id>")
@login_required
def edit_student(student_id): return render_template("new_admission.html", edit_id=student_id)

@app.route("/enquiries")
@login_required
def enquiries_page(): return render_template("enquiries.html")

@app.route("/fees")
@login_required
def fees_page(): return render_template("fees.html")

@app.route("/attendance")
@login_required
def attendance_page(): return render_template("attendance.html")

@app.route("/staff")
@login_required
def staff_page(): return render_template("staff.html")

@app.route("/exams")
@login_required
def exams_page(): return render_template("exams.html")

@app.route("/promotion")
@login_required
def promotion_page(): return render_template("promotion.html")

@app.route("/reports")
@login_required
def reports_page(): return render_template("reports.html")

@app.route("/settings")
@login_required
def settings_page(): return render_template("settings.html")

@app.get("/api/dashboard")
@login_required
def api_dashboard():
    students, fees, enquiries, staff = map(read_table, ["students","fees","enquiries","staff"])
    attendance = read_table("attendance")
    return jsonify({
        "students":len(students),
        "active_students":int((students.status=="Active").sum()) if not students.empty else 0,
        "enquiries":len(enquiries),
        "staff":len(staff),
        "fee_total":float(pd.to_numeric(fees.amount, errors="coerce").fillna(0).sum()) if not fees.empty else 0,
        "today_attendance":int((attendance.date==datetime.now().strftime("%Y-%m-%d")).sum()) if not attendance.empty else 0,
        "recent_students":students.tail(8).to_dict("records")[::-1] if not students.empty else []
    })

@app.get("/api/students")
@login_required
def api_students(): return jsonify(read_table("students").to_dict("records"))

@app.get("/api/students/<student_id>")
@login_required
def api_student(student_id):
    df=read_table("students"); row=df[df.id==student_id]
    if row.empty: return jsonify({"error":"Student not found"}),404
    return jsonify(row.iloc[0].to_dict())

@app.post("/api/students")
@login_required
def api_create_student():
    data=request.form.to_dict(); students=read_table("students")
    photo_name=data.get("existing_photo","")
    photo=request.files.get("photo")
    if photo and photo.filename:
        ext=Path(photo.filename).suffix.lower()
        if ext not in [".jpg",".jpeg",".png",".webp"]: return jsonify({"error":"Only JPG, JPEG, PNG or WEBP images are allowed."}),400
        photo_name=f"{uuid.uuid4().hex}{ext}"; photo.save(UPLOAD_DIR/photo_name)
    sid=data.get("id") or uuid.uuid4().hex[:12]
    row={c:data.get(c,"") for c in TABLES["students"]}
    row.update({"id":sid,"adm_no":data.get("adm_no") or next_number("ADM-26-","students","adm_no"),
                "status":data.get("status","Active"),"nationality":data.get("nationality","Indian"),
                "transport_fee":data.get("transport_fee","0"),"photo":photo_name,
                "created_at":data.get("created_at") or datetime.now().isoformat(timespec="seconds")})
    if sid in students.id.values:
        students.loc[students.id==sid,list(row.keys())]=list(row.values())
    else:
        students=pd.concat([students,pd.DataFrame([row])],ignore_index=True)
    write_table("students",students)
    return jsonify({"ok":True,"student":row})

@app.delete("/api/students/<student_id>")
@login_required
def api_delete_student(student_id):
    df=read_table("students")
    if student_id not in df.id.values:return jsonify({"error":"Student not found"}),404
    write_table("students",df[df.id!=student_id]); return jsonify({"ok":True})

@app.get("/api/enquiries")
@login_required
def api_enquiries(): return jsonify(read_table("enquiries").to_dict("records"))

@app.post("/api/enquiries")
@login_required
def api_create_enquiry():
    data=request.json or {}; df=read_table("enquiries")
    row={c:data.get(c,"") for c in TABLES["enquiries"]}
    row.update({"id":uuid.uuid4().hex[:12],"enquiry_no":data.get("enquiry_no") or next_number("ENQ-26-","enquiries","enquiry_no"),
                "created_at":datetime.now().isoformat(timespec="seconds")})
    write_table("enquiries",pd.concat([df,pd.DataFrame([row])],ignore_index=True)); return jsonify({"ok":True})

@app.get("/api/fees")
@login_required
def api_fees(): return jsonify(read_table("fees").to_dict("records"))

@app.post("/api/fees")
@login_required
def api_create_fee():
    data=request.json or {}; df=read_table("fees")
    row={c:data.get(c,"") for c in TABLES["fees"]}
    row.update({"id":uuid.uuid4().hex[:12],"receipt_no":data.get("receipt_no") or next_number("REC-","fees","receipt_no"),
                "payment_date":data.get("payment_date") or datetime.now().strftime("%Y-%m-%d")})
    write_table("fees",pd.concat([df,pd.DataFrame([row])],ignore_index=True)); return jsonify({"ok":True,"fee":row})

@app.get("/api/attendance")
@login_required
def api_attendance(): return jsonify(read_table("attendance").to_dict("records"))

@app.post("/api/attendance")
@login_required
def api_create_attendance():
    data=request.json or {}; df=read_table("attendance")
    row={c:data.get(c,"") for c in TABLES["attendance"]}
    row.update({"id":uuid.uuid4().hex[:12],"date":data.get("date") or datetime.now().strftime("%Y-%m-%d")})
    write_table("attendance",pd.concat([df,pd.DataFrame([row])],ignore_index=True)); return jsonify({"ok":True})

@app.get("/api/staff")
@login_required
def api_staff(): return jsonify(read_table("staff").to_dict("records"))

@app.post("/api/staff")
@login_required
def api_create_staff():
    data=request.json or {}; df=read_table("staff")
    row={c:data.get(c,"") for c in TABLES["staff"]}
    row.update({"id":uuid.uuid4().hex[:12],"staff_id":data.get("staff_id") or next_number("STF-","staff","staff_id"),
                "status":data.get("status","Active")})
    write_table("staff",pd.concat([df,pd.DataFrame([row])],ignore_index=True)); return jsonify({"ok":True})

@app.get("/api/marks")
@login_required
def api_marks(): return jsonify(read_table("marks").to_dict("records"))

@app.post("/api/marks")
@login_required
def api_create_mark():
    data=request.json or {}; df=read_table("marks")
    row={c:data.get(c,"") for c in TABLES["marks"]}
    row.update({"id":uuid.uuid4().hex[:12],"grade":grade_for(data.get("marks",""),data.get("max_marks","100"))})
    write_table("marks",pd.concat([df,pd.DataFrame([row])],ignore_index=True)); return jsonify({"ok":True})

@app.get("/api/student-report/<adm_no>")
@login_required
def student_report(adm_no):
    students=read_table("students"); s=students[students.adm_no==adm_no]
    if s.empty:return jsonify({"error":"Student not found"}),404
    marks=read_table("marks"); fees=read_table("fees"); att=read_table("attendance")
    m=marks[marks.adm_no==adm_no]; f=fees[fees.adm_no==adm_no]; a=att[att.adm_no==adm_no]
    return jsonify({"student":s.iloc[0].to_dict(),"marks":m.to_dict("records"),"fees":f.to_dict("records"),
                    "attendance":a.to_dict("records")})

@app.get("/export/<table>.xlsx")
@login_required
def export_excel(table):
    if table not in TABLES:return "Invalid table",404
    df=read_table(table); out=io.BytesIO()
    with pd.ExcelWriter(out,engine="openpyxl") as writer: df.to_excel(writer,index=False,sheet_name=table.title())
    out.seek(0); return send_file(out,as_attachment=True,download_name=f"{table}_export.xlsx",
                                  mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/receipt/<receipt_no>.pdf")
@login_required
def fee_receipt(receipt_no):
    df=read_table("fees"); r=df[df.receipt_no==receipt_no]
    if r.empty:return "Receipt not found",404
    x=r.iloc[0]; out=io.BytesIO(); c=canvas.Canvas(out,pagesize=A4); w,h=A4
    c.setFillColor(colors.HexColor("#064f47")); c.rect(0,h-90,w,90,fill=1,stroke=0)
    c.setFillColor(colors.white); c.setFont("Helvetica-Bold",20); c.drawString(45,h-55,"DATA CLICK EDUCATION")
    c.setFont("Helvetica",10); c.drawString(45,h-72,"Official Fee Receipt")
    c.setFillColor(colors.black); c.setFont("Helvetica-Bold",12); c.drawString(45,h-125,f"Receipt No: {x.receipt_no}")
    c.setFont("Helvetica",11); y=h-165
    for k,v in [("Admission No",x.adm_no),("Student",x.student_name),("Fee Type",x.fee_type),
                ("Amount",f"Rs. {x.amount}"),("Payment Mode",x.payment_mode),("Payment Date",x.payment_date)]:
        c.drawString(55,y,f"{k}: {v}"); y-=28
    c.line(55,y-10,w-55,y-10); c.setFont("Helvetica",9); c.drawString(55,y-35,"This is a computer-generated receipt.")
    c.save(); out.seek(0)
    return send_file(out,as_attachment=True,download_name=f"{receipt_no}.pdf",mimetype="application/pdf")

@app.get("/id-card/<adm_no>.pdf")
@login_required
def id_card(adm_no):
    df=read_table("students"); s=df[df.adm_no==adm_no]
    if s.empty:return "Student not found",404
    x=s.iloc[0]; out=io.BytesIO(); c=canvas.Canvas(out,pagesize=(340,215))
    c.setFillColor(colors.HexColor("#064f47")); c.roundRect(8,8,324,199,14,fill=1,stroke=0)
    c.setFillColor(colors.white); c.setFont("Helvetica-Bold",14); c.drawString(20,185,"DATA CLICK EDUCATION")
    c.setFont("Helvetica",8); c.drawString(20,171,"STUDENT IDENTITY CARD")
    c.setFillColor(colors.HexColor("#eafff9")); c.roundRect(20,45,78,105,7,fill=1,stroke=0)
    if x.photo and (UPLOAD_DIR/x.photo).exists():
        try: c.drawImage(str(UPLOAD_DIR/x.photo),24,49,width=70,height=97,preserveAspectRatio=True,anchor='c')
        except: pass
    c.setFillColor(colors.white); c.setFont("Helvetica-Bold",12); c.drawString(112,140,str(x.full_name)[:30])
    c.setFont("Helvetica",9)
    y=122
    for label,val in [("ADM NO",x.adm_no),("CLASS",f"{x.class_name} {x.section}"),("DOB",x.dob),("MOBILE",x.mobile),("FATHER",x.father_name)]:
        c.drawString(112,y,f"{label}: {val}"); y-=18
    c.setFont("Helvetica",7); c.drawString(20,25,"If found, please return this card to the institution.")
    c.save(); out.seek(0)
    return send_file(out,as_attachment=True,download_name=f"{adm_no}_id_card.pdf",mimetype="application/pdf")

@app.post("/api/promote")
@login_required
def promote():
    data=request.json or {}; df=read_table("students")
    ids=set(data.get("ids",[])); to_class=data.get("to_class","")
    if not ids or not to_class:return jsonify({"error":"Select students and destination class"}),400
    df.loc[df.id.isin(ids),"class_name"]=to_class
    write_table("students",df); return jsonify({"ok":True,"count":len(ids)})

@app.post("/api/backup")
@login_required
def backup():
    stamp=datetime.now().strftime("%Y%m%d_%H%M%S"); dest=BACKUP_DIR/f"backup_{stamp}"; dest.mkdir()
    for p in DATA_DIR.glob("*.csv"): shutil.copy2(p,dest/p.name)
    shutil.make_archive(str(dest), "zip", dest)
    return jsonify({"ok":True,"file":str(dest)+".zip"})

@app.get("/uploads/<path:filename>")
@login_required
def uploads(filename): return send_from_directory(UPLOAD_DIR,filename)

if __name__=="__main__":
    ensure_files()
    app.run(debug=True,host="127.0.0.1",port=5000)
