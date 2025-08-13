from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify, session
import sqlite3
from docx import Document
# ---- 发邮件相关 ----
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import formataddr
import traceback
import os
from functools import wraps

app = Flask(__name__)

# ========== 基本配置 ==========
app.secret_key = os.getenv("SECRET_KEY", "replace-this-in-prod")

# 直接使用固定密码（不再被环境变量覆盖）
ADMIN_PASSWORD = "maslandit33918"
print(">>> ADMIN_PASSWORD source: CODE")

# ========== 邮件配置（记得用应用专用密码）==========
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT   = int(os.getenv("SMTP_PORT", "587"))
SENDER_EMAIL    = os.getenv("SENDER_EMAIL", "qinmo840@gmail.com")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD", "izbw wime pzgn fyre")
ADMIN_EMAIL     = os.getenv("ADMIN_EMAIL", "lausukyork9@gmail.com")

# ========== 数据库路径（固定到用户目录，不受代码覆盖影响）==========
HOME_DIR = os.path.expanduser("~")
DATA_DIR = os.path.join(HOME_DIR, "masland-data")
os.makedirs(DATA_DIR, exist_ok=True)
# 允许可选用 DB_PATH 环境变量覆盖；未设置则用固定路径（满足你“push 不覆盖数据”的要求）
DB_PATH = os.getenv("DB_PATH") or os.path.join(DATA_DIR, "database.db")
print(">>> Using DB file:", DB_PATH)

# ========== 器材清单（与前端 index.html 的 equip_map 对应）==========
EQUIP_MAP = {
    "mic": "麦克风",
    "amp": "扩音器",
    "pa": "音响系统",
    "projector": "投影机",
    "screen": "投影屏幕",
    "ext": "延长线",
    "table": "桌子",
    "chair": "椅子",
    "podium": "讲台",
    "hdmi": "HDMI线",
}

# ===== 稳健版邮件发送函数（返回 (ok, err)）=====
def send_email(subject, content, to_email):
    msg = MIMEMultipart()
    msg['From'] = formataddr(("福源堂器材外借系统", SENDER_EMAIL))
    msg['To'] = to_email
    msg['Subject'] = Header(subject, "utf-8")
    msg.attach(MIMEText(content, "plain", "utf-8"))

    # 先尝试 SSL:465
    try:
        server = smtplib.SMTP_SSL(SMTP_SERVER, 465, timeout=20)
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, [to_email], msg.as_string())
        server.quit()
        print(f"✅ 邮件已发送至 {to_email}（SSL:465）")
        return True, None
    except Exception as e_ssl:
        print("⚠️ SSL(465) 发送失败：", e_ssl)
        print(traceback.format_exc())

    # 回退 TLS:587
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=20)
        server.ehlo(); server.starttls(); server.ehlo()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, [to_email], msg.as_string())
        server.quit()
        print(f"✅ 邮件已发送至 {to_email}（TLS:587）")
        return True, None
    except Exception as e_tls:
        print("❌ 邮件发送失败（TLS:587）：", e_tls)
        print(traceback.format_exc())
        return False, str(e_tls)

# ========================
# 数据库初始化（含自动补列 + 持久 WAL 模式）
# ========================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS submissions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        phone TEXT,
        email TEXT,
        group_name TEXT,
        event_name TEXT,
        start_date TEXT,
        start_time TEXT,
        end_date TEXT,
        end_time TEXT,
        location TEXT,
        event_type TEXT,
        participants TEXT,
        equipment TEXT,
        special_request TEXT,
        donation TEXT,
        donation_method TEXT,
        remarks TEXT,
        emergency_name TEXT,
        emergency_phone TEXT,
        status TEXT DEFAULT '待审核',
        review_comment TEXT
    )''')

    # —— 永久设置 WAL & 同步策略（更抗丢）——
    try:
        c.execute("PRAGMA journal_mode=WAL")
        c.execute("PRAGMA synchronous=NORMAL")
    except Exception:
        pass

    # —— 兼容旧库：补列 review_comment ——
    try:
        c.execute("PRAGMA table_info(submissions)")
        cols = [row[1] for row in c.fetchall()]
        if "review_comment" not in cols:
            c.execute("ALTER TABLE submissions ADD COLUMN review_comment TEXT")
    except Exception:
        pass

    conn.commit()
    conn.close()

init_db()

# ========================
# 禁止缓存（确保每次加载 /admin 都取数据库最新）
# ========================
@app.after_request
def add_no_cache_headers(resp):
    if request.path in ("/", "/admin"):
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0, private"
        resp.headers["Pragma"] = "no-cache"
        resp.headers["Expires"] = "0"
    return resp

# ========================
# 登录保护
# ========================
def login_required(view_func):
    @wraps(view_func)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login", next=request.path))
        return view_func(*args, **kwargs)
    return wrapper

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        password = request.form.get("password", "")
        if password == ADMIN_PASSWORD:
            session["logged_in"] = True
            return redirect(request.args.get("next") or url_for("admin"))
        else:
            error = "密码错误，请重试。"
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.pop("logged_in", None)
    return redirect(url_for("login"))

# ========================
# 前台
# ========================
@app.route("/")
def index():
    # 把 equip_map 传给模板（你的前端用它生成器材项）
    return render_template("index.html", equip_map=EQUIP_MAP)

@app.route("/submit", methods=["POST"])
def submit():
    data = request.form.to_dict(flat=True)

    # 读取 equip_*** + 数量，拼成 “中文名x数量”
    equip_items = []
    for key, cname in EQUIP_MAP.items():
        if data.get(f"equip_{key}") == "on":
            qty_str = (data.get(f"equip_{key}_qty") or "").strip()
            try:
                qty = int(qty_str)
            except:
                qty = 0
            if qty <= 0:
                qty = 1
            equip_items.append(f"{cname}x{qty}")
    equipment_str = ", ".join(equip_items)

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        INSERT INTO submissions (
            name, phone, email, group_name, event_name,
            start_date, start_time, end_date, end_time,
            location, event_type, participants, equipment,
            special_request, donation, donation_method,
            remarks, emergency_name, emergency_phone
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        data.get('name'), data.get('phone'), data.get('email'), data.get('group'),
        data.get('event_name'), data.get('start_date'), data.get('start_time'),
        data.get('end_date'), data.get('end_time'), data.get('location'),
        data.get('event_type'), data.get('participants'), equipment_str,
        data.get('special_request'), data.get('donation'), data.get('donation_method'),
        data.get('remarks'), data.get('emergency_name'), data.get('emergency_phone')
    ))
    conn.commit(); conn.close()

    send_email("【新申请】福源堂器材外借",
               f"申请人：{data.get('name')}\n活动：{data.get('event_name')}\n电话：{data.get('phone')}\n邮箱：{data.get('email')}",
               ADMIN_EMAIL)

    # 成功页：提示如何“查询状态”，并提供“返回首页”
    return """
<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>提交成功</title>
<meta http-equiv="Cache-Control" content="no-store, no-cache, must-revalidate, max-age=0"/>
<meta http-equiv="Pragma" content="no-cache"/>
<meta http-equiv="Expires" content="0"/>
<style>
  body{background:#f7f7f7;font-family:"Microsoft YaHei",sans-serif;color:#111827;padding:24px;}
  .card{max-width:600px;margin:60px auto;background:#fff;border:1px solid #e8ecf2;border-radius:14px;box-shadow:0 8px 20px rgba(0,0,0,.06);padding:22px;}
  h1{margin:0 0 8px;font-size:22px;font-weight:900;}
  p{margin:8px 0 0;line-height:1.7}
  .btns{margin-top:16px;display:flex;gap:10px}
  .btn{appearance:none;border:none;padding:10px 16px;border-radius:10px;cursor:pointer;font-weight:900}
  .btn-primary{background:#2563eb;color:#fff;box-shadow:0 3px 0 #1d4ed8,0 8px 16px rgba(0,0,0,.08)}
  .btn-primary:hover{filter:brightness(.95)}
</style>
</head>
<body>
  <div class="card">
    <h1>提交成功！</h1>
    <p>我们已收到您的申请，管理员会尽快处理。</p>
    <p>可回到填写页面，点击上方<strong>【查询状态】</strong>查看审核结果。</p>
    <div class="btns">
      <button class="btn btn-primary" onclick="location.href='/'">返回首页</button>
    </div>
  </div>
</body>
</html>
    """

# ========================
# 管理页 + 接口
# ========================
@app.route("/admin")
@login_required
def admin():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT * FROM submissions ORDER BY id DESC")
    submissions = c.fetchall()
    conn.close()
    return render_template("admin.html", submissions=submissions)

@app.route("/api/submission/<int:submission_id>")
@login_required
def api_submission(submission_id):
    """读取单条记录的最新状态，供前端在更新后再次拉取，确保显示数据库里的最终值。"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""SELECT id, name, email, event_name, status, review_comment
                 FROM submissions WHERE id=?""", (submission_id,))
    row = c.fetchone()
    conn.close()
    if not row:
        return jsonify({"success": False, "message": "记录不存在"}), 404
    return jsonify({
        "success": True,
        "data": {
            "id": row[0],
            "name": row[1],
            "email": row[2],
            "event_name": row[3],
            "status": row[4] or "待审核",
            "review_comment": row[5] or ""
        }
    })

@app.route("/update_status/<int:submission_id>/<string:new_status>", methods=["POST"])
@login_required
def update_status(submission_id, new_status):
    """异常也返回 JSON，且更新后尝试给申请人发邮件（失败不影响接口）"""
    try:
        data = request.get_json(silent=True) or {}
        comment = data.get("comment", "")

        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("UPDATE submissions SET status=?, review_comment=? WHERE id=?",
                  (new_status, comment, submission_id))
        conn.commit()
        affected = c.rowcount
        c.execute("SELECT name, email, status FROM submissions WHERE id=?", (submission_id,))
        row = c.fetchone()
        conn.close()

        if affected == 0:
            return jsonify({"success": False, "message": "未找到要更新的记录"}), 404

        # 发邮件通知申请人（可失败但不影响接口）
        try:
            if row and row[1]:
                send_email("【审核结果】福源堂器材外借申请",
                           f"您好 {row[0]}，您的申请已被审核为：{row[2]}\n审核说明：{comment or '无'}",
                           row[1])
        except Exception as mail_err:
            print("⚠️ 审核后通知申请人失败（忽略）：", mail_err)

        return jsonify({
            "success": True,
            "submission_id": submission_id,
            "name": row[0] if row else "",
            "status": row[2] if row else new_status
        })
    except Exception as e:
        print("❌ /update_status 出错：", e)
        print(traceback.format_exc())
        return jsonify({"success": False, "message": f"服务器错误：{e}"}), 500

@app.route("/send_review_email/<int:submission_id>", methods=["POST"])
@login_required
def send_review_email(submission_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT name, email, event_name, status, review_comment FROM submissions WHERE id=?",
              (submission_id,))
    row = c.fetchone()
    conn.close()

    if not row:
        return jsonify({"success": False, "message": "记录不存在"}), 404

    name, email, event_name, status, review_comment = row
    if not email:
        return jsonify({"success": False, "message": "该记录没有填写邮箱，无法发送"}), 400

    ok, err = send_email(f"【审核结果】{event_name or ''}",
                         f"您好 {name or ''}：\n\n您的申请（活动：{event_name or '-'}) 审核结果为：{status or '待审核'}\n审核说明：{review_comment or '无'}\n\n如有疑问请回复此邮件联系管理员。",
                         email)
    if ok:
        return jsonify({"success": True, "message": f"已发送到 {email}"})
    else:
        return jsonify({"success": False, "message": f"发送失败：{err}"}), 500

@app.route("/delete_submission/<int:submission_id>", methods=["POST"])
@login_required
def delete_submission(submission_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM submissions WHERE id=?", (submission_id,))
    affected = c.rowcount
    conn.commit(); conn.close()
    return jsonify({"success": True, "submission_id": submission_id, "deleted": affected})

@app.route("/download/<int:submission_id>")
@login_required
def download(submission_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT * FROM submissions WHERE id=?", (submission_id,))
    submission = c.fetchone()
    conn.close()
    if not submission:
        return "记录不存在"

    doc = Document()
    doc.add_heading('申请表详情', level=1)
    fields = [
        "ID","姓名","电话","邮箱","团体名称","活动名称",
        "开始日期","开始时间","结束日期","结束时间","地点","活动类型",
        "参与人数","器材","特别需求","捐款","捐款方式",
        "备注","紧急联系人","紧急联系电话","审核状态","审核说明"
    ]
    for i, field in enumerate(fields):
        doc.add_paragraph(f"{field}: {submission[i]}")

    file_path = f"submission_{submission_id}.docx"
    doc.save(file_path)
    return send_file(file_path, as_attachment=True)

@app.route("/check_status_api")
def check_status_api():
    name = request.args.get("name")
    if not name:
        return jsonify({"status": "error", "message": "Name is required"})
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT name, event_name, status, review_comment FROM submissions WHERE name = ?", (name,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return jsonify({"status": row[2],
                        "data": {"name": row[0], "event_name": row[1],
                                 "review_status": row[2], "review_comment": row[3] or ""}})
    else:
        return jsonify({"status": "not_found"})

@app.route("/_health")
def _health():
    return "ok", 200

if __name__ == "__main__":
    app.run(debug=True)
