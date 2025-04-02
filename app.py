import os
import uuid
import pandas as pd
import sqlite3
from flask import Flask, request, render_template, redirect, url_for, flash
from jinja2 import Environment, FileSystemLoader
from werkzeug.utils import secure_filename
import smtplib
from email.message import EmailMessage
from datetime import datetime

app = Flask(__name__)
app.secret_key = "your_secret_key"

UPLOAD_FOLDER = "uploads"
HTML_FOLDER = "static/confirm_pages"
DB_PATH = "submissions.db"
ALLOWED_EXTENSIONS = {"xlsx"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(HTML_FOLDER, exist_ok=True)

# 建立 SQLite 資料庫
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS submissions (
        id TEXT PRIMARY KEY,
        name TEXT,
        email TEXT,
        title TEXT,
        fee INTEGER,
        bank TEXT,
        account TEXT,
        submitted_at TEXT
    )""")
    conn.commit()
    conn.close()

init_db()

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        file = request.files["file"]
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(path)
            generate_confirm_pages(path)
            flash("成功上傳並產生確認頁！")
            return redirect(url_for("upload_file"))
        else:
            flash("請上傳 .xlsx 檔案")
            return redirect(url_for("upload_file"))
    return render_template("upload.html")

def generate_confirm_pages(filepath):
    df = pd.read_excel(filepath)
    env = Environment(loader=FileSystemLoader("templates"))
    template = env.get_template("confirm_template.html")

    for _, row in df.iterrows():
        unique_id = str(uuid.uuid4())[:8]
        filename = f"{unique_id}.html"
        html = template.render(
            id=unique_id,
            name=row["Author"],
            email=row["E-mail"],
            title=row["Title"],
            fee=row["Fee"]
        )
        with open(os.path.join(HTML_FOLDER, filename), "w", encoding="utf-8") as f:
            f.write(html)
        send_email(row["Author"], row["E-mail"], unique_id)

def send_email(name, recipient, page_id):
    msg = EmailMessage()
    msg["Subject"] = "【幻華創造】稿費資訊確認通知信"
    msg["From"] = os.getenv("EMAIL_ADDRESS")
    msg["To"] = recipient
    link = f"{os.getenv('BASE_URL')}/static/confirm_pages/{page_id}.html"
    msg.set_content(f"您好 {name}，\n\n請點選以下連結確認您的稿費資訊並填寫帳戶資料：\n{link}")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(os.getenv("EMAIL_ADDRESS"), os.getenv("EMAIL_PASSWORD"))
        smtp.send_message(msg)

@app.route("/submit", methods=["POST"])
def submit():
    data = request.form
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # 查詢該 ID 是否已經存在
    c.execute("SELECT * FROM submissions WHERE id = ?", (data["id"],))
    existing = c.fetchone()

    if existing:
        conn.close()
        return "⚠️ 您已經填寫過表單了，無需重複提交。"

    from datetime import datetime
    submitted_at = datetime.now().isoformat(timespec="seconds")

    c.execute("INSERT INTO submissions (id, name, email, title, fee, bank, account, submitted_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", (
        data["id"], data["name"], data["email"], data["title"],
        data["fee"], data["bank"], data["account"], submitted_at
    ))
    conn.commit()
    conn.close()
    return "✅ 表單已成功送出，感謝您！"

@app.route("/export")
def export():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM submissions", conn)
    conn.close()
    export_path = "submissions_export.xlsx"
    df.to_excel(export_path, index=False)
    from flask import send_file
    return send_file(export_path, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
