"""
批量邮件发送工具 - 企业微信邮箱版
通过企业微信邮箱 SMTP 批量发送个性化邮件
"""

import os
import json
import smtplib
import time
import uuid
import mimetypes
import traceback
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
from pathlib import Path

import xlrd
import openpyxl
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from dotenv import load_dotenv

# 加载 .env 文件中的环境变量
load_dotenv(Path(__file__).parent / ".env")

app = Flask(__name__, static_folder="static")
CORS(app)

UPLOAD_DIR = Path(__file__).parent / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

ATTACH_DIR = Path(__file__).parent / "attachments"
ATTACH_DIR.mkdir(exist_ok=True)

# 全局发送状态追踪
send_tasks = {}


# ── 静态页面 ──────────────────────────────────────────────────────────────────

@app.route("/api/defaults")
def get_defaults():
    """返回环境变量中的默认 SMTP 配置，前端自动填充"""
    return jsonify({
        "smtp_host": os.environ.get("SMTP_HOST", ""),
        "smtp_port": os.environ.get("SMTP_PORT", "465"),
        "email": os.environ.get("SENDER_EMAIL", ""),
        "password_set": bool(os.environ.get("SENDER_PASSWORD")),
        "sender_name": os.environ.get("SENDER_NAME", "CSDN"),
    })


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


# ── 上传并解析 Excel ──────────────────────────────────────────────────────────

@app.route("/api/upload", methods=["POST"])
def upload_excel():
    """上传 Excel 文件，返回列名和前几行预览数据"""
    if "file" not in request.files:
        return jsonify({"error": "未上传文件"}), 400

    f = request.files["file"]
    filename = f.filename
    if not filename:
        return jsonify({"error": "文件名为空"}), 400

    ext = Path(filename).suffix.lower()
    if ext not in (".xls", ".xlsx", ".csv"):
        return jsonify({"error": "仅支持 .xls / .xlsx / .csv 格式"}), 400

    save_path = UPLOAD_DIR / filename
    f.save(str(save_path))

    try:
        headers, rows = parse_excel(save_path)
    except Exception as e:
        return jsonify({"error": f"解析失败: {e}"}), 400

    return jsonify({
        "filename": filename,
        "headers": headers,
        "preview": rows[:10],
        "totalRows": len(rows),
    })


@app.route("/api/parse-all", methods=["POST"])
def parse_all_rows():
    """返回 Excel 全部数据行"""
    data = request.json
    filename = data.get("filename")
    if not filename:
        return jsonify({"error": "缺少 filename"}), 400

    save_path = UPLOAD_DIR / filename
    if not save_path.exists():
        return jsonify({"error": "文件不存在，请重新上传"}), 404

    try:
        headers, rows = parse_excel(save_path)
    except Exception as e:
        return jsonify({"error": f"解析失败: {e}"}), 400

    return jsonify({"headers": headers, "rows": rows})


# ── 附件管理 ──────────────────────────────────────────────────────────────────

@app.route("/api/upload-attachments", methods=["POST"])
def upload_attachments():
    """批量上传附件文件"""
    if "files" not in request.files:
        return jsonify({"error": "未上传文件"}), 400

    files = request.files.getlist("files")
    saved = []
    for f in files:
        if f.filename:
            save_path = ATTACH_DIR / f.filename
            f.save(str(save_path))
            saved.append(f.filename)

    return jsonify({
        "uploaded": len(saved),
        "files": saved,
        "all_files": _list_attachments(),
    })


@app.route("/api/attachments")
def list_attachments():
    """列出已上传的所有附件"""
    return jsonify({"files": _list_attachments()})


@app.route("/api/attachments/<filename>", methods=["DELETE"])
def delete_attachment(filename):
    """删除单个附件"""
    fpath = ATTACH_DIR / filename
    if fpath.exists():
        fpath.unlink()
    return jsonify({"success": True, "files": _list_attachments()})


@app.route("/api/clear-attachments", methods=["POST"])
def clear_attachments():
    """清空所有附件"""
    for f in ATTACH_DIR.iterdir():
        if f.is_file() and not f.name.startswith("."):
            f.unlink()
    return jsonify({"success": True, "files": []})


def _list_attachments():
    """返回附件目录中的文件列表"""
    files = []
    for f in sorted(ATTACH_DIR.iterdir()):
        if f.is_file() and not f.name.startswith("."):
            size = f.stat().st_size
            if size < 1024:
                size_str = f"{size} B"
            elif size < 1024 * 1024:
                size_str = f"{size / 1024:.1f} KB"
            else:
                size_str = f"{size / (1024 * 1024):.1f} MB"
            files.append({"name": f.name, "size": size_str, "size_bytes": size})
    return files


# ── 发送邮件 ──────────────────────────────────────────────────────────────────

@app.route("/api/send", methods=["POST"])
def send_emails():
    """
    批量发送邮件
    请求体:
    {
        "smtp_host": "smtp.exmail.qq.com",
        "smtp_port": 465,
        "email": "xxx@csdn.net",
        "password": "xxx",
        "subject": "邮件主题，支持 {{姓名}} 变量",
        "body_html": "<p>正文 HTML，支持 {{姓名}} 变量</p>",
        "recipients": [
            {"邮箱": "a@b.com", "姓名": "张三", ...},
            ...
        ],
        "email_col": "邮箱",
        "delay": 2
    }
    """
    data = request.json
    smtp_host = data.get("smtp_host") or os.environ.get("SMTP_HOST", "smtp.exmail.qq.com")
    smtp_port = int(data.get("smtp_port") or os.environ.get("SMTP_PORT", 465))
    sender_email = data.get("email") or os.environ.get("SENDER_EMAIL", "")
    sender_password = data.get("password") or os.environ.get("SENDER_PASSWORD", "")
    subject_tpl = data.get("subject", "")
    body_tpl = data.get("body_html", "")
    recipients = data.get("recipients", [])
    email_col = data.get("email_col", "邮箱")
    delay = float(data.get("delay", 2))
    sender_name = data.get("sender_name") or os.environ.get("SENDER_NAME", "CSDN")
    attach_col = data.get("attach_col", "")  # 附件文件名列（可选）

    if not sender_email or not sender_password:
        return jsonify({"error": "请填写发件人邮箱和密码"}), 400
    if not recipients:
        return jsonify({"error": "收件人列表为空"}), 400
    if not subject_tpl:
        return jsonify({"error": "请填写邮件主题"}), 400

    task_id = str(uuid.uuid4())[:8]
    send_tasks[task_id] = {
        "total": len(recipients),
        "sent": 0,
        "failed": 0,
        "results": [],
        "status": "running",
    }

    # 在后台线程中发送
    import threading
    t = threading.Thread(
        target=_send_worker,
        args=(task_id, smtp_host, smtp_port, sender_email, sender_password,
              sender_name, subject_tpl, body_tpl, recipients, email_col,
              attach_col, delay),
        daemon=True,
    )
    t.start()

    return jsonify({"task_id": task_id, "total": len(recipients)})


def _send_worker(task_id, smtp_host, smtp_port, sender_email, sender_password,
                 sender_name, subject_tpl, body_tpl, recipients, email_col,
                 attach_col, delay):
    """后台发送线程"""
    task = send_tasks[task_id]
    server = None

    try:
        # 连接 SMTP
        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30)
        else:
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=30)
            server.starttls()

        server.login(sender_email, sender_password)

        for i, recipient in enumerate(recipients):
            to_email = recipient.get(email_col, "").strip()
            if not to_email or "@" not in to_email:
                task["results"].append({
                    "index": i,
                    "email": to_email,
                    "status": "skipped",
                    "message": "邮箱地址无效",
                })
                task["failed"] += 1
                continue

            # 替换模板变量
            subject = _render_template(subject_tpl, recipient)
            body = _render_template(body_tpl, recipient)

            # 查找附件文件
            attach_files = []
            if attach_col:
                raw_names = recipient.get(attach_col, "").strip()
                if raw_names:
                    # 支持多个附件，用分号或逗号分隔
                    for name in raw_names.replace("；", ";").replace("，", ",").replace(";", ",").split(","):
                        name = name.strip()
                        if not name:
                            continue
                        fpath = ATTACH_DIR / name
                        if fpath.exists() and fpath.is_file():
                            attach_files.append(fpath)
                        else:
                            task["results"].append({
                                "index": i,
                                "email": to_email,
                                "status": "failed",
                                "message": f"附件未找到: {name}",
                            })
                            task["failed"] += 1
                            break
                    else:
                        # 所有附件都找到了，继续发送
                        pass

                    # 如果上面 break 了（有附件没找到），跳过这个收件人
                    if any(r["index"] == i and r["status"] == "failed" for r in task["results"]):
                        continue

            try:
                msg = MIMEMultipart("mixed")
                msg["From"] = f"{sender_name} <{sender_email}>"
                msg["To"] = to_email
                msg["Subject"] = Header(subject, "utf-8")

                # HTML 正文
                html_part = MIMEText(body, "html", "utf-8")
                msg.attach(html_part)

                # 添加附件
                for fpath in attach_files:
                    ctype, encoding = mimetypes.guess_type(str(fpath))
                    if ctype is None:
                        ctype = "application/octet-stream"
                    maintype, subtype = ctype.split("/", 1)

                    with open(fpath, "rb") as af:
                        part = MIMEBase(maintype, subtype)
                        part.set_payload(af.read())
                    encoders.encode_base64(part)
                    # 用 RFC 2231 编码中文文件名
                    part.add_header(
                        "Content-Disposition", "attachment",
                        filename=("utf-8", "", fpath.name),
                    )
                    msg.attach(part)

                server.sendmail(sender_email, [to_email], msg.as_string())

                attach_info = f"（附件: {', '.join(f.name for f in attach_files)}）" if attach_files else ""
                task["results"].append({
                    "index": i,
                    "email": to_email,
                    "status": "success",
                    "message": f"发送成功{attach_info}",
                })
                task["sent"] += 1

            except Exception as e:
                task["results"].append({
                    "index": i,
                    "email": to_email,
                    "status": "failed",
                    "message": str(e),
                })
                task["failed"] += 1

            # 发送间隔，避免被限流
            if i < len(recipients) - 1:
                time.sleep(delay)

    except Exception as e:
        task["results"].append({
            "index": -1,
            "email": "",
            "status": "error",
            "message": f"SMTP 连接/登录失败: {e}\n{traceback.format_exc()}",
        })
    finally:
        if server:
            try:
                server.quit()
            except Exception:
                pass
        task["status"] = "done"


@app.route("/api/send-status/<task_id>")
def send_status(task_id):
    """查询发送进度"""
    task = send_tasks.get(task_id)
    if not task:
        return jsonify({"error": "任务不存在"}), 404
    return jsonify(task)


@app.route("/api/test-smtp", methods=["POST"])
def test_smtp():
    """测试 SMTP 连接"""
    data = request.json
    smtp_host = data.get("smtp_host") or os.environ.get("SMTP_HOST", "smtp.exmail.qq.com")
    smtp_port = int(data.get("smtp_port") or os.environ.get("SMTP_PORT", 465))
    email = data.get("email") or os.environ.get("SENDER_EMAIL", "")
    password = data.get("password") or os.environ.get("SENDER_PASSWORD", "")

    if not email or not password:
        return jsonify({"error": "请填写邮箱和密码"}), 400

    try:
        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=15)
        else:
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=15)
            server.starttls()

        server.login(email, password)
        server.quit()
        return jsonify({"success": True, "message": "SMTP 连接成功！"})
    except Exception as e:
        return jsonify({"success": False, "message": f"连接失败: {e}"})


# ── 工具函数 ──────────────────────────────────────────────────────────────────

def parse_excel(filepath: Path):
    """解析 xls / xlsx / csv 文件，返回 (headers, rows)"""
    ext = filepath.suffix.lower()

    if ext == ".xls":
        # 先尝试用 xlrd 解析真正的 xls 格式
        try:
            wb = xlrd.open_workbook(str(filepath))
            ws = wb.sheets()[0]
            headers = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]
            rows = []
            for r in range(1, ws.nrows):
                row = {}
                for c in range(ws.ncols):
                    val = ws.cell_value(r, c)
                    if isinstance(val, float) and val == int(val):
                        val = int(val)
                    row[headers[c]] = str(val) if val else ""
                rows.append(row)
            return headers, rows
        except Exception:
            # 文件后缀是 .xls 但实际是 CSV/纯文本，回退到 CSV 解析
            return _parse_csv(filepath)

    elif ext == ".xlsx":
        wb = openpyxl.load_workbook(str(filepath), read_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))
        if not all_rows:
            return [], []
        headers = [str(h).strip() if h else f"列{i}" for i, h in enumerate(all_rows[0])]
        rows = []
        for r in all_rows[1:]:
            row = {}
            for c, val in enumerate(r):
                if c < len(headers):
                    if isinstance(val, float) and val == int(val):
                        val = int(val)
                    row[headers[c]] = str(val) if val is not None else ""
            rows.append(row)
        wb.close()
        return headers, rows

    elif ext == ".csv":
        return _parse_csv(filepath)

    else:
        raise ValueError(f"不支持的格式: {ext}")


def _parse_csv(filepath: Path):
    """解析 CSV 或纯文本格式的文件"""
    import csv

    # 尝试检测分隔符
    with open(filepath, "r", encoding="utf-8-sig") as f:
        sample = f.read(4096)

    # 判断分隔符：制表符、逗号、或纯换行（单列）
    if "\t" in sample:
        delimiter = "\t"
    elif "," in sample:
        delimiter = ","
    else:
        delimiter = ","  # 单列数据用逗号也能正常解析

    with open(filepath, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=delimiter)
        all_lines = [row for row in reader if any(cell.strip() for cell in row)]

    if not all_lines:
        return [], []

    headers = [h.strip() if h.strip() else f"列{i}" for i, h in enumerate(all_lines[0])]
    rows = []
    for line in all_lines[1:]:
        row = {}
        for c, val in enumerate(line):
            if c < len(headers):
                row[headers[c]] = val.strip()
        if any(row.values()):
            rows.append(row)
    return headers, rows


def _render_template(template: str, data: dict) -> str:
    """将 {{列名}} 替换为对应值"""
    result = template
    for key, value in data.items():
        result = result.replace("{{" + key + "}}", str(value))
    return result


# ── 启动 ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 50)
    print("  批量邮件发送工具")
    print("  打开浏览器访问: http://localhost:5050")
    print("=" * 50)
    app.run(host="0.0.0.0", port=5050, debug=True)
