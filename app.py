import os
import re
import json
import secrets
import time
import logging
import sqlite3
import socket
import ipaddress
import subprocess
import platform
import urllib.request
import urllib.error
from datetime import datetime
from urllib.parse import urlparse
from flask import Flask, render_template, request, redirect, session, abort, url_for

from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = os.urandom(24).hex()

# ── 安全配置 ──────────────────────────────────────────────
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "uploads")

app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=False,
    PERMANENT_SESSION_LIFETIME=1800,
    SESSION_REFRESH_EACH_REQUEST=True,
    MAX_CONTENT_LENGTH=16 * 1024 * 1024,
)

# ── 审计日志 ──────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [AUDIT] %(message)s",
    handlers=[logging.StreamHandler()],
)
audit_logger = logging.getLogger("audit")

# ── 用户数据（内存 + SQLite 双源）─────────────────────────
USERS = {
    "admin": {
        "username": "admin",
        "password": generate_password_hash("admin123"),
        "role": "admin",
        "email": "admin@example.com",
        "phone": "13800138000",
        "balance": 99999,
    },
    "alice": {
        "username": "alice",
        "password": generate_password_hash("alice2025"),
        "role": "user",
        "email": "alice@example.com",
        "phone": "13900139001",
        "balance": 100,
    },
}

# ── 暴力破解防护 ──────────────────────────────────────────
LOGIN_ATTEMPTS: dict[str, dict] = {}
LOGIN_ATTEMPTS_MAX = 10000
MAX_ATTEMPTS = 5
LOCKOUT_TIME = 300

_DUMMY_HASH = generate_password_hash("__dummy_timing_attack_defense__")

# ── SQLite 数据库 ──────────────────────────────────────────
DB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
DB_PATH = os.path.join(DB_DIR, "users.db")


def init_db():
    """初始化 SQLite 数据库，密码以 bcrypt 哈希存储"""
    os.makedirs(DB_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        email TEXT,
        phone TEXT,
        balance INTEGER DEFAULT 0
    )''')
    # [兼容] 为旧数据库补充 balance 列
    try:
        c.execute("ALTER TABLE users ADD COLUMN balance INTEGER DEFAULT 0")
    except sqlite3.OperationalError:
        pass  # 列已存在，忽略
    # [修复] 密码存储为 bcrypt 哈希，而非明文
    admin_hash = generate_password_hash("admin123")
    alice_hash = generate_password_hash("alice2025")
    c.execute("INSERT OR IGNORE INTO users (username, password, email, phone) VALUES (?, ?, ?, ?)",
              ("admin", admin_hash, "admin@example.com", "13800138000"))
    c.execute("INSERT OR IGNORE INTO users (username, password, email, phone) VALUES (?, ?, ?, ?)",
              ("alice", alice_hash, "alice@example.com", "13900139001"))
    conn.commit()
    conn.close()
    print(f"[DB] 数据库已初始化: {DB_PATH}")


def get_user_from_db(username):
    """从 SQLite 查询用户"""
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT username, password, email, phone, balance FROM users WHERE username = ?", (username,))
        row = c.fetchone()
        if row:
            result = {
                "username": row[0],
                "password": row[1],
                "email": row[2] or "",
                "phone": row[3] or "",
                "role": "user",
                "balance": row[4] or 0,
            }
            return result
        return None
    finally:
        if conn:
            conn.close()


def sanitize_user_info(user_dict: dict) -> dict:
    safe = dict(user_dict)
    safe.pop("password", None)
    return safe


def generate_csrf_token() -> str:
    if "_csrf_token" not in session:
        session["_csrf_token"] = secrets.token_hex(32)
    return session["_csrf_token"]


app.jinja_env.globals["csrf_token"] = generate_csrf_token


@app.context_processor
def inject_now():
    return {"now": datetime.now}


def login_required(f):
    from functools import wraps
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("username"):
            abort(401)
        return f(*args, **kwargs)
    return wrapper


def _prune_login_attempts() -> None:
    if len(LOGIN_ATTEMPTS) > LOGIN_ATTEMPTS_MAX:
        sorted_keys = sorted(
            LOGIN_ATTEMPTS,
            key=lambda k: LOGIN_ATTEMPTS[k]["first_fail"],
        )
        for key in sorted_keys[: len(sorted_keys) // 2]:
            del LOGIN_ATTEMPTS[key]


# ── HTTP 安全响应头 ──────────────────────────────────────
@app.after_request
def add_security_headers(response):
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; "
        "style-src 'self' 'unsafe-inline'; "
        "script-src 'self'; "
        "frame-ancestors 'none'; "
        "form-action 'self'; "
        "base-uri 'self';"
    )
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Permissions-Policy"] = (
        "camera=(), microphone=(), geolocation=(), payment=()"
    )
    return response


# ── 自定义错误页面 ──────────────────────────────────────
@app.errorhandler(401)
def unauthorized(e):
    return render_template("login.html", error="请先登录后再访问此页面"), 401


@app.errorhandler(403)
def forbidden(e):
    return render_template("base.html", error_code=403, error_msg="您没有权限访问此页面"), 403


@app.errorhandler(404)
def not_found(e):
    return render_template("base.html", error_code=404, error_msg="页面未找到"), 404


@app.errorhandler(405)
def method_not_allowed(e):
    return render_template("base.html", error_code=405, error_msg="请求方法不允许"), 405


@app.errorhandler(500)
def server_error(e):
    return render_template("base.html", error_code=500, error_msg="服务器内部错误"), 500


# ── 注册频率限制（防批量注册）────────────────────────────
REGISTER_COOLDOWN = {}  # username -> timestamp
REGISTER_COOLDOWN_SECONDS = 5


def check_register_cooldown(username):
    """检查同一用户名是否注册过于频繁"""
    if username in REGISTER_COOLDOWN:
        if time.time() - REGISTER_COOLDOWN[username] < REGISTER_COOLDOWN_SECONDS:
            return False
    REGISTER_COOLDOWN[username] = time.time()
    return True


# ── 路由 ──────────────────────────────────────────────────
@app.route("/")
def index():
    username = session.get("username")
    user_info = None
    if username and username in USERS:
        user_info = sanitize_user_info(USERS[username])
    return render_template("index.html", user=user_info)


@app.route("/login", methods=["GET", "POST"])
def login():
    client_ip = request.remote_addr or "unknown"

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        token = request.form.get("csrf_token", "")

        if not token or token != session.get("_csrf_token"):
            audit_logger.warning("CSRF校验失败 IP=%s", client_ip)
            return render_template("login.html", error="会话已过期，请刷新页面重试"), 400

        if len(username) > 50 or len(password) > 128:
            return render_template("login.html", error="输入内容不合法")

        if username in LOGIN_ATTEMPTS:
            rec = LOGIN_ATTEMPTS[username]
            if rec["count"] >= MAX_ATTEMPTS:
                if time.time() - rec["first_fail"] < LOCKOUT_TIME:
                    audit_logger.warning("账号锁定 username=%s IP=%s", username, client_ip)
                    return render_template("login.html", error="该账号登录失败次数过多，已被临时锁定，请 5 分钟后再试")
                else:
                    del LOGIN_ATTEMPTS[username]

        # [修复] 登录同时检查 USERS 字典和 SQLite 数据库
        user_record = USERS.get(username)
        pwd_hash = user_record["password"] if user_record else None

        # 如果在内存字典中没找到，去 SQLite 查
        if not user_record:
            db_user = get_user_from_db(username)
            if db_user:
                user_record = db_user
                pwd_hash = db_user["password"]

        # 如果都没找到，用虚拟哈希（防计时攻击）
        if not pwd_hash:
            pwd_hash = _DUMMY_HASH

        if pwd_hash and check_password_hash(pwd_hash, password) and user_record:
            session.clear()
            session["username"] = username
            session["_csrf_token"] = secrets.token_hex(32)
            session["_last_active"] = time.time()
            LOGIN_ATTEMPTS.pop(username, None)
            role = user_record.get("role", "user")
            audit_logger.info("登录成功 username=%s IP=%s role=%s", username, client_ip, role)
            return render_template("index.html", user=sanitize_user_info(user_record))
        else:
            now = time.time()
            if username in LOGIN_ATTEMPTS:
                LOGIN_ATTEMPTS[username]["count"] += 1
            else:
                LOGIN_ATTEMPTS[username] = {"count": 1, "first_fail": now}
            _prune_login_attempts()
            audit_logger.warning("登录失败 username=%s IP=%s", username, client_ip)

            remaining = MAX_ATTEMPTS - LOGIN_ATTEMPTS[username]["count"]
            if remaining <= 0:
                return render_template("login.html", error="该账号登录失败次数过多，已被临时锁定，请 5 分钟后再试")
            return render_template("login.html", error=f"用户名或密码错误，还剩 {remaining} 次尝试机会")

    msg = None
    if request.args.get("registered") == "1":
        msg = "注册成功，请登录"
    return render_template("login.html", success=msg)


@app.route("/logout", methods=["POST"])
def logout():
    username = session.get("username", "unknown")
    token = request.form.get("csrf_token", "")
    if not token or token != session.get("_csrf_token"):
        audit_logger.info("登出(CSRF跳过) username=%s", username)
    else:
        audit_logger.info("登出成功 username=%s", username)
    session.clear()
    return redirect("/")


# ── [修复] 注册 ──────────────────────────────────────────
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        email = request.form.get("email", "").strip()
        phone = request.form.get("phone", "").strip()
        token = request.form.get("csrf_token", "")

        # [修复] 添加 CSRF 校验
        if not token or token != session.get("_csrf_token"):
            return render_template("register.html", error="会话已过期，请刷新页面重试"), 400

        # [修复] 注册频率限制
        if not check_register_cooldown(username):
            return render_template("register.html", error="操作过于频繁，请稍后重试"), 429

        # [修复] 输入校验
        if not username or len(username) > 50:
            return render_template("register.html", error="用户名不合法（1-50字符）")
        if not password or len(password) > 128:
            return render_template("register.html", error="密码不合法")
        if email and len(email) > 100:
            return render_template("register.html", error="邮箱不合法")
        if phone and len(phone) > 20:
            return render_template("register.html", error="手机号不合法")

        # [修复] 使用参数化查询防止 SQL 注入
        # [修复] 密码存储为 bcrypt 哈希
        password_hash = generate_password_hash(password)

        try:
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute(
                "INSERT INTO users (username, password, email, phone) VALUES (?, ?, ?, ?)",
                (username, password_hash, email, phone)
            )
            conn.commit()
            conn.close()
            audit_logger.info("注册成功 username=%s", username)
            return redirect(url_for("login", registered="1"))
        except sqlite3.IntegrityError:
            return render_template("register.html", error="用户名已存在")
        except Exception as e:
            audit_logger.error("注册异常 username=%s error=%s", username, str(e))
            return render_template("register.html", error="注册失败，请稍后重试")

    return render_template("register.html")


# ── [修复] 搜索 ──────────────────────────────────────────
@app.route("/search")
@login_required
def search():
    keyword = request.args.get("keyword", "").strip()
    results = []

    if keyword:
        # [修复] 转义 LIKE 通配符，防止 % 和 _ 匹配全部数据
        safe_keyword = keyword.replace("%", "\\%").replace("_", "\\_")
        try:
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            like_pattern = f"%{safe_keyword}%"
            c.execute(
                "SELECT id, username, email, phone FROM users WHERE username LIKE ? OR email LIKE ? ESCAPE '\\'",
                (like_pattern, like_pattern)
            )
            results = c.fetchall()
            conn.close()
        except Exception as e:
            audit_logger.error("搜索异常 keyword=%s error=%s", keyword, str(e))

    username = session.get("username")
    user_info = None
    if username and username in USERS:
        user_info = sanitize_user_info(USERS[username])

    return render_template("index.html", user=user_info, search_results=results, search_keyword=keyword)


# ── 上传防护配置 ──────────────────────────────────────────
ALLOWED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"}


def safe_filename(filename: str) -> str:
    """[安全] 过滤路径穿越字符，生成安全文件名"""
    # 移除路径分隔符（防路径穿越）
    filename = filename.replace("\\", "/").split("/")[-1]
    # 清理特殊字符
    safe = ""
    for ch in filename:
        if ch.isalnum() or ch in "._- ":
            safe += ch
        else:
            safe += "_"
    # 防止空文件名
    safe = safe.strip()
    if not safe or safe == ".":
        return None
    return safe


def secure_file_extension(filename: str) -> bool:
    """[安全] 白名单校验文件扩展名"""
    _, ext = os.path.splitext(filename)
    return ext.lower() in ALLOWED_EXTENSIONS


# ── 上传 ──────────────────────────────────────────────────
@app.route("/upload", methods=["GET", "POST"])
@login_required
def upload():
    file_url = None
    error = None

    if request.method == "POST":
        # [安全] CSRF 校验
        token = request.form.get("csrf_token", "")
        if not token or token != session.get("_csrf_token"):
            return render_template("upload.html", error="会话已过期，请刷新页面重试"), 400

        file = request.files.get("file")
        if not file or not file.filename:
            error = "请选择一个文件"
        else:
            original_name = file.filename
            # [安全] 过滤路径穿越
            safe_name = safe_filename(original_name)
            if not safe_name:
                error = "文件名不合法"
            # [安全] 白名单校验扩展名
            elif not secure_file_extension(safe_name):
                error = "仅支持上传图片文件（png/jpg/gif/webp/bmp）"
            else:
                # [安全] 生成唯一文件名防覆盖
                timestamp = int(time.time() * 1000)
                rand_suffix = secrets.token_hex(4)
                name_part, ext_part = os.path.splitext(safe_name)
                unique_name = f"{name_part}_{timestamp}_{rand_suffix}{ext_part}"

                os.makedirs(UPLOAD_DIR, exist_ok=True)
                save_path = os.path.join(UPLOAD_DIR, unique_name)
                file.save(save_path)
                file_url = url_for("static", filename=f"uploads/{unique_name}")
                audit_logger.info("上传文件 username=%s original=%s saved=%s",
                                  session.get("username"), original_name, unique_name)

    username = session.get("username")
    user_info = None
    if username and username in USERS:
        user_info = sanitize_user_info(USERS[username])

    return render_template("upload.html", user=user_info, file_url=file_url, error=error)


# ── 个人中心与充值 ──────────────────────────────────────────
USER_ID_MAP = {"admin": 1, "alice": 2}
ID_USER_MAP = {1: "admin", 2: "alice"}


def get_user_id_by_username(username: str) -> int | None:
    if username in USER_ID_MAP:
        return USER_ID_MAP[username]
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT id FROM users WHERE username = ?", (username,))
        row = c.fetchone()
        return row[0] if row else None
    finally:
        if conn: conn.close()


def get_user_by_id(user_id: int) -> dict | None:
    if user_id in ID_USER_MAP:
        username = ID_USER_MAP[user_id]
        if username in USERS:
            return {"id": user_id, **sanitize_user_info(USERS[username])}
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT id, username, email, phone, balance FROM users WHERE id = ?", (user_id,))
        row = c.fetchone()
        if row:
            return {"id": row[0], "username": row[1], "email": row[2] or "",
                    "phone": row[3] or "", "balance": row[4] or 0}
        return None
    finally:
        if conn: conn.close()


def update_balance(user_id: int, amount: int) -> bool:
    MAX_BALANCE = 999999999
    if amount > MAX_BALANCE:
        amount = MAX_BALANCE
    if user_id in ID_USER_MAP:
        username = ID_USER_MAP[user_id]
        if username in USERS:
            USERS[username]["balance"] = min(USERS[username]["balance"] + amount, MAX_BALANCE)
            return True
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("UPDATE users SET balance = MIN(COALESCE(balance, 0) + ?, ?) WHERE id = ?",
                  (amount, MAX_BALANCE, user_id))
        conn.commit()
        return c.rowcount > 0
    except Exception:
        return False
    finally:
        if conn: conn.close()


@app.route("/profile")
@login_required
def profile():
    current_user = session.get("username")
    current_id = get_user_id_by_username(current_user) if current_user else None
    req_id = request.args.get("user_id", type=int)
    target_id = req_id if req_id else current_id
    if target_id != current_id:
        return render_template("profile.html", user_data=None, error="无权查看其他用户的资料")
    user_data = get_user_by_id(target_id)
    if not user_data:
        return render_template("profile.html", user_data=None, error="未找到用户信息")
    return render_template("profile.html", user_data=user_data)


@app.route("/recharge", methods=["POST"])
@login_required
def recharge():
    current_user = session.get("username")
    current_id = get_user_id_by_username(current_user) if current_user else None
    req_id = request.form.get("user_id", type=int)
    amount = request.form.get("amount", type=int, default=0)
    if req_id != current_id:
        return redirect("/")
    if amount <= 0:
        return redirect(f"/profile?user_id={req_id}")
    if amount > 1000000:
        amount = 1000000
    if req_id and amount > 0:
        update_balance(req_id, amount)
    return redirect(f"/profile?user_id={req_id}")


# ── 动态页面加载 ──────────────────────────────────────────
PAGES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "5-阶段五")
ALLOWED_PAGE_EXT = {".html"}


@app.route("/page")
@login_required
def dynamic_page():
    name = request.args.get("name", "")
    page_content = None
    error = None

    if name:
        # [安全] 防止路径穿越：只取文件名，丢弃路径部分
        safe_name = name.replace("\\", "/").split("/")[-1]

        # [安全] 只允许 .html 扩展名
        _, ext = os.path.splitext(safe_name)
        if ext == "":
            safe_name += ".html"

        _, ext = os.path.splitext(safe_name)
        if ext.lower() not in ALLOWED_PAGE_EXT:
            error = "不支持的页面类型"
        else:
            # [安全] 解析真实路径，确保仍在 pages/ 目录内
            file_path = os.path.join(PAGES_DIR, safe_name)
            real_path = os.path.realpath(file_path)
            real_base = os.path.realpath(PAGES_DIR)

            if not real_path.startswith(real_base + os.sep) and real_path != real_base:
                error = "非法的页面路径"
            elif not os.path.isfile(real_path):
                error = "页面不存在"
            else:
                with open(real_path, "r", encoding="utf-8") as f:
                    page_content = f.read()
    else:
        error = "请提供页面名称"

    username = session.get("username")
    user_info = None
    if username and username in USERS:
        user_info = sanitize_user_info(USERS[username])

    return render_template("index.html", user=user_info, page_content=page_content, page_error=error)


# ── 密码修改 ──────────────────────────────────────────
@app.route("/change-password", methods=["POST"])
@login_required
def change_password():
    current_user = session.get("username")
    old_password = request.form.get("old_password", "")
    new_password = request.form.get("new_password", "")
    confirm_password = request.form.get("confirm_password", "")
    token = request.form.get("csrf_token", "")

    user_info = None
    if current_user:
        user_info = sanitize_user_info(USERS.get(current_user, {}))

    # [修复] CSRF 校验
    if not token or token != session.get("_csrf_token"):
        return render_template("profile.html", user_data=get_user_by_id(get_user_id_by_username(current_user)) if current_user else None, error="会话已过期，请刷新页面重试"), 400

    # [修复] 只能修改自己的密码
    form_username = request.form.get("username", "").strip()
    if form_username != current_user:
        return render_template("profile.html", user_data=get_user_by_id(get_user_id_by_username(current_user)) if current_user else None, error="无权修改其他用户的密码"), 403

    # [修复] 验证原密码
    user_record = USERS.get(current_user)
    if user_record and not check_password_hash(user_record["password"], old_password):
        return render_template("profile.html", user_data=get_user_by_id(get_user_id_by_username(current_user)) if current_user else None, error="原密码错误")

    # [修复] 两次密码一致校验
    if new_password != confirm_password:
        return render_template("profile.html", user_data=get_user_by_id(get_user_id_by_username(current_user)) if current_user else None, error="两次输入的密码不一致")

    # 更新 USERS 字典中的密码
    if current_user in USERS:
        USERS[current_user]["password"] = generate_password_hash(new_password)

    # 更新 SQLite 中的密码
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("UPDATE users SET password = ? WHERE username = ?",
                  (generate_password_hash(new_password), current_user))
        conn.commit()
    except Exception:
        pass
    finally:
        if conn:
            conn.close()

    return redirect("/profile")


# ── URL 抓取 ──────────────────────────────────────────
ALLOWED_PROTOCOLS = {"http", "https"}
FORBIDDEN_IPS = {
    "127.0.0.1", "0.0.0.0", "localhost",
    "169.254.169.254", "metadata.google.internal",
    "100.100.100.200",  # 阿里云
}
FORBIDDEN_NETWORKS = [
    "127.0.0.0/8",
    "10.0.0.0/8",
    "172.16.0.0/12",
    "192.168.0.0/16",
    "169.254.0.0/16",
]


def _is_internal_ip(host: str) -> bool:
    """检查 host 是否为内网地址"""
    # 先检查已知的禁止 IP
    if host.lower() in FORBIDDEN_IPS:
        return True
    # 尝试解析域名
    try:
        ips = set()
        addr = socket.getaddrinfo(host, None)
        for info in addr:
            ip = info[4][0]
            ips.add(ip)
        for ip in ips:
            if ip in FORBIDDEN_IPS:
                return True
            ip_obj = ipaddress.ip_address(ip)
            for net in FORBIDDEN_NETWORKS:
                if ip_obj in ipaddress.ip_network(net):
                    return True
    except Exception:
        # 如果无法解析，保守起见拒绝
        return True
    return False


@app.route("/fetch-url", methods=["POST"])
@login_required
def fetch_url():
    target_url = request.form.get("url", "").strip()
    result = None
    status_code = None
    error = None

    if target_url:
        # [安全] 协议白名单
        parsed = urlparse(target_url)
        if parsed.scheme not in ALLOWED_PROTOCOLS:
            error = "不支持的协议类型，仅允许 http:// 和 https://"
        # [安全] 长度限制
        elif len(target_url) > 2048:
            error = "URL 过长"
        else:
            # [安全] 检查目标地址
            host = parsed.hostname or ""
            if _is_internal_ip(host):
                error = "不允许访问内网地址"
            else:
                try:
                    # [安全] 自定义 opener，不跟随重定向
                    class NoRedirectHandler(urllib.request.HTTPRedirectHandler):
                        def redirect_request(self, req, fp, code, msg, headers, newurl):
                            return None
                    opener = urllib.request.build_opener(NoRedirectHandler)
                    req = urllib.request.Request(target_url)
                    with opener.open(req, timeout=10) as response:
                        status_code = response.status
                        content = response.read().decode("utf-8", errors="replace")
                        result = content[:5000]
                except urllib.error.HTTPError as e:
                    status_code = e.code
                    result = str(e)
                except urllib.error.URLError as e:
                    error = f"URL 访问失败: {e.reason}"
                except Exception as e:
                    error = f"请求异常: {str(e)}"

    username = session.get("username")
    user_info = None
    if username and username in USERS:
        user_info = sanitize_user_info(USERS[username])

    return render_template("index.html", user=user_info,
                           fetch_result=result, fetch_status=status_code, fetch_error=error,
                           fetch_url=target_url)


# ── Ping 网络诊断 ──────────────────────────────────────
import re as _re

ALLOWED_PING_PATTERN = _re.compile(r"^(?:::)?[a-fA-F0-9](?:[a-fA-F0-9:]*[a-fA-F0-9])?$|^[a-zA-Z0-9](?:[a-zA-Z0-9\-\.]*[a-zA-Z0-9])?$")


def _validate_ip_or_domain(target: str) -> bool:
    """校验输入是否为合法的 IP 地址或域名"""
    if not target or len(target) > 255:
        return False
    # 检查是否包含 shell 特殊字符（冒号在 IPv6 中合法）
    forbidden = {";", "|", "&", "$", "`", "(", ")", "{", "}", "<", ">", "!", "#", "~", "%", "\"", "'", " ", "\t", "\n", "\r", "\\"}
    # 如果包含冒号，验证是否为纯 IPv6 地址格式
    if ":" in target:
        if not _re.match(r"^[a-fA-F0-9:]+$", target):
            return False
    elif any(ch in target for ch in forbidden):
        return False
    return bool(ALLOWED_PING_PATTERN.match(target))


@app.route("/ping", methods=["GET", "POST"])
@login_required
def ping():
    result = None
    ip = ""
    error = None

    if request.method == "POST":
        ip = request.form.get("ip", "").strip()
        if not ip:
            error = "请输入 IP 地址或域名"
        elif len(ip) > 255:
            error = "输入过长"
        elif not _validate_ip_or_domain(ip):
            error = "输入包含非法字符，仅允许字母、数字、点、中划线、下划线和冒号"
        else:
            try:
                # [安全] 使用列表参数，禁用 shell=True，杜绝命令注入
                cmd = ["ping", "-c", "3", ip]
                result = subprocess.check_output(cmd, shell=False, timeout=30, stderr=subprocess.STDOUT).decode("utf-8", errors="replace")
            except subprocess.CalledProcessError as e:
                result = e.output.decode("utf-8", errors="replace")
            except subprocess.TimeoutExpired:
                result = "Ping 超时"
            except Exception as e:
                result = f"执行失败: {str(e)}"

    username = session.get("username")
    user_info = None
    if username and username in USERS:
        user_info = sanitize_user_info(USERS[username])

    return render_template("ping.html", user=user_info, ping_result=result, ping_error=error, ping_ip=ip)


# ── XML 数据导入 ──────────────────────────────────────
@app.route("/xml-import", methods=["GET", "POST"])
@login_required
def xml_import():
    result = None
    error = None

    if request.method == "POST":
        xml_data = request.form.get("xml_data", "").strip()
        if xml_data:
            try:
                # [安全] 剥离 DTD 声明（<!DOCTYPE ...> 及内部定义），防止 XXE
                xml_data = re.sub(r'<!DOCTYPE[^]]*]>', "", xml_data, flags=re.DOTALL)
                xml_data = re.sub(r'<!ENTITY[^>]*>', "", xml_data)
                # [安全] 剥离未定义的实体引用（&xxx;），防止 DTD 剥离后残留
                xml_data = re.sub(r'&(?!(?:amp|lt|gt|quot|apos);)\w+;', "", xml_data)
                xml_data_no_dtd = xml_data

                import xml.etree.ElementTree as ET
                root = ET.fromstring(xml_data_no_dtd)
                users = []
                for user_elem in root.findall(".//user"):
                    name = user_elem.get("name", "") or user_elem.findtext("name", "")
                    email = user_elem.findtext("email", "")
                    users.append({"name": name, "email": email})

                if users:
                    result = json.dumps({"users": users, "total": len(users)}, ensure_ascii=False, indent=2)
                else:
                    error = "未找到 user 节点"
            except ET.ParseError as e:
                error = f"XML 解析失败: {str(e)}"
            except Exception as e:
                error = f"处理异常: {str(e)}"

    username = session.get("username")
    user_info = None
    if username and username in USERS:
        user_info = sanitize_user_info(USERS[username])

    return render_template("xml_import.html", user=user_info, xml_result=result, xml_error=error)


# ── 启动 ──────────────────────────────────────────────────
if __name__ == "__main__":
    init_db()
    import werkzeug.serving as _ws
    _ws.WSGIRequestHandler.version_string = lambda self: "WebServer"
    app.run(host="0.0.0.0", port=5000)
