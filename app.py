import os
import re
import sqlite3
import logging
import datetime
import urllib.request
import urllib.error
import urllib.parse
import socket
import subprocess
import platform
from datetime import timedelta
from flask import Flask, render_template, request, redirect, session, url_for
from flask_bcrypt import Bcrypt
from flask_wtf.csrf import CSRFProtect
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from werkzeug.serving import WSGIRequestHandler

# 隐藏 Werkzeug Server 版本信息
WSGIRequestHandler.server_version = ""
WSGIRequestHandler.sys_version = ""

app = Flask(__name__)
app.secret_key = os.environ.get(
    "SECRET_KEY",
    os.urandom(64).hex()
)


# CSRF 保护
csrf = CSRFProtect(app)


@app.route("/change-password", methods=["POST"])
def change_password():
    """修改密码"""
    username = request.form.get("username", "")
    new_password = request.form.get("new_password", "")

    if username and new_password and username in USERS:
        hashed = bcrypt.generate_password_hash(new_password).decode("utf-8")
        USERS[username]["password"] = hashed

    return redirect("/profile")


# Session 安全配置
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_PERMANENT"] = True
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=1)
# 仅在明确启用时设置 Secure 标志（需要 HTTPS）
secure_mode = os.environ.get("SESSION_COOKIE_SECURE", "false").lower() == "true"
if secure_mode:
    app.config["SESSION_COOKIE_SECURE"] = True

bcrypt = Bcrypt(app)

# 文件上传配置
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
UPLOAD_FOLDER = os.path.join(app.static_folder, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 日志配置
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger(__name__)

# 登录频率限制（基于 IP，支持代理转发）
def get_real_ip():
    """优先从 X-Forwarded-For 获取真实 IP（需反向代理配合）"""
    forwarded = request.headers.get("X-Forwarded-For")
    if forwarded:
        return forwarded.split(",")[0].strip()
    return get_remote_address()

limiter = Limiter(
    get_real_ip,
    app=app,
    default_limits=["200 per day", "50 per hour"],
    storage_uri="memory://",
)

# 用户数据库 - 密码为预生成的 bcrypt 哈希，源码中不包含明文密码
USERS = {
    "admin": {
        "id": 1,
        "username": "admin",
        "password": "$2b$12$.cLDo0EipcOk.1F6EedqdebiMwwCRZHAz84sMRu/6u5qwjE3v9k4e",
        "role": "admin",
        "email": "admin@example.com",
        "phone": "13800138000",
        "balance": 99999
    },
    "alice": {
        "id": 2,
        "username": "alice",
        "password": "$2b$12$pIqdI5oYxQD/XagaMngLqON5PzAlPqAaTO2WB/w4wyAS/GAZhgqVm",
        "role": "user",
        "email": "alice@example.com",
        "phone": "13900139001",
        "balance": 100
    }
}


def _get_current_user_id(username):
    """根据用户名获取对应的用户 ID"""
    user = USERS.get(username)
    return user["id"] if user else None

# 防时序攻击：启动时生成随机 dummy hash
_dummy_hash = bcrypt.generate_password_hash(os.urandom(32).hex()).decode("utf-8")

# 账号锁定机制：连续失败次数 → 锁定时间
_LOGIN_ATTEMPTS = {}  # {username: {"count": int, "locked_until": datetime}}
MAX_LOGIN_ATTEMPTS = 5
LOCKOUT_DURATION = timedelta(minutes=15)


def _check_account_locked(username):
    """检查账号是否被临时锁定"""
    record = _LOGIN_ATTEMPTS.get(username)
    if not record or not record.get("locked_until"):
        return False, 0
    if record["locked_until"] > datetime.datetime.now():
        remaining = int((record["locked_until"] - datetime.datetime.now()).total_seconds())
        logger.warning("账号 %s 已被锁定（剩余 %d 秒）", username, remaining)
        return True, remaining
    return False, 0


def _record_failed_attempt(username):
    """记录登录失败次数，达到阈值则锁定"""
    now = datetime.datetime.now()
    record = _LOGIN_ATTEMPTS.get(username)
    if record:
        # 如果锁定已过期，重置计数
        if record.get("locked_until") and record["locked_until"] < now:
            record["count"] = 1
            record["locked_until"] = None
        else:
            record["count"] += 1
    else:
        record = {"count": 1, "locked_until": None}
        _LOGIN_ATTEMPTS[username] = record

    if record["count"] >= MAX_LOGIN_ATTEMPTS:
        record["locked_until"] = now + LOCKOUT_DURATION
        logger.warning(
            "账号 %s 因连续 %d 次登录失败已被锁定 %d 分钟",
            username, MAX_LOGIN_ATTEMPTS, int(LOCKOUT_DURATION.total_seconds() / 60)
        )
    return record["count"], MAX_LOGIN_ATTEMPTS - record["count"]


def _reset_login_attempts(username):
    """登录成功后重置失败计数"""
    _LOGIN_ATTEMPTS.pop(username, None)


def _check_password_strength(password):
    """校验密码强度"""
    if len(password) < 8:
        return False, "密码长度至少 8 位"
    if not re.search(r"[A-Z]", password):
        return False, "密码需包含至少一个大写字母"
    if not re.search(r"[a-z]", password):
        return False, "密码需包含至少一个小写字母"
    if not re.search(r"\d", password):
        return False, "密码需包含至少一个数字"
    if not re.search(r"[!@#$%^&*(),.?\":{}|<>_\-+=]", password):
        return False, "密码需包含至少一个特殊字符"
    return True, "密码强度合格"


def _safe_user(username):
    """返回不包含密码字段的用户信息，防止哈希泄露到模板中"""
    user = USERS.get(username)
    if user is None:
        return None
    return {k: v for k, v in user.items() if k != "password"}


# 所有响应添加安全头
@app.after_request
def add_security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["X-XSS-Protection"] = "0"  # 已废弃但兼容旧浏览器
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; "
        "style-src 'self' 'unsafe-inline'; "
        "script-src 'self'; "
        "img-src 'self' data:; "
        "font-src 'self'; "
        "form-action 'self'; "
        "base-uri 'self'; "
        "frame-ancestors 'none'"
    )
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Permissions-Policy"] = (
        "camera=(), microphone=(), geolocation=(), interest-cohort=()"
    )
    # 隐藏后端技术信息
    response.headers.pop("Server", None)
    # HTTPS 模式下启用 HSTS
    if secure_mode:
        response.headers["Strict-Transport-Security"] = "max-age=63072000; includeSubDomains"
    return response


# ===== SQLite 数据库初始化 =====

def init_db():
    """初始化 SQLite 数据库，创建 users 表并插入默认用户"""
    os.makedirs("data", exist_ok=True)
    conn = sqlite3.connect("data/users.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            email TEXT,
            phone TEXT
        )
    """)
    # 插入默认用户，使用 INSERT OR IGNORE 防止重复
    c.execute("INSERT OR IGNORE INTO users (username, password, email, phone) VALUES ('admin', 'admin123', 'admin@example.com', '13800138000')")
    c.execute("INSERT OR IGNORE INTO users (username, password, email, phone) VALUES ('alice', 'alice2025', 'alice@example.com', '13900139001')")
    conn.commit()
    conn.close()
    logger.info("数据库初始化完成: data/users.db")


@app.route("/health")
def health():
    """健康检查端点"""
    return {"status": "ok"}


@app.route("/register", methods=["GET", "POST"])
def register():
    """用户注册"""
    message = None
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        email = request.form.get("email", "")
        phone = request.form.get("phone", "")

        conn = sqlite3.connect("data/users.db")
        c = conn.cursor()
        # 使用参数化查询，防止 SQL 注入
        sql = "INSERT INTO users (username, password, email, phone) VALUES (?, ?, ?, ?)"
        logger.info("执行 SQL: %s", sql)
        try:
            c.execute(sql, (username, password, email, phone))
            conn.commit()
            message = "注册成功，请登录"
        except Exception as e:
            message = f"注册失败: {e}"
        finally:
            conn.close()

        if message == "注册成功，请登录":
            return render_template("login.html", success=message)

    return render_template("register.html", message=message)


@app.route("/search")
def search():
    """搜索用户"""
    keyword = request.args.get("keyword", "")
    results = []
    if keyword:
        conn = sqlite3.connect("data/users.db")
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        # 使用参数化查询，防止 SQL 注入
        sql = "SELECT * FROM users WHERE username LIKE ? OR email LIKE ?"
        like_param = f"%{keyword}%"
        logger.info("执行 SQL: %s (参数: %s)", sql, like_param)
        try:
            c.execute(sql, (like_param, like_param))
            rows = c.fetchall()
            results = [{"id": r["id"], "username": r["username"], "email": r["email"], "phone": r["phone"]} for r in rows]
        except Exception as e:
            logger.error("搜索出错: %s", e)
        finally:
            conn.close()

    username = session.get("username")
    user_info = _safe_user(username)
    return render_template("index.html", user=user_info, results=results, keyword=keyword)


@app.route("/")
def index():
    username = session.get("username")
    user_info = _safe_user(username)
    return render_template("index.html", user=user_info, results=None, keyword="")


@app.route("/upload", methods=["GET", "POST"])
def upload():
    """用户头像上传"""
    username = session.get("username")
    if not username:
        return redirect("/login")

    message = None
    file_url = None
    filename = None

    if request.method == "POST":
        if "file" not in request.files:
            message = "没有选择文件"
        else:
            f = request.files["file"]
            if f.filename == "":
                message = "没有选择文件"
            else:
                # 检查文件扩展名
                allowed_extensions = {"png", "jpg", "jpeg", "gif", "webp"}
                ext = f.filename.rsplit(".", 1)[-1].lower() if "." in f.filename else ""
                if ext not in allowed_extensions:
                    message = "只允许上传图片文件（png、jpg、jpeg、gif、webp）"
                elif f.content_type not in ["image/png", "image/jpeg", "image/gif", "image/webp"]:
                    message = f"不支持的文件类型: {f.content_type}"
                else:
                    # 清理文件名：移除路径分隔符等不安全字符
                    import secrets
                    safe_name = re.sub(r"[^a-zA-Z0-9_.\-]", "", f.filename)
                    safe_name = safe_name if safe_name else f"avatar_{secrets.token_hex(4)}.{ext}"
                    # 防止重名覆盖
                    save_path = os.path.join(UPLOAD_FOLDER, safe_name)
                    counter = 1
                    while os.path.exists(save_path):
                        name_part = safe_name.rsplit(".", 1)[0]
                        safe_name = f"{name_part}_{counter}.{ext}"
                        save_path = os.path.join(UPLOAD_FOLDER, safe_name)
                        counter += 1
                    f.save(save_path)
                    file_url = url_for("static", filename=f"uploads/{safe_name}")
                    session["avatar"] = safe_name
                    message = "上传成功"
                    logger.info("用户 %s 上传文件: %s (类型: %s)", username, safe_name, f.content_type)

    return render_template("upload.html", message=message, file_url=file_url, filename=filename)


@app.route("/login", methods=["GET", "POST"])
@limiter.limit("10 per minute")
def login():
    error = None

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        # 输入验证
        if not username or not password:
            error = "用户名和密码不能为空"
        elif not re.match(r"^[a-zA-Z0-9_]{3,32}$", username):
            error = "用户名只能包含字母、数字和下划线（3-32位）"
        else:
            # 检查账号是否被锁定
            locked, remaining = _check_account_locked(username)
            if locked:
                logger.warning(
                    "被锁账号登录尝试: %s, IP: %s",
                    username, get_real_ip()
                )
                error = f"账号已被锁定，请 {remaining} 秒后再试"
                return render_template("login.html", error=error)

            # 防时序攻击：无论用户名是否存在都执行 bcrypt 验证
            real_user = USERS.get(username)
            check_hash = real_user["password"] if real_user else _dummy_hash
            password_valid = bcrypt.check_password_hash(check_hash, password)

            if real_user and password_valid:
                # 登录成功：更新 session，保留 CSRF token
                _reset_login_attempts(username)
                session.permanent = True
                session["username"] = username
                logger.info(
                    "用户 %s 登录成功, IP: %s",
                    username, get_real_ip()
                )
                return redirect("/")
            else:
                # 记录失败尝试
                attempt_count, remaining_attempts = _record_failed_attempt(username)
                client_ip = get_real_ip()
                logger.warning(
                    "登录失败: 用户名=%s, IP=%s, 连续失败=%d次",
                    username, client_ip, attempt_count
                )
                if remaining_attempts > 0:
                    error = f"用户名或密码错误，请重试（还剩 {remaining_attempts} 次机会）"
                else:
                    error = f"账号已被锁定，请 {int(LOCKOUT_DURATION.total_seconds() / 60)} 分钟后再试"

    return render_template("login.html", error=error)


@app.route("/logout", methods=["POST"])
def logout():
    username = session.get("username")
    session.clear()
    if username:
        logger.info("用户 %s 已登出, IP: %s", username, get_real_ip())
    return redirect("/")


def _get_username_by_id(user_id):
    """根据 user_id 从 SQLite 查询对应的用户名"""
    try:
        conn = sqlite3.connect("data/users.db")
        c = conn.cursor()
        c.execute("SELECT username FROM users WHERE id = ?", (user_id,))
        row = c.fetchone()
        conn.close()
        if row:
            return row[0]
    except Exception:
        pass
    return None


@app.route("/profile")
def profile():
    """个人中心 - 仅显示当前登录用户的资料"""
    username = session.get("username")
    if not username:
        return redirect("/login")

    user = USERS.get(username)
    if not user:
        return redirect("/login")

    user_data = {
        "id": user["id"],
        "username": user["username"],
        "email": user["email"],
        "phone": user["phone"],
        "balance": user["balance"]
    }

    error = request.args.get("error")
    return render_template("profile.html", user=user_data, error=error)


@app.route("/recharge", methods=["POST"])
def recharge():
    """充值 - 需要登录，仅限本人，校验金额为正数"""
    username = session.get("username")
    if not username:
        return redirect("/login")

    user = USERS.get(username)
    if not user:
        return redirect("/login")

    amount = request.form.get("amount", type=float, default=0)

    # 校验金额必须为正数
    if amount <= 0:
        return redirect("/profile?error=invalid_amount")

    # 限制单次充值上限
    if amount > 100000:
        return redirect("/profile?error=amount_too_large")

    user["balance"] = user["balance"] + amount
    logger.info("用户 %s 充值 %.2f 元，余额: %.2f", username, amount, user["balance"])

    return redirect("/profile")


@app.route("/fetch-url", methods=["POST"])
def fetch_url():
    """URL 抓取 - 修复 SSRF 漏洞"""
    username = session.get("username")
    if not username:
        return redirect("/login")

    target_url = request.form.get("url", "")
    status_code = None
    content_preview = None
    error_msg = None

    if target_url:
        # 第一层：只允许 http:// 和 https:// 协议
        if not target_url.startswith(("http://", "https://")):
            error_msg = "只允许访问 http:// 和 https:// 协议的 URL"
        else:
            try:
                # 第二层：解析 URL，提取主机名
                parsed = urllib.parse.urlparse(target_url)
                hostname = parsed.hostname

                # 第三层：禁止访问内网地址
                if _is_internal_ip(hostname):
                    error_msg = "不允许访问内网地址"
                else:
                    # 第四层：不跟随重定向，防止重定向到内网
                    class _NoRedirect(urllib.request.HTTPRedirectHandler):
                        def redirect_request(self, req, fp, code, msg, headers, newurl):
                            return None
                    # 第五层：设置 User-Agent
                    req = urllib.request.Request(
                        target_url,
                        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
                    )
                    opener = urllib.request.build_opener(_NoRedirect)
                    resp = opener.open(req, timeout=10)
                    code = resp.getcode()
                    # 检查是否是重定向响应
                    if 300 <= code < 400:
                        error_msg = f"目标服务器返回了重定向（{code}），已拦截"
                    else:
                        status_code = code if code else "OK"
                        raw = resp.read()
                        # 只返回文本类型内容
                        content_type = resp.headers.get("Content-Type", "")
                        if "text" in content_type or "json" in content_type or "xml" in content_type or "html" in content_type:
                            content_preview = raw.decode("utf-8", errors="replace")[:5000]
                        else:
                            content_preview = f"[二进制内容，Content-Type: {content_type}，大小: {len(raw)} 字节]"
            except urllib.error.HTTPError as e:
                if 300 <= e.code < 400:
                    error_msg = f"目标服务器返回了重定向（{e.code}），已拦截"
                else:
                    status_code = e.code
                    content_preview = str(e.reason)
            except urllib.error.URLError as e:
                error_msg = f"无法访问该 URL: {e.reason}"
            except Exception as e:
                error_msg = f"请求失败: {type(e).__name__}"

    user_info = _safe_user(username)
    return render_template("index.html", user=user_info, fetch_status=status_code, fetch_content=content_preview, fetch_error=error_msg, fetch_url=target_url, results=None, keyword="")


@app.route("/ping", methods=["GET", "POST"])
def ping():
    """Ping 网络诊断"""
    username = session.get("username")
    if not username:
        return redirect("/login")

    command = None
    output = None
    error = None

    if request.method == "POST":
        ip = request.form.get("ip", "")
        if ip:
            command = f"ping -c 3 {ip}"
            try:
                result = subprocess.check_output(command, shell=True, stderr=subprocess.STDOUT, timeout=30)
                output = result.decode("utf-8", errors="replace")
            except subprocess.CalledProcessError as e:
                output = e.output.decode("utf-8", errors="replace") if e.output else str(e)
            except Exception as e:
                error = str(e)

    return render_template("ping.html", command=command, output=output, error=error)


# 本机 IP 列表，启动时自动填充
_HOST_IPS = set()


def _is_internal_ip(hostname):
    """检查主机名是否为内网地址或本机地址"""
    if not hostname:
        return True
    hostname = hostname.lower()
    # 检查是否为域名形式的本地地址
    if hostname in ("localhost", "localhost.localdomain"):
        return True
    # 尝试解析 IP
    try:
        ip = socket.gethostbyname(hostname)
    except socket.gaierror:
        return False
    # 检查是否为本机 IP（启动时获取的本机所有 IP）
    if ip in _HOST_IPS:
        return True
    # 检查内网 IP 范围
    if ip.startswith("127.") or ip == "::1":
        return True
    if ip.startswith("10."):
        return True
    if ip.startswith("172."):
        parts = ip.split(".")
        if len(parts) == 4:
            second = int(parts[1])
            if 16 <= second <= 31:
                return True
    if ip.startswith("192.168."):
        return True
    if ip == "0.0.0.0":
        return True
    if ip.startswith("169.254."):
        return True
    return False


@app.route("/page")
def page():
    """动态页面加载 - 白名单机制，仅允许加载预设的合法页面"""
    name = request.args.get("name", "")
    page_content = None
    page_title = name

    # 白名单：仅允许加载这些预设页面
    allowed_pages = {"help", "about", "help.html", "about.html"}
    if name not in allowed_pages:
        page_content = "页面不存在"
    else:
        if not name.endswith(".html"):
            name += ".html"
        page_path = os.path.join("pages", name)
        if os.path.isfile(page_path):
            with open(page_path, "r", encoding="utf-8") as f:
                page_content = f.read()
                page_title = name
        else:
            page_content = "页面不存在"

    username = session.get("username")
    user_info = _safe_user(username)
    return render_template("index.html", user=user_info, page_content=page_content, page_title=page_title, results=None, keyword="")


if __name__ == "__main__":
    init_db()
    # 获取本机所有 IP 地址，用于 SSRF 防护
    try:
        import subprocess as _sp
        r = _sp.run(["hostname", "-I"], capture_output=True, text=True, timeout=5)
        if r.stdout:
            _HOST_IPS.update(r.stdout.strip().split())
        _HOST_IPS.add(socket.gethostbyname(socket.gethostname()))
        # 获取公网 IP
        try:
            r2 = _sp.run(["curl", "-s", "ifconfig.me"], capture_output=True, text=True, timeout=5)
            if r2.stdout and r2.stdout.strip():
                _HOST_IPS.add(r2.stdout.strip())
        except Exception:
            pass
    except Exception:
        pass
    _HOST_IPS.discard("")
    logger.info("本机 IP 列表: %s", _HOST_IPS)
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", 5000))
    logger.info("启动服务: %s:%d (debug=%s)", host, port, debug)
    app.run(debug=debug, host=host, port=port)
