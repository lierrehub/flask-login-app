import os
import re
from datetime import timedelta
from flask import Flask, render_template, request, redirect, session
from flask_bcrypt import Bcrypt
from flask_wtf.csrf import CSRFProtect
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

app = Flask(__name__)
app.secret_key = os.environ.get(
    "SECRET_KEY",
    os.urandom(64).hex()
)

# 登录频率限制（基于 IP）
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=["200 per day", "50 per hour"],
    storage_uri="memory://",
)

# CSRF 保护
csrf = CSRFProtect(app)

# Session 安全配置
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_PERMANENT"] = True
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=1)
# 仅在明确启用时设置 Secure 标志（需要 HTTPS）
if os.environ.get("SESSION_COOKIE_SECURE", "false").lower() == "true":
    app.config["SESSION_COOKIE_SECURE"] = True

bcrypt = Bcrypt(app)

# 用户数据库 - 密码为预生成的 bcrypt 哈希，源码中不包含明文密码
USERS = {
    "admin": {
        "username": "admin",
        "password": "$2b$12$.cLDo0EipcOk.1F6EedqdebiMwwCRZHAz84sMRu/6u5qwjE3v9k4e",
        "role": "admin",
        "email": "admin@example.com",
        "phone": "13800138000",
        "balance": 99999
    },
    "alice": {
        "username": "alice",
        "password": "$2b$12$pIqdI5oYxQD/XagaMngLqON5PzAlPqAaTO2WB/w4wyAS/GAZhgqVm",
        "role": "user",
        "email": "alice@example.com",
        "phone": "13900139001",
        "balance": 100
    }
}


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
    return response


@app.route("/")
def index():
    username = session.get("username")
    user_info = _safe_user(username)
    return render_template("index.html", user=user_info)


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
            # 防时序攻击：无论用户名是否存在都执行 bcrypt 验证
            real_user = USERS.get(username)
            dummy_hash = "$2b$12$ABCDEFGHIJKLMNOPQRSTUOhash4demoOnly1234567890123"
            check_hash = real_user["password"] if real_user else dummy_hash
            password_valid = bcrypt.check_password_hash(check_hash, password)

            if real_user and password_valid:
                # 登录成功后刷新 session，防止会话固定攻击
                session.clear()
                session.permanent = True
                session["username"] = username
                return render_template("index.html", user=_safe_user(username))
            else:
                error = "用户名或密码错误，请重试"

    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


if __name__ == "__main__":
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=debug, host=host, port=port)
