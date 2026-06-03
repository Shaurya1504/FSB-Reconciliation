"""
server.py
---------
Flask server with Excel-based session authentication.
Handles file upload, calls reconcile.run(), returns xlsx.

Setup (local):
  1. pip install -r requirements.txt
  2. python server.py

Deploy to Render:
  - Build command : pip install -r requirements.txt
  - Start command : gunicorn server:app
  - Environment vars (set in Render dashboard):
      FSB_SECRET   → any long random string (required in production)
      AUTH_DB_PATH → path to auth_db.xlsx if not bundled (optional)
"""

import io
import os
import secrets
import logging

import pandas as pd
from flask import (Flask, abort, jsonify, redirect, request,
                   send_file, send_from_directory, session, url_for)

import reconcile

# ── Logging ────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger(__name__)

# ── App init ───────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, static_folder=BASE_DIR)

# Secret key: MUST be set via environment variable in production (Render).
# Falling back to a random token means sessions are lost on every restart —
# acceptable locally, not in prod.
_secret = os.environ.get("FSB_SECRET", "")
if not _secret:
    _secret = secrets.token_hex(32)
    logger.warning(
        "FSB_SECRET env var not set. Using a random secret key — "
        "all sessions will be invalidated on restart. "
        "Set FSB_SECRET in your Render environment variables."
    )
app.secret_key = _secret

# Session cookie config — safe defaults for HTTPS deployments
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    # Set SECURE=True automatically when running behind Render's HTTPS proxy
    SESSION_COOKIE_SECURE=os.environ.get("RENDER", "") != "",
    MAX_CONTENT_LENGTH=200 * 1024 * 1024,   # 200 MB upload limit
)

# Auth DB location: can be overridden via env var for flexible deployments
AUTH_DB = os.environ.get("AUTH_DB_PATH", os.path.join(BASE_DIR, "auth_db.xlsx"))

PROTECTED = {"/", "/process"}


# ── Error handlers (return JSON, never HTML) ──────────────────────────────
@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "Files too large. Maximum upload size is 200 MB."}), 413

@app.errorhandler(500)
def server_error(e):
    return jsonify({"error": "Internal server error. Check logs."}), 500


# ── Auth guard ─────────────────────────────────────────────────────────────
@app.before_request
def require_login():
    if request.path in PROTECTED and not session.get("authenticated"):
        if request.path == "/process":
            return jsonify({"error": "Unauthorized. Please log in."}), 401
        return redirect(url_for("login_page"))


# ── Health check (required by Render) ─────────────────────────────────────
@app.route("/health")
def health():
    return jsonify({"status": "ok"}), 200


# ── Static routes ──────────────────────────────────────────────────────────
@app.route("/login")
def login_page():
    return send_file(os.path.join(BASE_DIR, "index.html"))


@app.route("/")
def index():
    return send_file(os.path.join(BASE_DIR, "index.html"))


# ── Auth API ───────────────────────────────────────────────────────────────
@app.route("/api/login", methods=["POST"])
def api_login():
    data   = request.get_json(silent=True) or {}
    logid  = str(data.get("logid", "")).strip()
    passwd = str(data.get("password", "")).strip()

    if not logid or not passwd:
        return jsonify({"error": "logid and password are required."}), 400

    if not os.path.exists(AUTH_DB):
        logger.error("Auth DB not found at: %s", AUTH_DB)
        return jsonify({"error": f"Auth database not found on server."}), 500

    try:
        db = pd.read_excel(AUTH_DB, engine="calamine", dtype=str)
        db.columns = [c.strip().lower() for c in db.columns]
        db = db.fillna("")
    except Exception as e:
        logger.exception("Failed to read auth DB")
        return jsonify({"error": f"Failed to read auth database: {e}"}), 500

    if "logid" not in db.columns or "password" not in db.columns:
        return jsonify({"error": "auth_db.xlsx must have 'logid' and 'password' columns."}), 500

    match = db[(db["logid"].str.strip() == logid) &
               (db["password"].str.strip() == passwd)]

    if match.empty:
        logger.warning("Failed login attempt for logid: %s", logid)
        return jsonify({"error": "Invalid credentials."}), 401

    session.permanent = False
    session["authenticated"] = True
    session["logid"] = logid
    logger.info("User logged in: %s", logid)
    return jsonify({"ok": True, "logid": logid})


@app.route("/api/logout", methods=["POST"])
def api_logout():
    logid = session.get("logid", "unknown")
    session.clear()
    logger.info("User logged out: %s", logid)
    return jsonify({"ok": True})


@app.route("/api/me")
def api_me():
    if session.get("authenticated"):
        return jsonify({"authenticated": True, "logid": session.get("logid")})
    return jsonify({"authenticated": False}), 401


# ── Reconciliation ─────────────────────────────────────────────────────────
@app.route("/process", methods=["POST"])
def process():
    if "invoice" not in request.files or "portal" not in request.files:
        return jsonify({"error": "Both files are required."}), 400

    invoice_file = request.files["invoice"]
    portal_file  = request.files["portal"]

    invoice_bytes = invoice_file.read()
    portal_bytes  = portal_file.read()

    if not invoice_bytes or not portal_bytes:
        return jsonify({"error": "One or both files are empty."}), 400

    logger.info(
        "Processing reconciliation for %s — invoice: %s (%d bytes), portal: %s (%d bytes)",
        session.get("logid", "unknown"),
        invoice_file.filename, len(invoice_bytes),
        portal_file.filename,  len(portal_bytes),
    )

    try:
        result_bytes = reconcile.run(invoice_bytes, portal_bytes)
    except Exception as e:
        logger.exception("Reconciliation failed")
        return jsonify({"error": str(e)}), 500

    logger.info("Reconciliation complete — output size: %d bytes", len(result_bytes))

    return send_file(
        io.BytesIO(result_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="FSB_Reconciliation_Master.xlsx"
    )


# ── Entry point ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    print("=" * 56)
    print("  FSB Reconciliation Server")
    print(f"  http://localhost:{port}")
    print(f"  Auth DB : {AUTH_DB}")
    print(f"  Debug   : {debug}")
    print("=" * 56)
    app.run(host="0.0.0.0", port=port, debug=debug)
