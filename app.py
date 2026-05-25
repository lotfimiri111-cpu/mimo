"""
مذكرتي Pro v18 Commercial — Flask Application
Full commercial system: Generate → Preview → Pay → Download
"""
import base64
import json
import logging
import os
import sys
import threading
import time
import unicodedata
import uuid

from flask import (Flask, Response, abort, jsonify, make_response,
                   request, send_file, send_from_directory, session)

_BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _BASE)

from core.models import PresentationRequest
from core.database import (
    init_db, create_order, set_order_files, submit_receipt,
    get_order, get_order_by_token, list_orders, approve_order,
    reject_order, validate_code, list_codes, deactivate_code,
    create_manual_code, get_settings, update_settings,
    check_admin_password, get_stats
)
from core.preview import (
    pptx_to_pdf, render_slide, render_thumbnail,
    prerender_all, get_slide_count_from_pdf
)
from engine.pipeline import get_pipeline

app = Flask(__name__, static_folder="public", static_url_path="")
app.secret_key = os.environ.get("SECRET_KEY", "mathkarati-v18-" + str(uuid.uuid4()))

logging.basicConfig(level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s", stream=sys.stdout)
log = logging.getLogger(__name__)

STORAGE_DIR = os.path.join(_BASE, "storage")
PPTX_DIR = os.path.join(STORAGE_DIR, "pptx")
RECEIPT_DIR = os.path.join(STORAGE_DIR, "receipts")
os.makedirs(PPTX_DIR, exist_ok=True)
os.makedirs(RECEIPT_DIR, exist_ok=True)


def _safe_filename(name):
    if not name: return f"prs_{int(time.time())}"
    normalized = unicodedata.normalize("NFKD", name)
    ascii_str = normalized.encode("ascii", "ignore").decode("ascii")
    safe = "".join(c if c.isalnum() else "_" for c in ascii_str).strip("_")
    if not safe: safe = f"student_{int(time.time()) % 100000}"
    return safe[:24]

def _cors(r):
    r.headers["Access-Control-Allow-Origin"] = "*"
    r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    r.headers["Access-Control-Allow-Headers"] = "Content-Type, X-Admin-Token"
    return r

def _no_cache(r):
    r.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    r.headers["Pragma"] = "no-cache"
    return r

def _check_admin():
    token = request.headers.get("X-Admin-Token") or request.args.get("admin_token")
    return token and token == session.get("admin_token")

def _client_ip():
    return (request.headers.get("X-Forwarded-For","") or request.remote_addr or "").split(",")[0].strip()


@app.after_request
def after_request(r): return _cors(r)

@app.before_request
def preflight():
    if request.method == "OPTIONS":
        return _cors(make_response("", 204))


# ── Pages ─────────────────────────────────────────────────────────────
@app.route("/")
def index(): return send_from_directory("public", "index.html")

@app.route("/preview/<token>")
def preview_page(token):
    if not get_order_by_token(token): abort(404)
    return send_from_directory("public", "preview.html")

@app.route("/download")
def download_page(): return send_from_directory("public", "download.html")

@app.route("/admin")
def admin_page(): return send_from_directory("admin", "index.html")

@app.route("/ping")
def ping(): return "pong", 200

@app.route("/health")
def health():
    pipeline = get_pipeline()
    return jsonify({"status": "ok", "version": "18.0", "font": pipeline._font}), 200


# ── Generate ──────────────────────────────────────────────────────────
@app.route("/api/generate", methods=["POST"])
def generate():
    t0 = time.monotonic()
    raw = request.get_json(force=True, silent=True)
    if not raw: return jsonify({"error": "بيانات غير صالحة"}), 400

    req = PresentationRequest.from_dict(raw)
    errors = req.validate()
    if errors: return jsonify({"error": " | ".join(errors)}), 400

    pipeline = get_pipeline()
    result = pipeline.build(req)
    if not result.success: return jsonify({"error": result.error}), 500

    order_info = create_order(
        student_name=req.student_name, title_ar=req.title_ar,
        engine=req.engine, theme=req.theme, request_json=raw,
        slide_count=result.slide_count,
    )
    order_id = order_info["id"]
    preview_token = order_info["preview_token"]

    safe = _safe_filename(req.student_name)
    pptx_filename = f"{order_id[:8]}_{safe}.pptx"
    pptx_path = os.path.join(PPTX_DIR, pptx_filename)
    with open(pptx_path, "wb") as f: f.write(result.data)

    def _bg():
        pdf_path = pptx_to_pdf(pptx_path, order_id)
        if pdf_path:
            set_order_files(order_id, pptx_path, pdf_path)
            prerender_all(pptx_path, order_id, result.slide_count)

    threading.Thread(target=_bg, daemon=True).start()
    set_order_files(order_id, pptx_path, "")

    settings = get_settings()
    log.info(f"Generated order={order_id[:8]} slides={result.slide_count} {time.monotonic()-t0:.2f}s")
    return jsonify({
        "ok": True, "order_id": order_id, "preview_token": preview_token,
        "preview_url": f"/preview/{preview_token}",
        "slides": result.slide_count, "font": result.font_used,
        "elapsed": round(time.monotonic()-t0, 2),
        "price": settings.get("price","800"), "currency": settings.get("currency","دج"),
    })


# ── Preview API ───────────────────────────────────────────────────────
@app.route("/api/preview/<token>/info")
def preview_info(token):
    order = get_order_by_token(token)
    if not order: return jsonify({"error": "not found"}), 404
    settings = get_settings()
    return jsonify({
        "order_id": order["id"], "student_name": order["student_name"],
        "title_ar": order["title_ar"], "engine": order["engine"],
        "theme": order["theme"], "slide_count": order["slide_count"],
        "status": order["status"], "price": settings.get("price","800"),
        "currency": settings.get("currency","دج"),
        "payment_ccp": settings.get("payment_ccp",""),
        "payment_baridimob": settings.get("payment_baridimob",""),
        "payment_name": settings.get("payment_name",""),
        "payment_info": settings.get("payment_info",""),
        "created_at": order["created_at"],
    })


@app.route("/api/preview/<token>/slide/<int:n>")
def preview_slide(token, n):
    order = get_order_by_token(token)
    if not order: abort(404)
    if n < 1 or n > order["slide_count"] + 2: abort(404)
    pptx_path = order.get("pptx_path","")
    if not pptx_path or not os.path.exists(pptx_path): abort(404)

    pdf_path = order.get("pdf_path","")
    if not pdf_path or not os.path.exists(pdf_path):
        new_pdf = pptx_to_pdf(pptx_path, order["id"])
        if new_pdf: set_order_files(order["id"], pptx_path, new_pdf)

    img_bytes = render_slide(pptx_path, order["id"], n, watermark=True)
    if not img_bytes: img_bytes = _placeholder(n)
    resp = make_response(img_bytes)
    resp.headers["Content-Type"] = "image/jpeg"
    resp.headers["X-Robots-Tag"] = "noindex"
    return _no_cache(resp)


@app.route("/api/preview/<token>/thumb/<int:n>")
def preview_thumb(token, n):
    order = get_order_by_token(token)
    if not order: abort(404)
    pptx_path = order.get("pptx_path","")
    if not pptx_path or not os.path.exists(pptx_path): abort(404)
    img_bytes = render_thumbnail(pptx_path, order["id"], n)
    if not img_bytes: img_bytes = _placeholder(n, thumb=True)
    resp = make_response(img_bytes)
    resp.headers["Content-Type"] = "image/jpeg"
    return _no_cache(resp)


def _placeholder(n, thumb=False):
    from PIL import Image, ImageDraw
    import io
    w, h = (300,169) if thumb else (1280,720)
    img = Image.new("RGB",(w,h),(15,30,65))
    draw = ImageDraw.Draw(img)
    draw.rectangle([0,0,w,4], fill=(198,160,60))
    msg = f"شريحة {n} — جاري التحضير..."
    draw.text((w//2-60, h//2-10), msg, fill=(198,160,60))
    buf = io.BytesIO(); img.save(buf,"JPEG",quality=70)
    return buf.getvalue()


# ── Payment / Receipt ─────────────────────────────────────────────────
@app.route("/api/orders/<order_id>/receipt", methods=["POST"])
def upload_receipt(order_id):
    order = get_order(order_id)
    if not order: return jsonify({"error": "الطلب غير موجود"}), 404
    if order["status"] not in ("preview","pending"):
        return jsonify({"error": f"حالة الطلب: {order['status']}"}), 400
    if "receipt" not in request.files:
        return jsonify({"error": "يرجى إرفاق وصل التسديد"}), 400
    file = request.files["receipt"]
    if not file.filename: return jsonify({"error": "ملف فارغ"}), 400
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in (".jpg",".jpeg",".png",".pdf",".webp"):
        return jsonify({"error": "صيغة الملف غير مدعومة"}), 400
    receipt_filename = f"{order_id[:8]}_receipt{ext}"
    receipt_path = os.path.join(RECEIPT_DIR, receipt_filename)
    file.save(receipt_path)
    submit_receipt(order_id, receipt_path, receipt_filename)
    return jsonify({"ok": True, "message": "تم استلام وصل التسديد. سيتم مراجعته خلال 24 ساعة."})


# ── Download ──────────────────────────────────────────────────────────
@app.route("/api/download/verify", methods=["POST"])
def verify_download():
    data = request.get_json(force=True, silent=True) or {}
    code = str(data.get("code","")).strip().upper()
    if not code: return jsonify({"error": "يرجى إدخال كود التحميل"}), 400

    result = validate_code(code, ip=_client_ip(), user_agent=request.headers.get("User-Agent",""))
    if not result["valid"]: return jsonify({"error": result["error"]}), 400

    order = get_order(result["order_id"])
    if not order: return jsonify({"error": "الطلب غير موجود"}), 404

    pptx_path = order.get("pptx_path","")
    if not pptx_path or not os.path.exists(pptx_path):
        return jsonify({"error": "الملف غير متوفر. تواصل مع الدعم."}), 500

    download_name = f"mathkarati_{_safe_filename(order['student_name'])}.pptx"
    log.info(f"Download: code={code} order={order['id'][:8]} ip={_client_ip()}")
    return send_file(pptx_path, as_attachment=True, download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")


# ── Admin Auth ────────────────────────────────────────────────────────
@app.route("/api/admin/login", methods=["POST"])
def admin_login():
    data = request.get_json(force=True, silent=True) or {}
    if check_admin_password(str(data.get("password",""))):
        token = str(uuid.uuid4())
        session["admin_token"] = token
        session.permanent = True
        return jsonify({"ok": True, "token": token})
    return jsonify({"error": "كلمة المرور غير صحيحة"}), 401

@app.route("/api/admin/logout", methods=["POST"])
def admin_logout():
    session.clear(); return jsonify({"ok": True})


# ── Admin Orders ──────────────────────────────────────────────────────
@app.route("/api/admin/orders")
def admin_orders():
    if not _check_admin(): return jsonify({"error": "غير مصرح"}), 401
    status = request.args.get("status")
    return jsonify({"orders": list_orders(status=status or None)})

@app.route("/api/admin/orders/<order_id>")
def admin_order_detail(order_id):
    if not _check_admin(): return jsonify({"error": "غير مصرح"}), 401
    order = get_order(order_id)
    if not order: return jsonify({"error": "not found"}), 404
    return jsonify({"order": order, "codes": list_codes(order_id=order_id)})

@app.route("/api/admin/orders/<order_id>/approve", methods=["POST"])
def admin_approve(order_id):
    if not _check_admin(): return jsonify({"error": "غير مصرح"}), 401
    if not get_order(order_id): return jsonify({"error": "not found"}), 404
    data = request.get_json(force=True, silent=True) or {}
    code = approve_order(order_id, note=str(data.get("note","")))
    log.info(f"Order approved: {order_id[:8]} code={code}")
    return jsonify({"ok": True, "code": code})

@app.route("/api/admin/orders/<order_id>/reject", methods=["POST"])
def admin_reject(order_id):
    if not _check_admin(): return jsonify({"error": "غير مصرح"}), 401
    if not get_order(order_id): return jsonify({"error": "not found"}), 404
    data = request.get_json(force=True, silent=True) or {}
    reject_order(order_id, note=str(data.get("note","")))
    return jsonify({"ok": True})

@app.route("/api/admin/receipt/<order_id>")
def admin_receipt(order_id):
    if not _check_admin(): return jsonify({"error": "غير مصرح"}), 401
    order = get_order(order_id)
    if not order or not order.get("receipt_path"): abort(404)
    if not os.path.exists(order["receipt_path"]): abort(404)
    return send_file(order["receipt_path"])

@app.route("/api/admin/stats")
def admin_stats():
    if not _check_admin(): return jsonify({"error": "غير مصرح"}), 401
    return jsonify(get_stats())

@app.route("/api/admin/settings", methods=["GET","POST"])
def admin_settings():
    if not _check_admin(): return jsonify({"error": "غير مصرح"}), 401
    if request.method == "GET":
        s = get_settings(); s.pop("admin_password", None)
        return jsonify(s)
    data = request.get_json(force=True, silent=True) or {}
    allowed = {"price","currency","payment_ccp","payment_baridimob","payment_name",
               "payment_info","contact_email","code_validity_days","site_name","admin_password"}
    updates = {k:v for k,v in data.items() if k in allowed}
    if updates: update_settings(updates)
    return jsonify({"ok": True})

@app.route("/api/admin/codes", methods=["POST"])
def admin_create_code():
    if not _check_admin(): return jsonify({"error": "غير مصرح"}), 401
    data = request.get_json(force=True, silent=True) or {}
    order_id = str(data.get("order_id",""))
    if not order_id or not get_order(order_id): return jsonify({"error": "طلب غير موجود"}), 404
    code = create_manual_code(order_id, max_uses=int(data.get("max_uses",1)), days=int(data.get("days",7)))
    return jsonify({"ok": True, "code": code})

@app.route("/api/admin/codes/<code>/deactivate", methods=["POST"])
def admin_deactivate_code(code):
    if not _check_admin(): return jsonify({"error": "غير مصرح"}), 401
    deactivate_code(code.upper()); return jsonify({"ok": True})


# ── Startup ───────────────────────────────────────────────────────────
if __name__ == "__main__":
    init_db()
    get_pipeline()
    port = int(os.environ.get("PORT", 5000))
    log.info(f"مذكرتي Pro v18 Commercial — port {port}")
    app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)
