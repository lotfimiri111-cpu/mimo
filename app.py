"""
Flask API — مذكرتي Pro v18 — Payment Edition
Adds: order creation, receipt upload, admin panel, code redemption, protected download.

ENV:
  ADMIN_PASSWORD_HASH  sha256 hex of admin password (default: sha256("admin1234"))
  DB_PATH              sqlite path (default: mathkarati_payments.db)
  STORAGE_DIR          where receipts & generated pptx are stored (default: ./storage)
  SECRET_KEY           flask secret for sessions
"""
import base64
import hashlib
import io
import logging
import os
import sys
import threading
import time
import unicodedata
import uuid
from functools import wraps
from pathlib import Path

from flask import (Flask, abort, jsonify, make_response, redirect,
                   request, send_file, send_from_directory)

_BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _BASE)

from core.models import PresentationRequest
from core.payment_models import (
    admin_check, admin_login, approve_order, attach_receipt,
    create_order, get_audit, get_order, get_stats, init_db,
    list_orders, redeem_code, reject_order, store_pptx,
    new_download_code, get_db,
)
from core.preview import get_cached_preview, pptx_to_preview_images, set_cached_preview
from engine.pipeline import get_pipeline

# ── App Setup ─────────────────────────────────────────────────────────────────
app = Flask(__name__, static_folder="public", static_url_path="")
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-in-prod-v18")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    stream=sys.stdout,
)
log = logging.getLogger(__name__)

STORAGE_DIR = Path(os.environ.get("STORAGE_DIR", os.path.join(_BASE, "storage")))
STORAGE_DIR.mkdir(parents=True, exist_ok=True)
(STORAGE_DIR / "receipts").mkdir(exist_ok=True)
(STORAGE_DIR / "pptx").mkdir(exist_ok=True)

ALLOWED_RECEIPT_MIME = {
    "image/jpeg", "image/jpg", "image/png", "image/webp",
    "application/pdf",
}
MAX_RECEIPT_SIZE = 10 * 1024 * 1024  # 10 MB

init_db()


# ── Helpers ──────────────────────────────────────────────────────────────────
def _safe_filename(name: str) -> str:
    if not name:
        return f"prs_{int(time.time())}"
    normalized = unicodedata.normalize("NFKD", name)
    ascii_str = normalized.encode("ascii", "ignore").decode("ascii")
    safe = "".join(c if c.isalnum() else "_" for c in ascii_str).strip("_")
    if not safe:
        safe = f"student_{int(time.time()) % 100000}"
    return safe[:24]


def _get_ip() -> str:
    return (
        request.headers.get("X-Forwarded-For", "").split(",")[0].strip()
        or request.remote_addr or "unknown"
    )


def _admin_token() -> str:
    return (
        request.headers.get("X-Admin-Token")
        or request.cookies.get("admin_token")
        or request.args.get("token") or ""
    )


def require_admin(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not admin_check(_admin_token()):
            return jsonify({"error": "Unauthorized"}), 401
        return f(*args, **kwargs)
    return wrapper


# ── CORS ──────────────────────────────────────────────────────────────────────
@app.after_request
def _cors(r):
    r.headers["Access-Control-Allow-Origin"] = "*"
    r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS, PUT, DELETE"
    r.headers["Access-Control-Allow-Headers"] = "Content-Type, X-Admin-Token"
    return r


@app.before_request
def _preflight():
    if request.method == "OPTIONS":
        r = make_response("", 204)
        r.headers["Access-Control-Allow-Origin"] = "*"
        r.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS, PUT, DELETE"
        r.headers["Access-Control-Allow-Headers"] = "Content-Type, X-Admin-Token"
        return r


# ── Static ────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory("public", "index.html")


@app.route("/admin")
@app.route("/admin/")
def admin_page():
    return send_from_directory("public", "admin.html")


# ── Health ────────────────────────────────────────────────────────────────────
@app.route("/ping")
def ping():
    return "pong", 200


@app.route("/health")
def health():
    pipeline = get_pipeline()
    return jsonify({"status": "ok", "version": "18.0", "font": pipeline._font})


@app.route("/warmup")
def warmup():
    get_pipeline()
    return jsonify({"status": "ready"})


# ── Background Preview Generator ──────────────────────────────────────────────
def _generate_preview_bg(presentation_id: str, pptx_path: Path):
    """يولّد المعاينة في الخلفية ويحفظها في الـ cache فور جاهزية الـ PPTX."""
    try:
        log.info(f"Preview BG start: {presentation_id}")
        slides = pptx_to_preview_images(str(pptx_path), watermark=True)
        set_cached_preview(presentation_id, slides)
        log.info(f"Preview BG done: {presentation_id} — {len(slides)} slides")
    except Exception as exc:
        log.warning(f"Preview BG failed for {presentation_id}: {exc}")


# ── Generate ──────────────────────────────────────────────────────────────────
@app.route("/generate", methods=["POST"])
def generate():
    t0 = time.monotonic()
    raw = request.get_json(force=True, silent=True)
    if not raw:
        return jsonify({"error": "بيانات غير صالحة"}), 400

    req = PresentationRequest.from_dict(raw)
    errors = req.validate()
    if errors:
        return jsonify({"error": " | ".join(errors)}), 400

    pipeline = get_pipeline()
    result = pipeline.build(req)

    if not result.success:
        log.error(f"Build failed: {result.error}")
        return jsonify({"error": result.error, "stages": result.stages}), 500

    presentation_id = str(uuid.uuid4())
    pptx_path = STORAGE_DIR / "pptx" / f"{presentation_id}.pptx"
    pptx_path.write_bytes(result.data)

    # ── توليد المعاينة متزامناً وإرسالها مع الـ response فوراً ─────────────
    preview_slides = []
    try:
        preview_slides = pptx_to_preview_images(str(pptx_path), watermark=True)
        set_cached_preview(presentation_id, preview_slides)
        log.info(f"Preview ready inline: {presentation_id} — {len(preview_slides)} slides")
    except Exception as exc:
        log.warning(f"Inline preview failed, falling back to BG thread: {exc}")
        threading.Thread(
            target=_generate_preview_bg,
            args=(presentation_id, pptx_path),
            daemon=True,
        ).start()

    b64 = base64.b64encode(result.data).decode("ascii")
    elapsed = time.monotonic() - t0

    log.info(f"Generated: id={presentation_id} slides={result.slide_count} {elapsed:.2f}s")
    return jsonify({
        "ok": True,
        "presentation_id": presentation_id,
        "slides": result.slide_count,
        "font": result.font_used,
        "elapsed": round(elapsed, 2),
        "stages": result.stages,
        "data": b64,
        "filename": f"mathkarati_{_safe_filename(req.student_name)}.pptx",
        "student_name": req.student_name,
        "title_ar": req.title_ar,
        "degree": raw.get("degree", "licence"),
        "preview_slides": preview_slides,
    })


# ── Preview endpoint (safe slide images with watermark) ───────────────────────
@app.route("/preview/<presentation_id>", methods=["GET"])
def get_preview(presentation_id):
    """
    Returns list of base64 JPEG slide images with watermark.
    These are safe to show — no raw PPTX data exposed.
    """
    # Validate ID format (UUID)
    import re
    if not re.match(r'^[0-9a-f\-]{36}$', presentation_id):
        return jsonify({"error": "معرف غير صالح"}), 400

    # Check cache first — if background thread already finished, return immediately
    cached = get_cached_preview(presentation_id)
    if cached:
        return jsonify({"ok": True, "slides": cached, "count": len(cached), "cached": True})

    # Find the PPTX file
    pptx_path = STORAGE_DIR / "pptx" / f"{presentation_id}.pptx"
    if not pptx_path.exists():
        return jsonify({"error": "العرض غير موجود أو انتهت صلاحيته"}), 404

    # Generate on-demand if not cached yet
    try:
        slides = pptx_to_preview_images(str(pptx_path), watermark=True)
        set_cached_preview(presentation_id, slides)
        if slides:
            return jsonify({"ok": True, "slides": slides, "count": len(slides), "cached": False})
    except Exception as exc:
        log.warning(f"On-demand preview failed: {exc}")

    return jsonify({"ok": False, "processing": True,
                    "message": "المعاينة قيد التجهيز، يرجى الانتظار..."}), 202


# ── Orders ────────────────────────────────────────────────────────────────────
@app.route("/orders", methods=["POST"])
def create_order_route():
    d = request.get_json(force=True, silent=True) or {}
    pid = d.get("presentation_id", "").strip()
    student_name = d.get("student_name", "").strip()
    email = d.get("email", "").strip()
    phone = d.get("phone", "").strip()
    degree = d.get("degree", "licence").strip()
    title_ar = d.get("title_ar", "").strip()
    payment_method = d.get("payment_method", "").strip()

    if not pid or not student_name or not payment_method:
        return jsonify({"error": "بيانات ناقصة"}), 400
    if payment_method not in ("ccp", "baridimob"):
        return jsonify({"error": "طريقة دفع غير صالحة"}), 400

    order = create_order(pid, student_name, email, phone, degree, title_ar, payment_method)
    log.info(f"Order created: {order['id']} student={student_name}")
    return jsonify({"ok": True, "order_id": order["id"]}), 201


@app.route("/orders/<order_id>", methods=["GET"])
def get_order_route(order_id):
    order = get_order(order_id)
    if not order:
        return jsonify({"error": "الطلب غير موجود"}), 404
    safe = {k: v for k, v in order.items()
            if k not in ("pptx_path", "receipt_path", "download_code", "download_ip")}
    return jsonify(safe)


@app.route("/orders/<order_id>/receipt", methods=["POST"])
def upload_receipt(order_id):
    order = get_order(order_id)
    if not order:
        return jsonify({"error": "الطلب غير موجود"}), 404
    if order["status"] not in ("pending",):
        return jsonify({"error": "لا يمكن رفع وصل لهذا الطلب"}), 400

    if "receipt" not in request.files:
        return jsonify({"error": "لم يتم إرسال الملف"}), 400

    file = request.files["receipt"]
    mime = file.content_type or "application/octet-stream"
    if mime not in ALLOWED_RECEIPT_MIME:
        return jsonify({"error": "نوع الملف غير مسموح (JPG، PNG، PDF فقط)"}), 400

    data = file.read(MAX_RECEIPT_SIZE + 1)
    if len(data) > MAX_RECEIPT_SIZE:
        return jsonify({"error": "حجم الملف يتجاوز 10MB"}), 400

    ext = "pdf" if "pdf" in mime else "jpg"
    fname = f"{order_id}_{int(time.time())}.{ext}"
    path = STORAGE_DIR / "receipts" / fname
    path.write_bytes(data)

    attach_receipt(order_id, str(path), mime)
    log.info(f"Receipt uploaded: {order_id} file={fname}")
    return jsonify({"ok": True, "message": "تم رفع الوصل بنجاح، سيتم مراجعته خلال 24 ساعة"})


@app.route("/redeem", methods=["POST"])
def redeem_by_code():
    d = request.get_json(force=True, silent=True) or {}
    code = d.get("code", "").strip().upper()
    if not code:
        return jsonify({"error": "الرجاء إدخال الكود"}), 400

    order = redeem_code(code, _get_ip())
    if not order:
        return jsonify({"error": "الكود غير صحيح أو منتهي الصلاحية أو تم استخدامه مسبقاً"}), 403

    pptx_path = order.get("pptx_path")
    if not pptx_path or not Path(pptx_path).exists():
        return jsonify({"error": "ملف العرض غير متوفر، تواصل مع الدعم"}), 500

    pptx_bytes = Path(pptx_path).read_bytes()
    b64 = base64.b64encode(pptx_bytes).decode("ascii")
    safe = _safe_filename(order["student_name"])
    filename = f"mathkarati_{safe}.pptx"

    log.info(f"Download via code: order={order['id']} ip={_get_ip()}")
    return jsonify({"ok": True, "data": b64, "filename": filename,
                    "order_id": order["id"], "student_name": order["student_name"]})


# ── Admin API ─────────────────────────────────────────────────────────────────
@app.route("/admin/login", methods=["POST"])
def admin_login_route():
    d = request.get_json(force=True, silent=True) or {}
    password = d.get("password", "")
    token = admin_login(password, _get_ip())
    if not token:
        return jsonify({"error": "كلمة المرور غير صحيحة"}), 401
    resp = jsonify({"ok": True, "token": token})
    resp.set_cookie("admin_token", token, httponly=True, samesite="Strict", max_age=8 * 3600)
    return resp


@app.route("/admin/stats", methods=["GET"])
@require_admin
def admin_stats():
    return jsonify(get_stats())


@app.route("/admin/orders", methods=["GET"])
@require_admin
def admin_orders():
    status = request.args.get("status")
    orders = list_orders(status=status or None)
    return jsonify(orders)


@app.route("/admin/orders/<order_id>", methods=["GET"])
@require_admin
def admin_order_detail(order_id):
    order = get_order(order_id)
    if not order:
        return jsonify({"error": "غير موجود"}), 404
    audit = get_audit(order_id)
    return jsonify({"order": order, "audit": audit})


@app.route("/admin/orders/<order_id>/receipt", methods=["GET"])
@require_admin
def admin_view_receipt(order_id):
    order = get_order(order_id)
    if not order or not order.get("receipt_path"):
        abort(404)
    path = Path(order["receipt_path"])
    if not path.exists():
        abort(404)
    mime = order.get("receipt_mime", "application/octet-stream")
    return send_file(path, mimetype=mime, as_attachment=False)


@app.route("/admin/orders/<order_id>/approve", methods=["POST"])
@require_admin
def admin_approve(order_id):
    order = get_order(order_id)
    if not order:
        return jsonify({"error": "غير موجود"}), 404
    if order["status"] not in ("pending", "rejected"):
        return jsonify({"error": f"الطلب في حالة {order['status']}"}), 400

    d = request.get_json(force=True, silent=True) or {}
    note = d.get("note", "")
    ttl = int(d.get("ttl_hours", 48))

    pid = order["presentation_id"]
    pptx_src = STORAGE_DIR / "pptx" / f"{pid}.pptx"
    if not pptx_src.exists():
        return jsonify({"error": "ملف العرض غير موجود على الخادم، ربما انتهت صلاحيته"}), 500

    store_pptx(order_id, str(pptx_src))
    code = approve_order(order_id, note, ttl)
    log.info(f"Approved: {order_id} code={code}")
    return jsonify({"ok": True, "download_code": code, "ttl_hours": ttl})


@app.route("/admin/orders/<order_id>/reject", methods=["POST"])
@require_admin
def admin_reject(order_id):
    d = request.get_json(force=True, silent=True) or {}
    note = d.get("note", "")
    reject_order(order_id, note)
    log.info(f"Rejected: {order_id}")
    return jsonify({"ok": True})


@app.route("/admin/orders/<order_id>/resend-code", methods=["POST"])
@require_admin
def admin_resend_code(order_id):
    order = get_order(order_id)
    if not order:
        return jsonify({"error": "غير موجود"}), 404
    if order["status"] not in ("approved", "downloaded"):
        return jsonify({"error": "يجب أن يكون الطلب معتمداً"}), 400

    d = request.get_json(force=True, silent=True) or {}
    ttl = int(d.get("ttl_hours", 24))
    code = new_download_code()
    expires = time.time() + ttl * 3600
    conn = get_db()
    try:
        conn.execute(
            """UPDATE orders SET download_code=?, code_expires=?, code_used=0,
               status='approved', updated_at=? WHERE id=?""",
            (code, expires, time.time(), order_id),
        )
        conn.commit()
    finally:
        conn.close()
    log.info(f"Code regenerated: {order_id} code={code}")
    return jsonify({"ok": True, "download_code": code, "ttl_hours": ttl})


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    get_pipeline()
    app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)
