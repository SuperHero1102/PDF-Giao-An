# -*- coding: utf-8 -*-
"""
PDF Giáo Án - máy chủ (Flask).

Luồng: tải PDF lên -> mỗi trang được vẽ thành ảnh -> Gemini đọc ảnh trang thành các "khối"
(chữ, công thức LaTeX, hình) -> người dùng xem/sửa trên web -> xuất Word (.docx).

Khóa Gemini: người dùng nhập trong Cài đặt (lưu trên máy họ) hoặc đặt biến môi trường
GEMINI_API_KEY trên máy chủ. Máy chủ KHÔNG lưu khóa.
"""
import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import quote

from flask import Flask, abort, jsonify, request, send_file, send_from_directory
from PIL import Image

import recognize
from docx_builder import Builder

BASE = os.path.dirname(os.path.abspath(__file__))
WORK = os.environ.get("PDFGA_WORKDIR", os.path.join(tempfile.gettempdir(), "pdfgiaoan"))
TTL_SECONDS = int(os.environ.get("PDFGA_TTL_HOURS", "3")) * 3600  # tự xóa file tạm sau N giờ
MAX_MB = int(os.environ.get("PDFGA_MAX_MB", "100"))
DEMO = os.environ.get("PDFGA_DEMO") == "1"  # chỉ dùng để thử máy chủ khi chưa có khóa Gemini

os.makedirs(WORK, exist_ok=True)
app = Flask(__name__, static_folder=os.path.join(BASE, "static"), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = MAX_MB * 1024 * 1024

LOCK = threading.Lock()
STATUS = {}  # job_id -> trạng thái nhận dạng (trong bộ nhớ)
GATES = {}   # job_id -> cổng chờ khi Gemini giới hạn tốc độ
EXPORT = {}  # job_id -> trạng thái tạo file Word chạy nền


class Gate:
    """Khi một trang bị Gemini giới hạn tốc độ (429), cả nhóm cùng chờ rồi mới gọi tiếp (tránh dồn dập)."""

    def __init__(self):
        self.until = 0.0

    def wait(self):
        while True:
            d = self.until - time.time()
            if d <= 0:
                return
            time.sleep(min(d, 1.0))

    def penalize(self, sec):
        self.until = max(self.until, time.time() + sec)

    def remaining(self):
        return max(0, int(self.until - time.time() + 0.99))
ID_RE = re.compile(r"^[a-f0-9]{16}$")


# ----------------------------------------------------------------------------
# Tiện ích
# ----------------------------------------------------------------------------
def err(msg, code=400):
    return jsonify({"error": msg}), code


def jdir(job):
    if not ID_RE.match(job or ""):
        abort(404)
    d = os.path.join(WORK, job)
    if not os.path.isdir(d):
        abort(404)
    os.utime(d, None)  # gia hạn thời gian giữ file
    return d


def read_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def write_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    os.replace(tmp, path)


def cleanup_old():
    now = time.time()
    for name in os.listdir(WORK):
        p = os.path.join(WORK, name)
        try:
            if os.path.isdir(p) and now - os.path.getmtime(p) > TTL_SECONDS:
                shutil.rmtree(p, ignore_errors=True)
                STATUS.pop(name, None)
        except OSError:
            pass


def page_img(d, n):
    """Vẽ trang n (1..) ra ảnh JPEG 200 DPI, chỉ vẽ khi cần. (JPEG nhẹ hơn PNG nhiều lần.)"""
    out = os.path.join(d, "p-%d.jpg" % n)
    if not os.path.exists(out):
        r = subprocess.run(["pdftoppm", "-r", "200", "-f", str(n), "-l", str(n), "-jpeg", "-jpegopt", "quality=92",
                            "-singlefile", os.path.join(d, "src.pdf"), os.path.join(d, "p-%d" % n)],
                           capture_output=True, timeout=180)
        if r.returncode != 0 or not os.path.exists(out):
            raise RuntimeError("Không đọc được trang %d của file PDF." % n)
    return out


def load_pages(d):
    raw = read_json(os.path.join(d, "blocks.json"), {})
    return {int(k): v for k, v in raw.items()}


def save_pages(d, pages):
    write_json(os.path.join(d, "blocks.json"), {str(k): v for k, v in pages.items()})


# ----------------------------------------------------------------------------
# Trang web
# ----------------------------------------------------------------------------
@app.route("/")
def index():
    return send_from_directory(os.path.join(BASE, "static"), "index.html")


@app.route("/print/<job>")
def print_page(job):
    """Trang in độc lập (mở ở thẻ mới) để lưu PDF: đủ mọi trang đã đọc."""
    resp = send_from_directory(os.path.join(BASE, "static"), "print.html")
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/sw.js")
def sw():  # service worker phải nằm ở gốc để phủ cả trang
    return send_from_directory(os.path.join(BASE, "static"), "sw.js", mimetype="application/javascript")


@app.route("/api/config")
def config():
    return jsonify({"server_key": bool(os.environ.get("GEMINI_API_KEY")), "max_mb": MAX_MB,
                    "model": os.environ.get("GEMINI_MODEL", recognize.DEFAULT_MODEL), "demo": DEMO})


# ----------------------------------------------------------------------------
# Tải PDF lên
# ----------------------------------------------------------------------------
@app.route("/api/upload", methods=["POST"])
def upload():
    cleanup_old()
    f = request.files.get("file")
    if not f or not f.filename:
        return err("Chưa chọn file. Hãy chọn một file PDF.")
    head = f.stream.read(5)
    f.stream.seek(0)
    if head != b"%PDF-":
        return err("File này không phải PDF. Hãy chọn file có đuôi .pdf.")
    job = uuid.uuid4().hex[:16]
    d = os.path.join(WORK, job)
    os.makedirs(d)
    f.save(os.path.join(d, "src.pdf"))
    r = subprocess.run(["pdfinfo", os.path.join(d, "src.pdf")], capture_output=True, text=True, timeout=60)
    m = re.search(r"^Pages:\s+(\d+)", r.stdout, re.M)
    if r.returncode != 0 or not m:
        shutil.rmtree(d, ignore_errors=True)
        low = (r.stderr or "").lower()
        if "password" in low or "encrypt" in low:
            return err("File PDF đang đặt mật khẩu. Hãy mở khóa rồi tải lại.")
        return err("Không đọc được file PDF này (có thể bị hỏng).")
    pages = int(m.group(1))
    sizes = {}  # khổ giấy từng trang (đã tính góc xoay) để Word có đúng khổ/hướng như PDF gốc
    try:
        info = subprocess.run(["pdfinfo", "-f", "1", "-l", str(pages), os.path.join(d, "src.pdf")],
                              capture_output=True, text=True, timeout=120).stdout
        rot = {int(n): int(v) for n, v in re.findall(r"^Page\s+(\d+) rot:\s+(\d+)", info, re.M)}
        for n, w, h in re.findall(r"^Page\s+(\d+) size:\s+([\d.]+) x ([\d.]+) pts", info, re.M):
            w, h = float(w), float(h)
            if rot.get(int(n), 0) in (90, 270):
                w, h = h, w
            sizes[n] = [w, h]
    except Exception:  # noqa
        pass
    write_json(os.path.join(d, "meta.json"), {"name": os.path.splitext(f.filename)[0], "pages": pages, "sizes": sizes})
    return jsonify({"job": job, "pages": pages, "name": os.path.splitext(f.filename)[0]})


@app.route("/api/job/<job>/page/<int:n>.jpg")
def page_view(job, n):
    d = jdir(job)
    meta = read_json(os.path.join(d, "meta.json"), {})
    if not 1 <= n <= meta.get("pages", 0):
        abort(404)
    view = os.path.join(d, "v-%d.jpg" % n)
    if not os.path.exists(view):
        im = Image.open(page_img(d, n)).convert("RGB")
        im.thumbnail((1100, 1600), Image.LANCZOS)
        im.save(view, "JPEG", quality=85)
    return send_file(view, mimetype="image/jpeg", max_age=3600)


@app.route("/api/job/<job>/crop/<int:n>.jpg")
def crop_view(job, n):
    """Xem trước vùng hình theo bbox (cho màn hình sửa)."""
    d = jdir(job)
    try:
        x1, y1, x2, y2 = [float(v) for v in request.args.get("bbox", "").split(",")]
    except Exception:
        abort(400)
    im = Image.open(page_img(d, n)).convert("RGB")
    W, H = im.size
    box = (max(0, int(x1 / 100 * W - 0.01 * W)), max(0, int(y1 / 100 * H - 0.006 * H)),
           min(W, int(x2 / 100 * W + 0.01 * W)), min(H, int(y2 / 100 * H + 0.006 * H)))
    if box[2] - box[0] < 5 or box[3] - box[1] < 5:
        abort(400)
    out = im.crop(box)
    out.thumbnail((900, 900), Image.LANCZOS)
    tmp = os.path.join(d, "crop-%s.jpg" % uuid.uuid4().hex[:8])
    out.save(tmp, "JPEG", quality=85)
    resp = send_file(tmp, mimetype="image/jpeg")
    resp.call_on_close(lambda: os.path.exists(tmp) and os.remove(tmp))
    return resp


# ----------------------------------------------------------------------------
# Nhận dạng (chạy nền, có tiến trình)
# ----------------------------------------------------------------------------
def _run_recognize(job, d, pages_list, key, model, accurate):
    """Đọc nhiều trang cùng lúc (mặc định 3); mỗi trang xong là lưu ngay.
    Dừng được (nút Dừng) và tự dừng sớm khi hết hạn mức MIỄN PHÍ trong ngày."""
    st, gate = STATUS[job], GATES[job]

    def work(n):
        if st.get("cancel") or st.get("stop_reason"):  # đã bị dừng: bỏ qua trang này, không tính là lỗi
            with LOCK:
                st["done"] += 1
                st["skipped"] += 1
            return
        st["current"] = n
        try:
            if DEMO:  # bản thử: trả dữ liệu mẫu, không gọi Gemini
                sample = read_json(os.path.join(BASE, "tests", "sample_blocks.json"), {})
                blocks = sample.get(str(n)) or [{"type": "paragraph", "text": "Nội dung trang %d: $x^{2}+%d$" % (n, n)}]
                time.sleep(float(os.environ.get("PDFGA_DEMO_DELAY", "0.3")))
            else:
                blocks = recognize.recognize_page(page_img(d, n), key, model, accurate, gate)
            with LOCK:
                pages = load_pages(d)
                pages[n] = blocks
                save_pages(d, pages)
        except recognize.DailyQuotaError as e:
            st["stop_reason"] = str(e)
            st["errors"][str(n)] = str(e)
        except recognize.RecognizeError as e:
            st["errors"][str(n)] = str(e)
        except Exception as e:  # noqa
            st["errors"][str(n)] = "Lỗi khi xử lý trang này: %s" % e
        with LOCK:
            st["done"] += 1

    workers = max(1, min(int(os.environ.get("PDFGA_PARALLEL", "3")), len(pages_list)))
    with ThreadPoolExecutor(max_workers=workers) as ex:
        list(ex.map(work, pages_list))
    st["state"] = "done"
    st["current"] = None


@app.route("/api/job/<job>/recognize", methods=["POST"])
def start_recognize(job):
    d = jdir(job)
    meta = read_json(os.path.join(d, "meta.json"), {})
    body = request.get_json(silent=True) or {}
    total = meta["pages"]
    a = max(1, int(body.get("from") or 1))
    b = min(total, int(body.get("to") or total))
    if a > b:
        return err("Khoảng trang không hợp lệ.")
    key = request.headers.get("X-Gemini-Key") or os.environ.get("GEMINI_API_KEY", "")
    if not key and not DEMO:
        return err("Chưa có khóa Gemini. Hãy mở Cài đặt và dán khóa (xem hướng dẫn trong ứng dụng).")
    model = request.headers.get("X-Gemini-Model") or os.environ.get("GEMINI_MODEL", recognize.DEFAULT_MODEL)
    accurate = request.headers.get("X-Gemini-Mode") == "accurate"
    with LOCK:
        if STATUS.get(job, {}).get("state") == "running":
            return err("Đang xử lý rồi, hãy đợi cho xong.")
        pages_list = list(range(a, b + 1))
        if body.get("only_missing"):  # tiếp tục: chỉ đọc các trang chưa có nội dung
            have = {n for n, blk in load_pages(d).items() if blk}
            pages_list = [n for n in pages_list if n not in have]
            if not pages_list:
                return err("Các trang trong khoảng này đều đã có nội dung.")
        STATUS[job] = {"state": "running", "done": 0, "total": len(pages_list), "current": pages_list[0], "errors": {},
                       "skipped": 0, "range": [a, b], "started": time.time()}
        GATES[job] = Gate()
    threading.Thread(target=_run_recognize, args=(job, d, pages_list, key, model, accurate), daemon=True).start()
    return jsonify({"ok": True, "total": len(pages_list)})


@app.route("/api/job/<job>/cancel", methods=["POST"])
def cancel(job):
    jdir(job)
    if job in STATUS and STATUS[job].get("state") == "running":
        STATUS[job]["cancel"] = True
    return jsonify({"ok": True})


@app.route("/api/job/<job>/status")
def status(job):
    jdir(job)
    st = dict(STATUS.get(job, {"state": "idle", "done": 0, "total": 0, "current": None, "errors": {}, "skipped": 0}))
    g = GATES.get(job)
    st["waiting"] = g.remaining() if g and st.get("state") == "running" else 0  # đang chờ Gemini (giây)
    st["elapsed"] = int(time.time() - st["started"]) if st.get("started") else 0
    return jsonify(st)


# ----------------------------------------------------------------------------
# Lấy / lưu nội dung đã nhận dạng
# ----------------------------------------------------------------------------
@app.route("/api/job/<job>/blocks", methods=["GET", "PUT"])
def blocks(job):
    d = jdir(job)
    if request.method == "GET":
        return jsonify({"pages": {str(k): v for k, v in load_pages(d).items()},
                        "meta": read_json(os.path.join(d, "meta.json"), {})})
    data = request.get_json(silent=True) or {}
    with LOCK:
        # GỘP chứ không ghi đè: chỉ cập nhật các trang được gửi lên, giữ nguyên mọi trang khác.
        # (Trước đây bản lưu cũ chỉ có 1 trang có thể làm mất các trang còn lại.)
        pages = load_pages(d)
        pages.update({int(k): recognize.normalize_blocks(v) for k, v in data.get("pages", {}).items()})
        save_pages(d, pages)
    return jsonify({"ok": True, "pages": sorted(n for n, b in pages.items() if b)})


# ----------------------------------------------------------------------------
# Xuất Word (chạy nền có tiến trình: file dài 100+ trang không bị ngắt giữa chừng)
# ----------------------------------------------------------------------------
def _merge_from_body(d, data):
    if data.get("pages"):  # lưu luôn các trang vừa sửa (gộp, không ghi đè các trang khác)
        with LOCK:
            cur = load_pages(d)
            cur.update({int(k): recognize.normalize_blocks(v) for k, v in data["pages"].items()})
            save_pages(d, cur)


def _build_docx(d, meta, pages, progress=None):
    out = os.path.join(d, "output.docx")
    sizes = {int(k): tuple(v) for k, v in (meta.get("sizes") or {}).items()}
    return Builder(lambda n: page_img(d, n), sizes).build(pages, out, progress), out


def _run_export(job, d, meta, pages):
    ex = EXPORT[job]
    try:
        warnings, _ = _build_docx(d, meta, pages, lambda i, t: ex.update(done=i))
        ex.update(state="done", warnings=warnings[:30], pages=sorted(pages))
    except Exception as e:  # noqa
        ex.update(state="error", error="Không tạo được file Word: %s" % e)


@app.route("/api/job/<job>/docx/start", methods=["POST"])
def export_start(job):
    d = jdir(job)
    meta = read_json(os.path.join(d, "meta.json"), {})
    _merge_from_body(d, request.get_json(silent=True) or {})
    pages = {n: b for n, b in load_pages(d).items() if b}
    if not pages:
        return err("Chưa có nội dung để xuất. Hãy bấm 'Đọc' trước.")
    with LOCK:
        if EXPORT.get(job, {}).get("state") == "running":
            return err("Đang tạo file Word rồi, hãy đợi cho xong.")
        EXPORT[job] = {"state": "running", "done": 0, "total": len(pages)}
    threading.Thread(target=_run_export, args=(job, d, meta, pages), daemon=True).start()
    return jsonify({"ok": True, "total": len(pages)})


@app.route("/api/job/<job>/docx/status")
def export_status(job):
    jdir(job)
    return jsonify(EXPORT.get(job, {"state": "idle"}))


@app.route("/api/job/<job>/docx/file")
def export_file(job):
    d = jdir(job)
    meta = read_json(os.path.join(d, "meta.json"), {})
    out = os.path.join(d, "output.docx")
    if EXPORT.get(job, {}).get("state") != "done" or not os.path.exists(out):
        return err("File Word chưa sẵn sàng.", 409)
    return send_file(out, as_attachment=True, download_name=(meta.get("name") or "tai-lieu") + ".docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/api/job/<job>/docx", methods=["POST"])
def export_docx(job):
    """Bản đồng bộ (giữ để tương thích/kiểm thử): tạo xong mới trả file."""
    d = jdir(job)
    meta = read_json(os.path.join(d, "meta.json"), {})
    _merge_from_body(d, request.get_json(silent=True) or {})
    pages = {n: b for n, b in load_pages(d).items() if b}
    if not pages:
        return err("Chưa có nội dung để xuất. Hãy bấm 'Nhận dạng' trước.")
    try:
        warnings, out = _build_docx(d, meta, pages)
    except Exception as e:  # noqa
        return err("Không tạo được file Word: %s" % e, 500)
    resp = send_file(out, as_attachment=True, download_name=(meta.get("name") or "tai-lieu") + ".docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    resp.headers["X-Warnings"] = quote(json.dumps(warnings[:30], ensure_ascii=False))
    resp.headers["X-Pages"] = ",".join(str(n) for n in sorted(pages))  # các trang có trong file Word
    resp.headers["Access-Control-Expose-Headers"] = "X-Warnings, X-Pages"
    return resp


@app.route("/api/models", methods=["POST"])
def models():
    key = request.headers.get("X-Gemini-Key") or os.environ.get("GEMINI_API_KEY", "")
    if not key:
        return err("Hãy dán khóa Gemini trước.")
    try:
        return jsonify({"models": recognize.list_models(key)})
    except recognize.RecognizeError as e:
        return err(str(e))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), debug=False, threaded=True)
