# -*- coding: utf-8 -*-
"""Nhận dạng một trang (ảnh) thành danh sách khối bằng Gemini (khóa lấy từ Google AI Studio)."""
import base64
import io
import json
import re
import socket
import time

import requests
import urllib3.util.connection as urllib3_conn
from PIL import Image

try:
    urllib3_conn.allowed_gai_family = lambda: socket.AF_INET
except Exception:
    pass

API = "https://generativelanguage.googleapis.com/v1beta"
DEFAULT_MODEL = "gemini-3.5-flash"
ALLOWED = {"paragraph", "columns", "boxed", "figure", "text_figure", "formula"}

PROMPT = """Bạn là công cụ chép lại đề/giáo án TOÁN tiếng Việt từ ảnh trang sang dữ liệu có cấu trúc.
Chép lại CHÍNH XÁC những gì nhìn thấy trên trang. TUYỆT ĐỐI không tóm tắt, không giải bài,
không thêm mô tả, chú thích hay lời dẫn nào KHÔNG có trong ảnh. Giữ nguyên dấu tiếng Việt,
số thứ tự (1), a), b)...), dấu câu và thứ tự đọc từ trên xuống dưới.

Trả về DUY NHẤT một JSON: {"blocks":[ ... ]}. Mỗi phần tử trong blocks là một trong các loại:

1. {"type":"paragraph","text":"...","align":"left|center|justify|right","indent":0,"first_line":false}
   - Một đoạn văn / một dòng. indent là cm thụt lề trái (0, 1, 1.5 ...), first_line=true nếu dòng đầu thụt vào.
   - Một đoạn văn nhiều dòng trên ảnh chỉ là MỘT khối (đừng tách theo dòng vật lý).
2. {"type":"columns","cells":["...","..."],"aligns":["center","center"],"widths":[45,55]}
   - Dùng khi trên ảnh có 2 (hoặc 3) cột đặt cạnh nhau trên cùng một hàng: phần đầu đề (trường | tên bài kiểm tra),
     hoặc các ý cạnh nhau như "1) (x+3)(2x-1)=0   2) phân số...". Trong một ô, xuống dòng bằng \\n.
   - Nếu một dòng trong ô nằm trong khung chữ nhật (vd ĐỀ CHÍNH THỨC) thì viết dòng đó bắt đầu bằng "[box] ".
3. {"type":"boxed","text":"..."}  - chữ nằm trong khung viền đứng riêng.
4. {"type":"figure","bbox":[x1,y1,x2,y2]}  - hình vẽ đứng riêng.
5. {"type":"text_figure","text":"...","bbox":[x1,y1,x2,y2],"side":"right"}
   - Đoạn chữ nằm cạnh một hình (hình bên phải thì side="right", bên trái thì "left").
6. {"type":"formula","latex":"..."}  - công thức đứng riêng một dòng giữa trang.

Quy tắc về bbox: x1,y1 là góc trên-trái, x2,y2 là góc dưới-phải, tính theo PHẦN TRĂM chiều rộng/cao trang (0-100).
Khung phải ôm sát TOÀN BỘ hình gồm cả nhãn điểm (A, B, C, H), số đo góc, độ dài, tiêu đề trong hình
(vd "Tường bao"), nhưng KHÔNG được dính chữ của đoạn văn bên cạnh/bên dưới.

Quy tắc về chữ và công thức trong text/cells:
- **đậm** viết giữa hai dấu ** ; *nghiêng* viết giữa hai dấu *.  Tiêu đề như "Bài 1. (1,5 điểm)" thường in đậm.
- MỌI biểu thức toán (như (x+3)(2x-1)=0, y=ax+b, 15^{0}, tọa độ (-2;1), tên góc \\widehat{ABC}, AE.EK=HE.EB, phân số, lũy thừa)
  phải viết bằng LaTeX đặt trong $...$ ngay giữa câu, đúng vị trí trên trang.
- Số thường và số kèm đơn vị trong câu văn (800 tấn thóc, 15%, 108m, 3m/giây, 120m, 1300 lít) là CHỮ THƯỜNG:
  để NGOÀI $...$. Tên điểm/đoạn thẳng đứng một mình trong câu (A, B, AC, CH) cũng để ngoài $...$.
- Quan trọng: bạn trả JSON nên MỌI dấu gạch chéo ngược trong LaTeX phải được viết gấp đôi
  (ví dụ viết \\\\frac, \\\\text, \\\\widehat) để sau khi đọc JSON còn đúng một dấu gạch chéo ngược.
- Phân số \\frac{a}{b}; mũ x^{2}; hệ phương trình \\begin{cases} dòng1 \\\\ dòng2 \\end{cases};
  góc có cung trên đầu \\widehat{ABC}; tam giác \\triangle ABC; độ: ^{0} hoặc ^{\\circ} đúng như ảnh;
  giữ nguyên dấu âm/trừ, dấu ngoặc, thứ tự số hạng, \\sin, \\cos, \\tan, \\sqrt{...}, \\le, \\ge, \\ne...
- Chữ thường tiếng Việt (câu hỏi, tên điểm A, B, C khi không có dấu toán) để ngoài $...$.
- Dấu gạch "----Hết----" chép như một paragraph căn giữa.

Chỉ trả JSON, không giải thích, không dùng ```.
"""


class RecognizeError(Exception):
    """Lỗi đã được dịch sang tiếng Việt để hiện cho người dùng."""


_MATH = re.compile(r"(\$[^$]+\$)")


def fix_escapes(text):
    r"""Khi mô hình viết \text / \frac / \beta / \right / \neq trong JSON mà không gấp đôi dấu \,
    JSON hiểu thành tab / form-feed / backspace / xuống dòng và làm mất chữ (vd '120extm').
    Hàm này khôi phục lại dấu \ trong công thức."""
    if not text:
        return text
    fix = {"\t": "\\t", "\x0c": "\\f", "\x08": "\\b", "\r": "\\r", "\n": "\\n"}
    out = []
    for part in _MATH.split(text):
        if part.startswith("$") and part.endswith("$") and len(part) > 1:
            for k, v in fix.items():
                part = part.replace(k, v)
        else:
            part = part.replace("\x0c", "").replace("\x08", "").replace("\t", " ")
        out.append(part)
    return "".join(out)


def _page_to_jpeg_b64(path, max_side=2000):
    im = Image.open(path).convert("RGB")
    if max(im.size) > max_side:
        k = max_side / max(im.size)
        im = im.resize((int(im.width * k), int(im.height * k)), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=92)
    return base64.b64encode(buf.getvalue()).decode()


def _translate_http_error(status, text):
    t = (text or "").lower()
    if status in (400, 401, 403) and ("api key" in t or "api_key" in t or "permission" in t):
        return "Khóa Gemini không đúng hoặc chưa được phép dùng. Hãy kiểm tra lại khóa trong phần Cài đặt."
    if status == 404:
        return "Không tìm thấy mô hình Gemini này. Hãy đổi tên mô hình trong Cài đặt (bấm 'Kiểm tra khóa' để xem danh sách)."
    if status == 429:
        return "Đã dùng hết lượt miễn phí của Gemini trong lúc này. Hãy đợi 1-2 phút rồi bấm thử lại."
    if status >= 500:
        return "Máy chủ Gemini đang bận. Hãy thử lại sau ít phút."
    return "Gemini báo lỗi %s." % status


def list_models(key):
    r = requests.get(API + "/models", headers={"x-goog-api-key": key},
                     params={"pageSize": 200}, timeout=30)
    if r.status_code != 200:
        raise RecognizeError(_translate_http_error(r.status_code, r.text))
    names = []
    for m in r.json().get("models", []):
        if "generateContent" in m.get("supportedGenerationMethods", []):
            names.append(m["name"].replace("models/", ""))
    return sorted(names)


def _normalize_bbox(bb):
    try:
        v = [float(x) for x in bb][:4]
    except Exception:
        return None
    if len(v) != 4:
        return None
    if max(v) > 100:  # có thể mô hình trả thang 0-1000
        v = [x / 10 for x in v]
    elif max(v) <= 1.0:  # hoặc thang 0-1
        v = [x * 100 for x in v]
    x1, y1, x2, y2 = v
    if x2 < x1:
        x1, x2 = x2, x1
    if y2 < y1:
        y1, y2 = y2, y1
    clamp = lambda z: max(0.0, min(100.0, z))  # noqa
    return [round(clamp(x1), 2), round(clamp(y1), 2), round(clamp(x2), 2), round(clamp(y2), 2)]


def normalize_blocks(data):
    blocks = data.get("blocks") if isinstance(data, dict) else data
    out = []
    for b in blocks or []:
        if not isinstance(b, dict) or b.get("type") not in ALLOWED:
            continue
        t = b["type"]
        nb = {"type": t}
        if t in ("paragraph", "boxed", "text_figure"):
            nb["text"] = fix_escapes(str(b.get("text", "")))
        if t == "paragraph":
            nb["align"] = b.get("align") if b.get("align") in ("left", "center", "right", "justify") else "left"
            nb["indent"] = float(b.get("indent") or 0)
            nb["first_line"] = bool(b.get("first_line"))
        if t == "columns":
            nb["cells"] = [fix_escapes(str(c)) for c in b.get("cells", [])]
            nb["aligns"] = [a if a in ("left", "center", "right", "justify") else "left"
                            for a in (b.get("aligns") or [])]
            if b.get("widths"):
                try:
                    nb["widths"] = [float(w) for w in b["widths"]]
                except Exception:
                    pass
        if t in ("figure", "text_figure"):
            bb = _normalize_bbox(b.get("bbox", []))
            if bb is None:
                continue
            nb["bbox"] = bb
            if t == "text_figure":
                nb["side"] = "left" if b.get("side") == "left" else "right"
        if t == "formula":
            nb["latex"] = fix_escapes("$" + str(b.get("latex", "")).strip("$ ") + "$")[1:-1]
        out.append(nb)
    return out


def recognize_page(image_path, key, model=DEFAULT_MODEL, accurate=False):
    body = {
        "contents": [{"parts": [
            {"text": PROMPT},
            {"inline_data": {"mime_type": "image/jpeg", "data": _page_to_jpeg_b64(image_path)}},
        ]}],
        "generationConfig": {"temperature": 0, "responseMimeType": "application/json"},
    }
    # Chế độ nhanh: tắt "suy nghĩ" dài của mô hình Flash (chép chữ không cần suy luận nhiều).
    # Nếu mô hình không nhận cấu hình này thì tự bỏ đi và thử lại bên dưới.
    fast_cfg = (not accurate) and "flash" in model.lower()
    if fast_cfg:
        body["generationConfig"]["thinkingConfig"] = {"thinkingBudget": 0}
    url = "%s/models/%s:generateContent" % (API, model)
    last = None
    for attempt in range(4):
        try:
            r = requests.post(url, json=body, headers={"x-goog-api-key": key}, timeout=180)
        except requests.RequestException:
            last = "Không kết nối được tới Gemini. Hãy kiểm tra mạng."
            time.sleep(3 * (attempt + 1))
            continue
        if r.status_code == 200:
            try:
                parts = r.json()["candidates"][0]["content"]["parts"]
                txt = "".join(p.get("text", "") for p in parts).strip()
                txt = re.sub(r"^```(?:json)?|```$", "", txt, flags=re.M).strip()
                blocks = normalize_blocks(json.loads(txt))
                if not blocks:
                    raise ValueError("rỗng")
                return blocks
            except Exception:
                last = "Gemini trả kết quả không đọc được cho trang này. Hãy thử lại."
                time.sleep(2)
                continue
        if r.status_code == 400 and fast_cfg and "thinking" in (r.text or "").lower():
            body["generationConfig"].pop("thinkingConfig", None)
            fast_cfg = False
            continue
        last = _translate_http_error(r.status_code, r.text)
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(8 * (attempt + 1))
            continue
        break
    raise RecognizeError(last or "Lỗi không rõ khi gọi Gemini.")
