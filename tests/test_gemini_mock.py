# -*- coding: utf-8 -*-
"""Thử phần gọi Gemini bằng máy chủ GIẢ (không dùng khóa thật): kiểm tra đọc kết quả, thử lại, dịch lỗi."""
import json, threading, time, sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from flask import Flask, request, jsonify
import recognize

mock = Flask("mock"); calls = {"n": 0}
SAMPLE = json.load(open(os.path.join(os.path.dirname(__file__), "sample_blocks.json"), encoding="utf-8"))

@mock.route("/v1beta/models/<m>:generateContent", methods=["POST"])
def gen(m):
    calls["n"] += 1
    key = request.headers.get("x-goog-api-key")
    if key == "sai": return jsonify({"error": {"message": "API key not valid"}}), 400
    if m == "khong-co": return jsonify({"error": {"message": "not found"}}), 404
    if key == "het-luot": return jsonify({"error": {"message": "quota"}}), 429
    if key == "loi-roi-ok" and calls["n"] < 3: return jsonify({"error": {}}), 503
    body = request.get_json()
    assert body["contents"][0]["parts"][1]["inline_data"]["mime_type"] == "image/jpeg"
    # trả JSON bọc ```json và bbox thang 0-1000 để thử bộ chuẩn hóa
    blocks = json.loads(json.dumps(SAMPLE["2"]))
    blocks[2]["bbox"] = [237, 222, 773, 404]
    txt = "```json\n" + json.dumps({"blocks": blocks + [{"type": "bay-ba"}]}, ensure_ascii=False) + "\n```"
    return jsonify({"candidates": [{"content": {"parts": [{"text": txt}]}}]})

threading.Thread(target=lambda: mock.run(port=9001), daemon=True).start(); time.sleep(1)
recognize.API = "http://127.0.0.1:9001/v1beta"
time.sleep(0)
orig_sleep = time.sleep; recognize.time.sleep = lambda s: None  # bỏ chờ khi thử lại

img = os.path.join(os.path.dirname(__file__), "..", "tests", "_p2.png")
os.system("pdftoppm -r 60 -f 2 -l 2 -png -singlefile %s %s" % (os.path.join(os.path.dirname(__file__), "de_toan_9.pdf"), img[:-4]))

b = recognize.recognize_page(img, "tot")
print("OK  đọc được", len(b), "khối; bbox chuẩn hóa:", b[2]["bbox"], "| khối lạ bị bỏ:", all(x["type"] != "bay-ba" for x in b))
assert b[2]["bbox"] == [23.7, 22.2, 77.3, 40.4]
for key, model, expect in [("sai", "m", "Khóa Gemini không đúng"), ("tot", "khong-co", "Không tìm thấy mô hình"),
                           ("het-luot", "m", "hết lượt")]:
    try: recognize.recognize_page(img, key, model); print("FAIL không báo lỗi", key)
    except recognize.RecognizeError as e:
        assert expect in str(e), str(e); print("OK  lỗi", key, "->", e)
calls["n"] = 0
print("OK  thử lại sau 503:", len(recognize.recognize_page(img, "loi-roi-ok")), "khối, số lần gọi", calls["n"])
os.remove(img)
