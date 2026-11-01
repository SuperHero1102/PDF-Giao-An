# PDF Giáo Án

Chuyển PDF toán (đề, giáo án; cả bản scan lẫn bản có chữ) sang Word (.docx):
đúng dấu tiếng Việt, **công thức là công thức Word thật (sửa được)**, hình vẽ cắt nguyên vẹn,
bố cục (khung, cột, thụt lề) giống bản gốc. Có màn hình xem trước để sửa trước khi tải.

## Cách hoạt động
PDF → từng trang thành ảnh → **Gemini** đọc ảnh, chép ra chữ + công thức LaTeX + vị trí hình →
bạn xem/sửa trên web → Pandoc đổi LaTeX thành công thức Word → file .docx.

## 1. Lấy khóa Gemini (miễn phí)
Vào https://aistudio.google.com/app/apikey → “Create API key” → chép khóa.
Dán vào ứng dụng ở **Cài đặt**, hoặc đặt biến môi trường `GEMINI_API_KEY` trên máy chủ (khi đó người dùng khỏi nhập).
Tên mô hình mặc định `gemini-2.5-flash`; nếu Google đổi tên, bấm “Kiểm tra khóa” trong Cài đặt để chọn tên mới.

## 2. Chạy thử trên máy tính
Cách dễ nhất (cần cài Docker Desktop):
```
docker build -t pdf-giao-an .
docker run -p 8000:8000 -e GEMINI_API_KEY=khoa_cua_ban pdf-giao-an
```
Mở http://localhost:8000.

Không dùng Docker: cài Python 3.10+, `poppler-utils`, `pandoc` (≥ 3), rồi
`pip install -r requirements.txt` và `python app.py`.

## 3. Đưa lên mạng để dùng trên điện thoại (ví dụ Render.com)

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/SuperHero1102/PDF-Giao-An)

Cách 1 (Nhanh nhất): Bấm trực tiếp vào nút **Deploy to Render** ở trên hoặc truy cập liên kết:
👉 [https://render.com/deploy?repo=https://github.com/SuperHero1102/PDF-Giao-An](https://render.com/deploy?repo=https://github.com/SuperHero1102/PDF-Giao-An)

Cách 2 (Thủ công):
1. Vào https://render.com → Đăng nhập bằng GitHub của bạn.
2. Bấm New → **Web Service** → chọn kho **SuperHero1102/PDF-Giao-An**.
3. Runtime chọn **Docker** (Render tự thấy Dockerfile). Chọn gói Free.
4. (Tùy chọn) Environment → thêm `GEMINI_API_KEY`.
5. Bấm Deploy, chờ vài phút → có địa chỉ dạng `https://ten-cua-ban.onrender.com`.

Gói miễn phí của nhiều nơi sẽ “ngủ” khi không dùng, lần mở đầu có thể chậm ~1 phút.

## 4. Cài như ứng dụng trên điện thoại
- iPhone (Safari): nút Chia sẻ → “Thêm vào Màn hình chính”.
- Android (Chrome): menu ⋮ → “Cài đặt ứng dụng”.
(Cần địa chỉ https, các nơi đưa lên mạng ở trên đều có sẵn.)

## Giới hạn cần biết
- Cần có mạng và khóa Gemini. Gói miễn phí có giới hạn số lượt/phút và/ngày; hết lượt thì đợi rồi bấm “Thử lại”.
- Đề thi gửi lên Gemini để đọc. Theo điều khoản gói miễn phí của Google, nội dung có thể được dùng cải thiện sản phẩm — đề chưa công bố nên cân nhắc (gói trả phí thì khác; hãy tự kiểm tra điều khoản hiện hành).
- Không công cụ nào đọc đúng 100% ảnh mờ/nghiêng: hãy xem lại các công thức khó (góc có cung, hệ phương trình) ở màn hình xem trước. Công thức đổi không được sẽ được tô vàng trong Word.
- Hình vẽ (hình học, sơ đồ) được giữ dưới dạng ảnh cắt từ trang gốc, không chuyển thành hình vẽ sửa được.
- Chưa có nút “Tải PDF mới”.
- File tạm tự xóa sau 3 giờ không dùng (đổi bằng biến `PDFGA_TTL_HOURS`). Khóa Gemini không lưu trên máy chủ.

## Các file
- `app.py` máy chủ · `recognize.py` gọi Gemini · `docx_builder.py` dựng Word
- `static/index.html` giao diện (+ `manifest.json`, `sw.js` cho PWA)
- `tests/` đề mẫu và các bài kiểm thử
