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
1. Tạo tài khoản GitHub, tạo kho mới và tải toàn bộ thư mục này lên.
2. Vào https://render.com → New → **Web Service** → chọn kho GitHub vừa tạo.
3. Runtime chọn **Docker** (Render tự thấy Dockerfile). Chọn gói Free.
4. (Tùy chọn) Environment → thêm `GEMINI_API_KEY`.
5. Bấm Deploy, chờ vài phút → có địa chỉ dạng `https://ten-cua-ban.onrender.com`.

Gói miễn phí của nhiều nơi sẽ “ngủ” khi không dùng, lần mở đầu có thể chậm ~1 phút.

## 4. Cài như ứng dụng trên điện thoại
- iPhone (Safari): nút Chia sẻ → “Thêm vào Màn hình chính”.
- Android (Chrome): menu ⋮ → “Cài đặt ứng dụng”.
(Cần địa chỉ https, các nơi đưa lên mạng ở trên đều có sẵn.)

## Cập nhật bản mới lên Render
Trên GitHub: vào kho của bạn → Add file → Upload files → kéo thả các tệp mới (ghi đè tệp cũ) → Commit.
Render tự build lại (vài phút). Khóa và cài đặt trên điện thoại vẫn giữ nguyên.

## File nhiều trang (100+ trang)
- Bấm **“Đọc toàn bộ file”**: máy đọc lần lượt, 3 trang cùng lúc, trang nào xong được lưu ngay; có thanh tiến trình, thời gian còn lại và nút **Dừng** (giữ các trang đã đọc).
- Gemini báo quá tải (429) → máy tự chờ đúng thời gian Gemini yêu cầu rồi đọc tiếp, không mất trang.
- Hết lượt **miễn phí trong ngày** → máy dừng gọn, giữ các trang đã đọc; bấm **“Đọc các trang còn thiếu”** vào ngày hôm sau (hạn mức đặt lại ≈ nửa đêm giờ Thái Bình Dương ≈ 14–15h Việt Nam) hoặc bật thanh toán cho khóa API.
- Từ ~12 trang, màn hình sửa hiện **từng trang một** (có nút ◀ ▶ và danh sách trang) để không bị nặng/đơ trên điện thoại.
- Tạo file Word chạy nền có tiến trình, không bị ngắt dù file dài.
- Gói Render miễn phí ngủ khi không dùng và xóa dữ liệu tạm: app sao lưu các trang đã đọc ngay trong trình duyệt; nếu mất phiên, chọn lại đúng file PDF cũ và bấm **“Khôi phục”**. Khi đọc file dài hãy để màn hình sáng (app tự xin giữ sáng màn hình).
- Đã thử bằng Gemini giả lập: 120 trang đọc + xuất Word đủ 120 trang (RAM ≈ 90MB), 429 tạm thời, hết hạn mức ngày, Dừng và đọc tiếp đều hoạt động.
- Word giữ đúng khổ giấy/hướng từng trang của PDF gốc (A4, A3, trang ngang…).
- Chế độ đọc (Cài đặt): “Nhanh” (mặc định) hoặc “Kỹ hơn” cho đề khó đọc.

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

## Lưu PDF
Nút “Lưu thành PDF” mở một trang in riêng (thẻ mới) gồm MỌI trang đã đọc, công thức hiển thị bằng KaTeX; bấm “In / Lưu thành PDF”
rồi chọn “Lưu dưới dạng PDF”. Cách này nhẹ, miễn phí, nhưng bố cục có thể hơi khác file Word.
Sửa đè trực tiếp lên PDF scan gốc chưa được hỗ trợ (chữ trong bản scan chỉ là hình).

## Lưu không làm mất trang
- Khi sửa, app chỉ gửi các trang vừa sửa; máy chủ GỘP vào bản đã có, không ghi đè các trang khác.
- Sau khi tải Word, app báo rõ file gồm những trang nào và trang nào chưa có nội dung.

## Công thức với MathType
File Word dùng công thức gốc của Word (OMML). Nếu cần chuyển sang MathType (Windows, MathType 7): Word → thẻ MathType →
Convert Equations → chọn “Word 2007 and later (OMML) equations” → đích “MathType equations (OLE objects)”.
