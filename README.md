Markdown
# 🧾 Invoice Approval Tool (Công cụ Phê duyệt Hóa đơn Tự động)

Chào mừng bạn đến với **Invoice Approval Tool**! Đây là trợ lý tự động giúp bạn quản lý, xử lý báo cáo Excel và tạo dự thảo phê duyệt hóa đơn qua Outlook chỉ với vài cú click chuột, giúp tiết kiệm hàng giờ làm việc thủ công.

Giao diện trực quan, dễ sử dụng và **không yêu cầu bất kỳ kiến thức lập trình nào** để cài đặt.

---

## ✨ Tính năng nổi bật
- 📂 **Đọc dữ liệu Excel:** Tự động quét và trích xuất dữ liệu hóa đơn (từ các file như `InvsummaryJun.xlsx`, `draftinv.xlsx`).
- 📧 **Kết nối Outlook tự động:** Tự động tạo bản nháp email phê duyệt dựa trên biểu mẫu có sẵn trong thư mục `templates/`.
- 📎 **Quản lý tệp đính kèm:** Sắp xếp gọn gàng toàn bộ chứng từ vào thư mục `attachments_root/`.
- 💻 **Giao diện hiện đại:** Thao tác nhấn nút dễ dàng, có thanh tiến trình (Progress Bar) không lo bị đơ máy.

---

## 🛠️ Hướng dẫn cài đặt từ A đến Z (Cho máy mới tinh)

Vui lòng làm theo đúng thứ tự 3 bước hướng dẫn chi tiết dưới đây để tránh gặp lỗi không tương thích phần mềm:

### Bước 1: Cài đặt phần mềm nền (Python)
Công cụ này chạy trên nền tảng Python. Nếu máy bạn chưa có, hãy cài đặt chính xác như sau:
1. Truy cập trang chủ: [Tải Python tại đây](https://www.python.org/downloads/) và bấm nút **Download Python** (phiên bản mới nhất).
2. Mở file vừa tải về lên. 
3. **⚠️ CỰC KỲ QUAN TRỌNG:** Ở ngay cửa sổ đầu tiên hiện ra, nhìn xuống dưới cùng và **ĐÁNH DẤU TÍCH vào ô "Add python.exe to PATH"**.
4. Sau đó bấm nút **Install Now** ở phía trên và đợi chạy xong.

### Bước 2: Tải công cụ này về máy
- **Cách 1 (Dễ nhất):** Nhìn lên góc trên bên phải trang web này, bấm vào nút màu xanh lá cây **`<> Code`** -> Chọn **`Download ZIP`**. Sau khi tải xong, nhấp chuột phải vào file và chọn **Extract Here** (Giải nén) ra một thư mục dễ tìm (ví dụ: Desktop hoặc ổ D).
- **Cách 2 (Cho ai dùng Git):** Mở CMD và gõ: `git clone https://github.com/Thinhuthao/invoice-approval-tool.git`

### Bước 3: Khởi tạo môi trường ảo & Cài đặt thư viện (Tránh lỗi xung đột)
Để tránh lỗi phiên bản thư viện không tương thích với máy tính của bạn, chúng ta sẽ tạo một môi trường chạy riêng cho ứng dụng:

1. Mở thư mục bạn vừa giải nén ở Bước 2 (thư mục có chứa file `main.py`).
2. Nhấp chuột vào thanh địa chỉ (thanh dài ở trên cùng hiển thị đường dẫn thư mục), xóa hết chữ trong đó, gõ `cmd` rồi nhấn **Enter**. Một cửa sổ màu đen sẽ hiện ra.
3. Copy toàn bộ đoạn lệnh dưới đây, dán vào cửa sổ đen đó (nhấp chuột phải để dán trên CMD) rồi nhấn **Enter**:
   ```cmd
   python -m venv venv && venv\Scripts\activate && python -m pip install --upgrade pip && pip install -r requirements.txt
(Quá trình này có thể mất 1-2 phút tùy tốc độ mạng, hãy đợi cho đến khi cửa sổ đen hiện lại dòng chữ chờ lệnh thông thường).

## 🚀 Hướng dẫn sử dụng hàng ngày
Sau khi đã cài đặt xong ở Bước 3, mỗi lần muốn mở phần mềm lên để làm việc, bạn chỉ cần thực hiện 2 bước cực kỳ nhanh sau:

1. Vào thư mục phần mềm, gõ `cmd` lên thanh địa chỉ trên cùng để mở cửa sổ đen.
2. Gõ (hoặc dán) dòng lệnh sau rồi nhấn **Enter**:
   ```cmd
   venv\Scripts\activate && python main.py
Giao diện phần mềm sẽ hiện lên. Bạn chỉ cần chọn file Excel, thư mục đính kèm và bấm nút chạy!
