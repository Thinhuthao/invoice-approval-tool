# 🧾 Invoice Approval Tool

Chào mừng bạn đến với **Invoice Approval Tool**! Đây là một công cụ tự động hóa giúp bạn xử lý, quản lý và phê duyệt hóa đơn một cách nhanh chóng, chính xác và giảm thiểu tối đa các thao tác thủ công.

Công cụ này được thiết kế để bất kỳ ai – kể cả những người chưa từng có kinh nghiệm lập trình – cũng có thể dễ dàng cài đặt và sử dụng.
---
## ✨ Tính năng nổi bật
- 📂 **Xử lý dữ liệu từ Excel:** Tự động trích xuất và đọc dữ liệu từ các báo cáo hóa đơn (ví dụ: `InvsummaryJun.xlsx`).
- 📝 **Tạo và quản lý bản nháp:** Xử lý linh hoạt các hóa đơn chưa hoàn thiện (`draftinv.xlsx`).
- 📎 **Quản lý tệp đính kèm tự động:** Sắp xếp gọn gàng các tài liệu liên quan vào thư mục `attachments_root`.
- 🚀 **Tiết kiệm thời gian:** Tự động hóa các quy trình phê duyệt lặp đi lặp lại.
---
## 🛠️ Hướng dẫn cài đặt (Dành cho người mới bắt đầu)

Đừng lo lắng nếu bạn chưa rành về máy tính, hãy làm theo đúng thứ tự 3 bước dưới đây:

### Bước 1: Cài đặt phần mềm nền (Python)
Công cụ này cần một phần mềm tên là Python để hoạt động. Nếu máy tính của bạn chưa có, hãy cài đặt theo cách sau:
1. Truy cập trang chủ tải phần mềm: [Tải Python tại đây](https://www.python.org/downloads/)
2. Bấm nút tải về phiên bản mới nhất.
3. **⚠️ QUAN TRỌNG:** Khi mở file vừa tải về để cài đặt, hãy nhìn xuống góc dưới cùng của cửa sổ và **ĐÁNH DẤU TÍCH vào ô "Add Python to PATH"** (hoặc "Add python.exe to PATH"). Sau đó mới bấm **Install Now**.

### Bước 2: Tải công cụ này về máy tính
- **Cách đơn giản nhất:** Kéo lên đầu trang web này, bấm vào nút màu xanh lá cây **`<> Code`** -> Chọn **`Download ZIP`**. Sau khi máy tính tải xong, hãy nhấp chuột phải vào file ZIP và chọn "Extract Here" (Giải nén) ra một thư mục dễ tìm.
- **Cách dành cho người dùng Git:** Mở Terminal / CMD và gõ lệnh: `git clone https://github.com/Thinhuthao/invoice-approval-tool.git`

### Bước 3: Khởi tạo và cài đặt tiện ích đi kèm
1. Mở thư mục bạn vừa giải nén ở Bước 2 (thư mục có chứa file `main.py`).
2. Nhấp chuột vào thanh địa chỉ (thanh hiển thị đường dẫn thư mục ở trên cùng), xóa hết chữ trong đó, gõ `cmd` rồi nhấn **Enter**. Một cửa sổ màu đen sẽ hiện ra.
3. Gõ dòng lệnh sau vào cửa sổ đen đó và nhấn **Enter**:
   ```cmd
   pip install -r requirements.txt

🚀 Hướng dẫn sử dụng
Sau khi hoàn tất cài đặt ở Bước 3, mỗi lần muốn sử dụng công cụ, bạn chỉ cần làm như sau:
Mở lại thư mục chứa công cụ, gõ cmd lên thanh địa chỉ để mở cửa sổ đen.
Gõ lệnh sau và nhấn Enter để phần mềm bắt đầu chạy:
python main.py
📌 Một số lưu ý quan trọng:
Đảm bảo các file dữ liệu (như InvsummaryJun.xlsx hoặc draftinv.xlsx) đã được bạn cập nhật đúng thông tin trước khi chạy phần mềm.
Các biểu mẫu (nếu có) được đặt trong thư mục templates/.
File kết quả hoặc tệp đính kèm trong quá trình chạy sẽ được đưa vào thư mục attachments_root/.
