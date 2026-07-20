import sys
import logging
import ttkbootstrap as ttk
from modules.gui import InvoiceApprovalApp

def setup_logging() -> None:
    """
    Thiết lập cấu hình logging cơ bản cho toàn bộ ứng dụng.
    Log sẽ được xuất ra console với định dạng dễ đọc.
    """
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - [%(module)s] - %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)]
    )

def main() -> None:
    """
    Điểm khởi chạy chính của ứng dụng Invoice Approval.
    Khởi tạo giao diện người dùng và vòng lặp sự kiện chính.
    """
    # Khởi tạo logging
    setup_logging()
    logger = logging.getLogger(__name__)
    logger.info("Starting Invoice Approval Application...")

    try:
        # Khởi tạo cửa sổ chính sử dụng ttkbootstrap để có giao diện Desktop hiện đại, sạch sẽ
        root = ttk.Window(
            title="Invoice Approval",
            themename="cosmo",  # Theme 'cosmo' mang lại cảm giác ứng dụng doanh nghiệp chuẩn mực
            resizable=(False, False)
        )
        
        # Gắn giao diện ứng dụng vào cửa sổ gốc
        app = InvoiceApprovalApp(root)
        
        # Bắt đầu vòng lặp sự kiện giao diện (Event Loop)
        root.mainloop()
        
    except Exception as e:
        logger.critical(f"Application crashed during initialization: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()