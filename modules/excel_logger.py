import os
from datetime import datetime
import pandas as pd


class ExcelEmailLogger:

    def __init__(self, log_dir="logs/excel"):
        self.log_dir = log_dir
        # Tự động tạo thư mục logs/excel nếu chưa tồn tại
        os.makedirs(self.log_dir, exist_ok=True)

    def _get_today_log_filepath(self) -> str:
        """Lấy đường dẫn file log dựa trên ngày hiện tại (Format: log_YYYY-MM-DD.xlsx)"""
        today_str = datetime.now().strftime("%Y-%m-%d")
        filename = f"log_{today_str}.xlsx"
        return os.path.join(self.log_dir, filename)

    def log_email_event(
        self,
        recipient_email: str,
        subject: str,
        status: str,
        error_message: str = "",
    ):
        """Ghi nhận 1 sự kiện gửi email vào file Excel của ngày hôm nay.

        :param recipient_email: Email người nhận
        :param subject: Tiêu đề email đã gửi
        :param status: Trạng thái ("SUCCESS" hoặc "FAILED")
        :param error_message: Thông báo lỗi (nếu có)
        """
        filepath = self._get_today_log_filepath()
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Tạo bản ghi mới
        new_row = {
            "Thời gian gửi": current_time,
            "Email người nhận": recipient_email,
            "Tiêu đề Mail": subject,
            "Trạng thái": status,
            "Ghi chú / Lỗi": error_message,
        }

        new_df = pd.DataFrame([new_row])

        # Nếu file log của ngày hôm nay đã tồn tại -> Ghi nối (Append)
        if os.path.exists(filepath):
            try:
                with pd.ExcelWriter(
                    filepath, mode="a", engine="openpyxl", if_sheet_exists="overlay"
                ) as writer:
                    # Lấy số dòng hiện tại để ghi tiếp vào bên dưới
                    existing_df = pd.read_excel(filepath)
                    start_row = len(existing_df) + 1

                    new_df.to_excel(
                        writer,
                        index=False,
                        header=False,
                        startrow=start_row,
                    )
            except Exception as e:
                print(f"[ERROR] Không thể ghi nối vào file log Excel: {e}")
        else:
            # File chưa tồn tại -> Tạo mới file và ghi cả dòng Tiêu đề (Header)
            try:
                new_df.to_excel(filepath, index=False, engine="openpyxl")
            except Exception as e:
                print(f"[ERROR] Không thể tạo file log Excel mới: {e}")