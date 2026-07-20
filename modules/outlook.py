import logging
from pathlib import Path
import win32com.client
from modules.email import EmailDraft

logger = logging.getLogger(__name__)

class OutlookError(Exception):
    """Ngoại lệ cơ sở cho các lỗi liên quan đến giao tiếp với Microsoft Outlook."""
    pass

class OutlookService:
    """
    Lớp chịu trách nhiệm điều khiển Microsoft Outlook Desktop COM.
    Chỉ thực hiện việc tạo và gửi email, không chứa logic nghiệp vụ khác.
    """

    @staticmethod
    def send_email(draft: EmailDraft) -> None:
        """
        Khởi tạo đối tượng Outlook MailItem, đính kèm file và gửi.
        
        Args:
            draft (EmailDraft): Đối tượng chứa toàn bộ thông tin email (To, CC, Subject, Body, Attachments).
            
        Raises:
            OutlookError: Nếu không thể kết nối tới Outlook hoặc có lỗi trong quá trình gửi.
        """
        try:
            # Khởi tạo kết nối COM tới Outlook
            # Sử dụng Dispatch thay vì DispatchEx để sử dụng instance Outlook đang chạy nếu có
            outlook = win32com.client.Dispatch("Outlook.Application")
            
            # 0 đại diện cho olMailItem (tạo một email mới)
            mail = outlook.CreateItem(0)
            
            # Thiết lập các trường cơ bản
            mail.To = draft.to_email
            mail.CC = draft.cc_email
            mail.Subject = draft.subject
            mail.HTMLBody = draft.html_body
            
            # Xử lý tệp đính kèm
            for attachment_path in draft.attachments:
                if attachment_path.exists():
                    # win32com yêu cầu đường dẫn tuyệt đối dưới dạng chuỗi (absolute string)
                    abs_path = str(attachment_path.resolve())
                    mail.Attachments.Add(abs_path)
                else:
                    logger.warning(f"Bỏ qua tệp đính kèm do không tồn tại trên hệ thống: {attachment_path}")
            
            # Thực thi lệnh gửi
            mail.Send()
            logger.info(f"Đã gửi thành công email tới '{draft.to_email}' (Subject: {draft.subject})")
            
        except Exception as e:
            error_msg = f"Lỗi giao tiếp với Outlook khi gửi email cho '{draft.to_email}': {str(e)}"
            logger.error(error_msg, exc_info=True)
            raise OutlookError(error_msg)