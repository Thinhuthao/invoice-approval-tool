import logging
from pathlib import Path
from typing import List

logger = logging.getLogger(__name__)

class AttachmentScanner:
    """
    Lớp chịu trách nhiệm định vị và quét các tệp đính kèm cho từng hóa đơn.
    Không xử lý giao diện hay logic gửi email.
    """

    @staticmethod
    def get_attachments_for_invoice(base_dir: Path | str, folder_name: str) -> List[Path]:
        """
        Tìm kiếm và trả về danh sách toàn bộ các tệp nằm trong thư mục con của hóa đơn.
        
        Args:
            base_dir: Đường dẫn đến thư mục gốc chứa tất cả tệp đính kèm (từ Cấu hình hệ thống).
            folder_name: Tên thư mục con tương ứng với dòng dữ liệu hiện tại (từ Excel).
            
        Returns:
            List[Path]: Danh sách các đường dẫn tới tệp (files). Trả về danh sách rỗng
                        nếu thư mục không tồn tại hoặc không có tệp nào.
        """
        try:
            base_path = Path(base_dir)
            
            # Làm sạch chuỗi folder_name để tránh lỗi do khoảng trắng thừa
            clean_folder_name = str(folder_name).strip()
            
            if not clean_folder_name:
                logger.warning("Cột 'Folder name' bị trống trong dữ liệu Excel.")
                return []

            target_folder = base_path / clean_folder_name
            
            # Kiểm tra sự tồn tại của thư mục
            if not target_folder.exists() or not target_folder.is_dir():
                logger.warning(f"Không tìm thấy thư mục đính kèm: {target_folder}")
                return []
            
            # Quét và lấy tất cả các tệp (loại trừ thư mục con nếu có)
            attachments = [f for f in target_folder.iterdir() if f.is_file()]
            
            if not attachments:
                logger.warning(f"Thư mục đính kèm tồn tại nhưng trống: {target_folder}")
            else:
                logger.info(f"Đã tìm thấy {len(attachments)} tệp đính kèm trong thư mục: {target_folder.name}")
                
            return attachments
            
        except Exception as e:
            logger.error(f"Lỗi hệ thống khi quét thư mục đính kèm '{folder_name}': {e}", exc_info=True)
            # Trả về danh sách rỗng thay vì làm gián đoạn toàn bộ quá trình
            return []