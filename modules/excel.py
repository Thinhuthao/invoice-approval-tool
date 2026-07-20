import logging
from pathlib import Path
from typing import List, Dict, Any
import pandas as pd

logger = logging.getLogger(__name__)

class ExcelError(Exception):
    """Lỗi cơ sở cho các ngoại lệ liên quan đến xử lý Excel."""
    pass

class ExcelValidationError(ExcelError):
    """Ngoại lệ được ném ra khi file Excel thiếu các cột bắt buộc."""
    pass

class ExcelFileError(ExcelError):
    """Ngoại lệ được ném ra khi không thể đọc file Excel (thiếu file, bị khóa, v.v.)."""
    pass

class ExcelProcessor:
    """
    Lớp chịu trách nhiệm xử lý dữ liệu từ file Excel.
    Đọc dữ liệu, xác thực cột và chuyển đổi thành đối tượng Python.
    """
    
    # Danh sách các cột bắt buộc theo đặc tả của hệ thống
    REQUIRED_COLUMNS = [
        "Oracle Project",
        "Project code Original",
        "Project Name",
        "Client Code",
        "Invoice#",
        "Invoice Date",
        "Amount for labor cost",
        "Tax",
        "Total Invoice Amount",
        "PO",
        "Note",
        "Quotation Date",
        "Folder name",
        "PM",
        "Email to",
        "Email cc"
    ]

    @classmethod
    def read_invoices(cls, file_path: Path | str) -> List[Dict[str, Any]]:
        """
        Đọc file Excel, kiểm tra cột và trả về danh sách các bản ghi (dòng).
        
        Args:
            file_path: Đường dẫn đến file Excel (.xlsx)
            
        Returns:
            List[Dict[str, Any]]: Danh sách các dictionary, mỗi dict đại diện cho một dòng.
            
        Raises:
            ExcelFileError: Nếu không thể đọc file hoặc file không tồn tại.
            ExcelValidationError: Nếu file thiếu một hoặc nhiều cột bắt buộc.
        """
        path = Path(file_path)
        logger.info(f"Bắt đầu đọc dữ liệu từ Excel: {path}")

        if not path.exists():
            error_msg = f"Không tìm thấy file Excel tại: {path}"
            logger.error(error_msg)
            raise ExcelFileError(error_msg)
        
        try:
            # Đọc file Excel sử dụng engine openpyxl
            df = pd.read_excel(path, engine="openpyxl")
        except PermissionError:
            error_msg = f"Không có quyền truy cập file. Vui lòng đóng Excel nếu file đang được mở: {path}"
            logger.error(error_msg)
            raise ExcelFileError(error_msg)
        except Exception as e:
            error_msg = f"Lỗi không xác định khi đọc file Excel: {str(e)}"
            logger.error(error_msg)
            raise ExcelFileError(error_msg)

        # Kiểm tra sự tồn tại của các cột bắt buộc
        missing_cols = [col for col in cls.REQUIRED_COLUMNS if col not in df.columns]
        if missing_cols:
            error_msg = f"File Excel bị thiếu các cột bắt buộc sau: {', '.join(missing_cols)}"
            logger.error(error_msg)
            raise ExcelValidationError(error_msg)

        # Làm sạch dữ liệu: Điền các ô trống (NaN/NaT) bằng chuỗi rỗng
        # Điều này tránh việc hiển thị chữ 'nan' trong nội dung email sau này.
        df = df.fillna("")

        # Chuyển đổi DataFrame thành danh sách các dictionary (Mỗi dòng là 1 dict)
        records = df.to_dict(orient="records")
        logger.info(f"Đã đọc và xác thực thành công {len(records)} dòng dữ liệu từ Excel.")
        
        return records