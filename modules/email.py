import datetime
from dataclasses import dataclass, field
import logging
import os
from pathlib import Path
from typing import Any, Dict, List, Optional
import jinja2
import pandas as pd

logger = logging.getLogger(__name__)


# --- 1. Custom Exceptions ---
class EmailError(Exception):
    pass


class TemplateError(EmailError):
    pass


# --- 2. Data Models ---
@dataclass
class EmailDraft:
    to_email: str
    cc_email: str
    subject: str
    html_body: str
    preview_body: str  # Hiển thị riêng trên giao diện
    attachments: List[Path] = field(default_factory=list)


# --- 3. Class Ghi Log Excel theo Ngày (MỚI) ---
class ExcelEmailLogger:

    def __init__(self, log_dir: str | Path = "logs/excel"):
        self.log_dir = Path(log_dir)
        # Tự động tạo thư mục log nếu chưa tồn tại
        self.log_dir.mkdir(parents=True, exist_ok=True)

    def _get_today_log_filepath(self) -> Path:
        """Tạo đường dẫn file log dựa theo ngày hiện tại: log_YYYY-MM-DD.xlsx"""
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        return self.log_dir / f"log_{today_str}.xlsx"

    def log_event(
        self,
        draft: EmailDraft,
        status: str,
        error_message: str = "",
        oracle_project: str = "",
    ):
        """Ghi nhận 1 sự kiện gửi mail vào file Excel của ngày hôm nay."""
        filepath = self._get_today_log_filepath()
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Cấu trúc dòng dữ liệu log
        new_row = {
            "Thời gian": current_time,
            "Oracle Project": oracle_project,
            "Email Đến (To)": draft.to_email,
            "Email CC": draft.cc_email,
            "Tiêu đề Mail": draft.subject,
            "Trạng thái": status,  # SUCCESS hoặc FAILED
            "Chi tiết / Lỗi": error_message,
            "Số file đính kèm": len(draft.attachments),
        }

        new_df = pd.DataFrame([new_row])

        try:
            # Nếu file log ngày hôm nay đã có -> Ghi nối (Append)
            if filepath.exists():
                with pd.ExcelWriter(
                    filepath,
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                ) as writer:
                    existing_df = pd.read_excel(filepath)
                    start_row = len(existing_df) + 1
                    new_df.to_excel(
                        writer, index=False, header=False, startrow=start_row
                    )
            else:
                # File chưa tồn tại -> Tạo mới kèm Tiêu đề (Header)
                new_df.to_excel(filepath, index=False, engine="openpyxl")

            logger.info(
                f"Đã ghi Excel log [{status}] cho email: {draft.to_email}"
            )

        except Exception as e:
            logger.error(
                f"Lỗi khi ghi file Excel log ({filepath.name}): {e}",
                exc_info=True,
            )


# --- 4. Class EmailBuilder (Đã tích hợp Logger) ---
class EmailBuilder:

    # Khởi tạo instance logger dùng chung
    logger_excel = ExcelEmailLogger()

    @staticmethod
    def _extract_month_year(date_value: Any) -> str:
        try:
            if pd.notna(date_value) and str(date_value).strip() != "":
                if hasattr(date_value, "strftime"):
                    return date_value.strftime("%b %Y")
                else:
                    return pd.to_datetime(date_value).strftime("%b %Y")
        except Exception:
            pass
        return "<Month> <Year>"

    @staticmethod
    def _format_money(val: Any) -> str:
        try:
            if pd.isna(val) or str(val).strip() == "":
                return ""
            if isinstance(val, str):
                val = val.replace(",", "")
            return f"{float(val):,.0f}"
        except Exception:
            return str(val).strip()

    @staticmethod
    def _build_preview_table(
        row_data: Dict[str, Any], invoice_date_str: str
    ) -> str:
        """Tạo bảng kẻ chữ thuần túy CHỈ DÀNH CHO GIAO DIỆN PREVIEW để dễ đọc"""
        oracle = (
            str(row_data.get("Oracle Project", "")).replace(".0", "").strip()[:6]
        )
        project = str(row_data.get("Project code Original", "")).strip()[:15]
        name = str(row_data.get("Project Name", "")).strip()
        name = (name[:17] + "...") if len(name) > 20 else name[:20]
        client = str(row_data.get("Client Code", "")).strip()[:6]
        inv = invoice_date_str[:11]

        amount = EmailBuilder._format_money(
            row_data.get("Amount for labor cost", "")
        )[:12]
        tax = EmailBuilder._format_money(row_data.get("Tax", ""))[:10]
        total = EmailBuilder._format_money(
            row_data.get("Total Invoice Amount", "")
        )[:12]

        w_ora = 7
        w_prj = 16
        w_name = 21
        w_cli = 7
        w_inv = 11
        w_amt = 12
        w_tax = 10
        w_tot = 12

        header = f"| {'Oracle':<{w_ora}}| {'Project #':<{w_prj}}| {'Project Name':<{w_name}}| {'Client':<{w_cli}}| {'Inv Date':<{w_inv}}| {'Amount':>{w_amt}} | {'Tax':>{w_tax}} | {'Total':>{w_tot}} |"
        row1 = f"| {oracle:<{w_ora}}| {project:<{w_prj}}| {name:<{w_name}}| {client:<{w_cli}}| {inv:<{w_inv}}| {amount:>{w_amt}} | {tax:>{w_tax}} | {total:>{w_tot}} |"

        border = f"+{'-'*(len(header)-2)}+"
        return f"{border}\n{header}\n{border}\n{row1}\n{border}"

    @classmethod
    def generate_subject(cls, row_data: Dict[str, Any]) -> str:
        client_code = str(row_data.get("Client Code", "")).strip()
        oracle_project = (
            str(row_data.get("Oracle Project", "")).replace(".0", "").strip()
        )
        month_year = cls._extract_month_year(row_data.get("Invoice Date", ""))

        return f"{client_code} - {oracle_project} - Approve invoice in {month_year}"

    @staticmethod
    def render_body(
        template_path: Path,
        row_data: Dict[str, Any],
        month_year: str,
        attachments: List[Path],
    ) -> str:
        try:
            env = jinja2.Environment(
                loader=jinja2.FileSystemLoader(template_path.parent),
                autoescape=jinja2.select_autoescape(["html", "xml"]),
            )
            template = env.get_template(template_path.name)

            raw_inv = row_data.get("Invoice Date", "")
            inv_str = ""
            if pd.notna(raw_inv) and str(raw_inv).strip() != "":
                if hasattr(raw_inv, "strftime"):
                    inv_str = raw_inv.strftime("%d-%b-%y")
                else:
                    try:
                        inv_str = pd.to_datetime(raw_inv).strftime("%d-%b-%y")
                    except Exception:
                        inv_str = str(raw_inv)

            context = {
                "PM": row_data.get("PM", ""),
                "Month_Year": month_year,
                "Oracle_Project": str(row_data.get("Oracle Project", ""))
                .replace(".0", "")
                .strip(),
                "Project_code_Original": str(
                    row_data.get("Project code Original", "")
                ).strip(),
                "Project_Name": str(row_data.get("Project Name", "")).strip(),
                "Client_Code": str(row_data.get("Client Code", "")).strip(),
                "Invoice_Date": inv_str,
                "Amount_for_labor_cost": EmailBuilder._format_money(
                    row_data.get("Amount for labor cost", "")
                ),
                "Tax": EmailBuilder._format_money(row_data.get("Tax", "")),
                "Total_Invoice_Amount": EmailBuilder._format_money(
                    row_data.get("Total Invoice Amount", "")
                ),
                "Attachment_Names": [a.name for a in attachments],
            }
            return template.render(**context)
        except Exception as e:
            raise TemplateError(
                f"Lỗi render template '{template_path.name}': {e}"
            )

    @classmethod
    def build_draft(
        cls,
        template_path: Path | str,
        row_data: Dict[str, Any],
        attachments: List[Path],
    ) -> EmailDraft:
        t_path = Path(template_path)
        if not t_path.exists():
            raise TemplateError(
                f"Không tìm thấy template HTML: {t_path}"
            )

        subject = cls.generate_subject(row_data)
        month_year = cls._extract_month_year(row_data.get("Invoice Date", ""))

        html_body = cls.render_body(t_path, row_data, month_year, attachments)

        raw_inv = row_data.get("Invoice Date", "")
        inv_str = ""
        if pd.notna(raw_inv) and str(raw_inv).strip() != "":
            if hasattr(raw_inv, "strftime"):
                inv_str = raw_inv.strftime("%d-%b-%y")
            else:
                try:
                    inv_str = pd.to_datetime(raw_inv).strftime("%d-%b-%y")
                except Exception:
                    inv_str = str(raw_inv)

        pm_name = row_data.get("PM", "")
        preview = f"Dear anh/chị @{pm_name},\n\n"
        preview += f"Could you please review and approve for the draft invoice which got Acceptance in {month_year} before we issue official as below:\n\n"
        preview += cls._build_preview_table(row_data, inv_str)
        preview += "\n\nAttachments:\n" + "\n".join(
            [f"✓ {a.name}" for a in attachments]
        )
        preview += "\n\nThank you.\nThanks and best regards,"

        return EmailDraft(
            to_email=str(row_data.get("Email to", "")).strip(),
            cc_email=str(row_data.get("Email cc", "")).strip(),
            subject=subject,
            html_body=html_body,
            preview_body=preview,
            attachments=attachments,
        )

    # --- HÀM THÊM MỚI: TÍCH HỢP LOGGING ---
    @classmethod
    def log_email(
        cls,
        draft: EmailDraft,
        status: str,
        error_message: str = "",
        oracle_project: str = "",
    ):
        """Hàm ghi log hỗ trợ gọi nhanh từ bên ngoài."""
        cls.logger_excel.log_event(
            draft=draft,
            status=status,
            error_message=error_message,
            oracle_project=oracle_project,
        )