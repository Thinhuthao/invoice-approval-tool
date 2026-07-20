import logging
import datetime
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, Any, List
import jinja2
import pandas as pd

logger = logging.getLogger(__name__)

class EmailError(Exception): pass
class TemplateError(EmailError): pass

@dataclass
class EmailDraft:
    to_email: str
    cc_email: str
    subject: str
    html_body: str
    preview_body: str  # THÊM TRƯỜNG NÀY ĐỂ HIỂN THỊ RIÊNG TRÊN GIAO DIỆN
    attachments: List[Path] = field(default_factory=list)


class EmailBuilder:

    @staticmethod
    def _extract_month_year(date_value: Any) -> str:
        try:
            if pd.notna(date_value) and str(date_value).strip() != "":
                if hasattr(date_value, "strftime"): return date_value.strftime("%b %Y")
                else: return pd.to_datetime(date_value).strftime("%b %Y")
        except Exception:
            pass
        return "<Month> <Year>"

    @staticmethod
    def _format_money(val: Any) -> str:
        try:
            if pd.isna(val) or str(val).strip() == "": return ""
            if isinstance(val, str): val = val.replace(",", "")
            return f"{float(val):,.0f}"
        except:
            return str(val).strip()

    @staticmethod
    def _build_preview_table(row_data: Dict[str, Any], invoice_date_str: str) -> str:
        """Tạo bảng kẻ chữ thuần túy CHỈ DÀNH CHO GIAO DIỆN PREVIEW để dễ đọc"""
        oracle = str(row_data.get("Oracle Project", "")).replace(".0", "").strip()[:6]
        project = str(row_data.get("Project code Original", "")).strip()[:15]
        name = str(row_data.get("Project Name", "")).strip()
        name = (name[:17] + "...") if len(name) > 20 else name[:20]
        client = str(row_data.get("Client Code", "")).strip()[:6]
        inv = invoice_date_str[:11]
        
        amount = EmailBuilder._format_money(row_data.get("Amount for labor cost", ""))[:12]
        tax = EmailBuilder._format_money(row_data.get("Tax", ""))[:10]
        total = EmailBuilder._format_money(row_data.get("Total Invoice Amount", ""))[:12]
        
        w_ora = 7; w_prj = 16; w_name = 21; w_cli = 7; w_inv = 11; w_amt = 12; w_tax = 10; w_tot = 12
        
        header = f"| {'Oracle':<{w_ora}}| {'Project #':<{w_prj}}| {'Project Name':<{w_name}}| {'Client':<{w_cli}}| {'Inv Date':<{w_inv}}| {'Amount':>{w_amt}} | {'Tax':>{w_tax}} | {'Total':>{w_tot}} |"
        row1   = f"| {oracle:<{w_ora}}| {project:<{w_prj}}| {name:<{w_name}}| {client:<{w_cli}}| {inv:<{w_inv}}| {amount:>{w_amt}} | {tax:>{w_tax}} | {total:>{w_tot}} |"
        
        border = f"+{'-'*(len(header)-2)}+"
        return f"{border}\n{header}\n{border}\n{row1}\n{border}"

    @classmethod
    def generate_subject(cls, row_data: Dict[str, Any]) -> str:
        # Cập nhật Subject theo mẫu: HTCL - 67646 - Approve invoice in Nov 2025
        client_code = str(row_data.get("Client Code", "")).strip()
        oracle_project = str(row_data.get("Oracle Project", "")).replace(".0", "").strip()
        month_year = cls._extract_month_year(row_data.get("Invoice Date", ""))
        
        return f"{client_code} - {oracle_project} - Approve invoice in {month_year}"

    @staticmethod
    def render_body(template_path: Path, row_data: Dict[str, Any], month_year: str, attachments: List[Path]) -> str:
        try:
            env = jinja2.Environment(
                loader=jinja2.FileSystemLoader(template_path.parent),
                autoescape=jinja2.select_autoescape(['html', 'xml'])
            )
            template = env.get_template(template_path.name)
            
            raw_inv = row_data.get("Invoice Date", "")
            inv_str = ""
            if pd.notna(raw_inv) and str(raw_inv).strip() != "":
                if hasattr(raw_inv, "strftime"): inv_str = raw_inv.strftime("%d-%b-%y")
                else:
                    try: inv_str = pd.to_datetime(raw_inv).strftime("%d-%b-%y")
                    except: inv_str = str(raw_inv)
            
            context = {
                "PM": row_data.get("PM", ""),
                "Month_Year": month_year,
                "Oracle_Project": str(row_data.get("Oracle Project", "")).replace(".0", "").strip(),
                "Project_code_Original": str(row_data.get("Project code Original", "")).strip(),
                "Project_Name": str(row_data.get("Project Name", "")).strip(),
                "Client_Code": str(row_data.get("Client Code", "")).strip(),
                "Invoice_Date": inv_str,
                "Amount_for_labor_cost": EmailBuilder._format_money(row_data.get("Amount for labor cost", "")),
                "Tax": EmailBuilder._format_money(row_data.get("Tax", "")),
                "Total_Invoice_Amount": EmailBuilder._format_money(row_data.get("Total Invoice Amount", "")),
                "Attachment_Names": [a.name for a in attachments]
            }
            return template.render(**context)
        except Exception as e:
            raise TemplateError(f"Lỗi render template '{template_path.name}': {e}")

    @classmethod
    def build_draft(cls, template_path: Path | str, row_data: Dict[str, Any], attachments: List[Path]) -> EmailDraft:
        t_path = Path(template_path)
        if not t_path.exists(): raise TemplateError(f"Không tìm thấy template HTML: {t_path}")
            
        subject = cls.generate_subject(row_data)
        month_year = cls._extract_month_year(row_data.get("Invoice Date", ""))
        
        # 1. Tạo HTML chuẩn để gửi qua Outlook
        html_body = cls.render_body(t_path, row_data, month_year, attachments)
        
        # Lấy ngày tháng để làm bảng preview
        raw_inv = row_data.get("Invoice Date", "")
        inv_str = ""
        if pd.notna(raw_inv) and str(raw_inv).strip() != "":
            if hasattr(raw_inv, "strftime"): inv_str = raw_inv.strftime("%d-%b-%y")
            else:
                try: inv_str = pd.to_datetime(raw_inv).strftime("%d-%b-%y")
                except: inv_str = str(raw_inv)
                
        # 2. Tạo Văn bản Preview gọn gàng để hiển thị trên Giao diện
        pm_name = row_data.get("PM", "")
        preview = f"Dear anh/chị @{pm_name},\n\n"
        preview += f"Could you please review and approve for the draft invoice which got Acceptance in {month_year} before we issue official as below:\n\n"
        preview += cls._build_preview_table(row_data, inv_str)
        preview += "\n\nAttachments:\n" + "\n".join([f"✓ {a.name}" for a in attachments])
        preview += "\n\nThank you.\nThanks and best regards,"
        
        return EmailDraft(
            to_email=str(row_data.get("Email to", "")).strip(),
            cc_email=str(row_data.get("Email cc", "")).strip(),
            subject=subject,
            html_body=html_body,
            preview_body=preview, # TRUYỀN VÀO EMAIL DRAFT
            attachments=attachments
        )