import logging
import threading
import queue
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from typing import List, Dict, Any

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame

from modules.excel import ExcelProcessor, ExcelError
from modules.attachment import AttachmentScanner
from modules.email import EmailBuilder, EmailDraft, EmailError
from modules.outlook import OutlookService, OutlookError

logger = logging.getLogger(__name__)

class InvoiceApprovalApp:
    def __init__(self, root: ttk.Window):
        self.root = root
        self.root.title("Invoice Approval")
        self.root.geometry("1050x750")
        
        # Thiết lập kích thước tối thiểu để khi thu nhỏ cửa sổ không bị vỡ giao diện
        self.root.minsize(800, 600)
        
        self.excel_path = tk.StringVar()
        self.attachment_dir = tk.StringVar()
        self.template_path = tk.StringVar()
        
        self.drafts: List[Dict[str, Any]] = []
        self.msg_queue = queue.Queue()
        
        self.main_frame = ttk.Frame(self.root, padding=20)
        self.main_frame.pack(fill=BOTH, expand=YES)
        
        self.root.after(100, self._process_queue)
        self.build_screen_config()

    def clear_screen(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    def _process_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                action = msg.get("action")
                
                if action == "error":
                    messagebox.showerror("Error", msg["message"])
                    if msg.get("return_to_config"):
                        self.build_screen_config()
                elif action == "review_ready":
                    self.build_screen_review()
                elif action == "progress_update":
                    if hasattr(self, 'progress_var'):
                        self.progress_var.set(msg["value"])
                        self.status_lbl.config(text=f"Sending {msg['value']} / {msg['total']}")
                        self.receiver_lbl.config(text=msg["receiver"])
                elif action == "send_complete":
                    messagebox.showinfo("Success", f"Successfully sent {msg['total']} emails!")
                    self.build_screen_config()
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._process_queue)

    def build_screen_config(self):
        self.clear_screen()
        
        ttk.Label(self.main_frame, text="Invoice Approval", font=("Helvetica", 24, "bold")).pack(anchor=W, pady=(0, 10))
        ttk.Label(self.main_frame, text="System configuration:", font=("Helvetica", 12, "bold")).pack(anchor=W, pady=(0, 20))
        
        # Đóng gói Bottom Frame trước để đảm bảo nó luôn dính dưới cùng khi Resize
        bottom_frame = ttk.Frame(self.main_frame)
        bottom_frame.pack(fill=X, side=BOTTOM, pady=20)
        
        # --- THÊM NHẬN DIỆN THƯƠNG HIỆU ---
        ttk.Label(bottom_frame, text="#designbyThaoVan", font=("Helvetica", 10, "italic"), foreground="gray").pack(side=LEFT, anchor=S)
        ttk.Button(bottom_frame, text="Next", command=self._validate_and_process, bootstyle=OUTLINE).pack(side=RIGHT)

        form_frame = ttk.Frame(self.main_frame)
        form_frame.pack(fill=X, pady=10)
        form_frame.columnconfigure(1, weight=1)
        
        ttk.Label(form_frame, text="File Data Excel").grid(row=0, column=0, sticky=W, pady=10, padx=(0, 10))
        ttk.Entry(form_frame, textvariable=self.excel_path).grid(row=0, column=1, sticky=EW, pady=10)
        ttk.Button(form_frame, text="Browse", command=self._browse_excel, bootstyle=OUTLINE).grid(row=0, column=2, padx=(10, 0))
        
        ttk.Label(form_frame, text="Folder Attachment").grid(row=1, column=0, sticky=W, pady=10, padx=(0, 10))
        ttk.Entry(form_frame, textvariable=self.attachment_dir).grid(row=1, column=1, sticky=EW, pady=10)
        ttk.Button(form_frame, text="Browse", command=self._browse_dir, bootstyle=OUTLINE).grid(row=1, column=2, padx=(10, 0))
        
        ttk.Label(form_frame, text="Template Email").grid(row=2, column=0, sticky=W, pady=10, padx=(0, 10))
        ttk.Entry(form_frame, textvariable=self.template_path).grid(row=2, column=1, sticky=EW, pady=10)
        ttk.Button(form_frame, text="Browse", command=self._browse_template, bootstyle=OUTLINE).grid(row=2, column=2, padx=(10, 0))

    def _browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path: self.excel_path.set(path)

    def _browse_dir(self):
        path = filedialog.askdirectory()
        if path: self.attachment_dir.set(path)

    def _browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("HTML files", "*.html")])
        if path: self.template_path.set(path)

    def _validate_and_process(self):
        ep, ad, tp = self.excel_path.get(), self.attachment_dir.get(), self.template_path.get()
        
        if not all([ep, ad, tp]):
            messagebox.showwarning("Validation Error", "All fields are required.")
            return
            
        if not Path(ep).exists() or not ep.endswith('.xlsx'):
            messagebox.showwarning("Validation Error", "Invalid Excel file.")
            return
            
        if not Path(ad).exists() or not Path(ad).is_dir():
            messagebox.showwarning("Validation Error", "Invalid Attachment Directory.")
            return
            
        if not Path(tp).exists() or not tp.endswith('.html'):
            messagebox.showwarning("Validation Error", "Invalid Template file.")
            return

        self.clear_screen()
        ttk.Label(self.main_frame, text="Processing Excel Data...", font=("Helvetica", 14)).pack(expand=YES)
        threading.Thread(target=self._process_data_thread, daemon=True).start()

    def _process_data_thread(self):
        try:
            records = ExcelProcessor.read_invoices(self.excel_path.get())
            self.drafts = []
            
            for row in records:
                attachments = AttachmentScanner.get_attachments_for_invoice(self.attachment_dir.get(), row.get("Folder name", ""))
                draft = EmailBuilder.build_draft(self.template_path.get(), row, attachments)
                
                self.drafts.append({
                    "draft": draft,
                    "approved": False,
                    "row_data": row
                })
            
            self.msg_queue.put({"action": "review_ready"})
            
        except (ExcelError, EmailError) as e:
            self.msg_queue.put({"action": "error", "message": str(e), "return_to_config": True})
        except Exception as e:
            logger.error("Unexpected error", exc_info=True)
            self.msg_queue.put({"action": "error", "message": "An unexpected error occurred. Check logs.", "return_to_config": True})

    def build_screen_review(self):
        self.clear_screen()
        
        ttk.Label(self.main_frame, text="Invoice Approval", font=("Helvetica", 24, "bold")).pack(anchor=W, pady=(0, 10))
        
        # Khung dưới cùng chứa chữ ký và các nút bấm
        bottom_frame = ttk.Frame(self.main_frame)
        bottom_frame.pack(fill=X, side=BOTTOM, pady=10)
        
        # --- THÊM NHẬN DIỆN THƯƠNG HIỆU ---
        ttk.Label(bottom_frame, text="#designbyThaoVan", font=("Helvetica", 10, "italic"), foreground="gray").pack(side=LEFT, anchor=S)
        
        # Khung nhóm 2 nút bấm ở bên phải
        btn_frame = ttk.Frame(bottom_frame)
        btn_frame.pack(side=RIGHT)
        ttk.Button(btn_frame, text="Back", command=self.build_screen_config, bootstyle=OUTLINE).pack(side=LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Send Selected Mails", command=self._confirm_send, bootstyle=SUCCESS).pack(side=LEFT)
        
        # Vùng cuộn (ScrolledFrame) để hiển thị danh sách Email, tự động co giãn
        scroll_frame = ScrolledFrame(self.main_frame, autohide=True)
        scroll_frame.pack(fill=BOTH, expand=YES, pady=10)
        
        for idx, item in enumerate(self.drafts):
            self._create_preview_card(scroll_frame, item, idx)

    def _create_preview_card(self, parent, item, idx):
        draft: EmailDraft = item["draft"]
        has_attachments = len(draft.attachments) > 0
        
        card_style = DANGER if not has_attachments else DEFAULT
        card = ttk.Frame(parent, borderwidth=1, relief=SOLID, padding=10, bootstyle=card_style)
        card.pack(fill=X, pady=5, padx=5)
        
        ttk.Label(card, text=f"Subject: {draft.subject}", font=("Helvetica", 10, "bold"), bootstyle=card_style+"-inverse" if not has_attachments else DEFAULT).pack(anchor=W)
        
        style_obj = ttk.Style()
        bg_color = style_obj.colors.danger if not has_attachments else style_obj.colors.bg
        fg_color = "white" if not has_attachments else style_obj.colors.fg
        
        line_count = len(draft.preview_body.split('\n')) + 1
        
        text_frame = ttk.Frame(card, bootstyle=card_style)
        text_frame.pack(fill=X, pady=5)
        
        body_text = tk.Text(
            text_frame, 
            font=("Consolas", 10), 
            bg=bg_color, 
            fg=fg_color, 
            wrap=tk.NONE,
            height=min(line_count, 18),
            borderwidth=0, 
            highlightthickness=0,
            relief=tk.FLAT
        )
        
        body_text.insert(tk.END, draft.preview_body)
        body_text.config(state=tk.DISABLED) 
        
        x_scroll = ttk.Scrollbar(text_frame, orient=tk.HORIZONTAL, command=body_text.xview)
        body_text.config(xscrollcommand=x_scroll.set)
        
        body_text.pack(fill=X, expand=True)
        x_scroll.pack(fill=X)
        
        if has_attachments:
            attach_text = "Status: Ready"
        else:
            attach_text = "Status: Missing Attachment"
            
        ttk.Label(card, text=attach_text, bootstyle=card_style+"-inverse" if not has_attachments else DEFAULT).pack(anchor=W, pady=(5,0))
        
        action_frame = ttk.Frame(card, bootstyle=card_style)
        action_frame.pack(fill=X, pady=(10,0))
        
        if not has_attachments:
            ttk.Button(action_frame, text="Add Attachment", bootstyle=WARNING, 
                       command=lambda: self._manual_add_attachment(item, card)).pack(side=LEFT)
            
        btn_text = "Approved" if item["approved"] else "Approve"
        btn_style = SUCCESS if item["approved"] else SECONDARY
        
        approve_btn = ttk.Button(action_frame, text=btn_text, bootstyle=btn_style, state=NORMAL if has_attachments else DISABLED,
                                 command=lambda i=item, b=None: self._toggle_approve(i, approve_btn))
        approve_btn.pack(side=RIGHT)

    def _manual_add_attachment(self, item, card_widget):
        files = filedialog.askopenfilenames(title="Select Attachments")
        if files:
            draft: EmailDraft = item["draft"]
            draft.attachments.extend([Path(f) for f in files])
            self.build_screen_review()

    def _toggle_approve(self, item, btn):
        item["approved"] = not item["approved"]
        if item["approved"]:
            btn.config(text="Approved", bootstyle=SUCCESS)
        else:
            btn.config(text="Approve", bootstyle=SECONDARY)

    def _confirm_send(self):
        approved_count = sum(1 for item in self.drafts if item["approved"])
        if approved_count == 0:
            messagebox.showinfo("Info", "No emails selected for sending.")
            return
            
        if messagebox.askyesno("Confirm", f"You are about to send {approved_count} emails.\nContinue?"):
            self.build_screen_processing(approved_count)

    def build_screen_processing(self, total_emails: int):
        self.clear_screen()
        
        ttk.Label(self.main_frame, text="Invoice Approval", font=("Helvetica", 24, "bold")).pack(anchor=W, pady=(0, 10))
        ttk.Label(self.main_frame, text="Processing:", font=("Helvetica", 12)).pack(anchor=W, pady=(20, 5))
        
        self.progress_var = tk.IntVar()
        self.progress_bar = ttk.Progressbar(self.main_frame, variable=self.progress_var, maximum=total_emails, bootstyle=PRIMARY)
        self.progress_bar.pack(fill=X, pady=10)
        
        self.status_lbl = ttk.Label(self.main_frame, text=f"Sending 0 / {total_emails}")
        self.status_lbl.pack(anchor=W)
        
        self.receiver_lbl = ttk.Label(self.main_frame, text="Initializing...", foreground="gray")
        self.receiver_lbl.pack(anchor=W)
        
        bottom_frame = ttk.Frame(self.main_frame)
        bottom_frame.pack(fill=X, side=BOTTOM, pady=10)
        
        # --- THÊM NHẬN DIỆN THƯƠNG HIỆU CẢ Ở MÀN HÌNH ĐANG GỬI ---
        ttk.Label(bottom_frame, text="#designbyThaoVan", font=("Helvetica", 10, "italic"), foreground="gray").pack(side=LEFT, anchor=S)
        
        threading.Thread(target=self._send_emails_thread, args=(total_emails,), daemon=True).start()

    def _send_emails_thread(self, total: int):
        sent_count = 0
        try:
            for item in self.drafts:
                if item["approved"]:
                    draft: EmailDraft = item["draft"]
                    self.msg_queue.put({
                        "action": "progress_update",
                        "value": sent_count,
                        "total": total,
                        "receiver": draft.to_email
                    })
                    
                    OutlookService.send_email(draft)
                    sent_count += 1
                    
                    self.msg_queue.put({
                        "action": "progress_update",
                        "value": sent_count,
                        "total": total,
                        "receiver": draft.to_email
                    })
                    
            self.msg_queue.put({"action": "send_complete", "total": sent_count})
            
        except OutlookError as e:
            self.msg_queue.put({"action": "error", "message": str(e), "return_to_config": False})
        except Exception as e:
            logger.error("Error during email sending", exc_info=True)
            self.msg_queue.put({"action": "error", "message": "An unexpected error occurred during sending.", "return_to_config": False})