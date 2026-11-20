import datetime
import os
import re
import sys
from tkinter import filedialog, messagebox
from typing import Any, Callable, Dict, List, Optional, Tuple
import customtkinter as ctk

from core.config import (
    backup_database,
    combine_receipt_no,
    export_full_zip_backup,
    format_next_receipt_number,
    format_receipt_date_thai,
    format_thai_date_short,
    normalize_template_style,
    restore_database,
    save_config,
    split_receipt_prefix_and_number,
    thai_baht_text,
)
from core.database import DatabaseManager, ReceiptNumberError
from core.models import UserSession
from core.security import SecurityManager
from exporters.docx_exporter import DocxReceiptExporter
from exporters.excel_exporter import ExcelFinanceExporter
from exporters.pdf_exporter import PdfReceiptExporter
from exporters.webhook_exporter import FinanceWebhookDispatcher
from ui.components import (
    DataTable,
    SearchableDropdown,
    SectionHeader,
    StatCard,
    ask_receipt_number_resolution,
    choose_export_folder,
    format_export_folder_button_text,
)
from ui.theme import AppTheme
from ui.notifications import highlight_missing_field

TEMPLATE_STYLE_OPTIONS = [
    "CLASSIC - ต้นฉบับทางการ (ตามแบบ EX 100%)",
    "MODERN - รูปแบบทันสมัย (อัปเกรด)",
]


def _style_label_from_code(raw_code: Any) -> str:
    if normalize_template_style(raw_code) == "MODERN":
        return TEMPLATE_STYLE_OPTIONS[1]
    return TEMPLATE_STYLE_OPTIONS[0]


def open_document_file(file_path: str, parent: Optional[Any] = None) -> bool:
    """
    Opens the exported document file (PDF, Word, Excel) with the OS default program
    so the cashier can immediately preview and print it.
    If opening fails, displays an informative error dialog instead of failing silently.
    """
    if not file_path or not os.path.exists(file_path):
        messagebox.showerror(
            "ไม่พบไฟล์เอกสาร",
            f"ระบบไม่พบไฟล์ที่ต้องการเปิดเพื่อสั่งพิมพ์:\n{file_path}\n\nกรุณาตรวจสอบว่าไฟล์ถูกลบหรือย้ายหรือไม่",
            parent=parent,
        )
        return False

    try:
        if sys.platform == "win32":
            os.startfile(os.path.abspath(file_path))
            return True
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", os.path.abspath(file_path)])
            return True
        else:
            import subprocess
            subprocess.Popen(["xdg-open", os.path.abspath(file_path)])
            return True
    except OSError as oe:
        messagebox.showwarning(
            "ไม่สามารถเปิดไฟล์อัตโนมัติได้",
            f"ไฟล์เอกสารถูกสร้างและบันทึกเรียบร้อยที่:\n{file_path}\n\n"
            f"⚠️ แต่ระบบไม่สามารถเปิดไฟล์ขึ้นมาอัตโนมัติได้:\n{oe}\n\n"
            f"(คำแนะนำ: เครื่องคอมพิวเตอร์นี้อาจยังไม่ได้ตั้งโปรแกรมเริ่มต้นสำหรับเปิดไฟล์ประเภทนี้ เช่น Adobe Acrobat, Microsoft Edge หรือ Word)\n"
            f"คุณสามารถเปิดไฟล์ดังกล่าวได้ด้วยตนเองจากโฟลเดอร์ที่บันทึกไว้ครับ",
            parent=parent,
        )
        return False
    except Exception as e:
        messagebox.showerror(
            "เกิดข้อผิดพลาดในการเปิดไฟล์",
            f"ไม่สามารถเปิดไฟล์เอกสารเพื่อสั่งพิมพ์ได้:\n{e}\n\nไฟล์อยู่ที่:\n{file_path}",
            parent=parent,
        )
        return False


def export_document_auto(
    save_path: str,
    payload: Any,
    college_name: str,
    metadata: Dict[str, Any],
    is_batch: bool = False,
) -> str:
    """Exports receipt/report to PDF (.pdf) by default, or Word (.docx) if selected with comprehensive error handling."""
    if not save_path:
        raise ValueError("ไม่ได้ระบุตำแหน่งไฟล์สำหรับบันทึกเอกสาร (save_path is empty)")

    # Ensure target directory exists
    target_dir = os.path.dirname(os.path.abspath(save_path))
    os.makedirs(target_dir, exist_ok=True)

    # Check if target file is locked by another process (e.g. Adobe Acrobat) before rendering
    if os.path.exists(save_path):
        try:
            with open(save_path, "r+b"):
                pass
        except PermissionError:
            err_msg = (
                f"ไม่สามารถบันทึกทับไฟล์เดิมได้:\n{save_path}\n\n"
                f"⚠️ สาเหตุ: ไฟล์นี้กำลังเปิดใช้งานอยู่ในโปรแกรมอื่น (เช่น Adobe Acrobat, Edge หรือ Word)\n"
                f"กรุณาปิดโปรแกรมดังกล่าวก่อน แล้วกดลองใหม่อีกครั้งครับ"
            )
            messagebox.showerror("ไฟล์ถูกเปิดค้างอยู่ (Permission Denied)", err_msg)
            raise PermissionError(err_msg)
        except Exception as e:
            err_msg = f"ไม่สามารถเข้าถึงไฟล์เดิมเพื่อเขียนทับได้:\n{save_path}\n({e})"
            messagebox.showerror("เกิดข้อผิดพลาดในการเข้าถึงไฟล์", err_msg)
            raise

    try:
        if save_path.lower().endswith(".docx"):
            exporter = DocxReceiptExporter(save_path, college_name=college_name)
        else:
            exporter = PdfReceiptExporter(save_path, college_name=college_name)

        if is_batch:
            res = exporter.export_batch(payload, metadata=metadata)
        else:
            res = exporter.export(payload, metadata=metadata)

        if not os.path.exists(save_path) or os.path.getsize(save_path) == 0:
            err_msg = f"คำสั่งสร้างเอกสารทำงานเสร็จสิ้น แต่ไม่พบไฟล์บนดิสก์หรือไฟล์มีขนาด 0 ไบต์ที่:\n{save_path}"
            messagebox.showerror("ไม่พบไฟล์เอกสารหลังสร้าง", err_msg)
            raise IOError(err_msg)

        return res
    except PermissionError:
        raise
    except Exception as e:
        err_msg = f"เกิดข้อผิดพลาดขณะสร้างเอกสาร ({os.path.basename(save_path)}):\n{e}"
        messagebox.showerror("เกิดข้อผิดพลาดในการสร้างเอกสาร", err_msg)
        raise RuntimeError(err_msg) from e


def make_dropdown(
    parent,
    values: List[str],
    variable: ctk.StringVar,
    width: int = 160,
    height: int = 38,
    font_size: int = 13,
    bold: bool = True,
    command: Optional[Callable] = None,
    allow_custom: bool = True,
    on_search: Optional[Callable[[str], Optional[List[str]]]] = None,
) -> SearchableDropdown:
    """Unified, clean, ultra-lightweight corporate searchable dropdown."""
    weight = "bold" if bold else "normal"
    return SearchableDropdown(
        parent,
        values=values,
        variable=variable,
        width=width,
        height=height,
        corner_radius=AppTheme.RADIUS_INPUT,
        fg_color=AppTheme.COLOR_DROPDOWN_FG,
        button_color=AppTheme.COLOR_DROPDOWN_BTN,
        button_hover_color=AppTheme.COLOR_DROPDOWN_HOVER,
        text_color=AppTheme.COLOR_DROPDOWN_TEXT,
        dropdown_fg_color=AppTheme.COLOR_DROPDOWN_MENU,
        dropdown_hover_color=AppTheme.COLOR_SECONDARY_BG,
        dropdown_text_color=AppTheme.COLOR_TEXT,
        font=(AppTheme.FONT_FAMILY, font_size, weight),
        dropdown_font=(AppTheme.FONT_FAMILY, 13),
        command=command,
        allow_custom=allow_custom,
        on_search=on_search,
    )


def make_entry(
    parent,
    width: int = 140,
    height: int = 38,
    font_size: int = 13,
    bold: bool = True,
    placeholder_text: str = "",
) -> ctk.CTkEntry:
    """Unified, clean corporate text entry control."""
    weight = "bold" if bold else "normal"
    kwargs: Dict[str, Any] = {
        "width": width,
        "height": height,
        "corner_radius": AppTheme.RADIUS_INPUT,
        "border_width": 1,
        "fg_color": AppTheme.COLOR_INPUT_BG,
        "border_color": AppTheme.COLOR_INPUT_BORDER,
        "text_color": AppTheme.COLOR_TEXT,
        "font": (AppTheme.FONT_FAMILY, font_size, weight),
    }
    if placeholder_text:
        kwargs["placeholder_text"] = placeholder_text
    return ctk.CTkEntry(parent, **kwargs)


# ==============================================================================
# VIEW 1: SMART RECEIPT POS (หน้า 1: ออกใบเสร็จรับเงิน - โทนสีสุภาพ เป็นมืออาชีพ)
# ==============================================================================
class SmartReceiptPOSView(ctk.CTkFrame):
    """
    Dual-Mode POS Receipt Issuing View matching official college specifications (EX 91693_0.jpg & EXofBill):
    - Mode 1: ส่วนปกติ (กรอกมือทั้งหมดตามแบบ EX)
      ฟอร์ม 2 คอลัมน์สมมาตร สวยงาม เรียงเท่าๆ กัน ดูง่าย ไม่มั่ว กรอกมือได้ทุกช่อง + ดึงรหัสนักศึกษาด่วน
    - Mode 2: ออกแบบใบเสร็จแบบแอดวานซ์ (Advanced Mode)
      รวมฟีเจอร์ทันสมัยครบครัน: ออกใบเสร็จยกห้อง, นำเข้าไฟล์ Excel รายชื่อเด็ก, ออกแบบสไตล์ใบเสร็จ, ระบบ QR PromptPay,
      ชุดค่าเทอมเหมาจ่ายด่วน, และ Webhook Alert
    """

    def __init__(
        self,
        parent,
        db: DatabaseManager,
        session: UserSession,
        config: Dict[str, Any],
        webhook: FinanceWebhookDispatcher,
        on_navigate: Optional[Callable[[str], None]] = None,
    ):
        super().__init__(parent, fg_color="transparent")
        self.db = db
        self.session = session
        self.config = config
        self.webhook = webhook
        self.on_navigate = on_navigate

        self.students: List[Dict[str, Any]] = []
        self.student_map: Dict[str, Dict[str, Any]] = {}
        self.cart_items: List[Dict[str, Any]] = []
        self.current_catalog: List[Dict[str, Any]] = []
        self.catalog_map: Dict[str, Dict[str, Any]] = {}
        self.active_mode = "normal"

        self.STUDENT_DROPDOWN_LIMIT = 200
        self._stu_search_job: Optional[str] = None
        self._batch_search_job: Optional[str] = None

        # Master & Dynamic Classroom State ("เลือกทั้งห้องได้ + dynamic เผื่อบางคนไม่เหมือนเพื่อน")
        self.batch_students_data: List[Dict[str, Any]] = []
        self.batch_check_vars: Dict[str, ctk.BooleanVar] = {}
        self.batch_custom_overrides: Dict[str, Dict[str, Any]] = {}
        self.batch_base_items: List[Dict[str, Any]] = [
            {
                "fee_code": "ENV",
                "curriculum_name": "หลักสูตรในระบบ-ปวช",
                "item_name": "ค่ารักษาภาพแวดล้อม",
                "quantity": 0.0,
                "unit_price": 330.0,
                "discount": 0.0,
                "line_total": 330.0,
            },
            {
                "fee_code": "NET",
                "curriculum_name": "หลักสูตรในระบบ-ปวช",
                "item_name": "ค่าใช้บริการอินเตอร์เน็ต",
                "quantity": 0.0,
                "unit_price": 300.0,
                "discount": 0.0,
                "line_total": 300.0,
            },
            {
                "fee_code": "TCH",
                "curriculum_name": "หลักสูตรในระบบ-ปวช",
                "item_name": "ค่าจ้างครูหรือผู้ที่มีความเชี่ยวชาญในสาขาเฉพาะ",
                "quantity": 0.0,
                "unit_price": 500.0,
                "discount": 0.0,
                "line_total": 500.0,
            },
        ]

        # Load master options
        self.curricula = self.db.get_options("CURRICULUM") or ["หลักสูตรในระบบ-ปวช"]
        self.faculties = self.db.get_options("FACULTY") or ["บริหารธุรกิจ / พาณิชยกรรม"]
        self.departments = self.db.get_options("DEPARTMENT") or ["การบัญชี"]
        self.class_levels = ["ทั้งหมด"] + (self.db.get_options("CLASS_LEVEL") or ["ปวช.1/1"])
        self.terms = self.db.get_options("TERM") or ["1/2569"]
        cls_options = [c for c in self.class_levels if c != "ทั้งหมด"] or ["ปวช.1/1"]

        # ------------------------------------------------------------------
        # TOP HEADER: Title + 3-Mode Switcher (1.รายคนกรอกมือ 2.ยกห้องDynamic 3.แอดวานซ์)
        # ------------------------------------------------------------------
        top_header = ctk.CTkFrame(
            self,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        top_header.pack(side="top", fill="x", pady=(0, 14))

        th_in = ctk.CTkFrame(top_header, fg_color="transparent")
        th_in.pack(fill="x", padx=20, pady=12)
        th_in.grid_columnconfigure(0, weight=1)
        th_in.grid_columnconfigure(1, weight=0)

        t_box = ctk.CTkFrame(th_in, fg_color="transparent")
        t_box.grid(row=0, column=0, sticky="w", pady=2)

        ctk.CTkLabel(
            t_box,
            text="ออกใบเสร็จรับเงินค่าเทอม",
            font=AppTheme.FONT_TITLE,
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        ).pack(anchor="w")

        ctk.CTkLabel(
            t_box,
            text=f"{self.config.get('college_name', 'วิทยาลัยอาชีวศึกษาสุพรรณบุรี')} • ระบบออกใบเสร็จรับเงินค่าเทอมมาตรฐาน",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color=AppTheme.COLOR_TEXT_MUTED,
            anchor="w",
        ).pack(anchor="w")

        mode_box = ctk.CTkFrame(th_in, fg_color="transparent")
        mode_box.grid(row=0, column=1, sticky="e", pady=2)
        self.mode_switcher = mode_box

        self.btn_tab_normal = ctk.CTkButton(
            mode_box,
            text="ออกใบเสร็จรายคน",
            width=140,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=lambda: self.set_mode("normal"),
        )
        self.btn_tab_normal.pack(side="left", padx=(0, 8))

        self.btn_tab_batch = ctk.CTkButton(
            mode_box,
            text="ออกใบเสร็จยกห้อง",
            width=140,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=lambda: self.set_mode("batch"),
        )
        self.btn_tab_batch.pack(side="left", padx=(0, 8))

        self.btn_tab_advanced = ctk.CTkButton(
            mode_box,
            text="ออกแบบ & แอดวานซ์",
            width=145,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=lambda: self.set_mode("advanced"),
        )
        self.btn_tab_advanced.pack(side="left")

        # ------------------------------------------------------------------
        # MAIN BODY CONTAINER (Holds Normal Panel, Batch Panel, & Advanced Panel)
        # ------------------------------------------------------------------
        self.body_container = ctk.CTkFrame(self, fg_color="transparent")
        self.body_container.pack(fill="both", expand=True)

        self.panel_normal = ctk.CTkFrame(self.body_container, fg_color="transparent")
        self.panel_batch = ctk.CTkFrame(self.body_container, fg_color="transparent")
        self.panel_advanced = ctk.CTkScrollableFrame(self.body_container, fg_color="transparent")

        # ==================================================================
        # 1. NORMAL MODE PANEL (ส่วนปกติ: จัดเรียงเป็นสัดส่วน ด้านล่างเป็นตารางรายละเอียด)
        # ==================================================================
        # 1.1 แถบสรุปยอดและปุ่มสั่งการด้านล่างสุด (Packed Bottom First เพื่อให้อยู่ล่างสุดเสมอ)
        bottom_bar = ctk.CTkFrame(
            self.panel_normal,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER_STRONG,
        )
        bottom_bar.pack(side="bottom", fill="x", pady=(10, 0))

        btm_inner = ctk.CTkFrame(bottom_bar, fg_color="transparent")
        btm_inner.pack(fill="x", padx=18, pady=10)

        tot_box = ctk.CTkFrame(btm_inner, fg_color="transparent")
        tot_box.pack(side="left", fill="x", expand=True)

        self.lbl_net_total = ctk.CTkLabel(
            tot_box,
            text="ยอดรวมสุทธิ: ฿0.00 บาท",
            font=(AppTheme.FONT_FAMILY, 21, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
            anchor="w",
        )
        self.lbl_net_total.pack(anchor="w")

        self.lbl_baht_text = ctk.CTkLabel(
            tot_box,
            text="ตัวอักษร: (ศูนย์บาทถ้วน)",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_MUTED,
            anchor="w",
        )
        self.lbl_baht_text.pack(anchor="w")

        btn_box = ctk.CTkFrame(btm_inner, fg_color="transparent")
        btn_box.pack(side="right")

        self.pay_method_var = ctk.StringVar(value=self.config.get("payment_method", "เงินสด (Cash)"))
        make_dropdown(
            btn_box,
            values=["เงินสด (Cash)", "สแกน QR PromptPay", "โอนเข้าบัญชีวิทยาลัย"],
            variable=self.pay_method_var,
            width=155,
            height=AppTheme.CTRL_H_MD,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_box,
            text="ลบแถวที่เลือก",
            width=105,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.remove_selected_item,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            btn_box,
            text="ล้างฟอร์ม",
            width=85,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_MUTED,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.reset_pos_form,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            btn_box,
            text="บันทึกข้อมูล",
            width=115,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.save_receipt_only,
        ).pack(side="left", padx=(0, 6))

        self.btn_export_folder_normal = ctk.CTkButton(
            btn_box,
            text=format_export_folder_button_text(self.config),
            width=155,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self._pick_export_folder,
        )
        self.btn_export_folder_normal.pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            btn_box,
            text="บันทึก & พิมพ์ใบเสร็จ (.pdf)",
            width=210,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.issue_receipt_and_export,
        ).pack(side="left")

        # ==================================================================
        # 1.2 พื้นที่ Scrollable Body (เลื่อนขึ้น-ลงอิสระเหมือน Landing Page)
        # ==================================================================
        self.scroll_normal = ctk.CTkScrollableFrame(self.panel_normal, fg_color="transparent")
        self.scroll_normal.pack(side="top", fill="both", expand=True)
        self.pos_form = self.scroll_normal

        # การ์ดข้อมูลหัวใบเสร็จรับเงิน (ส่วนที่ 1 - ด้านบน)
        head_card = ctk.CTkFrame(
            self.scroll_normal,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        head_card.pack(side="top", fill="x", pady=(0, 10))

        # Title + Template Style Selector (moved out of bottom bar for cleaner action bar)
        h_header_box = ctk.CTkFrame(head_card, fg_color="transparent")
        h_header_box.pack(fill="x", padx=20, pady=(14, 8))

        ctk.CTkLabel(
            h_header_box,
            text="ส่วนที่ 1 : ข้อมูลหัวใบเสร็จรับเงิน",
            font=AppTheme.FONT_SECTION,
            text_color=AppTheme.COLOR_ACCENT,
            anchor="w",
        ).pack(side="left")

        self.tpl_style_var = ctk.StringVar(value=_style_label_from_code(self.config.get("bill_template_style", "CLASSIC")))
        make_dropdown(
            h_header_box,
            values=TEMPLATE_STYLE_OPTIONS,
            variable=self.tpl_style_var,
            width=210,
            height=AppTheme.CTRL_H_SM,
            font_size=12,
            bold=False,
            command=self._on_change_template_style,
        ).pack(side="right")

        ctk.CTkLabel(
            h_header_box,
            text="รูปแบบใบเสร็จ:",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="right", padx=(0, 6))

        # Quick Search Toolbar (แถบค้นหาเรียบ ไร้กรอบซ้อน เว้นระยะเหมาะสม)
        search_toolbar = ctk.CTkFrame(head_card, fg_color="transparent")
        search_toolbar.pack(fill="x", padx=20, pady=(0, 12))

        ctk.CTkLabel(
            search_toolbar,
            text="ค้นหานักศึกษา:",
            font=AppTheme.FONT_LABEL,
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left", padx=(0, 8))

        self.filter_class_var = ctk.StringVar(value="ทั้งหมด")
        self.opt_filter_class = make_dropdown(
            search_toolbar,
            values=self.class_levels,
            variable=self.filter_class_var,
            width=115,
            height=AppTheme.CTRL_H_SM,
            font_size=12,
            bold=False,
            command=lambda _v: self.refresh_student_list(),
        )
        self.opt_filter_class.pack(side="left", padx=(0, 8))

        self.entry_stu_search = make_entry(
            search_toolbar,
            width=165,
            height=AppTheme.CTRL_H_SM,
            font_size=12,
            bold=False,
            placeholder_text="พิมพ์รหัส หรือ ชื่อ...",
        )
        self.entry_stu_search.pack(side="left", padx=(0, 8))
        self.entry_stu_search.bind("<KeyRelease>", self._schedule_student_search)

        self.stu_var = ctk.StringVar(value="")
        self.opt_student = make_dropdown(
            search_toolbar,
            values=["-"],
            variable=self.stu_var,
            width=260,
            height=AppTheme.CTRL_H_SM,
            font_size=12,
            bold=False,
            command=self.on_select_student,
            on_search=self._search_students_dropdown,
        )
        self.opt_student.pack(side="left", fill="x", expand=True)

        # ------------------------------------------------------------------
        # Header Fields Grid (เลย์เอาต์ตาราง 4 คอลัมน์สมมาตร ช่องกรอกสัดส่วนพอดี เป็นระเบียบ)
        # ------------------------------------------------------------------
        tbl_head = ctk.CTkFrame(head_card, fg_color="transparent")
        tbl_head.pack(fill="x", padx=18, pady=(2, 12))

        tbl_head.grid_columnconfigure(0, weight=0, minsize=145)
        tbl_head.grid_columnconfigure(1, weight=1, uniform="pos_grid")
        tbl_head.grid_columnconfigure(2, weight=0, minsize=145)
        tbl_head.grid_columnconfigure(3, weight=1, uniform="pos_grid")

        # Row 0: เลขที่ใบเสร็จรับเงิน  |  วันที่ออกใบเสร็จ
        ctk.CTkLabel(
            tbl_head,
            text="เลขที่ใบเสร็จรับเงิน:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)

        f_rcp = ctk.CTkFrame(tbl_head, fg_color="transparent")
        f_rcp.grid(row=0, column=1, sticky="w", padx=(0, 16), pady=5)

        self.entry_rcp_no = make_entry(
            f_rcp,
            width=210,
            height=32,
            font_size=13,
            bold=True,
            placeholder_text="เลขที่ใบเสร็จ (เช่น 700605433)",
        )
        self.entry_rcp_no.pack(side="left", padx=(0, 6))
        self.entry_rcp_no.bind("<FocusOut>", self._on_receipt_no_focusout)

        self.btn_rcp_cfg = ctk.CTkButton(
            f_rcp,
            text="⚙️ กำหนด",
            width=72,
            height=32,
            corner_radius=AppTheme.RADIUS_BTN,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            command=self.open_receipt_no_config_modal,
        )
        self.btn_rcp_cfg.pack(side="left")

        ctk.CTkLabel(
            tbl_head,
            text="วันที่ออกใบเสร็จ:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=0, column=2, sticky="w", padx=(0, 8), pady=5)

        self.entry_date = make_entry(tbl_head, width=160, height=32, font_size=13, bold=False)
        self.entry_date.grid(row=0, column=3, sticky="w", pady=5)
        self.entry_date.insert(0, format_thai_date_short())

        # Row 1: รหัสประจำตัวนักศึกษา  |  ได้รับเงินจาก (ชื่อ-สกุล)
        ctk.CTkLabel(
            tbl_head,
            text="รหัสประจำตัวนักศึกษา:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)

        f_code = ctk.CTkFrame(tbl_head, fg_color="transparent")
        f_code.grid(row=1, column=1, sticky="w", padx=(0, 16), pady=5)

        self.entry_stu_code = make_entry(f_code, width=140, height=32, font_size=13, bold=True)
        self.entry_stu_code.pack(side="left", padx=(0, 6))
        self.entry_stu_code.bind("<Return>", self._lookup_student_by_code_entry)
        self.entry_stu_code.bind("<FocusOut>", self._lookup_student_by_code_entry)

        ctk.CTkButton(
            f_code,
            text="🔍 ดึงข้อมูล",
            width=80,
            height=32,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self._lookup_student_by_code_entry,
        ).pack(side="left")

        ctk.CTkLabel(
            tbl_head,
            text="ได้รับเงินจาก (ชื่อ-สกุล):",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=1, column=2, sticky="w", padx=(0, 8), pady=5)

        self.entry_stu_name = make_entry(
            tbl_head,
            height=32,
            font_size=13,
            bold=True,
            placeholder_text="กรอกชื่อ-สกุลผู้ชำระเงิน หรือพิมพ์รหัสเพื่อค้นหา...",
        )
        self.entry_stu_name.grid(row=1, column=3, sticky="ew", pady=5)

        # Row 2: ภาค/ปีการศึกษา  |  คณะ
        ctk.CTkLabel(
            tbl_head,
            text="ภาค/ปีการศึกษา:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)

        self.term_var = ctk.StringVar(value=self.config.get("current_semester", self.terms[0]))
        self.opt_term = make_dropdown(
            tbl_head,
            values=self.terms,
            variable=self.term_var,
            width=160,
            height=32,
            font_size=13,
        )
        self.opt_term.grid(row=2, column=1, sticky="w", padx=(0, 16), pady=5)

        ctk.CTkLabel(
            tbl_head,
            text="คณะ:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=2, column=2, sticky="w", padx=(0, 8), pady=5)

        self.fac_var = ctk.StringVar(value=self.faculties[0])
        self.opt_faculty = make_dropdown(
            tbl_head,
            values=self.faculties,
            variable=self.fac_var,
            height=32,
            font_size=13,
            bold=False,
        )
        self.opt_faculty.grid(row=2, column=3, sticky="ew", pady=5)

        # Row 3: ชั้นปี / ห้อง  |  แผนกวิชา / สาขา
        ctk.CTkLabel(
            tbl_head,
            text="ชั้นปี / ห้องเรียน:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=3, column=0, sticky="w", padx=(0, 8), pady=5)

        cls_options = [c for c in self.class_levels if c != "ทั้งหมด"] or ["ปวช.1/1"]
        self.cls_var = ctk.StringVar(value=cls_options[0])
        self.opt_cls = make_dropdown(
            tbl_head,
            values=cls_options,
            variable=self.cls_var,
            width=160,
            height=32,
            font_size=13,
            bold=False,
        )
        self.opt_cls.grid(row=3, column=1, sticky="w", padx=(0, 16), pady=5)

        ctk.CTkLabel(
            tbl_head,
            text="แผนกวิชา / สาขา:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=3, column=2, sticky="w", padx=(0, 8), pady=5)

        self.dept_var = ctk.StringVar(value=self.departments[0])
        self.opt_dept = make_dropdown(
            tbl_head,
            values=self.departments,
            variable=self.dept_var,
            height=32,
            font_size=13,
            bold=False,
        )
        self.opt_dept.grid(row=3, column=3, sticky="ew", pady=5)

        # ==================================================================
        # 1.3 การ์ดเพิ่มรายการค่าธรรมเนียม (ส่วนที่ 2 - ตรงกลาง)
        # ==================================================================
        item_card = ctk.CTkFrame(
            self.scroll_normal,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        item_card.pack(side="top", fill="x", pady=(0, 10))

        i_header_box = ctk.CTkFrame(item_card, fg_color="transparent")
        i_header_box.pack(fill="x", padx=20, pady=(14, 6))

        ctk.CTkLabel(
            i_header_box,
            text="ส่วนที่ 2 : บันทึกรายการค่าธรรมเนียมลงใบเสร็จ",
            font=AppTheme.FONT_SECTION,
            text_color=AppTheme.COLOR_ACCENT,
            anchor="w",
        ).pack(side="left")

        # Quick Preset Buttons on Header
        p_quick = ctk.CTkFrame(i_header_box, fg_color="transparent")
        p_quick.pack(side="right")

        ctk.CTkButton(
            p_quick,
            text="ชุด ปวช. (฿1,130)",
            height=28,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=lambda: self.load_fee_preset("VC1"),
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            p_quick,
            text="⚡ ชุด ปวส. (฿1,550)",
            height=28,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=lambda: self.load_fee_preset("HVC1"),
        ).pack(side="left")

        # ------------------------------------------------------------------
        # Item Fields Grid (เลย์เอาต์ตาราง 4 คอลัมน์สมมาตร ตรงกับตารางหัวใบเสร็จ)
        # ------------------------------------------------------------------
        tbl_item = ctk.CTkFrame(item_card, fg_color="transparent")
        tbl_item.pack(fill="x", padx=18, pady=(2, 12))

        tbl_item.grid_columnconfigure(0, weight=0, minsize=145)
        tbl_item.grid_columnconfigure(1, weight=1, uniform="pos_grid")
        tbl_item.grid_columnconfigure(2, weight=0, minsize=145)
        tbl_item.grid_columnconfigure(3, weight=1, uniform="pos_grid")

        # Row 0: หลักสูตรการศึกษา  |  เลือกรายการจากคลัง (ด่วน)
        ctk.CTkLabel(
            tbl_item,
            text="หลักสูตรการศึกษา:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)

        self.curriculum_var = ctk.StringVar(value=self.curricula[0])
        self.opt_curriculum = make_dropdown(
            tbl_item,
            values=self.curricula + ["กำหนดเอง / อื่นๆ"],
            variable=self.curriculum_var,
            height=32,
            font_size=13,
            command=self.on_change_curriculum,
        )
        self.opt_curriculum.grid(row=0, column=1, sticky="ew", padx=(0, 16), pady=5)

        ctk.CTkLabel(
            tbl_item,
            text="เลือกรายการจากคลัง:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=0, column=2, sticky="w", padx=(0, 8), pady=5)

        self.fee_item_var = ctk.StringVar(value="")
        self.opt_fee_item = make_dropdown(
            tbl_item,
            values=["-"],
            variable=self.fee_item_var,
            height=32,
            font_size=13,
            command=self.on_select_fee_item,
        )
        self.opt_fee_item.grid(row=0, column=3, sticky="ew", pady=5)

        # Row 1: ชื่อรายการในใบเสร็จ  |  หน่วยกิต / จำนวน
        ctk.CTkLabel(
            tbl_item,
            text="ชื่อรายการในใบเสร็จ:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)

        self.entry_item_name = make_entry(
            tbl_item,
            height=32,
            font_size=13,
            bold=True,
            placeholder_text="พิมพ์ชื่อรายการ หรือเลือกจากคลัง...",
        )
        self.entry_item_name.grid(row=1, column=1, sticky="ew", padx=(0, 16), pady=5)

        ctk.CTkLabel(
            tbl_item,
            text="หน่วยกิต / จำนวน:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=1, column=2, sticky="w", padx=(0, 8), pady=5)

        self.entry_qty = make_entry(tbl_item, width=130, height=32, font_size=13, bold=True)
        self.entry_qty.grid(row=1, column=3, sticky="w", pady=5)
        self.entry_qty.insert(0, "0")
        self.entry_qty.bind("<KeyRelease>", self._update_line_total_preview)

        # Row 2: ราคาต่อหน่วย (บาท)  |  รวมเงิน & ปุ่มเพิ่มรายการ
        ctk.CTkLabel(
            tbl_item,
            text="ราคาต่อหน่วย (บาท):",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)

        self.entry_unit_price = make_entry(tbl_item, width=160, height=32, font_size=13, bold=True)
        self.entry_unit_price.grid(row=2, column=1, sticky="w", padx=(0, 16), pady=5)
        self.entry_unit_price.insert(0, "0.00")
        self.entry_unit_price.bind("<KeyRelease>", self._update_line_total_preview)
        self.entry_unit_price.bind("<Return>", lambda _e: self.add_selected_fee_to_cart())

        ctk.CTkLabel(
            tbl_item,
            text="รวมเงินคำนวณ:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=2, column=2, sticky="w", padx=(0, 8), pady=5)

        f_action = ctk.CTkFrame(tbl_item, fg_color="transparent")
        f_action.grid(row=2, column=3, sticky="w", pady=5)

        badge_tot = ctk.CTkFrame(
            f_action,
            height=32,
            width=115,
            corner_radius=AppTheme.RADIUS,
            fg_color="#EFF6FF",
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
        )
        badge_tot.pack(side="left", padx=(0, 8))
        badge_tot.pack_propagate(False)

        self.lbl_item_total = ctk.CTkLabel(
            badge_tot,
            text="฿0.00",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
            anchor="center",
        )
        self.lbl_item_total.pack(fill="both", expand=True)

        btn_add = ctk.CTkButton(
            f_action,
            text="เพิ่มรายการลงใบเสร็จ",
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=AppTheme.FONT_LABEL,
            command=self.add_selected_fee_to_cart,
        )
        btn_add.pack(side="left", fill="x", expand=True)

        # ==================================================================
        # 1.4 ส่วนที่แสดงรายละเอียดใบเสร็จรับเงิน (ตารางรายการค่าธรรมเนียม - ด้านล่าง)
        # ==================================================================
        table_card = ctk.CTkFrame(
            self.scroll_normal,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        table_card.pack(side="top", fill="x", pady=(0, 8))

        tbl_top = ctk.CTkFrame(table_card, fg_color="transparent")
        tbl_top.pack(fill="x", padx=18, pady=(12, 6))

        ctk.CTkLabel(
            tbl_top,
            text="📋 ส่วนที่ 3 : รายการค่าธรรมเนียมที่เลือกไว้ในใบเสร็จ",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
            anchor="w",
        ).pack(side="left")

        self.lbl_cart_count = ctk.CTkLabel(
            tbl_top,
            text="📋 0 รายการ",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        )
        self.lbl_cart_count.pack(side="right")

        cols = [
            ("line_no", "ลำดับ", 70, "center"),
            ("item_name", "รายการค่าธรรมเนียมในใบเสร็จ", 500, "w"),
            ("quantity", "หน่วยกิต/จำนวน", 120, "center"),
            ("unit_price", "หน่วยละ (บาท)", 150, "e"),
            ("line_total", "จำนวนเงิน (บาท)", 170, "e"),
        ]
        self.table = DataTable(table_card, columns=cols, height=8, show_inspector=False)
        self.table.pack(side="top", fill="both", expand=True, padx=8, pady=(0, 8))

        # ==================================================================
        # 2. DYNAMIC CLASSROOM BATCH PANEL (โหมดออกใบเสร็จยกห้อง & ปรับเฉพาะคน)
        # ==================================================================
        self._batch_panel_built = False

        # ==================================================================
        # 3. ADVANCED MODE PANEL (โหมดแอดวานซ์ & ออกแบบใบเสร็จรับเงิน)
        # ==================================================================
        self._advanced_panel_built = False

        # Start on Normal Mode by default
        self.set_mode("normal")

        # Initial data loading
        self.refresh_next_receipt_no()
        self.refresh_student_list()
        self.on_change_curriculum(self.curricula[0])
        self._last_data_version = getattr(self.db, "data_version", -1)

    def on_show(self):
        """Refreshes cached POS data only when underlying database data has changed."""
        if getattr(self.db, "data_version", None) == getattr(self, "_last_data_version", -1):
            return
        self.refresh_next_receipt_no()
        self.refresh_student_list()
        curr = self.curr_var.get() if hasattr(self, "curr_var") else self.curricula[0]
        self.on_change_curriculum(curr)
        if self._batch_panel_built and self.active_mode == "batch":
            self.refresh_batch_roster()
        elif self._advanced_panel_built and self.active_mode == "advanced":
            self.sync_advanced_panel_fields()
        self._last_data_version = getattr(self.db, "data_version", -1)

    def set_mode(self, mode: str):
        """Switches between Normal Mode (Manual entry matching EX), Batch Mode (Dynamic classroom roster), and Advanced Mode."""
        self.active_mode = mode

        self.panel_normal.pack_forget()
        self.panel_batch.pack_forget()
        self.panel_advanced.pack_forget()

        for btn, key in [
            (self.btn_tab_normal, "normal"),
            (self.btn_tab_batch, "batch"),
            (self.btn_tab_advanced, "advanced"),
        ]:
            if key == mode:
                btn.configure(
                    fg_color=AppTheme.COLOR_PRIMARY_BLUE,
                    text_color="#FFFFFF",
                    border_width=0,
                )
            else:
                btn.configure(
                    fg_color=AppTheme.COLOR_SECONDARY_BG,
                    border_width=1,
                    border_color=AppTheme.COLOR_SECONDARY_BORDER,
                    text_color=AppTheme.COLOR_SECONDARY_TEXT,
                )

        if mode == "normal":
            self.panel_normal.pack(fill="both", expand=True)
        elif mode == "batch":
            if not self._batch_panel_built:
                self._build_batch_panel()
                self._batch_panel_built = True
            self.panel_batch.pack(fill="both", expand=True)
            self.refresh_batch_roster()
        elif mode == "advanced":
            if not self._advanced_panel_built:
                self._build_advanced_panel()
                self._advanced_panel_built = True
                self.refresh_cart_ui()
            self.panel_advanced.pack(fill="both", expand=True)
            self.sync_advanced_panel_fields()

    # ------------------------------------------------------------------
    # DYNAMIC CLASSROOM BATCH PANEL (เลือกทั้งห้อง + ปรับแต่งเฉพาะบุคคล)
    # ------------------------------------------------------------------
    def _build_batch_panel(self):
        """Builds the dynamic interactive classroom roster for issuing batch receipts with per-student customization."""
        # Top Card: Room Selector & Base Fee Configuration
        b_head_card = ctk.CTkFrame(
            self.panel_batch,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        b_head_card.pack(side="top", fill="x", pady=(0, 6))

        bh_in = ctk.CTkFrame(b_head_card, fg_color="transparent")
        bh_in.pack(fill="x", padx=16, pady=8)

        # Row 1: Dropdown Filters & Actions
        r1 = ctk.CTkFrame(bh_in, fg_color="transparent")
        r1.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(
            r1,
            text="เลือกระดับชั้น / ห้องเรียน:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left", padx=(0, 6))

        cls_opts = [c for c in self.class_levels if c != "ทั้งหมด"] or ["ปวช.1/1"]
        self.batch_room_var = ctk.StringVar(value=cls_opts[0])
        self.opt_batch_room = make_dropdown(
            r1,
            values=cls_opts,
            variable=self.batch_room_var,
            width=150,
            height=36,
            font_size=13,
            command=lambda _v: self.refresh_batch_roster(),
        )
        self.opt_batch_room.pack(side="left", padx=(0, 12))

        ctk.CTkLabel(
            r1,
            text="แผนกวิชา / สาขา:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left", padx=(0, 6))

        self.batch_dept_var = ctk.StringVar(value="ทั้งหมด")
        self.opt_batch_dept = make_dropdown(
            r1,
            values=["ทั้งหมด"] + self.departments,
            variable=self.batch_dept_var,
            width=175,
            height=36,
            font_size=13,
            command=lambda _v: self.refresh_batch_roster(),
        )
        self.opt_batch_dept.pack(side="left", padx=(0, 12))

        ctk.CTkLabel(
            r1,
            text="ภาค/ปีการศึกษา:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left", padx=(0, 6))

        self.batch_term_var = ctk.StringVar(value=self.config.get("current_semester", self.terms[0]))
        self.opt_batch_term = make_dropdown(
            r1,
            values=self.terms,
            variable=self.batch_term_var,
            width=135,
            height=36,
            font_size=13,
        )
        self.opt_batch_term.pack(side="left", padx=(0, 12))

        ctk.CTkButton(
            r1,
            text="🔄 รีเฟรชรายชื่อ",
            width=120,
            height=36,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.refresh_batch_roster,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            r1,
            text="⚙️ ตั้งค่าเทอมมาตรฐานของห้อง",
            width=190,
            height=36,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self._open_batch_fee_preset_modal,
        ).pack(side="right")

        # Row 2: Base fee description & Excel import shortcut
        r2 = ctk.CTkFrame(bh_in, fg_color="transparent")
        r2.pack(fill="x")

        tot_base = sum(float(i["line_total"]) for i in self.batch_base_items)
        self.lbl_batch_base_fee_info = ctk.CTkLabel(
            r2,
            text=f"📌 ค่าเทอมมาตรฐานห้อง: ค่ารักษาภาพแวดล้อม (฿330), เน็ต (฿300), ครูผู้เชี่ยวชาญ (฿500) | รวม ฿{tot_base:,.2f} บาท",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
            anchor="w",
        )
        self.lbl_batch_base_fee_info.pack(side="left")

        ctk.CTkButton(
            r2,
            text="📥 นำเข้าไฟล์ Excel/CSV รายชื่อเด็กแทน",
            width=230,
            height=32,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12),
            command=self.import_and_issue_batch,
        ).pack(side="right")

        # Bottom Bar of Batch Panel (Summary & Issue CTA, Packed Bottom First)
        b_bottom_bar = ctk.CTkFrame(
            self.panel_batch,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER_STRONG,
        )
        b_bottom_bar.pack(side="bottom", fill="x", pady=(6, 0))

        bb_inner = ctk.CTkFrame(b_bottom_bar, fg_color="transparent")
        bb_inner.pack(fill="x", padx=16, pady=8)

        # Left: Totals & Receipts Count Info
        tot_b_box = ctk.CTkFrame(bb_inner, fg_color="transparent")
        tot_b_box.pack(side="left", fill="x", expand=True)

        self.lbl_batch_bottom_total = ctk.CTkLabel(
            tot_b_box,
            text="ยอดรวมทั้งห้องที่เลือก: ฿0.00 บาท",
            font=(AppTheme.FONT_FAMILY, 20, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
            anchor="w",
        )
        self.lbl_batch_bottom_total.pack(anchor="w")

        self.lbl_batch_bottom_baht = ctk.CTkLabel(
            tot_b_box,
            text="ตัวอักษร: (ศูนย์บาทถ้วน)",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_MUTED,
            anchor="w",
        )
        self.lbl_batch_bottom_baht.pack(anchor="w")

        self.lbl_batch_bottom_count = ctk.CTkLabel(
            tot_b_box,
            text="จำนวนใบเสร็จที่จะออก: 0 ฉบับ",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        )
        self.lbl_batch_bottom_count.pack(anchor="w")

        # Right: Payment method + Issue Batch Button
        act_b_box = ctk.CTkFrame(bb_inner, fg_color="transparent")
        act_b_box.pack(side="right")

        ctk.CTkLabel(
            act_b_box,
            text="ช่องทางชำระ:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))

        self.batch_pay_method_var = ctk.StringVar(value=self.config.get("payment_method", "เงินสด (Cash)"))
        make_dropdown(
            act_b_box,
            values=["เงินสด (Cash)", "สแกน QR PromptPay", "โอนเข้าบัญชีวิทยาลัย"],
            variable=self.batch_pay_method_var,
            width=155,
            height=38,
            font_size=12,
        ).pack(side="left", padx=(0, 10))

        self.btn_export_folder_batch = ctk.CTkButton(
            act_b_box,
            text=format_export_folder_button_text(self.config),
            width=165,
            height=42,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self._pick_export_folder,
        )
        self.btn_export_folder_batch.pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            act_b_box,
            text="🖨️ ออกใบเสร็จให้นักศึกษาที่เลือกทั้งหมด ➔ รวม PDF เดียว",
            height=42,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.execute_batch_roster_issue,
        ).pack(side="left")

        # Middle Toolbar: Quick Selection Tools + Search Filter + Stats Badge
        b_tool_card = ctk.CTkFrame(
            self.panel_batch,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE_ALT,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        b_tool_card.pack(side="top", fill="x", pady=(0, 6))

        bt_in = ctk.CTkFrame(b_tool_card, fg_color="transparent")
        bt_in.pack(fill="x", padx=14, pady=6)

        ctk.CTkButton(
            bt_in,
            text="☑️ เลือกทุกคนในห้อง (Select All)",
            width=180,
            height=34,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.select_all_batch_students,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            bt_in,
            text="⬜ ยกเลิกเลือกทั้งหมด",
            width=135,
            height=34,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.deselect_all_batch_students,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            bt_in,
            text="⚡ คืนค่ามาตรฐานทุกคน",
            width=150,
            height=34,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_MUTED,
            font=(AppTheme.FONT_FAMILY, 12),
            command=self.reset_all_batch_custom_overrides,
        ).pack(side="left", padx=(0, 14))

        self.entry_batch_search = make_entry(
            bt_in,
            width=190,
            height=34,
            font_size=12,
            bold=False,
            placeholder_text="🔍 ค้นหารายชื่อ/รหัสในห้อง...",
        )
        self.entry_batch_search.pack(side="left", padx=(0, 14))
        self.entry_batch_search.bind("<KeyRelease>", self._on_search_batch_roster)

        self.lbl_batch_stats_badge = ctk.CTkLabel(
            bt_in,
            text="👥 นักศึกษาในห้องนี้: 0 คน  |  ☑️ เลือก: 0 คน  |  💰 รวมทั้งสิ้น: ฿0.00 บาท",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
            anchor="e",
        )
        self.lbl_batch_stats_badge.pack(side="right")

        # Roster Area (Fixed Header + Scrollable Student Rows)
        b_roster_area = ctk.CTkFrame(
            self.panel_batch,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        b_roster_area.pack(side="top", fill="both", expand=True)

        # Header Row
        hdr_bar = ctk.CTkFrame(b_roster_area, fg_color="#1E293B", corner_radius=AppTheme.RADIUS, height=36)
        hdr_bar.pack(fill="x", padx=6, pady=(6, 2))
        hdr_bar.pack_propagate(False)

        ctk.CTkLabel(hdr_bar, text="เลือก", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color="#F8FAFC", width=60).pack(side="left", padx=(10, 0))
        ctk.CTkLabel(hdr_bar, text="ลำดับ", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color="#F8FAFC", width=50).pack(side="left")
        ctk.CTkLabel(hdr_bar, text="รหัสประจำตัว", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color="#F8FAFC", width=120, anchor="w").pack(side="left", padx=4)
        ctk.CTkLabel(hdr_bar, text="ชื่อ-สกุล นักศึกษา", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color="#F8FAFC", width=220, anchor="w").pack(side="left", padx=4)
        ctk.CTkLabel(hdr_bar, text="ห้อง / แผนกวิชา", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color="#F8FAFC", width=180, anchor="w").pack(side="left", padx=4)
        ctk.CTkLabel(hdr_bar, text="รูปแบบค่าธรรมเนียม / สถานะ (Dynamic)", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color="#F8FAFC", anchor="w").pack(side="left", fill="x", expand=True, padx=6)
        ctk.CTkLabel(hdr_bar, text="ยอดเงิน (บาท)", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color="#F8FAFC", width=125, anchor="e").pack(side="left", padx=8)
        ctk.CTkLabel(hdr_bar, text="ปรับเฉพาะคน", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color="#F8FAFC", width=140).pack(side="left", padx=(0, 6))

        # Scrollable rows frame
        self.batch_roster_scroll = ctk.CTkScrollableFrame(b_roster_area, fg_color="transparent")
        self.batch_roster_scroll.pack(fill="both", expand=True, padx=6, pady=(0, 6))

    def refresh_batch_roster(self):
        """Loads all students belonging to selected classroom/major and populates dynamic checklist."""
        room = self.batch_room_var.get() if hasattr(self, "batch_room_var") else "ALL"
        dept = self.batch_dept_var.get() if hasattr(self, "batch_dept_var") else "ทั้งหมด"

        lvl = room if room != "ทั้งหมด" else "ALL"
        raw_students = self.db.get_students(level_filter=lvl, limit=500)

        if dept != "ทั้งหมด":
            raw_students = [
                s for s in raw_students
                if str(s.get("major", "")).strip() == dept or str(s.get("department", "")).strip() == dept or str(s.get("faculty", "")).strip() == dept
            ]

        self.batch_students_data = raw_students

        for s in self.batch_students_data:
            code = s["student_code"]
            if code not in self.batch_check_vars:
                self.batch_check_vars[code] = ctk.BooleanVar(value=True)

        self.render_batch_student_rows()
        self.recalculate_batch_totals()

    def _on_search_batch_roster(self, _event=None):
        if self._batch_search_job is not None:
            try:
                self.after_cancel(self._batch_search_job)
            except Exception:
                pass
        self._batch_search_job = self.after(200, self.render_batch_student_rows)

    def render_batch_student_rows(self):
        """Renders interactive student rows with checkboxes, custom overrides badges, and edit triggers."""
        self._batch_search_job = None
        for w in self.batch_roster_scroll.winfo_children():
            w.destroy()

        search_txt = self.entry_batch_search.get().strip().lower() if hasattr(self, "entry_batch_search") else ""
        base_amount = sum(float(i["line_total"]) for i in self.batch_base_items)
        displayed_students = []

        for s in self.batch_students_data:
            code = str(s.get("student_code", "")).strip()
            name = str(s.get("full_name", "")).strip()
            dept = str(s.get("major", "")).strip()
            if search_txt:
                if search_txt not in code.lower() and search_txt not in name.lower() and search_txt not in dept.lower():
                    continue
            displayed_students.append(s)

        if not displayed_students:
            no_box = ctk.CTkFrame(self.batch_roster_scroll, fg_color="transparent")
            no_box.pack(fill="x", pady=40)
            ctk.CTkLabel(
                no_box,
                text="⚠️ ไม่พบรายชื่อนักศึกษาในห้องนี้ หรือไม่ตรงกับคำค้นหา",
                font=(AppTheme.FONT_FAMILY, 15, "bold"),
                text_color=AppTheme.COLOR_TEXT_MUTED,
            ).pack()
            ctk.CTkLabel(
                no_box,
                text="สามารถกดปุ่ม '📥 นำเข้าไฟล์ Excel/CSV รายชื่อเด็กแทน' ด้านบน หรือไปที่เมนู '3. รายชื่อนักศึกษา' เพื่อเพิ่มข้อมูล",
                font=(AppTheme.FONT_FAMILY, 12),
                text_color=AppTheme.COLOR_TEXT_MUTED,
            ).pack(pady=(4, 0))
            return

        for idx, s in enumerate(displayed_students, start=1):
            code = s["student_code"]
            name = s["full_name"]
            dept = s.get("major") or s.get("department") or "-"
            cls_grp = s.get("class_group") or "-"

            if code not in self.batch_check_vars:
                self.batch_check_vars[code] = ctk.BooleanVar(value=True)

            var = self.batch_check_vars[code]

            if code in self.batch_custom_overrides:
                ovr = self.batch_custom_overrides[code]
                amt = float(ovr.get("amount", base_amount))
                rem = ovr.get("remark", "ปรับแต่งเฉพาะบุคคล")
                mode = ovr.get("mode", "custom")
                if mode == "free" or amt == 0:
                    status_text = "🎓 ทุนการศึกษา 100% (ฟรี)"
                    status_color = "#16A34A"
                elif mode == "half":
                    status_text = f"🌟 ทุนการศึกษา 50% (฿{amt:,.2f})"
                    status_color = "#D97706"
                else:
                    status_text = f"✏️ กำหนดเอง: {rem} (฿{amt:,.2f})"
                    status_color = "#2563EB"
            else:
                amt = base_amount
                status_text = f"มาตรฐานตามห้อง ({len(self.batch_base_items)} รายการ)"
                status_color = AppTheme.COLOR_TEXT_SECONDARY

            row_bg = "#FFFFFF" if idx % 2 == 1 else AppTheme.COLOR_SURFACE_ALT
            row = ctk.CTkFrame(self.batch_roster_scroll, fg_color=row_bg, corner_radius=AppTheme.RADIUS, height=42)
            row.pack(fill="x", pady=2, padx=4)
            row.pack_propagate(False)

            # Checkbox
            f_chk = ctk.CTkFrame(row, fg_color="transparent", width=60)
            f_chk.pack(side="left", padx=(10, 0))
            f_chk.pack_propagate(False)
            ctk.CTkCheckBox(
                f_chk,
                text="",
                variable=var,
                width=24,
                height=24,
                checkbox_width=20,
                checkbox_height=20,
                command=self.recalculate_batch_totals,
            ).pack(anchor="center")

            # Number
            f_no = ctk.CTkFrame(row, fg_color="transparent", width=50)
            f_no.pack(side="left")
            f_no.pack_propagate(False)
            ctk.CTkLabel(f_no, text=str(idx), font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_MUTED).pack(anchor="center")

            # Student Code
            f_code = ctk.CTkFrame(row, fg_color="transparent", width=120)
            f_code.pack(side="left")
            f_code.pack_propagate(False)
            ctk.CTkLabel(f_code, text=code, font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT, anchor="w").pack(side="left", padx=4)

            # Student Name
            f_name = ctk.CTkFrame(row, fg_color="transparent", width=220)
            f_name.pack(side="left")
            f_name.pack_propagate(False)
            ctk.CTkLabel(f_name, text=name, font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT, anchor="w").pack(side="left", padx=4)

            # Room / Major
            f_maj = ctk.CTkFrame(row, fg_color="transparent", width=180)
            f_maj.pack(side="left")
            f_maj.pack_propagate(False)
            ctk.CTkLabel(f_maj, text=f"{cls_grp} • {dept}", font=(AppTheme.FONT_FAMILY, 12), text_color=AppTheme.COLOR_TEXT_MUTED, anchor="w").pack(side="left", padx=4)

            # Status Badge
            f_stat = ctk.CTkFrame(row, fg_color="transparent")
            f_stat.pack(side="left", fill="x", expand=True)
            ctk.CTkLabel(f_stat, text=status_text, font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=status_color, anchor="w").pack(side="left", padx=6)

            # Amount
            f_amt = ctk.CTkFrame(row, fg_color="transparent", width=125)
            f_amt.pack(side="left")
            f_amt.pack_propagate(False)
            ctk.CTkLabel(f_amt, text=f"฿{amt:,.2f}", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_ACCENT, anchor="e").pack(side="right", padx=8)

            # Action Button: Edit student specific fee
            f_btn = ctk.CTkFrame(row, fg_color="transparent", width=140)
            f_btn.pack(side="left", padx=(0, 6))
            f_btn.pack_propagate(False)
            ctk.CTkButton(
                f_btn,
                text="✏️ ปรับเฉพาะคน",
                width=125,
                height=30,
                corner_radius=AppTheme.RADIUS,
                fg_color=AppTheme.COLOR_SECONDARY_BG,
                hover_color=AppTheme.COLOR_SECONDARY_HOVER,
                border_width=1,
                border_color=AppTheme.COLOR_SECONDARY_BORDER,
                text_color=AppTheme.COLOR_SECONDARY_TEXT,
                font=(AppTheme.FONT_FAMILY, 12, "bold"),
                command=lambda st=s: self.open_student_custom_modal(st),
            ).pack(anchor="center")

    def select_all_batch_students(self):
        """Selects all students in the room with one click."""
        for v in self.batch_check_vars.values():
            v.set(True)
        self.recalculate_batch_totals()

    def deselect_all_batch_students(self):
        """Deselects all students in the room."""
        for v in self.batch_check_vars.values():
            v.set(False)
        self.recalculate_batch_totals()

    def reset_all_batch_custom_overrides(self):
        """Resets all students in the room to standard room tuition."""
        self.batch_custom_overrides.clear()
        self.render_batch_student_rows()
        self.recalculate_batch_totals()

    def recalculate_batch_totals(self):
        """Calculates total amount and count of selected students in real time."""
        base_amount = sum(float(i["line_total"]) for i in self.batch_base_items)
        selected_count = 0
        total_amount = 0.0

        for s in self.batch_students_data:
            code = s["student_code"]
            if self.batch_check_vars.get(code) and self.batch_check_vars[code].get():
                selected_count += 1
                if code in self.batch_custom_overrides:
                    amt = float(self.batch_custom_overrides[code]["amount"])
                else:
                    amt = base_amount
                total_amount += amt

        tot_students = len(self.batch_students_data)
        if hasattr(self, "lbl_batch_stats_badge"):
            self.lbl_batch_stats_badge.configure(
                text=f"👥 นักศึกษาในห้องนี้: {tot_students} คน  |  ☑️ เลือก: {selected_count} คน  |  💰 รวมทั้งสิ้น: ฿{total_amount:,.2f} บาท"
            )
        if hasattr(self, "lbl_batch_bottom_total"):
            self.lbl_batch_bottom_total.configure(
                text=f"ยอดรวมทั้งห้องที่เลือก: ฿{total_amount:,.2f} บาท"
            )
        if hasattr(self, "lbl_batch_bottom_baht"):
            self.lbl_batch_bottom_baht.configure(
                text=f"ตัวอักษร: ({thai_baht_text(total_amount)})"
            )
        if hasattr(self, "lbl_batch_bottom_count"):
            prefix = self.config.get("receipt_prefix", "")
            start_no = self.config.get("receipt_no_start", "700605433")
            end_no = self.config.get("receipt_no_end", "700999999")
            first_no = self.db.get_next_receipt_no(prefix, start_no, end_no)
            if first_no.isdigit() and selected_count > 0:
                last_no = str(int(first_no) + selected_count - 1)
                self.lbl_batch_bottom_count.configure(
                    text=f"จำนวนใบเสร็จที่จะออก: {selected_count} ฉบับ  (เลขที่ {first_no} ถึง {last_no})"
                )
            else:
                self.lbl_batch_bottom_count.configure(
                    text=f"จำนวนใบเสร็จที่จะออก: {selected_count} ฉบับ"
                )

    def open_student_custom_modal(self, student: Dict[str, Any]):
        """Opens dynamic per-student customization modal for scholarships, custom fees, or waivers."""
        modal = ctk.CTkToplevel(self)
        modal.title(f"ปรับแต่งค่าธรรมเนียมเฉพาะบุคคล - {student.get('full_name')}")
        w, h = 540, 440
        modal.geometry(f"{w}x{h}")
        modal.resizable(False, False)
        modal.configure(fg_color="#FFFFFF")
        modal.transient(self.winfo_toplevel())

        try:
            from ui.icon_manager import IconManager
            IconManager.set_window_icon(modal)
        except Exception:
            pass

        try:
            modal.update_idletasks()
            top = self.winfo_toplevel()
            top_w = top.winfo_width()
            top_h = top.winfo_height()
            top_x = top.winfo_rootx() if hasattr(top, "winfo_rootx") else top.winfo_x()
            top_y = top.winfo_rooty() if hasattr(top, "winfo_rooty") else top.winfo_y()
            if top_w > 200 and top_h > 200:
                x = max(0, top_x + (top_w - w) // 2)
                y = max(0, top_y + (top_h - h) // 2)
            else:
                x = max(0, (modal.winfo_screenwidth() - w) // 2)
                y = max(0, (modal.winfo_screenheight() - h) // 2)
            modal.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            modal.geometry(f"{w}x{h}")

        def close_modal():
            try:
                modal.grab_release()
            except Exception:
                pass
            modal.destroy()

        modal.protocol("WM_DELETE_WINDOW", close_modal)
        modal.bind("<Escape>", lambda e: close_modal())
        modal.grab_set()
        modal.after(50, lambda: [modal.lift(), modal.focus_force()])

        # Top Accent Strip (Sky Blue)
        accent_bar = ctk.CTkFrame(modal, height=4, fg_color="#0284C7", corner_radius=0)
        accent_bar.pack(fill="x", side="top")

        # Bottom Button Footer Bar
        b_act = ctk.CTkFrame(modal, fg_color="#F8FAFC", height=54, corner_radius=0)
        b_act.pack(fill="x", side="bottom")
        b_act.pack_propagate(False)

        sep = ctk.CTkFrame(modal, height=1, fg_color="#E2E8F0", corner_radius=0)
        sep.pack(fill="x", side="bottom")

        # Main Body
        c = ctk.CTkFrame(modal, fg_color="transparent")
        c.pack(fill="both", expand=True, padx=22, pady=(14, 10))

        code = student["student_code"]
        name = student["full_name"]
        cls_grp = student.get("class_group", "")
        major = student.get("major", "")

        ctk.CTkLabel(
            c,
            text="✏️ ปรับแต่งค่าธรรมเนียมเฉพาะบุคคล (Dynamic)",
            font=(AppTheme.FONT_FAMILY, 16, "bold"),
            text_color="#0F172A",
        ).pack(anchor="w", pady=(0, 2))

        ctk.CTkLabel(
            c,
            text=f"นักศึกษา: {name} ({code})  |  ห้อง: {cls_grp} - {major}",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color="#64748B",
        ).pack(anchor="w", pady=(0, 10))

        std_amount = sum(float(i["line_total"]) for i in self.batch_base_items)

        card = ctk.CTkFrame(c, fg_color="#F8FAFC", corner_radius=8, border_width=1, border_color="#E2E8F0")
        card.pack(fill="both", expand=True, pady=(0, 6))

        opt_mode_var = ctk.StringVar(value="standard")
        custom_amt_var = ctk.StringVar(value=f"{std_amount:.2f}")
        custom_remark_var = ctk.StringVar(value="")

        if code in self.batch_custom_overrides:
            ovr = self.batch_custom_overrides[code]
            opt_mode_var.set(ovr.get("mode", "custom"))
            custom_amt_var.set(f"{ovr.get('amount', std_amount):.2f}")
            custom_remark_var.set(ovr.get("remark", ""))

        c_in = ctk.CTkFrame(card, fg_color="transparent")
        c_in.pack(fill="both", expand=True, padx=16, pady=10)

        entry_amt = make_entry(c_in, width=130, height=34, font_size=13)
        entry_rem = make_entry(c_in, height=34, font_size=13, placeholder_text="เช่น ทุนเรียนดี, ผ่อนชำระงวดแรก...")

        def on_choose_radio():
            m = opt_mode_var.get()
            if m == "standard":
                entry_amt.delete(0, "end")
                entry_amt.insert(0, f"{std_amount:.2f}")
                entry_rem.delete(0, "end")
                entry_rem.insert(0, "ค่าธรรมเนียมการศึกษามาตรฐาน")
            elif m == "free":
                entry_amt.delete(0, "end")
                entry_amt.insert(0, "0.00")
                entry_rem.delete(0, "end")
                entry_rem.insert(0, "ทุนการศึกษาเต็มจำนวน (100% ฟรี)")
            elif m == "half":
                entry_amt.delete(0, "end")
                entry_amt.insert(0, f"{std_amount / 2:.2f}")
                entry_rem.delete(0, "end")
                entry_rem.insert(0, "ทุนการศึกษา 50%")
            elif m == "custom":
                pass

        ctk.CTkRadioButton(
            c_in,
            text=f"1. ค่าเทอมมาตรฐานตามห้อง (฿{std_amount:,.2f} บาท)",
            variable=opt_mode_var,
            value="standard",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=on_choose_radio,
        ).pack(anchor="w", pady=3)

        ctk.CTkRadioButton(
            c_in,
            text="2. ทุนการศึกษาเต็มจำนวน 100% (ยกเว้นค่าเทอม = ฿0.00 บาท)",
            variable=opt_mode_var,
            value="free",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=on_choose_radio,
        ).pack(anchor="w", pady=3)

        ctk.CTkRadioButton(
            c_in,
            text=f"3. ทุนการศึกษา 50% (ชำระ ฿{std_amount / 2:,.2f} บาท)",
            variable=opt_mode_var,
            value="half",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=on_choose_radio,
        ).pack(anchor="w", pady=3)

        ctk.CTkRadioButton(
            c_in,
            text="4. กำหนดจำนวนเงินและหมายเหตุเอง (Custom Amount)",
            variable=opt_mode_var,
            value="custom",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=on_choose_radio,
        ).pack(anchor="w", pady=3)

        f_amt = ctk.CTkFrame(c_in, fg_color="transparent")
        f_amt.pack(fill="x", padx=16, pady=(6, 4))
        ctk.CTkLabel(f_amt, text="จำนวนเงินที่ต้องชำระ (บาท):", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color="#334155").pack(side="left", padx=(0, 8))
        entry_amt.pack(side="left")
        entry_amt.delete(0, "end")
        entry_amt.insert(0, custom_amt_var.get())

        f_rem = ctk.CTkFrame(c_in, fg_color="transparent")
        f_rem.pack(fill="x", padx=16, pady=(0, 6))
        ctk.CTkLabel(f_rem, text="หมายเหตุในใบเสร็จ:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color="#334155").pack(side="left", padx=(0, 8))
        entry_rem.pack(side="left", fill="x", expand=True)
        entry_rem.delete(0, "end")
        entry_rem.insert(0, custom_remark_var.get())

        def save_custom():
            m = opt_mode_var.get()
            if m == "standard":
                if code in self.batch_custom_overrides:
                    del self.batch_custom_overrides[code]
            else:
                try:
                    amt = float(entry_amt.get().strip().replace(",", "") or "0")
                except ValueError:
                    highlight_missing_field(entry_amt, "กรุณาระบุจำนวนเงินเป็นตัวเลข", show_toast_msg=True)
                    return
                rem = entry_rem.get().strip() or ("ทุนการศึกษา" if m != "custom" else "ค่าธรรมเนียมเฉพาะบุคคล")
                self.batch_custom_overrides[code] = {
                    "mode": m,
                    "amount": amt,
                    "remark": rem,
                }
            close_modal()
            self.render_batch_student_rows()
            self.recalculate_batch_totals()

        modal.bind("<Return>", lambda e: save_custom())
        modal.bind("<KP_Enter>", lambda e: save_custom())

        btn_save = ctk.CTkButton(
            b_act,
            text="💾 บันทึกการปรับแต่งสำหรับเด็กคนนี้",
            width=230,
            height=36,
            corner_radius=8,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=save_custom,
        )
        btn_save.pack(side="right", padx=(12, 20), pady=9)

        btn_cancel = ctk.CTkButton(
            b_act,
            text="ยกเลิก",
            width=95,
            height=36,
            corner_radius=8,
            fg_color="#FFFFFF",
            hover_color="#F1F5F9",
            border_color="#CBD5E1",
            border_width=1,
            text_color="#475569",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=close_modal,
        )
        btn_cancel.pack(side="right", padx=(0, 0), pady=9)

    def _open_batch_fee_preset_modal(self):
        """Allows selecting or customizing standard room tuition fees."""
        modal = ctk.CTkToplevel(self)
        modal.title("⚙️ กำหนดค่าเทอมมาตรฐานของห้อง")
        w, h = 560, 420
        modal.geometry(f"{w}x{h}")
        modal.resizable(False, False)
        modal.configure(fg_color="#FFFFFF")
        modal.transient(self.winfo_toplevel())

        try:
            from ui.icon_manager import IconManager
            IconManager.set_window_icon(modal)
        except Exception:
            pass

        try:
            modal.update_idletasks()
            top = self.winfo_toplevel()
            top_w = top.winfo_width()
            top_h = top.winfo_height()
            top_x = top.winfo_rootx() if hasattr(top, "winfo_rootx") else top.winfo_x()
            top_y = top.winfo_rooty() if hasattr(top, "winfo_rooty") else top.winfo_y()
            if top_w > 200 and top_h > 200:
                x = max(0, top_x + (top_w - w) // 2)
                y = max(0, top_y + (top_h - h) // 2)
            else:
                x = max(0, (modal.winfo_screenwidth() - w) // 2)
                y = max(0, (modal.winfo_screenheight() - h) // 2)
            modal.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            modal.geometry(f"{w}x{h}")

        def close_modal():
            try:
                modal.grab_release()
            except Exception:
                pass
            modal.destroy()

        modal.protocol("WM_DELETE_WINDOW", close_modal)
        modal.bind("<Escape>", lambda e: close_modal())
        modal.grab_set()
        modal.after(50, lambda: [modal.lift(), modal.focus_force()])

        # Top Accent Strip (Navy)
        accent_bar = ctk.CTkFrame(modal, height=4, fg_color="#1E40AF", corner_radius=0)
        accent_bar.pack(fill="x", side="top")

        # Bottom Button Footer Bar
        b_act = ctk.CTkFrame(modal, fg_color="#F8FAFC", height=54, corner_radius=0)
        b_act.pack(fill="x", side="bottom")
        b_act.pack_propagate(False)

        sep = ctk.CTkFrame(modal, height=1, fg_color="#E2E8F0", corner_radius=0)
        sep.pack(fill="x", side="bottom")

        # Main Body
        c = ctk.CTkFrame(modal, fg_color="transparent")
        c.pack(fill="both", expand=True, padx=22, pady=(14, 10))

        ctk.CTkLabel(
            c,
            text="⚙️ กำหนดค่าเทอมมาตรฐานของห้อง",
            font=(AppTheme.FONT_FAMILY, 16, "bold"),
            text_color="#0F172A",
        ).pack(anchor="w", pady=(0, 2))

        ctk.CTkLabel(
            c,
            text="เลือกชุดค่าธรรมเนียมที่จะนำไปใช้กับนักศึกษาทุกคนในห้องนี้เป็นค่าเริ่มต้น:",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color="#64748B",
        ).pack(anchor="w", pady=(0, 10))

        card = ctk.CTkFrame(c, fg_color="#F8FAFC", corner_radius=8, border_width=1, border_color="#E2E8F0")
        card.pack(fill="both", expand=True, pady=(0, 6))

        preset_choice_var = ctk.StringVar(value="VC1")

        options = [
            ("VC1", "1. ค่าเทอม ปวช.1 (฿1,130.00 - สิ่งแวดล้อม 330, เน็ต 300, ครู 500)"),
            ("VC2", "2. ค่าเทอม ปวช.2-3 (฿1,130.00)"),
            ("HVC1", "3. ค่าเทอม ปวส.1 (฿1,550.00 - บำรุงการศึกษา 650, เน็ต 400, ครู 500)"),
            ("CART", f"4. ใช้รายการในตารางจากโหมดปกติ (มี {len(self.cart_items)} รายการ รวม ฿{sum(float(i['line_total']) for i in self.cart_items):,.2f} บ.)"),
            ("CUSTOM", "5. ระบุยอดเงินเหมาจ่ายเอง"),
        ]

        c_in = ctk.CTkFrame(card, fg_color="transparent")
        c_in.pack(fill="both", expand=True, padx=16, pady=10)

        for val, lbl in options:
            ctk.CTkRadioButton(
                c_in,
                text=lbl,
                variable=preset_choice_var,
                value=val,
                font=(AppTheme.FONT_FAMILY, 12, "bold"),
            ).pack(anchor="w", pady=3)

        f_cust = ctk.CTkFrame(c_in, fg_color="transparent")
        f_cust.pack(fill="x", padx=16, pady=(4, 2))
        ctk.CTkLabel(f_cust, text="ยอดเงิน (บาท):", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color="#334155").pack(side="left", padx=(0, 6))
        entry_c_amt = make_entry(f_cust, width=120, height=34, font_size=13)
        entry_c_amt.pack(side="left")
        entry_c_amt.insert(0, "1130.00")

        def apply_preset():
            ch = preset_choice_var.get()
            if ch in ("VC1", "VC2"):
                self.batch_base_items = [
                    {"fee_code": "ENV", "curriculum_name": "หลักสูตรในระบบ-ปวช", "item_name": "ค่ารักษาภาพแวดล้อม", "quantity": 0.0, "unit_price": 330.0, "discount": 0.0, "line_total": 330.0},
                    {"fee_code": "NET", "curriculum_name": "หลักสูตรในระบบ-ปวช", "item_name": "ค่าใช้บริการอินเตอร์เน็ต", "quantity": 0.0, "unit_price": 300.0, "discount": 0.0, "line_total": 300.0},
                    {"fee_code": "TCH", "curriculum_name": "หลักสูตรในระบบ-ปวช", "item_name": "ค่าจ้างครูหรือผู้ที่มีความเชี่ยวชาญในสาขาเฉพาะ", "quantity": 0.0, "unit_price": 500.0, "discount": 0.0, "line_total": 500.0},
                ]
            elif ch == "HVC1":
                self.batch_base_items = [
                    {"fee_code": "MAINT", "curriculum_name": "หลักสูตรในระบบ-ปวส", "item_name": "ค่าบำรุงการศึกษา", "quantity": 0.0, "unit_price": 650.0, "discount": 0.0, "line_total": 650.0},
                    {"fee_code": "NET_HVC", "curriculum_name": "หลักสูตรในระบบ-ปวส", "item_name": "ค่าบริการอินเตอร์เน็ตและเทคโนโลยี", "quantity": 0.0, "unit_price": 400.0, "discount": 0.0, "line_total": 400.0},
                    {"fee_code": "SPEC_TCH", "curriculum_name": "หลักสูตรในระบบ-ปวส", "item_name": "ค่าจ้างครูผู้เชี่ยวชาญเฉพาะทาง", "quantity": 0.0, "unit_price": 500.0, "discount": 0.0, "line_total": 500.0},
                ]
            elif ch == "CART":
                if not self.cart_items:
                    messagebox.showwarning("ไม่มีรายการ", "ยังไม่มีรายการในตารางโหมดปกติ กรุณาเพิ่มรายการก่อน")
                    return
                self.batch_base_items = [dict(i) for i in self.cart_items]
            elif ch == "CUSTOM":
                try:
                    amt = float(entry_c_amt.get().strip().replace(",", "") or "0")
                except ValueError:
                    highlight_missing_field(entry_c_amt, "กรุณาระบุจำนวนเงินเป็นตัวเลข", show_toast_msg=True)
                    return
                self.batch_base_items = [
                    {"fee_code": "CUSTOM_BATCH", "curriculum_name": self.batch_room_var.get(), "item_name": "ชำระค่าธรรมเนียมการศึกษา", "quantity": 1.0, "unit_price": amt, "discount": 0.0, "line_total": amt}
                ]

            tot = sum(float(i["line_total"]) for i in self.batch_base_items)
            items_desc = ", ".join(f"{i['item_name']} (฿{float(i['line_total']):,.0f})" for i in self.batch_base_items[:2])
            if len(self.batch_base_items) > 2:
                items_desc += f" และอีก {len(self.batch_base_items) - 2} รายการ"
            self.lbl_batch_base_fee_info.configure(
                text=f"📌 ค่าเทอมมาตรฐานห้อง: {items_desc} | รวม ฿{tot:,.2f} บาท"
            )
            close_modal()
            self.render_batch_student_rows()
            self.recalculate_batch_totals()

        modal.bind("<Return>", lambda e: apply_preset())
        modal.bind("<KP_Enter>", lambda e: apply_preset())

        btn_save = ctk.CTkButton(
            b_act,
            text="💾 ใช้ชุดค่าเทอมนี้กับห้อง",
            width=200,
            height=36,
            corner_radius=8,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=apply_preset,
        )
        btn_save.pack(side="right", padx=(12, 20), pady=9)

        btn_cancel = ctk.CTkButton(
            b_act,
            text="ยกเลิก",
            width=95,
            height=36,
            corner_radius=8,
            fg_color="#FFFFFF",
            hover_color="#F1F5F9",
            border_color="#CBD5E1",
            border_width=1,
            text_color="#475569",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=close_modal,
        )
        btn_cancel.pack(side="right", padx=(0, 0), pady=9)

    def execute_batch_roster_issue(self):
        """Issues official receipts for all checked students in the classroom roster with their specific dynamic fees."""
        base_amount = sum(float(i["line_total"]) for i in self.batch_base_items)
        selected_students = []
        for s in self.batch_students_data:
            code = s["student_code"]
            if self.batch_check_vars.get(code) and self.batch_check_vars[code].get():
                s_copy = dict(s)
                if code in self.batch_custom_overrides:
                    ovr = self.batch_custom_overrides[code]
                    if "items" in ovr and ovr["items"]:
                        s_copy["line_items"] = ovr["items"]
                    elif "amount" in ovr:
                        s_copy["custom_net_total"] = ovr["amount"]
                        s_copy["custom_remark"] = ovr.get("remark", "ค่าธรรมเนียมการศึกษาเฉพาะบุคคล")
                selected_students.append(s_copy)

        if not selected_students:
            messagebox.showwarning("ยังไม่ได้เลือกนักศึกษา", "กรุณากดเลือกนักศึกษาในห้องอย่างน้อย 1 คนก่อนออกใบเสร็จ")
            return

        room = self.batch_room_var.get()
        tot_amt = sum(float(s.get("custom_net_total", base_amount)) for s in selected_students)

        if not messagebox.askyesno(
            "ยืนยันการออกใบเสร็จยกห้อง",
            f"ยืนยันการออกใบเสร็จให้นักศึกษาห้อง '{room}'\n"
            f"จำนวนที่เลือก: {len(selected_students)} คน\n"
            f"ยอดเงินรวมทั้งสิ้น: ฿{tot_amt:,.2f} บาท ({thai_baht_text(tot_amt)})\n\n"
            f"ต้องการบันทึกและสร้างไฟล์ PDF รวมทั้งห้องใช่หรือไม่?",
        ):
            return

        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)
        safe_tag = room.replace("/", "-").replace(" ", "_")
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"BatchReceipts_{safe_tag}_{len(selected_students)}students_{stamp}.pdf"

        save_path = filedialog.asksaveasfilename(
            title="เลือกตำแหน่งบันทึกไฟล์ใบเสร็จยกห้อง (.pdf)",
            initialdir=out_dir,
            initialfile=default_filename,
            defaultextension=".pdf",
            filetypes=[
                ("PDF Document (*.pdf)", "*.pdf"),
                ("Word Document (*.docx)", "*.docx"),
            ],
        )
        if not save_path:
            messagebox.showinfo(
                "ยกเลิกการออกใบเสร็จยกห้อง",
                "คุณได้ยกเลิกการเลือกตำแหน่งบันทึกไฟล์ เอกสารจึงยังไม่ถูกบันทึกหรือพิมพ์",
                parent=self.winfo_toplevel(),
            )
            return

        self.config["default_export_folder"] = os.path.dirname(save_path)
        save_config(self.config)
        self._refresh_export_folder_buttons()

        while True:
            try:
                issued_list = self.db.issue_batch_class_receipts(
                    students_list=selected_students,
                    line_items=self.batch_base_items,
                    payment_method=self.batch_pay_method_var.get(),
                    actor=self.session.username,
                    cashier_name=self.session.full_name,
                    role=self.session.role,
                    term_code=self.batch_term_var.get(),
                    book_no=self.config.get("receipt_book_no", "7006"),
                    cashier_position=self.config.get("cashier_position", "ครู"),
                    receipt_prefix=self.config.get("receipt_prefix", ""),
                    start_no=self.config.get("receipt_no_start", "700605433"),
                    end_no=self.config.get("receipt_no_end", "700999999"),
                    receipt_date_override=format_thai_date_short(),
                )
                break
            except ReceiptNumberError as rne:
                if not self._apply_receipt_number_resolution(rne):
                    return
            except Exception as e:
                messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถออกใบเสร็จยกห้องได้:\n{e}", parent=self.winfo_toplevel())
                return

        try:
            export_document_auto(
                save_path,
                issued_list,
                college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
                metadata=self.config,
                is_batch=True,
            )

            first_no = issued_list[0]["receipt_no"]
            last_no = issued_list[-1]["receipt_no"]
            used_pfx, used_num = split_receipt_prefix_and_number(last_no)
            next_num = format_next_receipt_number(used_num)
            self.config["receipt_prefix"] = used_pfx
            self.config["receipt_no_start"] = next_num
            save_config(self.config)
            self.refresh_next_receipt_no()
            self.refresh_batch_roster()

            open_ok = open_document_file(save_path, parent=self.winfo_toplevel())

            messagebox.showinfo(
                "ออกใบเสร็จยกห้องสำเร็จ!",
                f"ออกใบเสร็จให้นักศึกษาห้อง {room} สำเร็จครบถ้วน!\n\n"
                f"จำนวนที่ออก: {len(issued_list)} ฉบับ\n"
                f"เลขที่ใบเสร็จ: {first_no} ถึง {last_no}\n"
                f"ยอดรวมทั้งสิ้น: ฿{tot_amt:,.2f} บาท\n\n"
                f"ไฟล์ใบเสร็จรวมพร้อมพิมพ์อยู่ที่:\n{save_path}\n\n"
                + ("(ระบบได้เปิดไฟล์เอกสารขึ้นมาให้พร้อมพิมพ์เรียบร้อยแล้ว)" if open_ok else ""),
                parent=self.winfo_toplevel(),
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถออกใบเสร็จยกห้องได้:\n{e}", parent=self.winfo_toplevel())

    # ------------------------------------------------------------------
    # ADVANCED MODE PANEL (ออกแบบแม่แบบ & ตั้งค่าระบบขั้นสูง)
    # ------------------------------------------------------------------

    def _build_advanced_panel(self):
        """Constructs modern smart features into clean, modular cards inside Advanced Mode."""
        # Banner Header
        banner = ctk.CTkFrame(self.panel_advanced, corner_radius=AppTheme.RADIUS, fg_color=AppTheme.COLOR_SURFACE, border_width=1, border_color=AppTheme.COLOR_BORDER)
        banner.pack(fill="x", padx=12, pady=(4, 10))
        b_in = ctk.CTkFrame(banner, fg_color="transparent")
        b_in.pack(fill="x", padx=16, pady=10)
        ctk.CTkLabel(
            b_in,
            text="⚡ ออกแบบใบเสร็จรับเงิน & ระบบจัดการการเงินขั้นสูง (Advanced System)",
            font=(AppTheme.FONT_FAMILY, 16, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        ).pack(anchor="w")
        ctk.CTkLabel(
            b_in,
            text="รวมฟีเจอร์ทันสมัย: ออกใบเสร็จยกห้อง, นำเข้าไฟล์ Excel, ปรับแต่งแม่แบบใบเสร็จ, สแกน QR PromptPay, และ Webhook",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color=AppTheme.COLOR_TEXT_MUTED,
        ).pack(anchor="w")

        # ------------------------------------------------------------------
        # CARD 1: 🏫 ออกใบเสร็จแบบเดียวยกห้อง (Batch Classroom Receipts)
        # ------------------------------------------------------------------
        card1 = ctk.CTkFrame(self.panel_advanced, corner_radius=AppTheme.RADIUS, fg_color=AppTheme.COLOR_SURFACE, border_width=1, border_color=AppTheme.COLOR_BORDER)
        card1.pack(fill="x", padx=12, pady=(0, 10))
        c1_in = ctk.CTkFrame(card1, fg_color="transparent")
        c1_in.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(
            c1_in,
            text="🏫 1. โหมดออกใบเสร็จแบบเดียวยกห้อง (Batch Classroom Receipts)",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(anchor="w")
        ctk.CTkLabel(
            c1_in,
            text="นำรายการค่าธรรมเนียมในตารางใบเสร็จปัจจุบัน ไปออกใบเสร็จให้นักเรียนทั้งห้องเรียนพร้อมกัน และรวมเป็นไฟล์ PDF เดียวกัน",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color=AppTheme.COLOR_TEXT_MUTED,
        ).pack(anchor="w", pady=(2, 8))

        self.lbl_batch_cart_info = ctk.CTkLabel(
            c1_in,
            text="รายการในตาราง: 0 รายการ  |  ยอดรวมคนละ: ฿0.00 บาท",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        )
        self.lbl_batch_cart_info.pack(anchor="w", pady=(0, 8))

        r_cls = ctk.CTkFrame(c1_in, fg_color="transparent")
        r_cls.pack(fill="x")

        ctk.CTkLabel(
            r_cls,
            text="เลือกห้องเรียน:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left", padx=(0, 6))

        cls_opts = [c for c in self.class_levels if c != "ทั้งหมด"] or ["ปวช.1/1"]
        self.batch_cls_var = ctk.StringVar(value=cls_opts[0])
        self.opt_batch_class = make_dropdown(
            r_cls,
            values=cls_opts,
            variable=self.batch_cls_var,
            width=160,
            height=38,
            font_size=13,
            command=self._on_change_batch_class,
        )
        self.opt_batch_class.pack(side="left", padx=(0, 12))

        self.lbl_batch_student_count = ctk.CTkLabel(
            r_cls,
            text="จำนวนนักศึกษาในห้องนี้: ... คน",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        )
        self.lbl_batch_student_count.pack(side="left", padx=(0, 16))

        ctk.CTkButton(
            r_cls,
            text="🖨️ ออกใบเสร็จให้นักศึกษาทั้งห้องทันที (.pdf)",
            height=38,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.run_batch_for_selected_room,
        ).pack(side="right")

        # ------------------------------------------------------------------
        # CARD 2: 📥 นำเข้าไฟล์รายชื่อนักศึกษาจาก Excel/CSV แล้วออกใบเสร็จทันที
        # ------------------------------------------------------------------
        card2 = ctk.CTkFrame(self.panel_advanced, corner_radius=AppTheme.RADIUS, fg_color=AppTheme.COLOR_SURFACE, border_width=1, border_color=AppTheme.COLOR_BORDER)
        card2.pack(fill="x", padx=12, pady=(0, 10))
        c2_in = ctk.CTkFrame(card2, fg_color="transparent")
        c2_in.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(
            c2_in,
            text="📥 2. นำเข้ารายชื่อนักศึกษาจากไฟล์ (Excel .xlsx / .csv) ➔ ออกใบเสร็จยกชุดทันที",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(anchor="w")
        ctk.CTkLabel(
            c2_in,
            text="รองรับไฟล์ Excel (.xlsx) และ CSV รายชื่อเด็ก (คอลัมน์: รหัสนักศึกษา, ชื่อ-สกุล, แผนกวิชา, ระดับชั้น)",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color=AppTheme.COLOR_TEXT_MUTED,
        ).pack(anchor="w", pady=(2, 8))

        ctk.CTkButton(
            c2_in,
            text="📥 เลือกไฟล์ Excel/CSV รายชื่อเด็ก ➔ บันทึกและออกใบเสร็จยกชุดทันที",
            height=38,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.import_and_issue_batch,
        ).pack(fill="x")

        # ------------------------------------------------------------------
        # CARD 3: 🎨 ออกแบบ & ตั้งค่าแม่แบบใบเสร็จ (Receipt Template Designer)
        # ------------------------------------------------------------------
        card3 = ctk.CTkFrame(self.panel_advanced, corner_radius=AppTheme.RADIUS, fg_color=AppTheme.COLOR_SURFACE, border_width=1, border_color=AppTheme.COLOR_BORDER)
        card3.pack(fill="x", padx=12, pady=(0, 10))
        c3_in = ctk.CTkFrame(card3, fg_color="transparent")
        c3_in.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(
            c3_in,
            text="🎨 3. ออกแบบและตั้งค่าแม่แบบเอกสารใบเสร็จ (Template Designer)",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(anchor="w", pady=(0, 8))

        grid3 = ctk.CTkFrame(c3_in, fg_color="transparent")
        grid3.pack(fill="x")
        grid3.grid_columnconfigure(0, weight=0, minsize=145)
        grid3.grid_columnconfigure(1, weight=1, uniform="adv_grid")
        grid3.grid_columnconfigure(2, weight=0, minsize=145)
        grid3.grid_columnconfigure(3, weight=1, uniform="adv_grid")

        # Row 0: สไตล์ใบเสร็จ + ชื่อสถาบัน
        ctk.CTkLabel(grid3, text="สไตล์ใบเสร็จ:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY, anchor="w").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.adv_tpl_style_var = ctk.StringVar(value=_style_label_from_code(self.config.get("bill_template_style", "CLASSIC")))
        self.opt_adv_tpl_style = make_dropdown(
            grid3,
            values=TEMPLATE_STYLE_OPTIONS,
            variable=self.adv_tpl_style_var,
            width=210,
            height=32,
            font_size=13,
            command=self._on_change_template_style,
        )
        self.opt_adv_tpl_style.grid(row=0, column=1, sticky="w", padx=(0, 16), pady=4)

        ctk.CTkLabel(grid3, text="ชื่อที่ทำการ/วิทยาลัย:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY, anchor="w").grid(row=0, column=2, sticky="w", padx=(0, 8), pady=4)
        self.adv_college_name = make_entry(grid3, height=32, font_size=13)
        self.adv_college_name.grid(row=0, column=3, sticky="ew", pady=4)
        self.adv_college_name.insert(0, self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"))

        # Row 1: สังกัดราชการ + เล่มที่ใบเสร็จ
        ctk.CTkLabel(grid3, text="สังกัดราชการ:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY, anchor="w").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.adv_gov_title = make_entry(grid3, height=32, font_size=13)
        self.adv_gov_title.grid(row=1, column=1, sticky="ew", padx=(0, 16), pady=4)
        self.adv_gov_title.insert(0, self.config.get("receipt_gov_title", "ในราชการกรมอาชีวศึกษา กระทรวงศึกษาธิการ"))

        ctk.CTkLabel(grid3, text="เล่มที่ใบเสร็จ (Book No):", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY, anchor="w").grid(row=1, column=2, sticky="w", padx=(0, 8), pady=4)
        self.adv_book_no = make_entry(grid3, width=130, height=32, font_size=13)
        self.adv_book_no.grid(row=1, column=3, sticky="w", pady=4)
        self.adv_book_no.insert(0, self.config.get("receipt_book_no", "7006"))

        # Row 2: ช่วงเลขที่เริ่มต้น + ช่วงเลขที่สิ้นสุด
        ctk.CTkLabel(grid3, text="เลขที่เริ่มต้น:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY, anchor="w").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        self.adv_no_start = make_entry(grid3, width=160, height=32, font_size=13)
        self.adv_no_start.grid(row=2, column=1, sticky="w", padx=(0, 16), pady=4)
        self.adv_no_start.insert(0, self.config.get("receipt_no_start", "700605433"))

        ctk.CTkLabel(grid3, text="เลขที่สิ้นสุด:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY, anchor="w").grid(row=2, column=2, sticky="w", padx=(0, 8), pady=4)
        self.adv_no_end = make_entry(grid3, width=160, height=32, font_size=13)
        self.adv_no_end.grid(row=2, column=3, sticky="w", pady=4)
        self.adv_no_end.insert(0, self.config.get("receipt_no_end", "700999999"))

        # Row 3: ชื่อผู้รับเงิน + ตำแหน่งผู้ลงนาม
        ctk.CTkLabel(grid3, text="ผู้รับเงินในใบเสร็จ:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY, anchor="w").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        self.adv_cashier_name = make_entry(grid3, height=32, font_size=13)
        self.adv_cashier_name.grid(row=3, column=1, sticky="ew", padx=(0, 16), pady=4)
        self.adv_cashier_name.insert(0, self.session.full_name)

        ctk.CTkLabel(grid3, text="ตำแหน่งผู้ลงนาม:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY, anchor="w").grid(row=3, column=2, sticky="w", padx=(0, 8), pady=4)
        self.adv_cashier_pos = make_entry(grid3, width=180, height=32, font_size=13)
        self.adv_cashier_pos.grid(row=3, column=3, sticky="w", pady=4)
        self.adv_cashier_pos.grid(row=3, column=3, sticky="ew", pady=4)
        self.adv_cashier_pos.insert(0, self.config.get("cashier_position", "ครู"))

        ctk.CTkButton(
            c3_in,
            text="💾 บันทึกการตั้งค่าแม่แบบใบเสร็จ",
            height=36,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.save_template_settings,
        ).pack(anchor="e", pady=(10, 0))

        # ------------------------------------------------------------------
        # CARD 4: 💳 ช่องทางชำระเงิน & QR PromptPay อัจฉริยะ
        # ------------------------------------------------------------------
        card4 = ctk.CTkFrame(self.panel_advanced, corner_radius=AppTheme.RADIUS, fg_color=AppTheme.COLOR_SURFACE, border_width=1, border_color=AppTheme.COLOR_BORDER)
        card4.pack(fill="x", padx=12, pady=(0, 10))
        c4_in = ctk.CTkFrame(card4, fg_color="transparent")
        c4_in.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(
            c4_in,
            text="💳 4. ช่องทางชำระเงิน & ระบบสแกน QR PromptPay",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(anchor="w", pady=(0, 8))

        grid4 = ctk.CTkFrame(c4_in, fg_color="transparent")
        grid4.pack(fill="x")
        grid4.grid_columnconfigure(0, weight=0, minsize=145)
        grid4.grid_columnconfigure(1, weight=1, uniform="adv_grid")
        grid4.grid_columnconfigure(2, weight=0, minsize=145)
        grid4.grid_columnconfigure(3, weight=1, uniform="adv_grid")

        ctk.CTkLabel(grid4, text="ช่องทางชำระเงิน:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY, anchor="w").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.opt_pay_method = make_dropdown(
            grid4,
            values=["เงินสด (Cash)", "สแกน QR PromptPay", "โอนเข้าบัญชีวิทยาลัย"],
            variable=self.pay_method_var,
            width=210,
            height=32,
            font_size=13,
        )
        self.opt_pay_method.grid(row=0, column=1, sticky="w", padx=(0, 16), pady=4)

        ctk.CTkLabel(grid4, text="หมายเลข PromptPay:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY, anchor="w").grid(row=0, column=2, sticky="w", padx=(0, 8), pady=4)
        self.adv_promptpay_no = make_entry(grid4, width=180, height=32, font_size=13)
        self.adv_promptpay_no.grid(row=0, column=3, sticky="w", pady=4)
        self.adv_promptpay_no.insert(0, self.config.get("promptpay_no", "0994000160000"))

        ctk.CTkLabel(grid4, text="เลขที่บัญชีวิทยาลัย:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY, anchor="w").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.adv_bank_account = make_entry(grid4, height=32, font_size=13)
        self.adv_bank_account.grid(row=1, column=1, sticky="ew", padx=(0, 16), pady=4)
        self.adv_bank_account.insert(0, self.config.get("bank_account", "ธนาคารกรุงไทย 710-0-12345-6"))

        self.adv_show_qr_var = ctk.BooleanVar(value=self.config.get("show_qr_on_receipt", True))
        ctk.CTkSwitch(
            grid4,
            text="พิมพ์ QR Code ชำระเงินบนใบเสร็จรับเงิน",
            variable=self.adv_show_qr_var,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).grid(row=1, column=3, sticky="w", pady=4)

        # ------------------------------------------------------------------
        # CARD 5: ⚡ จัดชุดค่าธรรมเนียมเหมาจ่ายด่วนตามหลักสูตร (Fee Preset Bundles)
        # ------------------------------------------------------------------
        card5 = ctk.CTkFrame(self.panel_advanced, corner_radius=AppTheme.RADIUS, fg_color=AppTheme.COLOR_SURFACE, border_width=1, border_color=AppTheme.COLOR_BORDER)
        card5.pack(fill="x", padx=12, pady=(0, 10))
        c5_in = ctk.CTkFrame(card5, fg_color="transparent")
        c5_in.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(
            c5_in,
            text="⚡ 5. ดึงชุดค่าธรรมเนียมเหมาจ่ายด่วนตามหลักสูตร (Auto-Fee Presets)",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(anchor="w")
        ctk.CTkLabel(
            c5_in,
            text="คลิกปุ่มด่วนเพื่อนำรายการค่าเทอมมาตรฐานของหลักสูตรมาใส่ในตารางใบเสร็จทันทีในคลิกเดียว",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color=AppTheme.COLOR_TEXT_MUTED,
        ).pack(anchor="w", pady=(2, 8))

        p_btns = ctk.CTkFrame(c5_in, fg_color="transparent")
        p_btns.pack(fill="x")

        presets = [
            ("⚡ ค่าเทอม ปวช.1 (฿1,130 - ตามแบบ EX)", "VC1"),
            ("⚡ ค่าเทอม ปวช.2-3 (฿1,130)", "VC2"),
            ("⚡ ค่าเทอม ปวส.1 (฿1,550)", "HVC1"),
            ("⚡ ค่าวัสดุฝึกสำเร็จรูป (฿600)", "WORKSHOP"),
        ]
        for lbl, pkey in presets:
            ctk.CTkButton(
                p_btns,
                text=lbl,
                height=36,
                corner_radius=AppTheme.RADIUS,
                fg_color=AppTheme.COLOR_SECONDARY_BG,
                hover_color=AppTheme.COLOR_SECONDARY_HOVER,
                border_width=1,
                border_color=AppTheme.COLOR_SECONDARY_BORDER,
                text_color=AppTheme.COLOR_SECONDARY_TEXT,
                font=(AppTheme.FONT_FAMILY, 12, "bold"),
                command=lambda k=pkey: self.load_fee_preset(k),
            ).pack(side="left", padx=(0, 8), fill="x", expand=True)

        # ------------------------------------------------------------------
        # CARD 6: 🔔 ระบบส่งแจ้งเตือน Webhook อัตโนมัติ (Automated Webhook Dispatcher)
        # ------------------------------------------------------------------
        card6 = ctk.CTkFrame(self.panel_advanced, corner_radius=AppTheme.RADIUS, fg_color=AppTheme.COLOR_SURFACE, border_width=1, border_color=AppTheme.COLOR_BORDER)
        card6.pack(fill="x", padx=12, pady=(0, 12))
        c6_in = ctk.CTkFrame(card6, fg_color="transparent")
        c6_in.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(
            c6_in,
            text="🔔 6. ระบบส่งแจ้งเตือนธุรกรรมการเงินอัตโนมัติ (Webhook & Alert)",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(anchor="w", pady=(0, 6))

        grid6 = ctk.CTkFrame(c6_in, fg_color="transparent")
        grid6.pack(fill="x")
        grid6.columnconfigure(1, weight=1)

        self.adv_webhook_enabled_var = ctk.BooleanVar(value=self.config.get("webhook_enabled", False))
        ctk.CTkSwitch(
            grid6,
            text="เปิดใช้งานระบบส่งแจ้งเตือน Webhook (LINE Notify / Discord / n8n)",
            variable=self.adv_webhook_enabled_var,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

        ctk.CTkLabel(grid6, text="Webhook URL:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT, anchor="e", width=120).grid(row=1, column=0, sticky="e", padx=(0, 8), pady=4)
        self.adv_webhook_url = make_entry(grid6, height=36, font_size=13)
        self.adv_webhook_url.grid(row=1, column=1, sticky="ew", pady=4)
        self.adv_webhook_url.insert(0, self.config.get("webhook_url", ""))

        ctk.CTkLabel(grid6, text="เกณฑ์แจ้งเตือน (บาท):", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT, anchor="e", width=120).grid(row=2, column=0, sticky="e", padx=(0, 8), pady=4)
        w_sub = ctk.CTkFrame(grid6, fg_color="transparent")
        w_sub.grid(row=2, column=1, sticky="ew", pady=4)

        self.adv_webhook_threshold = make_entry(w_sub, width=130, height=36, font_size=13)
        self.adv_webhook_threshold.pack(side="left", padx=(0, 10))
        self.adv_webhook_threshold.insert(0, f"{float(self.config.get('webhook_alert_threshold', 50000.0)):,.2f}")

        ctk.CTkButton(
            w_sub,
            text="🔔 ทดสอบส่งสัญญาณ Webhook (Ping Test)",
            height=36,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.test_webhook_ping,
        ).pack(side="left")

        # Bottom Action: Return to Normal Mode
        ctk.CTkButton(
            self.panel_advanced,
            text="⬅️ บันทึกการตั้งค่า & สลับกลับไปหน้าออกใบเสร็จปกติ (กรอกมือตามแบบ EX)",
            height=42,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=lambda: self.set_mode("normal"),
        ).pack(fill="x", padx=12, pady=(10, 24))

    # ------------------------------------------------------------------
    # EVENT HANDLERS & CORE POS LOGIC
    # ------------------------------------------------------------------
    def _on_change_template_style(self, selected_label: str):
        self.config["bill_template_style"] = normalize_template_style(selected_label)
        save_config(self.config)

    def get_current_receipt_no(self) -> str:
        """Returns the current receipt number string from the single receipt entry."""
        if hasattr(self, "entry_rcp_no"):
            return self.entry_rcp_no.get().strip()
        pfx = self.config.get("receipt_prefix", "")
        num = self.config.get("receipt_no_start", "700605433")
        return combine_receipt_no(pfx, num)

    def _on_receipt_no_focusout(self, _event=None):
        """When cashier edits receipt number, parse prefix/number and remember in config."""
        if not hasattr(self, "entry_rcp_no"):
            return
        raw_val = self.entry_rcp_no.get().strip()
        if not raw_val:
            return
        pfx, num = split_receipt_prefix_and_number(raw_val)
        self.config["receipt_prefix"] = pfx
        if num:
            self.config["receipt_no_start"] = num
        save_config(self.config)

    def refresh_next_receipt_no(self):
        prefix = self.config.get("receipt_prefix", "")
        start_no = self.config.get("receipt_no_start", "700605433")
        end_no = self.config.get("receipt_no_end", "700999999")
        nxt = self.db.get_next_receipt_no(prefix, start_no=start_no, end_no=end_no)

        if hasattr(self, "entry_rcp_no"):
            self.entry_rcp_no.delete(0, "end")
            self.entry_rcp_no.insert(0, nxt)

    def _apply_receipt_number_resolution(self, err: ReceiptNumberError, parent=None) -> bool:
        """Shows one-click resolution dialog for ReceiptNumberError and applies the fix. Returns True to retry."""
        action, value = ask_receipt_number_resolution(parent or self.winfo_toplevel(), err)
        if action == "cancel" or not value:
            return False
        if action == "extend_range":
            self.config["receipt_no_end"] = value
            save_config(self.config)
            if hasattr(self, "adv_no_end"):
                try:
                    self.adv_no_end.delete(0, "end")
                    self.adv_no_end.insert(0, value)
                except Exception:
                    pass
            return True
        if action in ("use_next", "reset_low"):
            pfx, num = split_receipt_prefix_and_number(value)
            self.config["receipt_prefix"] = pfx
            self.config["receipt_no_start"] = num or value
            save_config(self.config)
            if hasattr(self, "entry_rcp_no"):
                self.entry_rcp_no.delete(0, "end")
                self.entry_rcp_no.insert(0, value)
            if hasattr(self, "adv_no_start"):
                try:
                    self.adv_no_start.delete(0, "end")
                    self.adv_no_start.insert(0, num or value)
                except Exception:
                    pass
            return True
        return False

    def open_receipt_no_config_modal(self):
        """Opens modal dialog to customize receipt prefix, starting number, and auto-increment."""
        modal = ctk.CTkToplevel(self)
        modal.title("⚙️ กำหนดรูปแบบเลขที่และตัวอักษรออกบิล (Receipt Numbering Setup)")
        w, h = 540, 395
        modal.geometry(f"{w}x{h}")
        modal.resizable(False, False)
        modal.configure(fg_color="#FFFFFF")
        modal.transient(self.winfo_toplevel())

        try:
            from ui.icon_manager import IconManager
            IconManager.set_window_icon(modal)
        except Exception:
            pass

        try:
            modal.update_idletasks()
            top = self.winfo_toplevel()
            top_w = top.winfo_width()
            top_h = top.winfo_height()
            top_x = top.winfo_rootx() if hasattr(top, "winfo_rootx") else top.winfo_x()
            top_y = top.winfo_rooty() if hasattr(top, "winfo_rooty") else top.winfo_y()
            if top_w > 200 and top_h > 200:
                x = max(0, top_x + (top_w - w) // 2)
                y = max(0, top_y + (top_h - h) // 2)
            else:
                x = max(0, (modal.winfo_screenwidth() - w) // 2)
                y = max(0, (modal.winfo_screenheight() - h) // 2)
            modal.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            modal.geometry(f"{w}x{h}")

        def close_modal():
            try:
                modal.grab_release()
            except Exception:
                pass
            modal.destroy()

        modal.protocol("WM_DELETE_WINDOW", close_modal)
        modal.bind("<Escape>", lambda e: close_modal())
        modal.grab_set()
        modal.after(50, lambda: [modal.lift(), modal.focus_force()])

        # Top Accent Strip (4px)
        accent_bar = ctk.CTkFrame(modal, height=4, fg_color="#1E40AF", corner_radius=0)
        accent_bar.pack(fill="x", side="top")

        # Bottom Button Footer Bar
        b_bar = ctk.CTkFrame(modal, fg_color="#F8FAFC", height=54, corner_radius=0)
        b_bar.pack(fill="x", side="bottom")
        b_bar.pack_propagate(False)

        sep = ctk.CTkFrame(modal, height=1, fg_color="#E2E8F0", corner_radius=0)
        sep.pack(fill="x", side="bottom")

        # Main Body
        c = ctk.CTkFrame(modal, fg_color="transparent")
        c.pack(fill="both", expand=True, padx=24, pady=(16, 10))

        ctk.CTkLabel(
            c,
            text="⚙️ กำหนดรูปแบบเลขที่และตัวอักษรออกบิล",
            font=(AppTheme.FONT_FAMILY, 16, "bold"),
            text_color="#0F172A",
        ).pack(anchor="w", pady=(0, 2))

        ctk.CTkLabel(
            c,
            text="คุณสามารถเลือกตัวอักษรนำหน้า (Prefix) และเลขเริ่มต้น ระบบจะรันเลขอัตโนมัติเมื่อออกบิล",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color="#64748B",
            justify="left",
        ).pack(anchor="w", pady=(0, 12))

        f_grid = ctk.CTkFrame(c, fg_color="transparent")
        f_grid.pack(fill="x", pady=(0, 10))
        f_grid.columnconfigure(1, weight=1)

        # Row 0: Prefix
        ctk.CTkLabel(
            f_grid,
            text="ตัวอักษร / คำนำหน้า:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color="#334155",
            anchor="e",
            width=155,
        ).grid(row=0, column=0, sticky="e", padx=(0, 12), pady=4)

        cur_full = self.get_current_receipt_no()
        parsed_pfx, parsed_num = split_receipt_prefix_and_number(cur_full)
        cur_pfx = parsed_pfx or self.config.get("receipt_prefix", "")
        ent_m_pfx = make_entry(f_grid, width=120, height=34, font_size=13, bold=True)
        ent_m_pfx.grid(row=0, column=1, sticky="w", pady=4)
        ent_m_pfx.insert(0, cur_pfx)

        # Row 1: Start No
        ctk.CTkLabel(
            f_grid,
            text="เลขที่เริ่มต้น / รันต่อจาก:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color="#334155",
            anchor="e",
            width=155,
        ).grid(row=1, column=0, sticky="e", padx=(0, 12), pady=4)

        cur_num = parsed_num or self.config.get("receipt_no_start", "700605433")
        ent_m_num = make_entry(f_grid, width=180, height=34, font_size=13, bold=True)
        ent_m_num.grid(row=1, column=1, sticky="w", pady=4)
        ent_m_num.insert(0, cur_num)

        # Row 2: End No
        ctk.CTkLabel(
            f_grid,
            text="เลขที่สิ้นสุด (ตัวเลือก):",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color="#334155",
            anchor="e",
            width=155,
        ).grid(row=2, column=0, sticky="e", padx=(0, 12), pady=4)

        cur_end = str(self.config.get("receipt_no_end", "700999999"))
        ent_m_end = make_entry(f_grid, width=180, height=34, font_size=13)
        ent_m_end.grid(row=2, column=1, sticky="w", pady=4)
        ent_m_end.insert(0, cur_end)

        # Live Preview Box
        prev_box = ctk.CTkFrame(
            c,
            fg_color="#F8FAFC",
            corner_radius=8,
            border_width=1,
            border_color="#E2E8F0",
        )
        prev_box.pack(fill="x", pady=(2, 0))

        lbl_prev1 = ctk.CTkLabel(
            prev_box,
            text="",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color="#1E40AF",
        )
        lbl_prev1.pack(anchor="w", padx=14, pady=(8, 2))

        lbl_prev2 = ctk.CTkLabel(
            prev_box,
            text="",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color="#64748B",
        )
        lbl_prev2.pack(anchor="w", padx=14, pady=(0, 8))

        def update_preview(*_):
            p = ent_m_pfx.get().strip()
            n = ent_m_num.get().strip() or "1"
            _, n_digits = split_receipt_prefix_and_number(n)
            eff_n = n_digits or n
            b1 = combine_receipt_no(p, eff_n)
            n_next = format_next_receipt_number(eff_n)
            b2 = combine_receipt_no(p, n_next)
            n_next2 = format_next_receipt_number(eff_n, step=2)
            b3 = combine_receipt_no(p, n_next2)
            lbl_prev1.configure(text=f"📌 ตัวอย่างบิลนี้: {b1}")
            lbl_prev2.configure(text=f"➡️ บิลถัดไปจะรันอัตโนมัติ: {b2}  ➔  {b3}  ➔  ...")

        ent_m_pfx.bind("<KeyRelease>", update_preview)
        ent_m_num.bind("<KeyRelease>", update_preview)
        update_preview()

        def save_and_apply():
            new_pfx = ent_m_pfx.get().strip()
            new_num = ent_m_num.get().strip()
            new_end = ent_m_end.get().strip()

            if not new_num:
                highlight_missing_field(ent_m_num, "กรุณากรอกเลขที่เริ่มต้น", show_toast_msg=True)
                return

            sub_p, sub_n = split_receipt_prefix_and_number(new_num)
            if sub_p and not new_pfx:
                new_pfx = sub_p
                new_num = sub_n

            self.config["receipt_prefix"] = new_pfx
            self.config["receipt_no_start"] = new_num
            if new_end:
                self.config["receipt_no_end"] = new_end
            save_config(self.config)

            if hasattr(self, "entry_rcp_no"):
                self.entry_rcp_no.delete(0, "end")
                self.entry_rcp_no.insert(0, combine_receipt_no(new_pfx, new_num))

            if hasattr(self, "adv_no_start"):
                try:
                    self.adv_no_start.delete(0, "end")
                    self.adv_no_start.insert(0, new_num)
                except Exception:
                    pass
            if hasattr(self, "adv_no_end") and new_end:
                try:
                    self.adv_no_end.delete(0, "end")
                    self.adv_no_end.insert(0, new_end)
                except Exception:
                    pass

            if hasattr(self, "recalculate_batch_totals"):
                try:
                    self.recalculate_batch_totals()
                except Exception:
                    pass

            close_modal()
            messagebox.showinfo(
                "บันทึกสำเร็จ",
                f"ตั้งค่าเลขที่และตัวอักษรออกบิลเรียบร้อยแล้ว!\n\n"
                f"• ตัวอักษร/คำนำหน้า: '{new_pfx or '(ไม่มี - ตัวเลขล้วน)'}'\n"
                f"• เลขที่เริ่มต้น: '{new_num}'\n"
                f"• บิลฉบับนี้: '{combine_receipt_no(new_pfx, new_num)}'\n"
                f"ระบบจะรันต่อจากเลขนี้ให้อัตโนมัติเมื่อออกบิล",
                parent=self.winfo_toplevel(),
            )

        modal.bind("<Return>", lambda e: save_and_apply())
        modal.bind("<KP_Enter>", lambda e: save_and_apply())

        btn_save = ctk.CTkButton(
            b_bar,
            text="💾 บันทึกและใช้เลขนี้ทันที",
            width=190,
            height=36,
            corner_radius=8,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=save_and_apply,
        )
        btn_save.pack(side="right", padx=(12, 20), pady=9)

        btn_cancel = ctk.CTkButton(
            b_bar,
            text="ยกเลิก",
            width=95,
            height=36,
            corner_radius=8,
            fg_color="#FFFFFF",
            hover_color="#F1F5F9",
            border_color="#CBD5E1",
            border_width=1,
            text_color="#475569",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=close_modal,
        )
        btn_cancel.pack(side="right", padx=(0, 0), pady=9)

    def _schedule_student_search(self, _event=None):
        if self._stu_search_job is not None:
            try:
                self.after_cancel(self._stu_search_job)
            except Exception:
                pass
        self._stu_search_job = self.after(250, self.refresh_student_list)

    def _lookup_student_by_code_entry(self, _event=None):
        code = self.entry_stu_code.get().strip()
        if not code:
            return
        stu = self.db.get_student_by_code(code)
        if stu:
            self.entry_stu_name.delete(0, "end")
            self.entry_stu_name.insert(0, stu.get("full_name", ""))
            fac = stu.get("faculty") or self.faculties[0]
            dept = stu.get("major") or self.departments[0]
            cls_g = stu.get("class_group") or "ปวช.1/1"
            self.fac_var.set(fac)
            self.dept_var.set(dept)
            self.cls_var.set(cls_g)
            if hasattr(self, "opt_faculty"):
                self.opt_faculty.set(fac)
            if hasattr(self, "opt_dept"):
                self.opt_dept.set(dept)
            if hasattr(self, "opt_cls"):
                self.opt_cls.set(cls_g)

    def _search_students_dropdown(self, query: str) -> List[str]:
        q = (query or "").strip()
        cls_filter = self.filter_class_var.get() if hasattr(self, "filter_class_var") else "ทั้งหมด"
        lvl = cls_filter if cls_filter != "ทั้งหมด" else "ALL"
        yr_filter = self.filter_year_var.get() if hasattr(self, "filter_year_var") else "ทั้งหมด"
        yr = yr_filter if yr_filter != "ทั้งหมด" else "ALL"

        stus = self.db.get_students(
            search=q,
            level_filter=lvl,
            year_level_filter=yr,
            limit=self.STUDENT_DROPDOWN_LIMIT,
            offset=0,
        )
        labels = []
        for s in stus:
            lbl = f"{s['student_code']}  —  {s['full_name']} ({s['class_group']})"
            self.student_map[lbl] = s
            labels.append(lbl)
        return labels or ["(ไม่พบรายชื่อนักศึกษาที่ค้นหา)"]

    def refresh_student_list(self):
        self._stu_search_job = None
        cls_filter = self.filter_class_var.get() if hasattr(self, "filter_class_var") else "ทั้งหมด"
        lvl = cls_filter if cls_filter != "ทั้งหมด" else "ALL"
        yr_filter = self.filter_year_var.get() if hasattr(self, "filter_year_var") else "ทั้งหมด"
        yr = yr_filter if yr_filter != "ทั้งหมด" else "ALL"
        search_txt = self.entry_stu_search.get().strip() if hasattr(self, "entry_stu_search") else ""

        total_matching = self.db.count_students(search=search_txt, level_filter=lvl, year_level_filter=yr)
        self.students = self.db.get_students(
            search=search_txt,
            level_filter=lvl,
            year_level_filter=yr,
            limit=self.STUDENT_DROPDOWN_LIMIT,
            offset=0,
        )
        self.student_map.clear()
        labels = []
        for s in self.students:
            lbl = f"{s['student_code']}  —  {s['full_name']} ({s['class_group']})"
            self.student_map[lbl] = s
            labels.append(lbl)

        if not labels:
            labels = ["ไม่พบรายชื่อนักศึกษาที่ค้นหา"]
            if hasattr(self, "opt_student"):
                self.opt_student.configure(values=labels)
                self.stu_var.set(labels[0])
            return

        if total_matching > len(self.students):
            labels.append(
                f"--- แสดง {len(self.students):,} จาก {total_matching:,} คน (พิมพ์ค้นหาหรือเลือกห้องเรียนเพื่อกรองเพิ่ม) ---"
            )

        if hasattr(self, "opt_student"):
            self.opt_student.configure(values=labels)
            self.stu_var.set(labels[0])

    def on_select_student(self, label: str):
        if not label or label.startswith("---") or "(ไม่พบ" in label:
            return
        stu = self.student_map.get(label)
        if not stu:
            code_cand = label.split("—")[0].split("-")[0].strip()
            stu = self.db.get_student_by_code(code_cand)
        if not stu:
            return
        self.entry_stu_name.delete(0, "end")
        self.entry_stu_name.insert(0, stu.get("full_name", ""))
        self.entry_stu_code.delete(0, "end")
        self.entry_stu_code.insert(0, stu.get("student_code", ""))

        fac = stu.get("faculty") or self.faculties[0]
        dept = stu.get("major") or self.departments[0]
        cls_g = stu.get("class_group") or "ปวช.1/1"
        self.fac_var.set(fac)
        self.dept_var.set(dept)
        self.cls_var.set(cls_g)
        if hasattr(self, "opt_faculty"):
            self.opt_faculty.set(fac)
        if hasattr(self, "opt_dept"):
            self.opt_dept.set(dept)
        if hasattr(self, "opt_cls"):
            self.opt_cls.set(cls_g)

    def on_change_curriculum(self, curriculum_name: str):
        if "กำหนดเอง" in curriculum_name or "อื่นๆ" in curriculum_name:
            labels = ["(กำหนดชื่อรายการและจำนวนเงินเอง)"]
            self.catalog_map.clear()
            if hasattr(self, "opt_fee_item"):
                self.opt_fee_item.configure(values=labels)
                self.fee_item_var.set(labels[0])
            return

        self.current_catalog = self.db.get_fee_catalog(curriculum_filter=curriculum_name)
        self.catalog_map.clear()
        labels = []
        for f in self.current_catalog:
            lbl = f"{f['fee_name']} ({float(f['default_price']):,.2f} บ.)"
            self.catalog_map[lbl] = f
            self.catalog_map[f["fee_name"]] = f
            labels.append(lbl)

        if not labels:
            labels = ["ไม่มีรายการในหลักสูตรนี้"]
            if hasattr(self, "opt_fee_item"):
                self.opt_fee_item.configure(values=labels)
                self.fee_item_var.set(labels[0])
            return

        if hasattr(self, "opt_fee_item"):
            self.opt_fee_item.configure(values=labels)
            self.fee_item_var.set(labels[0])
            self.on_select_fee_item(labels[0])

    def on_select_fee_item(self, label: str):
        fee = self.catalog_map.get(label)
        if not fee:
            return
        item_name = fee.get("fee_name", "")
        if hasattr(self, "entry_item_name"):
            self.entry_item_name.delete(0, "end")
            self.entry_item_name.insert(0, item_name)
        cred = float(fee.get("default_credits", 0))
        self.entry_qty.delete(0, "end")
        self.entry_qty.insert(0, f"{cred:.0f}" if cred > 0 else "0")
        self.entry_unit_price.delete(0, "end")
        self.entry_unit_price.insert(0, f"{float(fee.get('default_price', 0)):.2f}")
        self._update_line_total_preview()

    def _update_line_total_preview(self, _event=None):
        try:
            qty_s = self.entry_qty.get().strip().replace(",", "") if hasattr(self, "entry_qty") else "0"
            price_s = self.entry_unit_price.get().strip().replace(",", "") if hasattr(self, "entry_unit_price") else "0"
            q = float(qty_s or "0")
            p = float(price_s or "0")
            eff = q if q > 0 else 1.0
            tot = round(eff * p, 2)
            if hasattr(self, "lbl_item_total"):
                self.lbl_item_total.configure(text=f"฿{tot:,.2f}")
        except Exception:
            if hasattr(self, "lbl_item_total"):
                self.lbl_item_total.configure(text="฿0.00")

    def add_selected_fee_to_cart(self):
        item_name = ""
        if hasattr(self, "entry_item_name"):
            item_name = self.entry_item_name.get().strip()
        if not item_name and hasattr(self, "fee_item_var"):
            val = self.fee_item_var.get().strip()
            if val not in ("-", "ไม่มีรายการในหลักสูตรนี้", "(กำหนดชื่อรายการและจำนวนเงินเอง)"):
                fee = self.catalog_map.get(val)
                item_name = fee["fee_name"] if fee else val.split(" (")[0]

        if not item_name:
            target = getattr(self, "entry_item_name", getattr(self, "opt_fee_item", None))
            if target:
                highlight_missing_field(target, "กรุณาระบุชื่อรายการค่าธรรมเนียม", show_toast_msg=True)
            return

        try:
            qty = float(self.entry_qty.get().strip().replace(",", "") or "0")
            unit_price = float(self.entry_unit_price.get().strip().replace(",", "") or "0")
        except ValueError:
            highlight_missing_field(self.entry_unit_price, "กรุณาระบุจำนวนหน่วยกิตและราคาเป็นตัวเลข", show_toast_msg=True)
            return

        if unit_price < 0:
            highlight_missing_field(self.entry_unit_price, "ราคาไม่สามารถติดลบได้", show_toast_msg=True)
            return

        fee = self.catalog_map.get(item_name) or self.catalog_map.get(self.fee_item_var.get())
        fee_code = fee["fee_code"] if fee else "CUSTOM"
        curr_name = fee.get("curriculum_name", self.curriculum_var.get()) if fee else self.curriculum_var.get()

        eff_qty = qty if qty > 0 else 1.0
        line_total = round(eff_qty * unit_price, 2)

        self.cart_items.append(
            {
                "fee_code": fee_code,
                "curriculum_name": curr_name,
                "item_name": item_name,
                "quantity": qty,
                "unit_price": unit_price,
                "discount": 0.0,
                "line_total": line_total,
            }
        )
        self.refresh_cart_ui()
        if hasattr(self, "entry_item_name"):
            self.entry_item_name.delete(0, "end")
        self.entry_qty.delete(0, "end")
        self.entry_qty.insert(0, "0")
        self.entry_unit_price.delete(0, "end")
        self.entry_unit_price.insert(0, "0.00")
        if hasattr(self, "lbl_item_total"):
            self.lbl_item_total.configure(text="฿0.00")
        if hasattr(self, "entry_item_name"):
            self.entry_item_name.focus()

    def remove_selected_item(self):
        vals = self.table.get_selected_values()
        if not vals:
            if self.cart_items:
                self.cart_items.pop()
                self.refresh_cart_ui()
            return
        try:
            idx = int(vals[0]) - 1
            if 0 <= idx < len(self.cart_items):
                self.cart_items.pop(idx)
                self.refresh_cart_ui()
        except Exception:
            pass

    def reset_pos_form(self):
        self.cart_items.clear()
        self.refresh_cart_ui()
        self.refresh_next_receipt_no()
        self.entry_stu_code.delete(0, "end")
        self.entry_stu_name.delete(0, "end")
        if hasattr(self, "entry_item_name"):
            self.entry_item_name.delete(0, "end")
        if hasattr(self, "entry_qty"):
            self.entry_qty.delete(0, "end")
            self.entry_qty.insert(0, "0")
        if hasattr(self, "entry_unit_price"):
            self.entry_unit_price.delete(0, "end")
            self.entry_unit_price.insert(0, "0.00")

    def refresh_cart_ui(self):
        rows = []
        net_total = 0.0
        for idx, it in enumerate(self.cart_items, start=1):
            it["line_no"] = idx
            eff_q = float(it["quantity"]) if float(it["quantity"]) > 0 else 1.0
            line_tot = round(eff_q * float(it["unit_price"]), 2)
            it["line_total"] = line_tot
            net_total += line_tot
            q_display = f"{float(it['quantity']):,.0f}" if float(it["quantity"]) > 0 else "0"
            rows.append(
                (
                    [
                        str(idx),
                        it["item_name"],
                        q_display,
                        f"฿{float(it['unit_price']):,.2f}",
                        f"฿{line_tot:,.2f}",
                    ],
                    "PAID",
                )
            )
        self.table.set_rows(rows)
        self.lbl_net_total.configure(text=f"ยอดรวมสุทธิ: ฿{net_total:,.2f} บาท")
        self.lbl_baht_text.configure(text=f"ตัวอักษร: ({thai_baht_text(net_total)})")
        if hasattr(self, "lbl_cart_count"):
            self.lbl_cart_count.configure(text=f"📋 {len(self.cart_items)} รายการ")
        if hasattr(self, "lbl_batch_cart_info"):
            self.lbl_batch_cart_info.configure(
                text=f"รายการในตาราง: {len(self.cart_items)} รายการ  |  ยอดรวมคนละ: ฿{net_total:,.2f} บาท ({thai_baht_text(net_total)})"
            )

    def save_receipt_only(self):
        """Saves receipt record into the database without prompting for file export."""
        stu_code = self.entry_stu_code.get().strip()
        stu_name = self.entry_stu_name.get().strip()
        if not stu_name:
            highlight_missing_field(self.entry_stu_name, "กรุณากรอกชื่อผู้ชำระเงินในช่อง 'ได้รับเงินจาก'", show_toast_msg=True)
            return
        if not self.cart_items:
            messagebox.showwarning("ยังไม่มีรายการ", "กรุณากดปุ่ม '➕ เพิ่มรายการลงใบเสร็จ' อย่างน้อย 1 รายการ", parent=self.winfo_toplevel())
            return

        while True:
            rcp_no = self.get_current_receipt_no()
            rcp_prefix, _ = split_receipt_prefix_and_number(rcp_no)
            try:
                saved = self.db.issue_flexible_receipt(
                    student_code=stu_code or "EXTERNAL",
                    line_items=self.cart_items,
                    payment_method=self.pay_method_var.get(),
                    actor=self.session.username,
                    cashier_name=self.session.full_name,
                    role=self.session.role,
                    book_no=self.config.get("receipt_book_no", "7006"),
                    term_code=self.term_var.get(),
                    receipt_prefix=rcp_prefix,
                    receipt_no_override=rcp_no,
                    receipt_date_override=self.entry_date.get().strip(),
                    cashier_position=self.config.get("cashier_position", "ครู"),
                    custom_payer_name=stu_name,
                    custom_faculty=self.fac_var.get(),
                    custom_major=self.dept_var.get(),
                    custom_class_group=self.cls_var.get(),
                    start_no=self.config.get("receipt_no_start", "700605433"),
                    end_no=self.config.get("receipt_no_end", "700999999"),
                )
                break
            except ReceiptNumberError as rne:
                if not self._apply_receipt_number_resolution(rne):
                    return
            except Exception as e:
                messagebox.showerror("เกิดข้อผิดพลาดในการบันทึก", f"ไม่สามารถบันทึกข้อมูลได้:\n{e}", parent=self.winfo_toplevel())
                return

        try:
            # Auto-advance sequence and remember user-chosen prefix & number
            issued_no = saved.get("receipt_no", rcp_no)
            used_pfx, used_num = split_receipt_prefix_and_number(issued_no)
            next_num = format_next_receipt_number(used_num)
            self.config["receipt_prefix"] = used_pfx
            self.config["receipt_no_start"] = next_num
            save_config(self.config)

            want_print = messagebox.askyesno(
                "บันทึกข้อมูลสำเร็จ",
                f"บันทึกข้อมูลใบเสร็จรับเงินเรียบร้อยแล้ว!\n"
                f"• เลขที่ใบเสร็จ: {saved['receipt_no']}\n"
                f"• ผู้ชำระ: {saved['full_name']}\n"
                f"• ยอดสุทธิ: ฿{float(saved['net_total_amount']):,.2f} บาท\n\n"
                f"🖨️ ต้องการสร้างไฟล์และเปิดพิมพ์ใบเสร็จ (.pdf) ตอนนี้เลยหรือไม่?",
                parent=self.winfo_toplevel(),
            )
            if want_print:
                self._export_and_open_saved_receipt(saved)
            self.reset_pos_form()
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการบันทึก", f"ไม่สามารถบันทึกข้อมูลได้:\n{e}", parent=self.winfo_toplevel())

    def _export_and_open_saved_receipt(self, saved: Dict[str, Any]):
        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)
        rcp_no = saved.get("receipt_no", "Receipt")
        stu_name = saved.get("full_name", "")
        safe_name = re.sub(r"[^\w\-_]", "_", stu_name) if stu_name else "Student"
        default_filename = f"Receipt_{rcp_no}_{safe_name}.pdf"

        save_path = filedialog.asksaveasfilename(
            parent=self.winfo_toplevel(),
            title="เลือกตำแหน่งบันทึกไฟล์ใบเสร็จรับเงิน (.pdf)",
            initialdir=out_dir,
            initialfile=default_filename,
            defaultextension=".pdf",
            filetypes=[
                ("PDF Document (*.pdf)", "*.pdf"),
                ("Word Document (*.docx)", "*.docx"),
            ],
        )
        if not save_path:
            messagebox.showinfo(
                "ยกเลิกการพิมพ์",
                "ยกเลิกการส่งออกไฟล์ใบเสร็จ (ข้อมูลยังคงถูกบันทึกไว้ในระบบเรียบร้อยแล้ว สามารถสั่งพิมพ์ย้อนหลังได้ทุกเมื่อ)",
                parent=self.winfo_toplevel(),
            )
            return

        self.config["default_export_folder"] = os.path.dirname(save_path)
        save_config(self.config)
        self._refresh_export_folder_buttons()

        try:
            export_document_auto(
                save_path,
                saved,
                college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
                metadata=self.config,
                is_batch=False,
            )
            open_ok = open_document_file(save_path, parent=self.winfo_toplevel())
            messagebox.showinfo(
                "ส่งออกเอกสารสำเร็จ",
                f"สร้างไฟล์ใบเสร็จเลขที่ {rcp_no} เรียบร้อยแล้วที่:\n{save_path}\n\n"
                + ("(ระบบได้เปิดไฟล์เอกสารขึ้นมาให้พร้อมพิมพ์เรียบร้อยแล้ว)" if open_ok else "(กรุณาเปิดไฟล์ดังกล่าวเพื่อสั่งพิมพ์)"),
                parent=self.winfo_toplevel(),
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการพิมพ์ใบเสร็จ", f"ไม่สามารถสร้างหรือเปิดไฟล์เอกสารได้:\n{e}", parent=self.winfo_toplevel())

    def _refresh_export_folder_buttons(self):
        """Updates the export-folder button labels in both normal and batch POS bars."""
        label = format_export_folder_button_text(self.config)
        for attr in ("btn_export_folder_normal", "btn_export_folder_batch"):
            btn = getattr(self, attr, None)
            if btn is not None:
                btn.configure(text=label)

    def _pick_export_folder(self):
        """Lets the cashier pick the default folder for saving exported receipts."""
        choose_export_folder(
            self.winfo_toplevel(),
            self.config,
            on_updated=lambda _: self._refresh_export_folder_buttons(),
        )

    def issue_receipt_and_export(self):
        """Saves receipt record into the database AND exports official PDF matching EX."""
        stu_code = self.entry_stu_code.get().strip()
        stu_name = self.entry_stu_name.get().strip()
        if not stu_name:
            highlight_missing_field(self.entry_stu_name, "กรุณากรอกชื่อผู้ชำระเงินในช่อง 'ได้รับเงินจาก'", show_toast_msg=True)
            return
        if not self.cart_items:
            messagebox.showwarning("ยังไม่มีรายการในใบเสร็จ", "กรุณากดปุ่ม '➕ เพิ่มรายการลงใบเสร็จ' อย่างน้อย 1 รายการ", parent=self.winfo_toplevel())
            return

        rcp_no = self.get_current_receipt_no()
        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)
        safe_name = re.sub(r"[^\w\-_]", "_", stu_name) if stu_name else "Student"
        default_filename = f"Receipt_{rcp_no}_{safe_name}.pdf"

        save_path = filedialog.asksaveasfilename(
            parent=self.winfo_toplevel(),
            title="เลือกตำแหน่งบันทึกไฟล์ใบเสร็จรับเงิน (.pdf)",
            initialdir=out_dir,
            initialfile=default_filename,
            defaultextension=".pdf",
            filetypes=[
                ("PDF Document (*.pdf)", "*.pdf"),
                ("Word Document (*.docx)", "*.docx"),
            ],
        )
        if not save_path:
            messagebox.showinfo(
                "ยกเลิกการออกใบเสร็จ",
                "คุณได้ยกเลิกการเลือกตำแหน่งบันทึกไฟล์ ข้อมูลใบเสร็จจึงยังไม่ถูกบันทึกและยังไม่มีการพิมพ์เอกสาร",
                parent=self.winfo_toplevel(),
            )
            return

        self.config["default_export_folder"] = os.path.dirname(save_path)
        save_config(self.config)
        self._refresh_export_folder_buttons()

        while True:
            rcp_no = self.get_current_receipt_no()
            rcp_prefix, _ = split_receipt_prefix_and_number(rcp_no)
            try:
                saved = self.db.issue_flexible_receipt(
                    student_code=stu_code or "EXTERNAL",
                    line_items=self.cart_items,
                    payment_method=self.pay_method_var.get(),
                    actor=self.session.username,
                    cashier_name=self.session.full_name,
                    role=self.session.role,
                    book_no=self.config.get("receipt_book_no", "7006"),
                    term_code=self.term_var.get(),
                    receipt_prefix=rcp_prefix,
                    receipt_no_override=rcp_no,
                    receipt_date_override=self.entry_date.get().strip(),
                    cashier_position=self.config.get("cashier_position", "ครู"),
                    custom_payer_name=stu_name,
                    custom_faculty=self.fac_var.get(),
                    custom_major=self.dept_var.get(),
                    custom_class_group=self.cls_var.get(),
                    start_no=self.config.get("receipt_no_start", "700605433"),
                    end_no=self.config.get("receipt_no_end", "700999999"),
                )
                break
            except ReceiptNumberError as rne:
                if not self._apply_receipt_number_resolution(rne):
                    return
            except Exception as e:
                messagebox.showerror("เกิดข้อผิดพลาดในการออกใบเสร็จ", f"ไม่สามารถออกใบเสร็จได้:\n{e}", parent=self.winfo_toplevel())
                return

        try:
            # Auto-advance sequence and remember user-chosen prefix & number
            issued_no = saved.get("receipt_no", rcp_no)
            used_pfx, used_num = split_receipt_prefix_and_number(issued_no)
            next_num = format_next_receipt_number(used_num)
            self.config["receipt_prefix"] = used_pfx
            self.config["receipt_no_start"] = next_num
            save_config(self.config)

            export_document_auto(
                save_path,
                saved,
                college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
                metadata=self.config,
                is_batch=False,
            )

            open_ok = open_document_file(save_path, parent=self.winfo_toplevel())

            messagebox.showinfo(
                "ออกใบเสร็จสำเร็จ",
                f"เลขที่ใบเสร็จ: {saved['receipt_no']}\n"
                f"ผู้ชำระ: {saved['full_name']}\n"
                f"ยอดสุทธิ: ฿{float(saved['net_total_amount']):,.2f} บาท\n\n"
                f"ไฟล์ใบเสร็จบันทึกอยู่ที่:\n{save_path}\n\n"
                + ("(ระบบได้เปิดไฟล์เอกสารขึ้นมาให้พร้อมพิมพ์เรียบร้อยแล้ว)" if open_ok else "(กรุณาเปิดไฟล์ดังกล่าวเพื่อสั่งพิมพ์)"),
                parent=self.winfo_toplevel(),
            )
            self.reset_pos_form()
        except PermissionError as pe:
            messagebox.showerror(
                "บันทึกไฟล์ไม่สำเร็จ (ไฟล์ถูกเปิดใช้งานอยู่)",
                f"ไม่สามารถเขียนทับไฟล์ได้:\n{save_path}\n\n"
                f"⚠️ ไฟล์นี้กำลังถูกเปิดใช้งานอยู่ในโปรแกรมอื่น (เช่น Adobe Acrobat, Edge, หรือ Word)\n"
                f"กรุณาปิดโปรแกรมดังกล่าวก่อน แล้วกดออกใบเสร็จใหม่อีกครั้ง",
                parent=self.winfo_toplevel(),
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการออกใบเสร็จ", f"ไม่สามารถออกใบเสร็จได้:\n{e}", parent=self.winfo_toplevel())

    # ------------------------------------------------------------------
    # ADVANCED MODE HELPERS & ACTIONS
    # ------------------------------------------------------------------
    def sync_advanced_panel_fields(self):
        """Synchronizes configuration and cart state into Advanced Mode fields."""
        if hasattr(self, "opt_adv_tpl_style"):
            self.adv_tpl_style_var.set(_style_label_from_code(self.config.get("bill_template_style", "CLASSIC")))
        if hasattr(self, "adv_college_name"):
            self.adv_college_name.delete(0, "end")
            self.adv_college_name.insert(0, self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"))
        if hasattr(self, "adv_gov_title"):
            self.adv_gov_title.delete(0, "end")
            self.adv_gov_title.insert(0, self.config.get("receipt_gov_title", "ในราชการกรมอาชีวศึกษา กระทรวงศึกษาธิการ"))
        if hasattr(self, "adv_book_no"):
            self.adv_book_no.delete(0, "end")
            self.adv_book_no.insert(0, self.config.get("receipt_book_no", "7006"))
        if hasattr(self, "adv_no_start"):
            self.adv_no_start.delete(0, "end")
            self.adv_no_start.insert(0, self.config.get("receipt_no_start", "700605433"))
        if hasattr(self, "adv_no_end"):
            self.adv_no_end.delete(0, "end")
            self.adv_no_end.insert(0, self.config.get("receipt_no_end", "700999999"))
        if hasattr(self, "adv_cashier_name"):
            self.adv_cashier_name.delete(0, "end")
            self.adv_cashier_name.insert(0, self.session.full_name)
        if hasattr(self, "adv_cashier_pos"):
            self.adv_cashier_pos.delete(0, "end")
            self.adv_cashier_pos.insert(0, self.config.get("cashier_position", "ครู"))
        if hasattr(self, "adv_promptpay_no"):
            self.adv_promptpay_no.delete(0, "end")
            self.adv_promptpay_no.insert(0, self.config.get("promptpay_no", "0994000160000"))
        if hasattr(self, "adv_bank_account"):
            self.adv_bank_account.delete(0, "end")
            self.adv_bank_account.insert(0, self.config.get("bank_account", "ธนาคารกรุงไทย 710-0-12345-6"))
        self._on_change_batch_class()

    def _on_change_batch_class(self, _val=None):
        if not hasattr(self, "batch_cls_var") or not hasattr(self, "lbl_batch_student_count"):
            return
        room = self.batch_cls_var.get()
        students_in_room = self.db.get_students(level_filter=room)
        self.lbl_batch_student_count.configure(text=f"จำนวนนักศึกษาในห้อง {room}: {len(students_in_room)} คน")

    def run_batch_for_selected_room(self):
        if not self.cart_items:
            messagebox.showwarning(
                "ไม่มีรายการในใบเสร็จ",
                "กรุณาเพิ่มรายการค่าธรรมเนียมในตาราง หรือกดปุ่มชุดค่าธรรมเนียมด่วน (หัวข้อ 5) ก่อนออกใบเสร็จยกห้อง",
            )
            return
        room = self.batch_cls_var.get()
        students_in_room = self.db.get_students(level_filter=room)
        if not students_in_room:
            messagebox.showwarning("ไม่พบนักศึกษา", f"ไม่พบรายชื่อนักศึกษาในห้อง {room}")
            return
        net_total = sum(float(i["line_total"]) for i in self.cart_items)
        if not messagebox.askyesno(
            "ยืนยันออกใบเสร็จยกห้อง",
            f"ออกใบเสร็จให้ห้อง '{room}' จำนวน {len(students_in_room)} คน\nยอดคนละ ฿{net_total:,.2f} บาท ({thai_baht_text(net_total)})\nต้องการดำเนินการใช่หรือไม่?",
        ):
            return
        self._execute_batch_receipts(self, students_in_room, label_tag=room)

    def import_and_issue_batch(self):
        if not self.cart_items:
            messagebox.showwarning(
                "ไม่มีรายการในใบเสร็จ",
                "กรุณาเพิ่มรายการค่าธรรมเนียมในตาราง หรือกดปุ่มชุดค่าธรรมเนียมด่วน (หัวข้อ 5) ก่อนนำเข้าเพื่อออกใบเสร็จ",
            )
            return
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์รายชื่อนักศึกษา (Excel .xlsx หรือ .csv)",
            filetypes=[("Excel / CSV Files", "*.xlsx *.csv")],
        )
        if not file_path:
            return
        res = self.db.bulk_import_students(
            file_path=file_path,
            actor=self.session.username,
            role=self.session.role,
            default_faculty=self.fac_var.get(),
            default_major=self.dept_var.get(),
            default_class_group=self.cls_var.get(),
        )
        codes = res.get("student_codes", [])
        if not codes:
            messagebox.showwarning("ไม่พบข้อมูล", "ไม่พบรายชื่อนักศึกษาในไฟล์ที่เลือก")
            return
        self.refresh_student_list()
        imported_students = self.db.get_students_by_codes(codes)
        net_total = sum(float(i["line_total"]) for i in self.cart_items)
        if not messagebox.askyesno(
            "นำเข้ารายชื่อสำเร็จ",
            f"นำเข้ารายชื่อสำเร็จ {len(imported_students)} คน!\nยอดคนละ ฿{net_total:,.2f} บาท\nต้องการออกใบเสร็จให้ทุกคนทันทีหรือไม่?",
        ):
            return
        self._execute_batch_receipts(self, imported_students, label_tag="ImportedClass")

    def save_template_settings(self):
        style_label = self.adv_tpl_style_var.get()
        self.config["bill_template_style"] = normalize_template_style(style_label)
        self.config["college_name"] = self.adv_college_name.get().strip()
        self.config["receipt_gov_title"] = self.adv_gov_title.get().strip()
        self.config["receipt_book_no"] = self.adv_book_no.get().strip()
        self.config["receipt_no_start"] = self.adv_no_start.get().strip()
        self.config["receipt_no_end"] = self.adv_no_end.get().strip()
        self.config["cashier_position"] = self.adv_cashier_pos.get().strip()

        self.config["payment_method"] = self.pay_method_var.get()
        self.config["promptpay_no"] = self.adv_promptpay_no.get().strip()
        self.config["bank_account"] = self.adv_bank_account.get().strip()
        self.config["show_qr_on_receipt"] = bool(self.adv_show_qr_var.get())
        self.config["webhook_enabled"] = bool(self.adv_webhook_enabled_var.get())
        self.config["webhook_url"] = self.adv_webhook_url.get().strip()
        try:
            self.config["webhook_alert_threshold"] = float(self.adv_webhook_threshold.get().strip().replace(",", "") or "50000")
        except Exception:
            pass

        save_config(self.config)
        self.refresh_next_receipt_no()
        messagebox.showinfo("บันทึกสำเร็จ", "บันทึกการตั้งค่าแม่แบบใบเสร็จและระบบแอดวานซ์เรียบร้อยแล้ว!")

    def load_fee_preset(self, preset_key: str):
        self.cart_items.clear()
        if preset_key in ("VC1", "VC2"):
            # Official SPVC Receipt Template matching EX 95b999ca-2a23-4412-bbee-a3b29b8aa4e6.jpeg
            self.cart_items = [
                {
                    "fee_code": "ENV",
                    "curriculum_name": "หลักสูตรในระบบ-ปวช",
                    "item_name": "ค่ารักษาภาพแวดล้อม",
                    "quantity": 0.0,
                    "unit_price": 330.0,
                    "discount": 0.0,
                    "line_total": 330.0,
                },
                {
                    "fee_code": "NET",
                    "curriculum_name": "หลักสูตรในระบบ-ปวช",
                    "item_name": "ค่าใช้บริการอินเตอร์เน็ต",
                    "quantity": 0.0,
                    "unit_price": 300.0,
                    "discount": 0.0,
                    "line_total": 300.0,
                },
                {
                    "fee_code": "TCH",
                    "curriculum_name": "หลักสูตรในระบบ-ปวช",
                    "item_name": "ค่าจ้างครูหรือผู้ที่มีความเชี่ยวชาญในสาขาเฉพาะ",
                    "quantity": 0.0,
                    "unit_price": 500.0,
                    "discount": 0.0,
                    "line_total": 500.0,
                },
            ]
        elif preset_key == "HVC1":
            self.cart_items = [
                {
                    "fee_code": "MAINT",
                    "curriculum_name": "หลักสูตรในระบบ-ปวส",
                    "item_name": "ค่าบำรุงการศึกษา",
                    "quantity": 0.0,
                    "unit_price": 650.0,
                    "discount": 0.0,
                    "line_total": 650.0,
                },
                {
                    "fee_code": "NET_HVC",
                    "curriculum_name": "หลักสูตรในระบบ-ปวส",
                    "item_name": "ค่าบริการอินเตอร์เน็ตและเทคโนโลยี",
                    "quantity": 0.0,
                    "unit_price": 400.0,
                    "discount": 0.0,
                    "line_total": 400.0,
                },
                {
                    "fee_code": "SPEC_TCH",
                    "curriculum_name": "หลักสูตรในระบบ-ปวส",
                    "item_name": "ค่าจ้างครูผู้เชี่ยวชาญเฉพาะทาง",
                    "quantity": 0.0,
                    "unit_price": 500.0,
                    "discount": 0.0,
                    "line_total": 500.0,
                },
            ]
        elif preset_key == "WORKSHOP":
            self.cart_items = [
                {
                    "fee_code": "MAT",
                    "curriculum_name": "ค่าวัสดุฝึกสำเร็จรูป",
                    "item_name": "ค่าวัสดุฝึกสำเร็จรูปและอุปกรณ์ภาคปฏิบัติ",
                    "quantity": 1.0,
                    "unit_price": 600.0,
                    "discount": 0.0,
                    "line_total": 600.0,
                }
            ]

        self.refresh_cart_ui()
        net = sum(i["line_total"] for i in self.cart_items)
        messagebox.showinfo(
            "โหลดชุดค่าธรรมเนียมสำเร็จ",
            f"โหลดรายการค่าธรรมเนียม {len(self.cart_items)} รายการเรียบร้อยแล้ว!\n"
            f"ยอดรวม: ฿{net:,.2f} บาท ({thai_baht_text(net)})\n\n"
            f"รายการพร้อมใช้งานในตารางแล้ว สามารถสลับกลับไปหน้าปกติ (กรอกมือ) เพื่อตรวจสอบหรือสั่งพิมพ์ได้ทันที",
        )

    def test_webhook_ping(self):
        url = self.adv_webhook_url.get().strip()
        if not url:
            messagebox.showwarning("ระบุ Webhook URL", "กรุณาระบุ Webhook URL ก่อนทดสอบ")
            return
        try:
            self.webhook.webhook_url = url
            self.webhook.enabled = True
            success = self.webhook.ping_test()
            if success:
                messagebox.showinfo("ส่งสัญญาณสำเร็จ", "ส่งสัญญาณทดสอบไปยัง Webhook เรียบร้อยแล้ว!")
            else:
                messagebox.showwarning("ส่งสัญญาณไม่สำเร็จ", "ไม่สามารถส่งสัญญาณไปยัง Webhook URL ได้ กรุณาตรวจสอบ URL")
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถส่ง Webhook ได้:\n{e}")

    def open_batch_class_modal(self):
        """Switches to Dedicated Batch Classroom View if available, or Batch Mode."""
        if self.on_navigate:
            self.on_navigate("batch_pos")
            return
        self.set_mode("batch")

    def _execute_batch_receipts(self, parent_widget, students_list: List[Dict[str, Any]], label_tag: str):
        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)
        safe_tag = label_tag.replace("/", "-").replace(" ", "_")
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"BatchReceipts_{safe_tag}_{len(students_list)}students_{stamp}.pdf"

        save_path = filedialog.asksaveasfilename(
            parent=parent_widget,
            title="เลือกตำแหน่งบันทึกไฟล์ใบเสร็จยกห้อง (.pdf)",
            initialdir=out_dir,
            initialfile=default_filename,
            defaultextension=".pdf",
            filetypes=[
                ("PDF Document (*.pdf)", "*.pdf"),
                ("Word Document (*.docx)", "*.docx"),
            ],
        )
        if not save_path:
            messagebox.showinfo(
                "ยกเลิกการออกใบเสร็จยกห้อง",
                "คุณได้ยกเลิกการเลือกตำแหน่งบันทึกไฟล์ เอกสารจึงยังไม่ถูกบันทึกหรือพิมพ์",
                parent=parent_widget,
            )
            return

        self.config["default_export_folder"] = os.path.dirname(save_path)
        save_config(self.config)
        self._refresh_export_folder_buttons()

        while True:
            try:
                issued_list = self.db.issue_batch_class_receipts(
                    students_list=students_list,
                    line_items=self.cart_items,
                    payment_method=self.pay_method_var.get(),
                    actor=self.session.username,
                    cashier_name=self.session.full_name,
                    role=self.session.role,
                    term_code=self.term_var.get(),
                    book_no=self.config.get("receipt_book_no", "7006"),
                    cashier_position=self.config.get("cashier_position", "ครู"),
                    receipt_prefix=self.config.get("receipt_prefix", ""),
                    start_no=self.config.get("receipt_no_start", "700605433"),
                    end_no=self.config.get("receipt_no_end", "700999999"),
                    receipt_date_override=self.entry_date.get().strip(),
                )
                break
            except ReceiptNumberError as rne:
                if not self._apply_receipt_number_resolution(rne, parent=parent_widget):
                    return
            except Exception as e:
                messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถออกใบเสร็จยกห้องได้:\n{e}", parent=parent_widget)
                return

        try:
            export_document_auto(
                save_path,
                issued_list,
                college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
                metadata=self.config,
                is_batch=True,
            )

            first_no = issued_list[0]["receipt_no"]
            last_no = issued_list[-1]["receipt_no"]
            used_pfx, used_num = split_receipt_prefix_and_number(last_no)
            next_num = format_next_receipt_number(used_num)
            self.config["receipt_prefix"] = used_pfx
            self.config["receipt_no_start"] = next_num
            save_config(self.config)
            self.reset_pos_form()

            open_ok = open_document_file(save_path, parent=parent_widget)

            messagebox.showinfo(
                "ออกใบเสร็จยกห้องสำเร็จ!",
                f"ออกใบเสร็จครบทั้งห้องจำนวน: {len(issued_list)} ฉบับ\n"
                f"เลขที่ใบเสร็จ: {first_no} ถึง {last_no}\n\n"
                f"ไฟล์ใบเสร็จรวมทั้งห้องพร้อมพิมพ์บันทึกอยู่ที่:\n{save_path}\n\n"
                + ("(ระบบได้เปิดไฟล์เอกสารขึ้นมาให้พร้อมพิมพ์เรียบร้อยแล้ว)" if open_ok else ""),
                parent=parent_widget,
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถออกใบเสร็จยกห้องได้:\n{e}", parent=parent_widget)


# ==============================================================================
# VIEW 2: STUDENT MANAGEMENT VIEW (หน้า 2: รายชื่อนักศึกษา - โทนสีสุภาพ สะอาดตา)
# ==============================================================================
class StudentAndFeeCatalogView(ctk.CTkFrame):
    """
    Minimal Student View with Unified Corporate Palette:
    - Top bar: เลือกห้องเรียน + ช่องค้นหา + ปุ่มนำเข้า Excel
    - Center: ตารางรายชื่อนักศึกษา (คลิกแถวเพื่อแก้ไข)
    - Bottom bar: ช่องกรอกข้อมูล + ปุ่มหลัก '💾 บันทึกข้อมูลนักศึกษา' + ปุ่มลบ
    """

    def __init__(self, parent, db: DatabaseManager, session: UserSession, config: Dict[str, Any]):
        super().__init__(parent, fg_color="transparent")
        self.db = db
        self.session = session
        self.config = config

        self.faculties = self.db.get_options("FACULTY") or ["บริหารธุรกิจ / พาณิชยกรรม"]
        self.departments = self.db.get_options("DEPARTMENT") or ["การบัญชี"]
        self.class_levels = self.db.get_options("CLASS_LEVEL") or ["ปวช.1/1"]

        header = SectionHeader(
            self,
            "รายชื่อนักศึกษา",
            "คลิกเลือกแถวในตารางเพื่อแก้ไข หรือกดนำเข้ารายชื่อจากไฟล์ Excel",
        )
        header.pack(fill="x", pady=(0, 10))
        header.add_action_button(
            "นำเข้า Excel / CSV",
            self.import_students_dialog,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
        )

        # ------------------------------------------------------------------
        # BOTTOM EDITOR BAR (จัดฟอร์มอย่างเป็นระเบียบ 2 แถว อ่านง่าย ชัดเจน)
        # ------------------------------------------------------------------
        form_card = ctk.CTkFrame(
            self,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        form_card.pack(side="bottom", fill="x", pady=(12, 0))

        # แถวที่ 1: รหัส (120px) | ชื่อ-สกุล (ขยายกว้าง) | ชั้นปี (80px) | ห้อง (130px)
        f_row1 = ctk.CTkFrame(form_card, fg_color="transparent")
        f_row1.pack(fill="x", padx=20, pady=(14, 6))
        f_row1.grid_columnconfigure(1, weight=0)
        f_row1.grid_columnconfigure(3, weight=1)
        f_row1.grid_columnconfigure(5, weight=0)
        f_row1.grid_columnconfigure(7, weight=0)

        ctk.CTkLabel(
            f_row1,
            text="รหัสนักศึกษา:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 6), pady=2)
        self.ent_code = make_entry(f_row1, width=120, height=38, font_size=14, bold=True)
        self.ent_code.grid(row=0, column=1, sticky="w", padx=(0, 16), pady=2)

        ctk.CTkLabel(
            f_row1,
            text="ชื่อ-นามสกุล:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        ).grid(row=0, column=2, sticky="w", padx=(0, 6), pady=2)
        self.ent_name = make_entry(f_row1, height=38, font_size=14, bold=True)
        self.ent_name.grid(row=0, column=3, sticky="ew", padx=(0, 16), pady=2)

        ctk.CTkLabel(
            f_row1,
            text="ชั้นปี:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        ).grid(row=0, column=4, sticky="w", padx=(0, 6), pady=2)
        self.form_yr_var = ctk.StringVar(value="1")
        self.opt_form_yr = make_dropdown(
            f_row1,
            values=["1", "2", "3", "4"],
            variable=self.form_yr_var,
            width=75,
            height=38,
        )
        self.opt_form_yr.grid(row=0, column=5, sticky="w", padx=(0, 16), pady=2)

        ctk.CTkLabel(
            f_row1,
            text="ระดับชั้น / ห้อง:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        ).grid(row=0, column=6, sticky="w", padx=(0, 6), pady=2)
        self.form_cls_var = ctk.StringVar(value=self.class_levels[0])
        self.opt_form_cls = make_dropdown(
            f_row1,
            values=self.class_levels,
            variable=self.form_cls_var,
            width=130,
            height=38,
            command=self._on_form_cls_change,
        )
        self.opt_form_cls.grid(row=0, column=7, sticky="e", pady=2)

        # แถวที่ 2: แผนกวิชา/สาขา (50%) | คณะวิชา (50%)
        f_row2 = ctk.CTkFrame(form_card, fg_color="transparent")
        f_row2.pack(fill="x", padx=20, pady=(4, 8))
        f_row2.grid_columnconfigure(1, weight=1)
        f_row2.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(
            f_row2,
            text="แผนกวิชา / สาขา:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 6), pady=2)
        self.form_dept_var = ctk.StringVar(value=self.departments[0])
        make_dropdown(
            f_row2,
            values=self.departments,
            variable=self.form_dept_var,
            height=38,
        ).grid(row=0, column=1, sticky="ew", padx=(0, 20), pady=2)

        ctk.CTkLabel(
            f_row2,
            text="คณะวิชา:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        ).grid(row=0, column=2, sticky="w", padx=(0, 6), pady=2)
        self.form_fac_var = ctk.StringVar(value=self.faculties[0])
        make_dropdown(
            f_row2,
            values=self.faculties,
            variable=self.form_fac_var,
            height=38,
        ).grid(row=0, column=3, sticky="ew", pady=2)

        # แถวที่ 3: ปุ่มดำเนินการ
        f_row3 = ctk.CTkFrame(form_card, fg_color="transparent")
        f_row3.pack(fill="x", padx=20, pady=(4, 14))

        ctk.CTkButton(
            f_row3,
            text="บันทึกข้อมูลนักศึกษา",
            width=200,
            height=40,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.save_or_update_student,
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            f_row3,
            text="ล้างช่องกรอก",
            width=115,
            height=40,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.clear_form,
        ).pack(side="left")

        ctk.CTkButton(
            f_row3,
            text="ลบรายชื่อนี้",
            width=120,
            height=40,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_DANGER_BG,
            hover_color="#FEE2E2",
            border_width=1,
            border_color=AppTheme.COLOR_DANGER_BORDER,
            text_color=AppTheme.COLOR_DANGER,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.delete_selected_student,
        ).pack(side="right")

        # ------------------------------------------------------------------
        # TOP FILTER BAR
        # ------------------------------------------------------------------
        filter_bar = ctk.CTkFrame(
            self,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        filter_bar.pack(side="top", fill="x", pady=(0, 10))

        fb_in = ctk.CTkFrame(filter_bar, fg_color="transparent")
        fb_in.pack(fill="x", padx=16, pady=10)

        self.page_size = 100
        self.current_page = 1
        self._search_job: Optional[str] = None

        # 1. กรองปีที่เข้า / รุ่น
        ctk.CTkLabel(
            fb_in,
            text="กรองปี (รุ่น):",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))
        self.flt_year_cohort_var = ctk.StringVar(value="ทั้งหมด")
        self.opt_flt_cohort = make_dropdown(
            fb_in,
            values=["ทั้งหมด"] + self.db.get_distinct_student_years(),
            variable=self.flt_year_cohort_var,
            width=140,
            height=38,
            command=lambda _v: self._reset_and_refresh(),
        )
        self.opt_flt_cohort.pack(side="left", padx=(0, 12))

        # 2. กรองชั้นปี
        ctk.CTkLabel(
            fb_in,
            text="กรองชั้นปี:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))
        self.flt_year_level_var = ctk.StringVar(value="ทั้งหมด")
        make_dropdown(
            fb_in,
            values=["ทั้งหมด", "ปี 1", "ปี 2", "ปี 3", "ปี 4"],
            variable=self.flt_year_level_var,
            width=100,
            height=38,
            command=lambda _v: self._reset_and_refresh(),
        ).pack(side="left", padx=(0, 12))

        # 3. กรองสาขา
        ctk.CTkLabel(
            fb_in,
            text="กรองสาขา:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))
        self.flt_dept_var = ctk.StringVar(value="ทั้งหมด")
        make_dropdown(
            fb_in,
            values=["ทั้งหมด"] + self.departments,
            variable=self.flt_dept_var,
            width=150,
            height=38,
            command=lambda _v: self._reset_and_refresh(),
        ).pack(side="left", padx=(0, 12))

        # 4. กรองห้องเรียน
        ctk.CTkLabel(
            fb_in,
            text="กรองห้อง:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))
        self.flt_cls_var = ctk.StringVar(value="ทั้งหมด")
        make_dropdown(
            fb_in,
            values=["ทั้งหมด"] + self.class_levels,
            variable=self.flt_cls_var,
            width=115,
            height=38,
            command=lambda _v: self._reset_and_refresh(),
        ).pack(side="left", padx=(0, 12))

        # 5. ค้นหา
        self.ent_search = make_entry(
            fb_in,
            height=38,
            bold=False,
            placeholder_text="🔍 ค้นหารหัสนักศึกษา หรือ ชื่อ-นามสกุล...",
        )
        self.ent_search.pack(side="left", fill="x", expand=True)
        self.ent_search.bind("<KeyRelease>", self._schedule_search)

        # 6. ปุ่มรีเซ็ตตัวกรอง
        ctk.CTkButton(
            fb_in,
            text="🔄 รีเซ็ต",
            width=75,
            height=38,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.reset_filters,
        ).pack(side="left", padx=(8, 0))

        cols = [
            ("student_code", "รหัสนักศึกษา", 130, "center"),
            ("full_name", "ชื่อ-นามสกุล", 230, "w"),
            ("year_level", "ชั้นปี", 75, "center"),
            ("class_group", "ระดับชั้น / ห้อง", 125, "center"),
            ("major", "แผนกวิชา / สาขา", 190, "w"),
            ("faculty", "คณะ", 160, "w"),
            ("paid_amount", "ยอดชำระสะสม (บาท)", 150, "e"),
        ]
        self.table = DataTable(
            self,
            columns=cols,
            height=12,
            on_select=self.on_select_row,
            show_inspector=False,
            page_size=self.page_size,
            on_page_change=self.on_page_change,
            show_pagination=True,
        )
        self.table.pack(side="top", fill="both", expand=True)
        self.students_table = self.table

        self.refresh_table()

    def on_show(self):
        """Reloads student catalog only if database data has changed."""
        if getattr(self.db, "data_version", None) == getattr(self, "_last_data_version", -1):
            return
        self.refresh_table()

    def _on_form_cls_change(self, val: str):
        m = re.search(r"(\d+)", str(val))
        if m:
            y = m.group(1)
            if y in ("1", "2", "3", "4"):
                self.form_yr_var.set(y)

    def reset_filters(self):
        self.flt_year_cohort_var.set("ทั้งหมด")
        self.flt_year_level_var.set("ทั้งหมด")
        self.flt_dept_var.set("ทั้งหมด")
        self.flt_cls_var.set("ทั้งหมด")
        self.ent_search.delete(0, "end")
        self._reset_and_refresh()

    def _schedule_search(self, _event=None):
        if self._search_job is not None:
            try:
                self.after_cancel(self._search_job)
            except Exception:
                pass
        self._search_job = self.after(250, self._reset_and_refresh)

    def _reset_and_refresh(self):
        self._search_job = None
        self.current_page = 1
        self.refresh_table()

    def on_page_change(self, new_page: int):
        self.current_page = max(1, int(new_page))
        self.refresh_table()

    def refresh_table(self):
        s_txt = self.ent_search.get().strip()
        f_cohort = self.flt_year_cohort_var.get()
        f_yr_lvl = self.flt_year_level_var.get()
        f_dept = self.flt_dept_var.get()
        f_cls = self.flt_cls_var.get()

        cohort_flt = f_cohort if f_cohort != "ทั้งหมด" else "ALL"
        yr_lvl_flt = f_yr_lvl if f_yr_lvl != "ทั้งหมด" else "ALL"
        maj_flt = f_dept if f_dept != "ทั้งหมด" else "ALL"
        lvl_flt = f_cls if f_cls != "ทั้งหมด" else "ALL"

        total_count = self.db.count_students(
            search=s_txt,
            major_filter=maj_flt,
            level_filter=lvl_flt,
            year_level_filter=yr_lvl_flt,
            admission_year_filter=cohort_flt,
        )
        max_pages = max(1, (total_count + self.page_size - 1) // self.page_size)
        if self.current_page > max_pages:
            self.current_page = max_pages

        offset = (self.current_page - 1) * self.page_size
        students = self.db.get_students(
            search=s_txt,
            major_filter=maj_flt,
            level_filter=lvl_flt,
            year_level_filter=yr_lvl_flt,
            admission_year_filter=cohort_flt,
            limit=self.page_size,
            offset=offset,
        )
        rows = []
        for s in students:
            yr_val = s.get("year_level") or 1
            rows.append(
                (
                    [
                        s["student_code"],
                        s["full_name"],
                        f"ปี {yr_val}",
                        s.get("class_group") or "-",
                        s.get("major") or "-",
                        s.get("faculty") or "บริหารธุรกิจ / พาณิชยกรรม",
                        f"฿{float(s.get('paid_amount', 0)):,.2f}",
                    ],
                    "PAID",
                )
            )
        self.table.set_rows(rows, total_count=total_count, page=self.current_page)
        self._last_data_version = getattr(self.db, "data_version", -1)

    def on_select_row(self, vals: List[str]):
        if not vals:
            return
        code = vals[0]
        stu = self.db.get_student_by_code(code)
        if stu:
            self.ent_code.delete(0, "end")
            self.ent_code.insert(0, stu.get("student_code", ""))
            self.ent_name.delete(0, "end")
            self.ent_name.insert(0, stu.get("full_name", ""))
            self.form_yr_var.set(str(stu.get("year_level", 1)))
            if stu.get("faculty"):
                self.form_fac_var.set(stu["faculty"])
            if stu.get("major"):
                self.form_dept_var.set(stu["major"])
            if stu.get("class_group"):
                self.form_cls_var.set(stu["class_group"])
        else:
            self.ent_code.delete(0, "end")
            self.ent_code.insert(0, vals[0])
            self.ent_name.delete(0, "end")
            self.ent_name.insert(0, vals[1])

    def clear_form(self):
        self.ent_code.delete(0, "end")
        self.ent_name.delete(0, "end")
        self.form_yr_var.set("1")

    def save_or_update_student(self):
        code = self.ent_code.get().strip()
        name = self.ent_name.get().strip()
        if not code or not name:
            if not code:
                highlight_missing_field(self.ent_code, "กรุณากรอกรหัสนักศึกษา", show_toast_msg=True)
            elif not name:
                highlight_missing_field(self.ent_name, "กรุณากรอกชื่อ-นามสกุล", show_toast_msg=True)
            return
        existing = self.db.get_student_by_code(code)
        is_exact = existing is not None
        yr_val = int(self.form_yr_var.get()) if self.form_yr_var.get().isdigit() else 1
        try:
            if is_exact:
                self.db.update_student(
                    student_code=code,
                    full_name=name,
                    faculty=self.form_fac_var.get(),
                    major=self.form_dept_var.get(),
                    class_group=self.form_cls_var.get(),
                    year_level=yr_val,
                    actor=self.session.username,
                    role=self.session.role,
                )
                msg = f"อัปเดตข้อมูลนักศึกษา {code} ({name}) เรียบร้อยแล้ว"
            else:
                self.db.add_student(
                    student_code=code,
                    full_name=name,
                    faculty=self.form_fac_var.get(),
                    major=self.form_dept_var.get(),
                    class_group=self.form_cls_var.get(),
                    education_level=self.form_cls_var.get(),
                    year_level=yr_val,
                    actor=self.session.username,
                    role=self.session.role,
                )
                msg = f"เพิ่มข้อมูลนักศึกษาใหม่ {code} ({name}) เรียบร้อยแล้ว"
            if hasattr(self, "opt_flt_cohort"):
                self.opt_flt_cohort.configure(values=["ทั้งหมด"] + self.db.get_distinct_student_years())
            self.refresh_table()
            self.clear_form()
            messagebox.showinfo("บันทึกสำเร็จ", msg)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def delete_selected_student(self):
        code = self.ent_code.get().strip()
        if not code:
            messagebox.showwarning("เลือกรายชื่อ", "กรุณาคลิกเลือกนักศึกษาในตารางก่อนลบ", parent=self.winfo_toplevel())
            return
        if not messagebox.askyesno("ยืนยันการลบ", f"ต้องการลบข้อมูลนักศึกษารหัส {code} ใช่หรือไม่?", parent=self.winfo_toplevel()):
            return
        try:
            self.db.delete_student(code, actor=self.session.username, role=self.session.role)
            if hasattr(self, "opt_flt_cohort"):
                self.opt_flt_cohort.configure(values=["ทั้งหมด"] + self.db.get_distinct_student_years())
            self.refresh_table()
            self.clear_form()
            messagebox.showinfo("ลบสำเร็จ", f"ลบข้อมูลนักศึกษารหัส {code} เรียบร้อยแล้ว", parent=self.winfo_toplevel())
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการลบ", f"ไม่สามารถลบข้อมูลนักศึกษาได้:\n{e}", parent=self.winfo_toplevel())

    def import_students_dialog(self):
        file_path = filedialog.askopenfilename(
            parent=self.winfo_toplevel(),
            title="เลือกไฟล์รายชื่อนักศึกษา (Excel .xlsx หรือ .csv)",
            filetypes=[("Excel / CSV Files", "*.xlsx *.csv")],
        )
        if not file_path:
            messagebox.showinfo("ยกเลิกการนำเข้า", "คุณได้ยกเลิกการเลือกไฟล์นำเข้า", parent=self.winfo_toplevel())
            return
        try:
            res = self.db.bulk_import_students(
                file_path=file_path,
                actor=self.session.username,
                role=self.session.role,
                default_faculty=self.form_fac_var.get(),
                default_major=self.form_dept_var.get(),
                default_class_group=self.form_cls_var.get(),
            )
            if hasattr(self, "opt_flt_cohort"):
                self.opt_flt_cohort.configure(values=["ทั้งหมด"] + self.db.get_distinct_student_years())
            self.refresh_table()
            messagebox.showinfo(
                "นำเข้ารายชื่อสำเร็จ",
                f"นำเข้าข้อมูลนักศึกษาสำเร็จทั้งหมด {res['total']} คน\n"
                f"• เพิ่มใหม่: {res['inserted']} คน\n"
                f"• อัปเดตข้อมูลเดิม: {res['updated']} คน",
                parent=self.winfo_toplevel(),
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการนำเข้า", f"ไม่สามารถนำเข้าไฟล์ได้:\n{e}", parent=self.winfo_toplevel())


# ==============================================================================
# VIEW 3: RECEIPT REGISTER & DAILY REPORT (หน้า 3: ประวัติ & ยกเลิกใบเสร็จ)
# ==============================================================================
class ReceiptRegisterView(ctk.CTkFrame):
    """
    Minimal Receipt History & Void View with Unified Corporate Palette:
    - Top bar: เลือกวันที่ออกใบเสร็จ + ค้นหา + ยอดรวมสุทธิ + ปุ่ม 'พิมพ์รายงานประจำวัน'
    - Center: ตารางใบเสร็จ
    - Bottom bar: เลขใบเสร็จที่เลือก + ปุ่ม 'พิมพ์ใบเสร็จซ้ำ (.docx)' + ปุ่ม 'ยกเลิกใบเสร็จ (Void)'
    """

    def __init__(self, parent, db: DatabaseManager, session: UserSession, config: Dict[str, Any]):
        super().__init__(parent, fg_color="transparent")
        self.db = db
        self.session = session
        self.config = config

        header = SectionHeader(
            self,
            "ประวัติใบเสร็จ & รายงานประจำวัน",
            "เลือกวันที่เพื่อดูสรุปยอดรายวัน หรือคลิกใบเสร็จในตารางเพื่อพิมพ์ซ้ำ / ยกเลิก",
        )
        header.pack(fill="x", pady=(0, 10))
        header.add_action_button(
            "ส่งออกรายงาน Excel",
            self.export_excel_report,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        )
        header.add_action_button(
            "พิมพ์รายงานประจำวัน (.pdf)",
            self.export_daily_closing_docx,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
        )

        # Sub-tab switcher: ประวัติใบเสร็จ & รายงาน | ประวัติการทำงาน (Audit Log)
        subtab_bar = ctk.CTkFrame(self, fg_color="transparent")
        subtab_bar.pack(fill="x", pady=(0, 8))

        self.active_subtab = "receipts"
        self._audit_panel_built = False
        self._audit_has_perm = False
        self._audit_search_job: Optional[str] = None

        self.btn_subtab_receipts = ctk.CTkButton(
            subtab_bar,
            text="ประวัติใบเสร็จ & รายงาน",
            height=36,
            corner_radius=AppTheme.RADIUS_BTN,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            command=lambda: self.switch_subtab("receipts"),
        )
        self.btn_subtab_receipts.pack(side="left", padx=(0, 8))

        self.btn_subtab_audit = ctk.CTkButton(
            subtab_bar,
            text="ประวัติการทำงาน (Audit Log)",
            height=36,
            corner_radius=AppTheme.RADIUS_BTN,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            command=lambda: self.switch_subtab("audit"),
        )
        self.btn_subtab_audit.pack(side="left")

        self.subtab_container = ctk.CTkFrame(self, fg_color="transparent")
        self.subtab_container.pack(fill="both", expand=True)

        self.panel_receipts = ctk.CTkFrame(self.subtab_container, fg_color="transparent")
        self.panel_audit = ctk.CTkFrame(self.subtab_container, fg_color="transparent")
        self.panel_receipts.pack(fill="both", expand=True)

        # ------------------------------------------------------------------
        # BOTTOM ACTION BAR: Reprint or Void Selected Receipt
        # ------------------------------------------------------------------
        action_bar = ctk.CTkFrame(
            self.panel_receipts,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        action_bar.pack(side="bottom", fill="x", pady=(12, 0))

        ab_in = ctk.CTkFrame(action_bar, fg_color="transparent")
        ab_in.pack(fill="x", padx=20, pady=14)

        ctk.CTkLabel(
            ab_in,
            text="เลขที่ใบเสร็จ:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))

        self.ent_void_no = make_entry(
            ab_in,
            width=175,
            height=AppTheme.CTRL_H_MD,
            font_size=13,
            placeholder_text="คลิกเลือกในตาราง...",
        )
        self.ent_void_no.pack(side="left", padx=(0, 12))

        self.btn_reprint = ctk.CTkButton(
            ab_in,
            text="พิมพ์ใบเสร็จซ้ำ (.pdf)",
            width=185,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.reprint_selected_receipt,
        )
        self.btn_reprint.pack(side="left", padx=(0, 16))

        ctk.CTkLabel(
            ab_in,
            text="สาเหตุยกเลิก:",
            font=(AppTheme.FONT_FAMILY, 13),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))
        self.ent_void_reason = make_entry(ab_in, height=AppTheme.CTRL_H_MD, bold=False)
        self.ent_void_reason.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.ent_void_reason.insert(0, "รายการคลาดเคลื่อน / ออกใบเสร็จใหม่")

        self.btn_void = ctk.CTkButton(
            ab_in,
            text="ยกเลิกใบเสร็จ (Void)",
            width=185,
            height=40,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_DANGER_BG,
            hover_color="#FEE2E2",
            border_width=1,
            border_color=AppTheme.COLOR_DANGER_BORDER,
            text_color=AppTheme.COLOR_DANGER,
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.void_by_number,
        )
        self.btn_void.pack(side="right")

        # ------------------------------------------------------------------
        # TOP DATE, PERIOD & UNIVERSAL SEARCH BAR
        # ------------------------------------------------------------------
        top_bar = ctk.CTkFrame(
            self.panel_receipts,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        top_bar.pack(side="top", fill="x", pady=(0, 10))

        tb_in = ctk.CTkFrame(top_bar, fg_color="transparent")
        tb_in.pack(fill="x", padx=16, pady=10)

        # Row 1: Universal Search (expand=True) + Date Scope Dropdown + Dynamic Date Container + Status Dropdown + Reset Button
        row1 = ctk.CTkFrame(tb_in, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 8))

        self.page_size = 100
        self.current_page = 1
        self._search_job: Optional[str] = None

        self.ent_search = make_entry(
            row1,
            height=36,
            bold=False,
            placeholder_text="🔍 ค้นหาเลขที่ใบเสร็จ / รหัสนักศึกษา / ชื่อ-นามสกุล / แผนกวิชา / วันที่...",
        )
        self.ent_search.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.ent_search.bind("<KeyRelease>", self._schedule_search)

        # Scope Mode Dropdown
        self.period_mode_options = [
            "ทั้งหมด (ทุกวัน)",
            "วันนี้",
            "เมื่อวาน",
            "7 วันล่าสุด",
            "เดือนนี้",
            "เลือกวันที่ระบุ...",
            "กำหนดช่วงวันที่...",
        ]
        self.period_mode_var = ctk.StringVar(value="ทั้งหมด (ทุกวัน)")
        self.opt_period_mode = make_dropdown(
            row1,
            values=self.period_mode_options,
            variable=self.period_mode_var,
            width=150,
            height=36,
            font_size=12,
            command=self._on_period_mode_change,
        )
        self.opt_period_mode.pack(side="left", padx=(0, 6))

        # Dynamic Container for Specific Date or Custom Date Range (Only packed when needed)
        self.date_ctrl_frame = ctk.CTkFrame(row1, fg_color="transparent", width=0, height=36)

        # Cached Formatted Distinct Dates
        self.formatted_dates: List[Dict[str, str]] = []
        self.date_label_to_iso: Dict[str, str] = {}
        self.date_var = ctk.StringVar(value="ทั้งหมด")

        # Specific Date Dropdown (unpacked initially)
        self.opt_date = make_dropdown(
            self.date_ctrl_frame,
            values=["ทั้งหมด"],
            variable=self.date_var,
            width=175,
            height=36,
            font_size=12,
            command=lambda _v: self._reset_and_refresh(),
        )

        # Custom Range Sub-frame (unpacked initially)
        self.range_ctrl_frame = ctk.CTkFrame(self.date_ctrl_frame, fg_color="transparent")
        ctk.CTkLabel(
            self.range_ctrl_frame,
            text="จาก:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(2, 4))
        self.ent_date_from = make_entry(
            self.range_ctrl_frame,
            width=95,
            height=36,
            font_size=12,
            bold=False,
            placeholder_text="วว/ดด/ปปปป",
        )
        self.ent_date_from.pack(side="left", padx=(0, 4))
        self.ent_date_from.bind("<Return>", lambda _e: self._reset_and_refresh())

        ctk.CTkLabel(
            self.range_ctrl_frame,
            text="ถึง:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(2, 4))
        self.ent_date_to = make_entry(
            self.range_ctrl_frame,
            width=95,
            height=36,
            font_size=12,
            bold=False,
            placeholder_text="วว/ดด/ปปปป",
        )
        self.ent_date_to.pack(side="left", padx=(0, 4))
        self.ent_date_to.bind("<Return>", lambda _e: self._reset_and_refresh())

        self.btn_apply_range = ctk.CTkButton(
            self.range_ctrl_frame,
            text="🔍 กรอง",
            width=58,
            height=36,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self._reset_and_refresh,
        )
        self.btn_apply_range.pack(side="left")

        self._refresh_date_options()

        # Status Filter Dropdown
        self.status_var = ctk.StringVar(value="สถานะ: ทั้งหมด")
        self.opt_status = make_dropdown(
            row1,
            values=["สถานะ: ทั้งหมด", "ปกติ (ISSUED)", "ยกเลิกแล้ว (VOID)"],
            variable=self.status_var,
            width=140,
            height=36,
            font_size=12,
            command=lambda _v: self._reset_and_refresh(),
        )
        self.opt_status.pack(side="left", padx=(0, 6))

        # Clear All Filters Button
        self.btn_reset_filters = ctk.CTkButton(
            row1,
            text="🔄 ล้างตัวกรอง",
            width=90,
            height=36,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.reset_all_filters,
        )
        self.btn_reset_filters.pack(side="left")

        # Row 2: Quick Preset Chips + Live Multi-Selection Stats + Financial Summary
        row2 = ctk.CTkFrame(tb_in, fg_color="transparent")
        row2.pack(fill="x")

        preset_box = ctk.CTkFrame(row2, fg_color="transparent")
        preset_box.pack(side="left")

        self.PERIOD_PRESETS = [
            ("ALL", "ทั้งหมด"),
            ("TODAY", "วันนี้"),
            ("YESTERDAY", "เมื่อวาน"),
            ("THIS_WEEK", "7 วันล่าสุด"),
            ("THIS_MONTH", "เดือนนี้"),
        ]
        self.active_preset_key = "ALL"
        self.chip_buttons: Dict[str, ctk.CTkButton] = {}
        for key, lbl in self.PERIOD_PRESETS:
            btn = ctk.CTkButton(
                preset_box,
                text=lbl,
                width=64 if len(lbl) <= 4 else 80,
                height=30,
                corner_radius=AppTheme.RADIUS,
                font=(AppTheme.FONT_FAMILY, 12),
                fg_color=AppTheme.COLOR_PRIMARY_BLUE if key == "ALL" else AppTheme.COLOR_SECONDARY_BG,
                hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER if key == "ALL" else AppTheme.COLOR_SECONDARY_HOVER,
                text_color="#FFFFFF" if key == "ALL" else AppTheme.COLOR_SECONDARY_TEXT,
                command=lambda k=key: self._on_preset_click(k),
            )
            btn.pack(side="left", padx=(0, 6))
            self.chip_buttons[key] = btn

        self.lbl_summary = ctk.CTkLabel(
            row2,
            text="ยอดรับสุทธิ: ฿0.00 บาท (0 ฉบับ)",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        )
        self.lbl_summary.pack(side="right")

        self.lbl_selection = ctk.CTkLabel(
            row2,
            text="",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_PRIMARY_BLUE,
        )
        self.lbl_selection.pack(side="right", padx=(0, 16))

        cols = [
            ("receipt_no", "เลขที่ใบเสร็จ", 130, "center"),
            ("receipt_date", "วันที่", 120, "center"),
            ("student_code", "รหัสนักศึกษา", 120, "center"),
            ("snapshot_student_name", "ชื่อ-นามสกุล", 220, "w"),
            ("snapshot_class_group", "ชั้น/ห้อง", 100, "center"),
            ("item_summary", "รายการค่าธรรมเนียม", 270, "w"),
            ("net_total_amount", "จำนวนเงินรวม", 130, "e"),
            ("status", "สถานะ", 100, "center"),
        ]
        self.table = DataTable(
            self.panel_receipts,
            columns=cols,
            height=12,
            on_select=self.on_select_row,
            show_inspector=False,
            page_size=self.page_size,
            on_page_change=self.on_page_change,
            show_pagination=True,
        )
        self.table.pack(side="top", fill="both", expand=True)
        self.receipts_table = self.table

        self.refresh_receipts()

    def on_show(self):
        """Reloads receipt register (and audit log if built) only if database data has changed."""
        if getattr(self.db, "data_version", None) == getattr(self, "_last_data_version", -1):
            return
        self._refresh_date_options()
        self.refresh_receipts()
        if self._audit_panel_built:
            self.refresh_audit_logs()

    def switch_subtab(self, subtab_key: str):
        """Switches between Receipt Register and Audit Log sub-tabs."""
        self.active_subtab = subtab_key
        self.panel_receipts.pack_forget()
        self.panel_audit.pack_forget()

        for btn, key in [
            (self.btn_subtab_receipts, "receipts"),
            (self.btn_subtab_audit, "audit"),
        ]:
            if key == subtab_key:
                btn.configure(
                    fg_color=AppTheme.COLOR_PRIMARY_BLUE,
                    hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
                    text_color="#FFFFFF",
                    border_width=0,
                )
            else:
                btn.configure(
                    fg_color=AppTheme.COLOR_SECONDARY_BG,
                    hover_color=AppTheme.COLOR_SECONDARY_HOVER,
                    border_width=1,
                    border_color=AppTheme.COLOR_SECONDARY_BORDER,
                    text_color=AppTheme.COLOR_SECONDARY_TEXT,
                )

        if subtab_key == "receipts":
            self.panel_receipts.pack(fill="both", expand=True)
        elif subtab_key == "audit":
            if not self._audit_panel_built:
                self._build_audit_panel()
                self._audit_panel_built = True
            self.panel_audit.pack(fill="both", expand=True)
            self.refresh_audit_logs()

    def _build_audit_panel(self):
        """Lazily builds the Audit Log inspection tab inside ReceiptRegisterView."""
        role = getattr(self.session, "role", "") if self.session is not None else ""
        if not SecurityManager.has_permission(role, "audit.view"):
            self._audit_has_perm = False
            card = ctk.CTkFrame(
                self.panel_audit,
                corner_radius=AppTheme.RADIUS,
                fg_color=AppTheme.COLOR_SURFACE,
                border_width=1,
                border_color=AppTheme.COLOR_BORDER,
            )
            card.pack(fill="both", expand=True, pady=12)
            ctk.CTkLabel(
                card,
                text="คุณไม่มีสิทธิ์เข้าถึงประวัติการทำงาน (Audit Log)",
                font=(AppTheme.FONT_FAMILY, 15, "bold"),
                text_color=AppTheme.COLOR_TEXT_SECONDARY,
            ).pack(expand=True, pady=40)
            return

        self._audit_has_perm = True

        top_bar = ctk.CTkFrame(
            self.panel_audit,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        top_bar.pack(side="top", fill="x", pady=(0, 10))

        tb_in = ctk.CTkFrame(top_bar, fg_color="transparent")
        tb_in.pack(fill="x", padx=16, pady=10)

        row1 = ctk.CTkFrame(tb_in, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 6))

        self.ent_audit_search = make_entry(
            row1,
            height=36,
            bold=False,
            placeholder_text="ค้นหาผู้ทำรายการ / กิจกรรม / เป้าหมาย / รายละเอียด...",
        )
        self.ent_audit_search.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.ent_audit_search.bind("<KeyRelease>", self._schedule_audit_search)
        self.ent_audit_search.bind("<Return>", lambda _e: self.refresh_audit_logs())

        action_opts = ["ทั้งหมด"] + self.db.get_audit_action_types()
        self.audit_action_var = ctk.StringVar(value="ทั้งหมด")
        self.opt_audit_action = make_dropdown(
            row1,
            values=action_opts,
            variable=self.audit_action_var,
            width=180,
            height=36,
            font_size=12,
            command=lambda _v: self.refresh_audit_logs(),
        )
        self.opt_audit_action.pack(side="left", padx=(0, 8))

        ctk.CTkLabel(
            row1,
            text="จาก:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.ent_audit_date_from = make_entry(
            row1,
            width=105,
            height=36,
            font_size=12,
            bold=False,
            placeholder_text="YYYY-MM-DD",
        )
        self.ent_audit_date_from.pack(side="left", padx=(0, 6))
        self.ent_audit_date_from.bind("<Return>", lambda _e: self.refresh_audit_logs())

        ctk.CTkLabel(
            row1,
            text="ถึง:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.ent_audit_date_to = make_entry(
            row1,
            width=105,
            height=36,
            font_size=12,
            bold=False,
            placeholder_text="YYYY-MM-DD",
        )
        self.ent_audit_date_to.pack(side="left", padx=(0, 8))
        self.ent_audit_date_to.bind("<Return>", lambda _e: self.refresh_audit_logs())

        ctk.CTkButton(
            row1,
            text="รีเฟรช",
            width=85,
            height=36,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.refresh_audit_logs,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            row1,
            text="ตรวจสอบความถูกต้อง (Hash Chain)",
            width=215,
            height=36,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.verify_audit_chain,
        ).pack(side="left")

        row2 = ctk.CTkFrame(tb_in, fg_color="transparent")
        row2.pack(fill="x")
        self.lbl_audit_count = ctk.CTkLabel(
            row2,
            text="แสดง 0 รายการล่าสุด",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        )
        self.lbl_audit_count.pack(side="left")

        audit_cols = [
            ("id", "ID", 65, "center"),
            ("timestamp", "วัน-เวลา", 165, "center"),
            ("actor", "ผู้ทำรายการ", 120, "center"),
            ("role", "สิทธิ์", 90, "center"),
            ("action", "กิจกรรม", 150, "center"),
            ("target", "เป้าหมาย", 160, "w"),
            ("detail", "รายละเอียด", 320, "w"),
        ]
        self.audit_table = DataTable(
            self.panel_audit,
            columns=audit_cols,
            height=14,
            show_inspector=True,
        )
        self.audit_table.pack(side="top", fill="both", expand=True)

    def _schedule_audit_search(self, _event=None):
        if self._audit_search_job is not None:
            try:
                self.after_cancel(self._audit_search_job)
            except Exception:
                pass
        self._audit_search_job = self.after(250, self.refresh_audit_logs)

    def refresh_audit_logs(self):
        """Loads filtered audit log entries into the Audit Log DataTable."""
        self._audit_search_job = None
        if not self._audit_panel_built or not self._audit_has_perm:
            return
        action_opts = ["ทั้งหมด"] + self.db.get_audit_action_types()
        self.opt_audit_action.configure(values=action_opts)

        search_txt = self.ent_audit_search.get().strip()
        action_flt = self.audit_action_var.get()
        date_from = self.ent_audit_date_from.get().strip()
        date_to = self.ent_audit_date_to.get().strip()

        logs = self.db.get_audit_logs(
            limit=200,
            search=search_txt,
            action_filter=action_flt,
            date_from=date_from,
            date_to=date_to,
        )
        rows = []
        for entry in logs:
            action_str = str(entry.get("action") or "")
            tag = "VOIDED" if "VOID" in action_str or "DELETE" in action_str else "PAID"
            rows.append(
                (
                    [
                        str(entry.get("id", "")),
                        str(entry.get("timestamp", "")),
                        str(entry.get("actor", "")),
                        str(entry.get("role", "")),
                        action_str,
                        str(entry.get("target", "")),
                        str(entry.get("detail", "")),
                    ],
                    tag,
                )
            )
        self.audit_table.set_rows(rows)
        self.lbl_audit_count.configure(text=f"แสดง {len(rows):,} รายการ")

    def verify_audit_chain(self):
        """Verifies the SHA-256 audit log hash chain and displays the result."""
        ok, count, msg = self.db.verify_audit_integrity()
        if ok:
            messagebox.showinfo(
                "การตรวจสอบความสมบูรณ์สำเร็จ",
                f"สถานะ: {msg}\nจำนวนรายการ Audit Logs: {count:,} รายการ\nลูกโซ่ SHA-256 Hash Chain สมบูรณ์ 100%",
                parent=self.winfo_toplevel(),
            )
        else:
            messagebox.showerror("ตรวจพบความผิดปกติ", msg, parent=self.winfo_toplevel())

    def _refresh_date_options(self):
        try:
            self.formatted_dates = self.db.get_distinct_receipt_dates_formatted()
            self.date_label_to_iso = {item["label"]: item["iso"] for item in self.formatted_dates}
            date_options = ["ทั้งหมด"] + [item["label"] for item in self.formatted_dates]
            if hasattr(self, "opt_date"):
                self.opt_date.configure(values=date_options)
        except Exception:
            pass

    def _update_chip_styles(self):
        for k, btn in self.chip_buttons.items():
            if k == self.active_preset_key:
                btn.configure(
                    fg_color=AppTheme.COLOR_PRIMARY_BLUE,
                    text_color="#FFFFFF",
                    hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
                )
            else:
                btn.configure(
                    fg_color=AppTheme.COLOR_SECONDARY_BG,
                    text_color=AppTheme.COLOR_SECONDARY_TEXT,
                    hover_color=AppTheme.COLOR_SECONDARY_HOVER,
                )

    def _on_preset_click(self, key: str):
        self.active_preset_key = key
        label_map = {
            "ALL": "ทั้งหมด (ทุกวัน)",
            "TODAY": "วันนี้",
            "YESTERDAY": "เมื่อวาน",
            "THIS_WEEK": "7 วันล่าสุด",
            "THIS_MONTH": "เดือนนี้",
        }
        mode_label = label_map.get(key, "ทั้งหมด (ทุกวัน)")
        self.period_mode_var.set(mode_label)
        self.opt_date.pack_forget()
        self.range_ctrl_frame.pack_forget()
        self.date_ctrl_frame.pack_forget()
        self._update_chip_styles()
        self._reset_and_refresh()

    def _on_period_mode_change(self, mode: str):
        rev_presets = {
            "ทั้งหมด (ทุกวัน)": "ALL",
            "ทั้งหมด": "ALL",
            "วันนี้": "TODAY",
            "เมื่อวาน": "YESTERDAY",
            "7 วันล่าสุด": "THIS_WEEK",
            "เดือนนี้": "THIS_MONTH",
        }
        if mode in rev_presets:
            self.active_preset_key = rev_presets[mode]
            self.opt_date.pack_forget()
            self.range_ctrl_frame.pack_forget()
            self.date_ctrl_frame.pack_forget()
        elif mode == "เลือกวันที่ระบุ...":
            self.active_preset_key = "CUSTOM_DATE"
            self.range_ctrl_frame.pack_forget()
            self.date_ctrl_frame.pack(side="left", padx=(0, 6), after=self.opt_period_mode)
            self.opt_date.pack(side="left")
        elif mode == "กำหนดช่วงวันที่...":
            self.active_preset_key = "CUSTOM_RANGE"
            self.opt_date.pack_forget()
            self.date_ctrl_frame.pack(side="left", padx=(0, 6), after=self.opt_period_mode)
            self.range_ctrl_frame.pack(side="left")
        else:
            self.active_preset_key = "ALL"
            self.opt_date.pack_forget()
            self.range_ctrl_frame.pack_forget()
            self.date_ctrl_frame.pack_forget()
        self._update_chip_styles()
        self._reset_and_refresh()

    def reset_all_filters(self):
        self.active_preset_key = "ALL"
        self.period_mode_var.set("ทั้งหมด (ทุกวัน)")
        self.status_var.set("สถานะ: ทั้งหมด")
        self.date_var.set("ทั้งหมด")
        self.ent_search.delete(0, "end")
        self.ent_date_from.delete(0, "end")
        self.ent_date_to.delete(0, "end")
        self.opt_date.pack_forget()
        self.range_ctrl_frame.pack_forget()
        self.date_ctrl_frame.pack_forget()
        self._update_chip_styles()
        self._reset_and_refresh()

    def _get_active_date_filter(self) -> Tuple[str, str, str]:
        if self.active_preset_key in ("TODAY", "YESTERDAY", "THIS_WEEK", "THIS_MONTH"):
            return self.active_preset_key, "", ""
        if self.active_preset_key == "CUSTOM_DATE":
            val = self.date_var.get()
            if not val or val == "ทั้งหมด":
                return "", "", ""
            iso_d = self.date_label_to_iso.get(val, val)
            return iso_d, "", ""
        if self.active_preset_key == "CUSTOM_RANGE":
            d_from = self.ent_date_from.get().strip()
            d_to = self.ent_date_to.get().strip()
            return "", d_from, d_to
        val = self.date_var.get()
        if val and val != "ทั้งหมด":
            iso_d = self.date_label_to_iso.get(val, val)
            return iso_d, "", ""
        return "", "", ""

    def _schedule_search(self, _event=None):
        if self._search_job is not None:
            try:
                self.after_cancel(self._search_job)
            except Exception:
                pass
        self._search_job = self.after(250, self._reset_and_refresh)

    def _reset_and_refresh(self):
        self._search_job = None
        self.current_page = 1
        self.refresh_receipts()

    def on_page_change(self, new_page: int):
        self.current_page = max(1, int(new_page))
        self.refresh_receipts()

    def refresh_receipts(self):
        date_pref, d_from, d_to = self._get_active_date_filter()
        search_txt = self.ent_search.get().strip()
        status_txt = self.status_var.get()
        if status_txt.startswith("สถานะ: "):
            status_txt = status_txt[7:].strip()
        if status_txt == "ทั้งหมด":
            status_txt = "ทั้งหมด"

        stats = self.db.get_receipts_stats(
            status_filter=status_txt,
            search=search_txt,
            date_prefix=date_pref,
            date_from=d_from,
            date_to=d_to,
        )
        total_count = stats["total_count"]
        valid_count = stats["valid_count"]
        void_count = stats.get("void_count", 0)
        total_amt = stats["total_amount"]

        max_pages = max(1, (total_count + self.page_size - 1) // self.page_size)
        if self.current_page > max_pages:
            self.current_page = max_pages

        offset = (self.current_page - 1) * self.page_size
        receipts = self.db.get_receipts(
            status_filter=status_txt,
            search=search_txt,
            date_prefix=date_pref,
            date_from=d_from,
            date_to=d_to,
            limit=self.page_size,
            offset=offset,
        )
        rows = []
        for r in receipts:
            st = r.get("status", "ISSUED")
            amt = float(r.get("net_total_amount", 0))
            tag = "PAID" if st == "ISSUED" else "VOIDED"
            rows.append(
                (
                    [
                        r["receipt_no"],
                        format_receipt_date_thai(r["receipt_date"]),
                        r["student_code"],
                        r["snapshot_student_name"],
                        r["snapshot_class_group"],
                        r.get("item_summary", "-"),
                        f"฿{amt:,.2f}",
                        "ปกติ" if st == "ISSUED" else "ยกเลิกแล้ว",
                    ],
                    tag,
                )
            )
        self.table.set_rows(rows, total_count=total_count, page=self.current_page)
        self._last_data_version = getattr(self.db, "data_version", -1)
        summary_str = f"ยอดรับสุทธิ: ฿{total_amt:,.2f} บาท ({valid_count:,} ฉบับ)"
        if void_count > 0:
            summary_str += f" • ยกเลิก {void_count:,} ฉบับ"
        self.lbl_summary.configure(text=summary_str)
        self.on_select_row([])
        self._last_data_version = getattr(self.db, "data_version", -1)

    def on_select_row(self, vals: Optional[List[str]] = None):
        selected_rows = self.table.get_all_selected_values()
        count = len(selected_rows)
        if count == 0:
            self.ent_void_no.delete(0, "end")
            self.lbl_selection.configure(text="")
            self.btn_reprint.configure(text="🖨️ พิมพ์ใบเสร็จซ้ำ (.pdf)")
            self.btn_void.configure(text="🚫 ยกเลิกใบเสร็จ (Void)")
            return
        if count == 1:
            rcp_no = selected_rows[0][0]
            amt_str = selected_rows[0][6]
            self.ent_void_no.delete(0, "end")
            self.ent_void_no.insert(0, rcp_no)
            self.lbl_selection.configure(text=f"📌 เลือก 1 ฉบับ: {rcp_no} ({amt_str})")
            self.btn_reprint.configure(text="🖨️ พิมพ์ใบเสร็จใบนี้ซ้ำ (.pdf)")
            self.btn_void.configure(text="🚫 ยกเลิกใบเสร็จนี้ (Void)")
        else:
            rcp_nos = [r[0] for r in selected_rows]
            total_sel = 0.0
            for r in selected_rows:
                try:
                    amt_cleaned = r[6].replace("฿", "").replace(",", "").strip()
                    total_sel += float(amt_cleaned)
                except Exception:
                    pass
            self.ent_void_no.delete(0, "end")
            self.ent_void_no.insert(0, ", ".join(rcp_nos))
            self.lbl_selection.configure(
                text=f"📌 เลือกอยู่ {count:,} ฉบับ (รวม ฿{total_sel:,.2f})"
            )
            self.btn_reprint.configure(text=f"🖨️ พิมพ์ใบเสร็จที่เลือก ({count:,} ฉบับ)")
            self.btn_void.configure(text=f"🚫 ยกเลิกที่เลือก ({count:,} ฉบับ)")

    def void_by_number(self):
        selected_rows = self.table.get_all_selected_values()
        if selected_rows:
            rcp_nos = [r[0] for r in selected_rows if r and r[0]]
        else:
            raw = self.ent_void_no.get().strip()
            rcp_nos = [x.strip() for x in raw.split(",") if x.strip()] if raw else []

        reason = self.ent_void_reason.get().strip() or "ยกเลิกใบเสร็จ"
        if not rcp_nos:
            messagebox.showwarning("ระบุเลขที่ใบเสร็จ", "กรุณาคลิกเลือกใบเสร็จในตาราง หรือพิมพ์เลขใบเสร็จก่อนยกเลิก")
            return

        if len(rcp_nos) == 1:
            rcp_no = rcp_nos[0]
            if not messagebox.askyesno(
                "ยืนยันยกเลิกใบเสร็จ",
                f"ต้องการยกเลิกใบเสร็จเลขที่ '{rcp_no}' ใช่หรือไม่?\nสาเหตุ: {reason}",
            ):
                return
            try:
                self.db.void_receipt(rcp_no, reason, self.session.username, self.session.role)
                self._refresh_date_options()
                self.refresh_receipts()
                messagebox.showinfo("ยกเลิกสำเร็จ", f"ยกเลิกใบเสร็จเลขที่ {rcp_no} เรียบร้อยแล้ว")
            except Exception as e:
                messagebox.showerror("ไม่สามารถยกเลิกได้", str(e))
        else:
            sample_str = ", ".join(rcp_nos[:4]) + ("..." if len(rcp_nos) > 4 else "")
            if not messagebox.askyesno(
                "ยืนยันยกเลิกหลายใบเสร็จ",
                f"คุณกำลังจะยกเลิกใบเสร็จทั้งหมด {len(rcp_nos)} ฉบับ:\n({sample_str})\n\nสาเหตุ: {reason}\n\nต้องการดำเนินการต่อหรือไม่?",
            ):
                return
            success_count = 0
            fail_count = 0
            for no in rcp_nos:
                try:
                    self.db.void_receipt(no, reason, self.session.username, self.session.role)
                    success_count += 1
                except Exception:
                    fail_count += 1
            self._refresh_date_options()
            self.refresh_receipts()
            msg = f"ยกเลิกใบเสร็จเรียบร้อยแล้ว {success_count} ฉบับ"
            if fail_count > 0:
                msg += f" (ไม่สำเร็จ {fail_count} ฉบับ)"
            messagebox.showinfo("ผลการยกเลิกใบเสร็จ", msg)

    def reprint_selected_receipt(self):
        selected_rows = self.table.get_all_selected_values()
        if selected_rows:
            rcp_nos = [r[0] for r in selected_rows if r and r[0]]
        else:
            raw = self.ent_void_no.get().strip()
            rcp_nos = [x.strip() for x in raw.split(",") if x.strip()] if raw else []

        if not rcp_nos:
            messagebox.showwarning("เลือกใบเสร็จ", "กรุณาคลิกเลือกใบเสร็จในตารางก่อนสั่งพิมพ์ซ้ำ")
            return

        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)

        if len(rcp_nos) == 1:
            rcp_no = rcp_nos[0]
            data = self.db.get_receipt_by_no(rcp_no)
            if not data:
                messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบใบเสร็จเลขที่ {rcp_no}", parent=self.winfo_toplevel())
                return
            path = filedialog.asksaveasfilename(
                parent=self.winfo_toplevel(),
                title="เลือกตำแหน่งบันทึกไฟล์ใบเสร็จซ้ำ (.pdf)",
                initialdir=out_dir,
                initialfile=f"Reprint_Receipt_{rcp_no}.pdf",
                defaultextension=".pdf",
                filetypes=[
                    ("PDF Document (*.pdf)", "*.pdf"),
                    ("Word Document (*.docx)", "*.docx"),
                ],
            )
            if not path:
                messagebox.showinfo("ยกเลิกการพิมพ์ซ้ำ", "คุณได้ยกเลิกการเลือกตำแหน่งบันทึกไฟล์ เอกสารจึงยังไม่ถูกส่งออก", parent=self.winfo_toplevel())
                return
            self.config["default_export_folder"] = os.path.dirname(path)
            save_config(self.config)
            try:
                export_document_auto(
                    path,
                    data,
                    college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
                    metadata=self.config,
                    is_batch=False,
                )
                open_ok = open_document_file(path, parent=self.winfo_toplevel())
                messagebox.showinfo(
                    "พิมพ์ใบเสร็จซ้ำสำเร็จ",
                    f"บันทึกไฟล์ใบเสร็จเลขที่ {rcp_no} เรียบร้อยแล้วที่:\n{path}\n\n"
                    + ("(ระบบได้เปิดไฟล์เอกสารขึ้นมาให้พร้อมพิมพ์เรียบร้อยแล้ว)" if open_ok else ""),
                    parent=self.winfo_toplevel(),
                )
            except Exception as e:
                messagebox.showerror("เกิดข้อผิดพลาดในการพิมพ์ซ้ำ", f"ไม่สามารถสร้างหรือเปิดไฟล์เอกสารได้:\n{e}", parent=self.winfo_toplevel())
        else:
            receipts_data = []
            for no in rcp_nos:
                d = self.db.get_receipt_by_no(no)
                if d:
                    receipts_data.append(d)
            if not receipts_data:
                messagebox.showerror("ไม่พบข้อมูล", "ไม่พบข้อมูลใบเสร็จที่เลือก", parent=self.winfo_toplevel())
                return
            path = filedialog.asksaveasfilename(
                parent=self.winfo_toplevel(),
                title=f"เลือกตำแหน่งบันทึกไฟล์ใบเสร็จรวม {len(receipts_data)} ฉบับ (.pdf)",
                initialdir=out_dir,
                initialfile=f"Batch_Receipts_{len(receipts_data)}_items.pdf",
                defaultextension=".pdf",
                filetypes=[
                    ("PDF Document (*.pdf)", "*.pdf"),
                    ("Word Document (*.docx)", "*.docx"),
                ],
            )
            if not path:
                messagebox.showinfo("ยกเลิกการพิมพ์ซ้ำ", "คุณได้ยกเลิกการเลือกตำแหน่งบันทึกไฟล์ เอกสารจึงยังไม่ถูกส่งออก", parent=self.winfo_toplevel())
                return
            self.config["default_export_folder"] = os.path.dirname(path)
            save_config(self.config)
            try:
                export_document_auto(
                    path,
                    receipts_data,
                    college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
                    metadata=self.config,
                    is_batch=True,
                )
                open_ok = open_document_file(path, parent=self.winfo_toplevel())
                messagebox.showinfo(
                    "พิมพ์ใบเสร็จซ้ำสำเร็จ",
                    f"รวมไฟล์ใบเสร็จจำนวน {len(receipts_data)} ฉบับ บันทึกอยู่ที่:\n{path}\n\n"
                    + ("(ระบบได้เปิดไฟล์เอกสารขึ้นมาให้พร้อมพิมพ์เรียบร้อยแล้ว)" if open_ok else ""),
                    parent=self.winfo_toplevel(),
                )
            except Exception as e:
                messagebox.showerror("เกิดข้อผิดพลาดในการพิมพ์ซ้ำ", f"ไม่สามารถสร้างหรือเปิดไฟล์เอกสารได้:\n{e}", parent=self.winfo_toplevel())

    def export_daily_closing_docx(self):
        date_pref, d_from, d_to = self._get_active_date_filter()
        target_date = date_pref if date_pref else datetime.date.today().isoformat()
        closing = self.db.get_daily_closing_summary(
            date_prefix=date_pref if not (d_from or d_to) else None,
            date_from=d_from,
            date_to=d_to,
        )
        closing["report_type"] = "DAILY_CLOSING"
        disp_date = closing.get("date", target_date)
        closing["doc_title"] = f"รายงานสรุปการรับเงินค่าเทอมประจำวัน ({disp_date})"
        closing["recorded_by"] = self.session.full_name

        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)
        safe_tag = str(disp_date).replace("/", "-").replace(" ", "_")
        path = filedialog.asksaveasfilename(
            parent=self.winfo_toplevel(),
            title="เลือกตำแหน่งบันทึกรายงานประจำวัน (.pdf)",
            initialdir=out_dir,
            initialfile=f"SPVC_DailyReport_{safe_tag}.pdf",
            defaultextension=".pdf",
            filetypes=[
                ("PDF Document (*.pdf)", "*.pdf"),
                ("Word Document (*.docx)", "*.docx"),
            ],
        )
        if not path:
            messagebox.showinfo("ยกเลิกการพิมพ์รายงาน", "คุณได้ยกเลิกการเลือกตำแหน่งบันทึกไฟล์", parent=self.winfo_toplevel())
            return
        self.config["default_export_folder"] = os.path.dirname(path)
        save_config(self.config)
        try:
            export_document_auto(
                path,
                closing,
                college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
                metadata=self.config,
                is_batch=False,
            )
            open_ok = open_document_file(path, parent=self.winfo_toplevel())
            messagebox.showinfo(
                "พิมพ์รายงานประจำวันสำเร็จ",
                f"สร้างไฟล์รายงานสรุปยอดประจำวัน ({disp_date}) เรียบร้อยที่:\n{path}\n\n"
                + ("(ระบบได้เปิดไฟล์เอกสารขึ้นมาให้พร้อมพิมพ์เรียบร้อยแล้ว)" if open_ok else ""),
                parent=self.winfo_toplevel(),
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการพิมพ์รายงาน", f"ไม่สามารถสร้างหรือพิมพ์รายงานได้:\n{e}", parent=self.winfo_toplevel())

    def export_excel_report(self):
        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = filedialog.asksaveasfilename(
            parent=self.winfo_toplevel(),
            title="เลือกตำแหน่งบันทึกรายงาน Excel (.xlsx)",
            initialdir=out_dir,
            initialfile=f"SPVC_Receipt_Report_{stamp}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook (*.xlsx)", "*.xlsx")],
        )
        if not path:
            messagebox.showinfo("ยกเลิกการส่งออก Excel", "คุณได้ยกเลิกการเลือกตำแหน่งบันทึกไฟล์", parent=self.winfo_toplevel())
            return
        self.config["default_export_folder"] = os.path.dirname(path)
        save_config(self.config)
        date_pref, d_from, d_to = self._get_active_date_filter()
        search_txt = self.ent_search.get().strip()
        st_val = self.status_var.get()
        if st_val.startswith("สถานะ: "):
            st_val = st_val[7:].strip()
        payload = {
            "summary": self.db.get_dashboard_summary(),
            "transactions": self.db.get_receipts(
                status_filter=st_val,
                search=search_txt,
                date_prefix=date_pref,
                date_from=d_from,
                date_to=d_to,
            ),
            "students": self.db.get_students(search=search_txt),
            "budgets": [],
        }
        try:
            ExcelFinanceExporter(path, college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")).export(
                payload, metadata=self.config
            )
            open_ok = open_document_file(path, parent=self.winfo_toplevel())
            messagebox.showinfo(
                "ส่งออก Excel สำเร็จ",
                f"บันทึกไฟล์รายงาน Excel เรียบร้อยที่:\n{path}\n\n"
                + ("(ระบบได้เปิดไฟล์เอกสารขึ้นมาเรียบร้อยแล้ว)" if open_ok else ""),
                parent=self.winfo_toplevel(),
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการส่งออก Excel", f"ไม่สามารถสร้างไฟล์ Excel ได้:\n{e}", parent=self.winfo_toplevel())


# ==============================================================================
# VIEW 4: SETTINGS, DROPDOWN MASTER & FEE CATALOG (หน้า 4: ตั้งค่าระบบ)
# ==============================================================================
class SettingsAndBackupView(ctk.CTkFrame):
    """
    Enterprise-Grade Settings, Fee Catalog, Dropdown Master & System Backup View:
    - แท็บ 1: 🏛️ ข้อมูลสถานศึกษา & เลขที่ใบเสร็จ (General & Receipt Settings)
    - แท็บ 2: 💰 รายการค่าธรรมเนียมตามหลักสูตร (Fee Catalog Management)
    - แท็บ 3: 📋 จัดการตัวเลือก Dropdown (Master Dropdown Options)
    - แท็บ 4: 🛡️ สำรองข้อมูล & ความปลอดภัยระบบ (Database Backup & Security & Webhook)
    """

    def __init__(
        self,
        parent,
        db: DatabaseManager,
        session: UserSession,
        config: Dict[str, Any],
        webhook: Optional[FinanceWebhookDispatcher] = None,
        on_config_updated: Optional[Callable[[], None]] = None,
    ):
        super().__init__(parent, fg_color="transparent")
        self.db = db
        self.session = session
        self.config = config
        self.on_config_updated = on_config_updated
        self.webhook = webhook or FinanceWebhookDispatcher(
            webhook_url=self.config.get("webhook_url", ""),
            enabled=self.config.get("webhook_enabled", False),
            alert_threshold=self.config.get("webhook_alert_threshold", 50000.0),
        )

        self._fee_search_job: Optional[str] = None
        self._selected_dropdown_val: Optional[str] = None

        header = SectionHeader(
            self,
            "ตั้งค่าระบบ",
            "ข้อมูลสถานศึกษา, เลขที่ใบเสร็จ, รายการค่าธรรมเนียม, ตัวเลือก Dropdown และระบบสำรองข้อมูล",
        )
        header.pack(fill="x", pady=(0, 10))

        header.add_action_button(
            "บันทึกการตั้งค่าทั้งหมด",
            self.save_all_settings_quick,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
        )
        header.add_action_button(
            "สำรองข้อมูลด่วน (.zip)",
            self.run_zip_backup,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
        )

        # ------------------------------------------------------------------
        # TOP 4-TAB NAVIGATION BAR
        # ------------------------------------------------------------------
        tab_nav_card = ctk.CTkFrame(
            self,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        tab_nav_card.pack(fill="x", pady=(0, 12))
        self.settings_tabs = tab_nav_card

        tab_nav_grid = ctk.CTkFrame(tab_nav_card, fg_color="transparent")
        tab_nav_grid.pack(fill="x", padx=8, pady=8)
        for i in range(4):
            tab_nav_grid.grid_columnconfigure(i, weight=1, uniform="settings_tab")

        self.tab_defs = [
            ("general", "ข้อมูลสถานศึกษา & ใบเสร็จ"),
            ("fees", "รายการค่าธรรมเนียม"),
            ("dropdowns", "จัดการตัวเลือก Dropdown"),
            ("backup", "สำรองข้อมูล & ความปลอดภัย"),
        ]

        self.tab_buttons: Dict[str, ctk.CTkButton] = {}
        for idx, (key, label) in enumerate(self.tab_defs):
            btn = ctk.CTkButton(
                tab_nav_grid,
                text=label,
                height=40,
                corner_radius=AppTheme.RADIUS,
                font=(AppTheme.FONT_FAMILY, 13, "bold"),
                command=lambda k=key: self.switch_settings_tab(k),
            )
            btn.grid(row=0, column=idx, sticky="nsew", padx=4)
            self.tab_buttons[key] = btn

        # Container for tab content frames
        self.tab_container = ctk.CTkFrame(self, fg_color="transparent")
        self.tab_container.pack(fill="both", expand=True)

        self.tab_general = ctk.CTkFrame(self.tab_container, fg_color="transparent")
        self.tab_fees = ctk.CTkFrame(self.tab_container, fg_color="transparent")
        self.tab_dropdowns = ctk.CTkFrame(self.tab_container, fg_color="transparent")
        self.tab_backup = ctk.CTkFrame(self.tab_container, fg_color="transparent")

        self.tab_frames: Dict[str, ctk.CTkFrame] = {
            "general": self.tab_general,
            "fees": self.tab_fees,
            "dropdowns": self.tab_dropdowns,
            "backup": self.tab_backup,
        }

        self._built_tabs: Dict[str, bool] = {
            "general": True,
            "fees": False,
            "dropdowns": False,
            "backup": False,
        }
        self._build_tab_general()

        # Start on General settings tab
        self.switch_settings_tab("general")
        self._last_data_version = getattr(self.db, "data_version", -1)

    def on_show(self):
        """Refreshes active settings tables only if database data has changed."""
        if getattr(self.db, "data_version", None) == getattr(self, "_last_data_version", -1):
            return
        if self.active_tab == "fees" and self._built_tabs.get("fees"):
            self.refresh_fee_table()
        elif self.active_tab == "dropdowns" and self._built_tabs.get("dropdowns"):
            self.refresh_dropdown_table()
        self._last_data_version = getattr(self.db, "data_version", -1)

    def switch_settings_tab(self, tab_key: str):
        self.active_tab = tab_key
        for key, btn in self.tab_buttons.items():
            if key == tab_key:
                btn.configure(
                    fg_color=AppTheme.COLOR_PRIMARY_BLUE,
                    hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
                    text_color="#FFFFFF",
                    border_width=0,
                )
            else:
                btn.configure(
                    fg_color=AppTheme.COLOR_SURFACE_ALT,
                    hover_color=AppTheme.COLOR_SECONDARY_HOVER,
                    text_color=AppTheme.COLOR_TEXT_SECONDARY,
                    border_width=1,
                    border_color=AppTheme.COLOR_BORDER,
                )

        for f in self.tab_frames.values():
            f.pack_forget()

        if tab_key in self.tab_frames:
            was_built = self._built_tabs.get(tab_key, False)
            if not was_built:
                if tab_key == "fees":
                    self._build_tab_fees()
                elif tab_key == "dropdowns":
                    self._build_tab_dropdowns()
                elif tab_key == "backup":
                    self._build_tab_backup()
                self._built_tabs[tab_key] = True
            self.tab_frames[tab_key].pack(fill="both", expand=True)
            if was_built:
                if tab_key == "fees":
                    self.refresh_fee_table()
                elif tab_key == "dropdowns":
                    self.refresh_dropdown_table()

    @property
    def tabs(self):
        outer = self
        class _TabviewCompat:
            def set(self, val: str):
                for k, label in outer.tab_defs:
                    if val in label or label in val or k == val:
                        outer.switch_settings_tab(k)
                        break
            def get(self) -> str:
                for k, label in outer.tab_defs:
                    if k == outer.active_tab:
                        return label
                return outer.active_tab
        return _TabviewCompat()

    # ==========================================================================
    # TAB 1: ข้อมูลสถานศึกษา & เลขที่ใบเสร็จ
    # ==========================================================================
    def _build_tab_general(self):
        grid = ctk.CTkFrame(self.tab_general, fg_color="transparent")
        grid.pack(fill="both", expand=True, padx=14, pady=10)
        grid.grid_columnconfigure((0, 1), weight=1)
        grid.grid_rowconfigure(0, weight=1)

        # Left Card: ข้อมูลสถานศึกษา & เอกสารส่งออก
        c_left = ctk.CTkFrame(
            grid,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        c_left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        ctk.CTkLabel(
            c_left,
            text="🏛️ ข้อมูลสถานศึกษา & เอกสารส่งออก",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        ).pack(anchor="w", padx=16, pady=(14, 8))

        self.ent_college = self._labeled_entry(
            c_left, "ที่ทำการ (ชื่อสถานศึกษา):", self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")
        )
        self.ent_college_address = self._labeled_entry(
            c_left,
            "ที่อยู่สถานศึกษา:",
            self.config.get(
                "college_address", "455 ถ.พระพันวษา ต.ท่าพี่เลี้ยง อ.เมืองสุพรรณบุรี จ.สุพรรณบุรี 72000"
            ),
        )
        self.ent_tax_id = self._labeled_entry(
            c_left, "เลขประจำตัวผู้เสียภาษี:", self.config.get("college_tax_id", "0994000159721")
        )

        style_row = ctk.CTkFrame(c_left, fg_color="transparent")
        style_row.pack(fill="x", padx=16, pady=5)
        ctk.CTkLabel(
            style_row,
            text="รูปแบบใบเสร็จ/อศท.28:",
            width=160,
            anchor="w",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left")
        self.sys_style_var = ctk.StringVar(value=_style_label_from_code(self.config.get("bill_template_style", "CLASSIC")))
        make_dropdown(
            style_row,
            values=TEMPLATE_STYLE_OPTIONS,
            variable=self.sys_style_var,
            height=36,
        ).pack(side="left", fill="x", expand=True)

        exp_row = ctk.CTkFrame(c_left, fg_color="transparent")
        exp_row.pack(fill="x", padx=16, pady=5)
        ctk.CTkLabel(
            exp_row,
            text="โฟลเดอร์ส่งออกเอกสาร:",
            width=160,
            anchor="w",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left")
        self.ent_export_folder = make_entry(exp_row, height=36, bold=False)
        self.ent_export_folder.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.ent_export_folder.insert(0, self.config.get("default_export_folder", os.getcwd()))
        ctk.CTkButton(
            exp_row,
            text="📁 เลือก...",
            width=70,
            height=36,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.browse_export_folder,
        ).pack(side="left")

        # Right Card: ข้อมูลเล่มที่ & เลขที่ใบเสร็จ
        c_right = ctk.CTkFrame(
            grid,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        c_right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        ctk.CTkLabel(
            c_right,
            text="🧾 ข้อมูลเล่มที่ & เลขที่ใบเสร็จรับเงิน",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        ).pack(anchor="w", padx=16, pady=(14, 8))

        term_row = ctk.CTkFrame(c_right, fg_color="transparent")
        term_row.pack(fill="x", padx=16, pady=5)
        ctk.CTkLabel(
            term_row,
            text="ปีการศึกษา / ปีงบ:",
            width=140,
            anchor="w",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left")
        self.ent_fiscal_year = make_entry(term_row, width=85, height=36)
        self.ent_fiscal_year.pack(side="left", padx=(0, 14))
        self.ent_fiscal_year.insert(0, str(self.config.get("fiscal_year", "2569")))

        ctk.CTkLabel(
            term_row,
            text="ภาค/ปีการศึกษาปัจจุบัน:",
            width=140,
            anchor="w",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))
        terms = self.db.get_options("TERM") or ["1/2569", "2/2569"]
        self.term_var = ctk.StringVar(value=self.config.get("current_semester", terms[0]))
        make_dropdown(
            term_row,
            values=terms,
            variable=self.term_var,
            width=135,
            height=36,
        ).pack(side="left")

        pos_row = ctk.CTkFrame(c_right, fg_color="transparent")
        pos_row.pack(fill="x", padx=16, pady=5)
        ctk.CTkLabel(
            pos_row,
            text="ตำแหน่งผู้รับเงิน:",
            width=140,
            anchor="w",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left")
        positions = self.db.get_options("CASHIER_POSITION") or ["ครู", "เจ้าหน้าที่การเงิน"]
        self.pos_var = ctk.StringVar(value=self.config.get("cashier_position", positions[0]))
        make_dropdown(
            pos_row,
            values=positions,
            variable=self.pos_var,
            width=210,
            height=36,
        ).pack(side="left")

        book_row = ctk.CTkFrame(c_right, fg_color="transparent")
        book_row.pack(fill="x", padx=16, pady=5)
        ctk.CTkLabel(
            book_row,
            text="เล่มที่ใบเสร็จ:",
            width=140,
            anchor="w",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left")
        self.ent_book = make_entry(book_row, width=85, height=36)
        self.ent_book.pack(side="left", padx=(0, 14))
        self.ent_book.insert(0, str(self.config.get("receipt_book_no", "7006")))

        ctk.CTkLabel(
            book_row,
            text="คำนำหน้า (Prefix):",
            width=110,
            anchor="w",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))
        self.ent_prefix = make_entry(book_row, width=120, height=36, bold=False)
        self.ent_prefix.pack(side="left")
        self.ent_prefix.insert(0, str(self.config.get("receipt_prefix", "")))

        range_row = ctk.CTkFrame(c_right, fg_color="transparent")
        range_row.pack(fill="x", padx=16, pady=5)
        ctk.CTkLabel(
            range_row,
            text="ช่วงเลขที่ใบเสร็จ:",
            width=140,
            anchor="w",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left")
        self.ent_start_no = make_entry(range_row, width=120, height=36)
        self.ent_start_no.pack(side="left", padx=(0, 6))
        self.ent_start_no.insert(0, str(self.config.get("receipt_no_start", "700605433")))
        ctk.CTkLabel(
            range_row,
            text="ถึง",
            font=(AppTheme.FONT_FAMILY, 13),
            text_color=AppTheme.COLOR_TEXT_MUTED,
        ).pack(side="left", padx=(0, 6))
        self.ent_end_no = make_entry(range_row, width=120, height=36)
        self.ent_end_no.pack(side="left")
        self.ent_end_no.insert(0, str(self.config.get("receipt_no_end", "700999999")))

        auto_bk_row = ctk.CTkFrame(c_right, fg_color="transparent")
        auto_bk_row.pack(fill="x", padx=16, pady=(10, 5))
        self.auto_backup_var = ctk.BooleanVar(value=bool(self.config.get("auto_backup_on_exit", True)))
        self.chk_auto_backup = ctk.CTkCheckBox(
            auto_bk_row,
            text="สำรองฐานข้อมูลอัตโนมัติทุกครั้งเมื่อปิดโปรแกรม (Auto-backup on exit)",
            variable=self.auto_backup_var,
            font=(AppTheme.FONT_FAMILY, 13),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
        )
        self.chk_auto_backup.pack(side="left")

        reg_menu_card = ctk.CTkFrame(
            c_right,
            fg_color="#F8FAFC",
            corner_radius=AppTheme.RADIUS,
            border_width=1,
            border_color="#CBD5E1",
        )
        reg_menu_card.pack(fill="x", padx=16, pady=(12, 6))

        reg_card_inner = ctk.CTkFrame(reg_menu_card, fg_color="transparent")
        reg_card_inner.pack(fill="x", padx=14, pady=10)

        self.show_registration_var = ctk.BooleanVar(value=bool(self.config.get("show_course_registration", False)))
        self.chk_show_registration = ctk.CTkCheckBox(
            reg_card_inner,
            text="แสดงเมนู 'ทะเบียนวิชา (อศท.28)' บนแถบเมนูหลักด้านข้าง (Sidebar)",
            variable=self.show_registration_var,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            command=self.on_toggle_course_registration,
        )
        self.chk_show_registration.pack(anchor="w")

        ctk.CTkLabel(
            reg_card_inner,
            text="ติ๊กเลือกเพื่อแสดง หรือเอาติ๊กออกเพื่อซ่อนเมนูทะเบียนวิชาและการพิมพ์บัตร อศท.28 จากแถบข้างทันที",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color=AppTheme.COLOR_TEXT_MUTED,
        ).pack(anchor="w", padx=(28, 0), pady=(3, 0))

        # Action Bar ด้านล่าง
        btn_bar = ctk.CTkFrame(self.tab_general, fg_color="transparent")
        btn_bar.pack(fill="x", padx=14, pady=(6, 10))

        ctk.CTkButton(
            btn_bar,
            text="💾 บันทึกการตั้งค่าข้อมูลสถานศึกษา & ใบเสร็จ",
            height=42,
            width=320,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.save_general_config,
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_bar,
            text="🔄 คืนค่าเริ่มต้น",
            height=42,
            width=150,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.reset_general_config_defaults,
        ).pack(side="left")

    def _labeled_entry(self, parent, label_text: str, initial_val: str) -> ctk.CTkEntry:
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=5)
        ctk.CTkLabel(
            row,
            text=label_text,
            width=160,
            anchor="w",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left")
        ent = make_entry(row, height=36)
        ent.pack(side="left", fill="x", expand=True)
        ent.insert(0, str(initial_val))
        return ent

    def browse_export_folder(self):
        folder = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับบันทึกเอกสารส่งออก")
        if folder:
            self.ent_export_folder.delete(0, "end")
            self.ent_export_folder.insert(0, folder)

    def on_toggle_course_registration(self):
        """Immediately toggles course registration in config and live updates the sidebar navigation."""
        val = bool(self.show_registration_var.get())
        self.config["show_course_registration"] = val
        try:
            save_config(self.config)
        except Exception:
            pass
        if hasattr(self, "on_config_updated") and self.on_config_updated:
            try:
                self.on_config_updated()
            except Exception:
                pass

    def save_general_config(self):
        self.config["college_name"] = self.ent_college.get().strip()
        self.config["college_address"] = self.ent_college_address.get().strip()
        self.config["college_tax_id"] = self.ent_tax_id.get().strip()
        self.config["bill_template_style"] = normalize_template_style(self.sys_style_var.get())
        self.config["default_export_folder"] = self.ent_export_folder.get().strip()
        self.config["fiscal_year"] = self.ent_fiscal_year.get().strip()
        self.config["current_semester"] = self.term_var.get().strip()
        self.config["cashier_position"] = self.pos_var.get().strip()
        self.config["receipt_book_no"] = self.ent_book.get().strip()
        self.config["receipt_prefix"] = self.ent_prefix.get().strip()
        self.config["receipt_no_start"] = self.ent_start_no.get().strip()
        self.config["receipt_no_end"] = self.ent_end_no.get().strip()
        self.config["auto_backup_on_exit"] = bool(self.auto_backup_var.get())
        self.config["show_course_registration"] = bool(self.show_registration_var.get())

        try:
            save_config(self.config)
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการบันทึก", f"ไม่สามารถบันทึกการตั้งค่าลงไฟล์ได้:\n{e}")
            return
        if hasattr(self, "on_config_updated") and self.on_config_updated:
            try:
                self.on_config_updated()
            except Exception:
                pass
        messagebox.showinfo("บันทึกสำเร็จ", "บันทึกการตั้งค่าข้อมูลสถานศึกษาและใบเสร็จเรียบร้อยแล้ว!")

    def reset_general_config_defaults(self):
        if not messagebox.askyesno("ยืนยันคืนค่าเริ่มต้น", "ต้องการคืนค่าการตั้งค่าทั่วไปเป็นค่าเริ่มต้นใช่หรือไม่?"):
            return
        self.ent_college.delete(0, "end")
        self.ent_college.insert(0, "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")
        self.ent_college_address.delete(0, "end")
        self.ent_college_address.insert(0, "455 ถ.พระพันวษา ต.ท่าพี่เลี้ยง อ.เมืองสุพรรณบุรี จ.สุพรรณบุรี 72000")
        self.ent_tax_id.delete(0, "end")
        self.ent_tax_id.insert(0, "0994000159721")
        self.sys_style_var.set(TEMPLATE_STYLE_OPTIONS[0])
        self.ent_fiscal_year.delete(0, "end")
        self.ent_fiscal_year.insert(0, "2569")
        self.term_var.set("1/2569")
        self.pos_var.set("ครู")
        self.ent_book.delete(0, "end")
        self.ent_book.insert(0, "7006")
        self.ent_prefix.delete(0, "end")
        self.ent_start_no.delete(0, "end")
        self.ent_start_no.insert(0, "700605433")
        self.ent_end_no.delete(0, "end")
        self.ent_end_no.insert(0, "700999999")
        self.auto_backup_var.set(True)
        self.show_registration_var.set(False)
        self.save_general_config()

    def save_all_settings_quick(self):
        self.save_general_config()

    # ==========================================================================
    # TAB 2: รายการค่าธรรมเนียมตามหลักสูตร
    # ==========================================================================
    def _build_tab_fees(self):
        curricula = self.db.get_options("CURRICULUM") or ["หลักสูตรในระบบ-ปวช"]

        # Top Bar: ตัวกรองหลักสูตร + ค้นหา
        top_bar = ctk.CTkFrame(self.tab_fees, fg_color="transparent")
        top_bar.pack(side="top", fill="x", padx=12, pady=(8, 8))

        ctk.CTkLabel(
            top_bar,
            text="กรองตามหลักสูตร:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))

        self.fee_filter_curr_var = ctk.StringVar(value="ทุกหลักสูตร")
        self.opt_fee_filter_curr = make_dropdown(
            top_bar,
            values=["ทุกหลักสูตร"] + curricula,
            variable=self.fee_filter_curr_var,
            width=210,
            height=36,
            command=lambda _v: self.refresh_fee_table(),
        )
        self.opt_fee_filter_curr.pack(side="left", padx=(0, 14))

        self.ent_fee_search = make_entry(
            top_bar,
            width=260,
            height=36,
            bold=False,
            placeholder_text="🔍 ค้นหารหัส หรือ ชื่อรายการค่าธรรมเนียม...",
        )
        self.ent_fee_search.pack(side="left", fill="x", expand=True, padx=(0, 14))
        self.ent_fee_search.bind("<KeyRelease>", self._schedule_fee_search)

        self.lbl_fee_count = ctk.CTkLabel(
            top_bar,
            text="พบ 0 รายการ",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        )
        self.lbl_fee_count.pack(side="right")

        # Center Table
        cols = [
            ("fee_code", "รหัสรายการ", 110, "center"),
            ("curriculum_name", "หลักสูตร", 200, "w"),
            ("fee_name", "ชื่อรายการค่าธรรมเนียม", 390, "w"),
            ("default_credits", "หน่วยกิต", 90, "center"),
            ("default_price", "ราคาต่อหน่วย (บาท)", 150, "e"),
            ("calc_type", "ประเภทคิดเงิน", 120, "center"),
        ]
        self.fee_table = DataTable(
            self.tab_fees,
            columns=cols,
            height=6,
            on_select=self.on_select_fee_row,
            show_inspector=False,
        )
        self.fee_table.pack(side="top", fill="both", expand=True, padx=12, pady=(0, 8))

        # Bottom Card: Form Editor
        edit_card = ctk.CTkFrame(
            self.tab_fees,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        edit_card.pack(side="bottom", fill="x", padx=12, pady=(0, 10))

        # Row 1
        r1 = ctk.CTkFrame(edit_card, fg_color="transparent")
        r1.pack(fill="x", padx=14, pady=(10, 4))

        ctk.CTkLabel(
            r1, text="รหัสรายการ:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY
        ).pack(side="left", padx=(0, 4))
        self.ent_fee_code = make_entry(r1, width=110, height=36)
        self.ent_fee_code.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            r1, text="ชื่อรายการค่าธรรมเนียม:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY
        ).pack(side="left", padx=(0, 4))
        self.ent_fee_name = make_entry(r1, height=36)
        self.ent_fee_name.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkLabel(
            r1, text="หลักสูตร:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY
        ).pack(side="left", padx=(0, 4))
        self.fee_form_curr_var = ctk.StringVar(value=curricula[0])
        self.opt_fee_form_curr = make_dropdown(
            r1,
            values=curricula,
            variable=self.fee_form_curr_var,
            width=210,
            height=36,
        )
        self.opt_fee_form_curr.pack(side="left")

        # Row 2
        r2 = ctk.CTkFrame(edit_card, fg_color="transparent")
        r2.pack(fill="x", padx=14, pady=(4, 10))

        ctk.CTkLabel(
            r2, text="ราคา (บาท):", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY
        ).pack(side="left", padx=(0, 4))
        self.ent_fee_price = make_entry(r2, width=105, height=36)
        self.ent_fee_price.pack(side="left", padx=(0, 10))
        self.ent_fee_price.insert(0, "0.00")

        ctk.CTkLabel(
            r2, text="หน่วยกิต:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY
        ).pack(side="left", padx=(0, 4))
        self.ent_fee_credits = make_entry(r2, width=65, height=36)
        self.ent_fee_credits.pack(side="left", padx=(0, 10))
        self.ent_fee_credits.insert(0, "0")

        ctk.CTkLabel(
            r2, text="ประเภท:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY
        ).pack(side="left", padx=(0, 4))
        self.calc_type_var = ctk.StringVar(value="FIXED - เหมาจ่าย")
        self.opt_calc_type = make_dropdown(
            r2,
            values=["FIXED - เหมาจ่าย", "PER_CREDIT - คิดตามหน่วยกิต"],
            variable=self.calc_type_var,
            width=185,
            height=36,
        )
        self.opt_calc_type.pack(side="left", padx=(0, 14))

        ctk.CTkButton(
            r2,
            text="💾 บันทึกรายการ (เพิ่มใหม่ / แก้ไข)",
            width=215,
            height=36,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.save_fee_item,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            r2,
            text="➕ ล้างฟอร์มสร้างใหม่",
            width=140,
            height=36,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.clear_fee_form,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            r2,
            text="🗑 ลบรายการนี้",
            width=115,
            height=36,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_DANGER_BG,
            hover_color="#FEE2E2",
            border_width=1,
            border_color=AppTheme.COLOR_DANGER_BORDER,
            text_color=AppTheme.COLOR_DANGER,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.delete_fee_item,
        ).pack(side="left")

        self.refresh_fee_table()

    def _schedule_fee_search(self, _event=None):
        if self._fee_search_job is not None:
            try:
                self.after_cancel(self._fee_search_job)
            except Exception:
                pass
        self._fee_search_job = self.after(250, self.refresh_fee_table)

    def refresh_fee_table(self):
        self._fee_search_job = None
        curr = self.fee_filter_curr_var.get()
        flt = curr if curr != "ทุกหลักสูตร" else "ALL"
        search_txt = self.ent_fee_search.get().strip().lower() if hasattr(self, "ent_fee_search") else ""

        items = self.db.get_fee_catalog(curriculum_filter=flt)
        rows = []
        for it in items:
            name = it["fee_name"]
            code = it["fee_code"]
            if search_txt and (search_txt not in name.lower() and search_txt not in code.lower()):
                continue
            ctype = "เหมาจ่าย" if it.get("calc_type") == "FIXED" else "ตามหน่วยกิต"
            creds = f"{float(it.get('default_credits', 0)):.0f}" if float(it.get('default_credits', 0)) > 0 else "-"
            rows.append(
                (
                    [
                        code,
                        it["curriculum_name"],
                        name,
                        creds,
                        f"฿{float(it['default_price']):,.2f}",
                        ctype,
                    ],
                    "PAID",
                )
            )
        self.fee_table.set_rows(rows)
        self.lbl_fee_count.configure(text=f"พบ {len(rows):,} รายการ")

    def on_select_fee_row(self, vals: List[str]):
        if len(vals) < 5:
            return
        self.ent_fee_code.delete(0, "end")
        self.ent_fee_code.insert(0, vals[0])
        self.fee_form_curr_var.set(vals[1])
        self.ent_fee_name.delete(0, "end")
        self.ent_fee_name.insert(0, vals[2])
        self.ent_fee_credits.delete(0, "end")
        self.ent_fee_credits.insert(0, vals[3] if vals[3] != "-" else "0")
        self.ent_fee_price.delete(0, "end")
        self.ent_fee_price.insert(0, vals[4].replace("฿", "").replace(",", ""))
        if "หน่วยกิต" in vals[5]:
            self.calc_type_var.set("PER_CREDIT - คิดตามหน่วยกิต")
        else:
            self.calc_type_var.set("FIXED - เหมาจ่าย")

    def clear_fee_form(self):
        self.ent_fee_code.delete(0, "end")
        self.ent_fee_name.delete(0, "end")
        self.ent_fee_credits.delete(0, "end")
        self.ent_fee_credits.insert(0, "0")
        self.ent_fee_price.delete(0, "end")
        self.ent_fee_price.insert(0, "0.00")
        self.calc_type_var.set("FIXED - เหมาจ่าย")

    def save_fee_item(self):
        code = self.ent_fee_code.get().strip()
        name = self.ent_fee_name.get().strip()
        curr = self.fee_form_curr_var.get()
        if not name:
            highlight_missing_field(self.ent_fee_name, "กรุณาระบุชื่อรายการค่าธรรมเนียม", show_toast_msg=True)
            return
        try:
            price = float(self.ent_fee_price.get().strip().replace(",", "") or "0")
            credits = float(self.ent_fee_credits.get().strip().replace(",", "") or "0")
        except ValueError:
            highlight_missing_field(self.ent_fee_price, "กรุณากรอกจำนวนเงินเป็นตัวเลข", show_toast_msg=True)
            return

        calc_type = "PER_CREDIT" if "PER_CREDIT" in self.calc_type_var.get() else "FIXED"

        try:
            self.db.upsert_fee_item(
                fee_code=code,
                fee_name=name,
                default_price=price,
                curriculum_name=curr,
                default_credits=credits,
                calc_type=calc_type,
                actor=self.session.username,
                role=self.session.role,
            )
            self.refresh_fee_table()
            messagebox.showinfo("บันทึกสำเร็จ", f"บันทึกรายการ '{name}' ({price:,.2f} บาท) เรียบร้อยแล้ว!", parent=self.winfo_toplevel())
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการบันทึก", f"ไม่สามารถบันทึกรายการค่าธรรมเนียมได้:\n{e}", parent=self.winfo_toplevel())

    def delete_fee_item(self):
        code = self.ent_fee_code.get().strip()
        if not code:
            messagebox.showwarning("เลือกรายการ", "กรุณาคลิกเลือกรายการในตารางก่อนลบ", parent=self.winfo_toplevel())
            return
        if not messagebox.askyesno("ยืนยันการลบ", f"ต้องการลบรายการค่าธรรมเนียมรหัส '{code}' ใช่หรือไม่?", parent=self.winfo_toplevel()):
            return
        try:
            self.db.delete_fee_item(code, actor=self.session.username, role=self.session.role)
            self.clear_fee_form()
            self.refresh_fee_table()
            messagebox.showinfo("ลบสำเร็จ", f"ลบรายการรหัส {code} เรียบร้อยแล้ว", parent=self.winfo_toplevel())
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการลบ", f"ไม่สามารถลบรายการค่าธรรมเนียมได้:\n{e}", parent=self.winfo_toplevel())

    # ==========================================================================
    # TAB 3: จัดการตัวเลือก Dropdown
    # ==========================================================================
    def _build_tab_dropdowns(self):
        self.category_labels = {
            "หลักสูตร (CURRICULUM)": "CURRICULUM",
            "คณะ / ประเภทวิชา (FACULTY)": "FACULTY",
            "แผนกวิชา / สาขาวิชา (DEPARTMENT)": "DEPARTMENT",
            "ชั้นปี / ห้องเรียน (CLASS_LEVEL)": "CLASS_LEVEL",
            "ภาค / ปีการศึกษา (TERM)": "TERM",
            "ตำแหน่งผู้รับเงิน (CASHIER_POSITION)": "CASHIER_POSITION",
        }

        # Top Bar: หมวดหมู่
        top_bar = ctk.CTkFrame(self.tab_dropdowns, fg_color="transparent")
        top_bar.pack(side="top", fill="x", padx=12, pady=(8, 8))

        ctk.CTkLabel(
            top_bar,
            text="เลือกหมวดหมู่ Dropdown:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 8))

        cat_keys = list(self.category_labels.keys())
        self.opt_cat_var = ctk.StringVar(value=cat_keys[0])
        make_dropdown(
            top_bar,
            values=cat_keys,
            variable=self.opt_cat_var,
            width=290,
            height=36,
            command=lambda _v: self.refresh_dropdown_table(),
        ).pack(side="left", padx=(0, 14))

        self.lbl_dd_count = ctk.CTkLabel(
            top_bar,
            text="พบ 0 ตัวเลือก",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        )
        self.lbl_dd_count.pack(side="right")

        # Center Table
        cols = [
            ("idx", "ลำดับที่", 80, "center"),
            ("cat", "หมวดหมู่ Dropdown", 250, "w"),
            ("val", "ค่าตัวเลือกที่แสดงในระบบ", 560, "w"),
        ]
        self.dd_table = DataTable(
            self.tab_dropdowns,
            columns=cols,
            height=6,
            on_select=self.on_select_dd_row,
            show_inspector=False,
        )
        self.dd_table.pack(side="top", fill="both", expand=True, padx=12, pady=(0, 8))

        # Bottom Card: Form Editor
        edit_card = ctk.CTkFrame(
            self.tab_dropdowns,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        edit_card.pack(side="bottom", fill="x", padx=12, pady=(0, 10))

        bd_in = ctk.CTkFrame(edit_card, fg_color="transparent")
        bd_in.pack(fill="x", padx=14, pady=10)

        ctk.CTkLabel(
            bd_in,
            text="ชื่อตัวเลือก:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 8))

        self.ent_new_opt = make_entry(
            bd_in,
            height=38,
            bold=False,
            placeholder_text="พิมพ์ชื่อตัวเลือกที่ต้องการเพิ่ม เช่น ปวช.1/5 หรือ คอมพิวเตอร์กราฟิก...",
        )
        self.ent_new_opt.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkButton(
            bd_in,
            text="➕ เพิ่มตัวเลือกใหม่",
            width=150,
            height=38,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.add_dropdown_option,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            bd_in,
            text="✏️ อัปเดตตัวเลือกนี้",
            width=140,
            height=38,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.update_dropdown_option,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            bd_in,
            text="🗑 ลบตัวเลือกนี้",
            width=120,
            height=38,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_DANGER_BG,
            hover_color="#FEE2E2",
            border_width=1,
            border_color=AppTheme.COLOR_DANGER_BORDER,
            text_color=AppTheme.COLOR_DANGER,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.delete_dropdown_option,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            bd_in,
            text="🧹 ล้าง",
            width=70,
            height=38,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.clear_dd_form,
        ).pack(side="left")

        self.refresh_dropdown_table()

    def refresh_dropdown_table(self):
        lbl = self.opt_cat_var.get()
        cat_code = self.category_labels[lbl]
        opts = self.db.get_options(cat_code)
        rows = []
        for idx, val in enumerate(opts, start=1):
            rows.append(([str(idx), lbl, val], "PAID"))
        self.dd_table.set_rows(rows)
        self.lbl_dd_count.configure(text=f"พบ {len(rows):,} ตัวเลือก")

    def on_select_dd_row(self, vals: List[str]):
        if len(vals) >= 3:
            self._selected_dropdown_val = vals[2]
            self.ent_new_opt.delete(0, "end")
            self.ent_new_opt.insert(0, vals[2])

    def clear_dd_form(self):
        self._selected_dropdown_val = None
        self.ent_new_opt.delete(0, "end")

    def add_dropdown_option(self):
        val = self.ent_new_opt.get().strip()
        if not val:
            messagebox.showwarning("ระบุข้อมูล", "กรุณาพิมพ์ชื่อตัวเลือกที่ต้องการเพิ่ม", parent=self.winfo_toplevel())
            return
        cat_code = self.category_labels[self.opt_cat_var.get()]
        try:
            self.db.add_option(cat_code, val, actor=self.session.username, role=self.session.role)
            self.clear_dd_form()
            self.refresh_dropdown_table()
            if cat_code == "CURRICULUM":
                currs = self.db.get_options("CURRICULUM")
                if hasattr(self, "opt_fee_filter_curr"):
                    self.opt_fee_filter_curr.configure(values=["ทุกหลักสูตร"] + currs)
                if hasattr(self, "opt_fee_form_curr"):
                    self.opt_fee_form_curr.configure(values=currs)
            messagebox.showinfo("เพิ่มสำเร็จ", f"เพิ่ม '{val}' ลงใน Dropdown เรียบร้อยแล้ว!", parent=self.winfo_toplevel())
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการเพิ่มตัวเลือก", f"ไม่สามารถเพิ่มตัวเลือกได้:\n{e}", parent=self.winfo_toplevel())

    def update_dropdown_option(self):
        old_val = self._selected_dropdown_val
        new_val = self.ent_new_opt.get().strip()
        if not old_val:
            messagebox.showwarning("เลือกแถว", "กรุณาคลิกเลือกตัวเลือกในตารางที่ต้องการแก้ไขก่อน", parent=self.winfo_toplevel())
            return
        if not new_val:
            messagebox.showwarning("ระบุข้อมูล", "กรุณาระบุชื่อตัวเลือกใหม่", parent=self.winfo_toplevel())
            return
        cat_code = self.category_labels[self.opt_cat_var.get()]
        try:
            self.db.update_option(cat_code, old_val, new_val, actor=self.session.username, role=self.session.role)
            self.clear_dd_form()
            self.refresh_dropdown_table()
            if cat_code == "CURRICULUM":
                currs = self.db.get_options("CURRICULUM")
                if hasattr(self, "opt_fee_filter_curr"):
                    self.opt_fee_filter_curr.configure(values=["ทุกหลักสูตร"] + currs)
                if hasattr(self, "opt_fee_form_curr"):
                    self.opt_fee_form_curr.configure(values=currs)
            messagebox.showinfo("อัปเดตสำเร็จ", f"อัปเดตตัวเลือกเป็น '{new_val}' เรียบร้อยแล้ว!", parent=self.winfo_toplevel())
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการแก้ไขตัวเลือก", f"ไม่สามารถแก้ไขตัวเลือกได้:\n{e}", parent=self.winfo_toplevel())

    def delete_dropdown_option(self):
        val = self.ent_new_opt.get().strip() or self._selected_dropdown_val
        if not val:
            messagebox.showwarning("เลือกรายการ", "กรุณาคลิกเลือกแถวในตารางก่อนลบ", parent=self.winfo_toplevel())
            return
        cat_code = self.category_labels[self.opt_cat_var.get()]
        if not messagebox.askyesno("ยืนยันการลบ", f"ต้องการลบตัวเลือก '{val}' ออกจาก Dropdown ใช่หรือไม่?", parent=self.winfo_toplevel()):
            return
        try:
            self.db.delete_option(cat_code, val, actor=self.session.username, role=self.session.role)
            self.clear_dd_form()
            self.refresh_dropdown_table()
            messagebox.showinfo("ลบสำเร็จ", f"ลบตัวเลือก '{val}' เรียบร้อยแล้ว", parent=self.winfo_toplevel())
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการลบตัวเลือก", f"ไม่สามารถลบตัวเลือกได้:\n{e}", parent=self.winfo_toplevel())

    # ==========================================================================
    # TAB 4: สำรองข้อมูล & ความปลอดภัยระบบ
    # ==========================================================================
    def _build_tab_backup(self):
        summary = self.db.get_dashboard_summary()

        # Row 1: Stat Cards Summary
        stats_frame = ctk.CTkFrame(self.tab_backup, fg_color="transparent")
        stats_frame.pack(fill="x", padx=14, pady=(10, 8))
        stats_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.card_students = StatCard(
            stats_frame,
            title="👥 นักศึกษาในระบบ",
            value=f"{summary['student_count']:,} คน",
            accent_color=AppTheme.COLOR_PRIMARY_BLUE,
        )
        self.card_students.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        self.card_receipts = StatCard(
            stats_frame,
            title="🧾 ใบเสร็จรับเงินที่ออกแล้ว",
            value=f"{summary['receipt_count']:,} ฉบับ",
            accent_color=AppTheme.COLOR_SUCCESS,
        )
        self.card_receipts.grid(row=0, column=1, sticky="nsew", padx=6)

        self.card_income = StatCard(
            stats_frame,
            title="💰 ยอดเงินรับชำระสะสม",
            value=f"฿{summary['total_income']:,.2f}",
            accent_color=AppTheme.COLOR_ACCENT,
        )
        self.card_income.grid(row=0, column=2, sticky="nsew", padx=6)

        self.card_db_status = StatCard(
            stats_frame,
            title="🔒 ความสมบูรณ์ของฐานข้อมูล",
            value="SHA-256 Verified",
            accent_color=AppTheme.COLOR_PRIMARY_BLUE,
        )
        self.card_db_status.grid(row=0, column=3, sticky="nsew", padx=(6, 0))

        # Row 2: 2 Column Layout (Left: Database Backup / Right: Webhook Integration)
        body_grid = ctk.CTkFrame(self.tab_backup, fg_color="transparent")
        body_grid.pack(fill="both", expand=True, padx=14, pady=(0, 10))
        body_grid.grid_columnconfigure((0, 1), weight=1)
        body_grid.grid_rowconfigure(0, weight=1)

        # Left Card: Backup & Restore
        c_left = ctk.CTkFrame(
            body_grid,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        c_left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        ctk.CTkLabel(
            c_left,
            text="🛡️ การสำรองและกู้คืนฐานข้อมูล (Database)",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        ).pack(anchor="w", padx=16, pady=(14, 2))

        ctk.CTkLabel(
            c_left,
            text="ฐานข้อมูลระบบทำงานบน SQLite WAL ปลอดภัยต่อธุรกรรมการเงินและรองรับ Hot Backup",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color=AppTheme.COLOR_TEXT_MUTED,
        ).pack(anchor="w", padx=16, pady=(0, 12))

        ctk.CTkButton(
            c_left,
            text="📦 สำรองฐานข้อมูลด่วนทันที (.db)",
            height=38,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.run_quick_db_backup,
        ).pack(fill="x", padx=16, pady=4)

        ctk.CTkButton(
            c_left,
            text="🗂️ สำรองข้อมูลครบวงจร (.zip Full Backup + สลิปหลักฐาน)",
            height=38,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.run_zip_backup,
        ).pack(fill="x", padx=16, pady=4)

        ctk.CTkButton(
            c_left,
            text="♻️ กู้คืนฐานข้อมูลจากไฟล์สำรอง (.db)",
            height=38,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.run_restore_db,
        ).pack(fill="x", padx=16, pady=4)

        ctk.CTkButton(
            c_left,
            text="🔍 ตรวจสอบความสมบูรณ์ของลูกโซ่ประวัติ (Verify Audit Chain)",
            height=38,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.test_audit_integrity,
        ).pack(fill="x", padx=16, pady=(4, 14))

        # Right Card: Webhook Integration
        c_right = ctk.CTkFrame(
            body_grid,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        c_right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        ctk.CTkLabel(
            c_right,
            text="🔔 ระบบแจ้งเตือนธุรกรรมอัตโนมัติ (Webhook)",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        ).pack(anchor="w", padx=16, pady=(14, 2))

        ctk.CTkLabel(
            c_right,
            text="ส่งข้อมูลการออกใบเสร็จแบบ Real-time ไปยัง n8n / LINE Bot / Discord",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color=AppTheme.COLOR_TEXT_MUTED,
        ).pack(anchor="w", padx=16, pady=(0, 10))

        self.webhook_enabled_var = ctk.BooleanVar(value=bool(self.config.get("webhook_enabled", False)))
        self.chk_webhook_enabled = ctk.CTkCheckBox(
            c_right,
            text="เปิดใช้งานระบบส่งสัญญาณ Webhook อัตโนมัติ",
            variable=self.webhook_enabled_var,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
        )
        self.chk_webhook_enabled.pack(anchor="w", padx=16, pady=(0, 8))

        wh_row = ctk.CTkFrame(c_right, fg_color="transparent")
        wh_row.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(
            wh_row,
            text="Webhook URL:",
            width=135,
            anchor="w",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left")
        self.ent_webhook_url = make_entry(wh_row, height=36, bold=False, placeholder_text="https://webhook.site/...")
        self.ent_webhook_url.pack(side="left", fill="x", expand=True)
        self.ent_webhook_url.insert(0, str(self.config.get("webhook_url", "")))

        thresh_row = ctk.CTkFrame(c_right, fg_color="transparent")
        thresh_row.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(
            thresh_row,
            text="ยอดแจ้งเตือนขั้นต่ำ (บาท):",
            width=135,
            anchor="w",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left")
        self.ent_webhook_thresh = make_entry(thresh_row, height=36)
        self.ent_webhook_thresh.pack(side="left", fill="x", expand=True)
        self.ent_webhook_thresh.insert(0, str(self.config.get("webhook_alert_threshold", 50000.0)))

        wb_btn_row = ctk.CTkFrame(c_right, fg_color="transparent")
        wb_btn_row.pack(fill="x", padx=16, pady=(12, 14))

        ctk.CTkButton(
            wb_btn_row,
            text="💾 บันทึกการตั้งค่า Webhook",
            height=38,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.save_webhook_config,
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))

        ctk.CTkButton(
            wb_btn_row,
            text="🔔 ทดสอบส่งสัญญาณ",
            height=38,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.test_webhook_connection,
        ).pack(side="left", padx=(6, 0))

    def run_quick_db_backup(self):
        dest = backup_database(self.db.db_path)
        if dest:
            messagebox.showinfo("สำรองข้อมูลสำเร็จ", f"สำรองไฟล์ฐานข้อมูล (.db) เรียบร้อยแล้วที่:\n{dest}")
        else:
            messagebox.showerror("ผิดพลาด", "ไม่สามารถสำรองฐานข้อมูลได้")

    def run_zip_backup(self):
        out_folder = filedialog.askdirectory(parent=self.winfo_toplevel(), title="เลือกโฟลเดอร์สำหรับเก็บไฟล์สำรองข้อมูล (.zip)")
        if not out_folder:
            messagebox.showinfo("ยกเลิกการสำรองข้อมูล", "คุณได้ยกเลิกการเลือกโฟลเดอร์สำหรับเก็บไฟล์สำรองข้อมูล", parent=self.winfo_toplevel())
            return
        try:
            zip_path = export_full_zip_backup(
                self.db.db_path,
                self.config.get("evidence_folder", ""),
                out_folder,
            )
            messagebox.showinfo("สำรองข้อมูลสำเร็จ", f"บันทึกไฟล์สำรองข้อมูล (.zip) เรียบร้อยที่:\n{zip_path}", parent=self.winfo_toplevel())
        except Exception as e:
            messagebox.showerror("สำรองข้อมูลไม่สำเร็จ", f"เกิดข้อผิดพลาดในการสร้างไฟล์สำรอง Zip:\n{e}", parent=self.winfo_toplevel())

    def run_restore_db(self):
        src = filedialog.askopenfilename(
            title="เลือกไฟล์ฐานข้อมูลสำรอง (.db) เพื่อกู้คืน",
            filetypes=[("SQLite Database", "*.db")],
        )
        if not src:
            return
        if not messagebox.askyesno(
            "ยืนยันการกู้คืนฐานข้อมูล",
            "ระบบจะตรวจสอบความสมบูรณ์ของไฟล์สำรองและสำรองข้อมูลปัจจุบันไว้ก่อนกู้คืน\nต้องการดำเนินการต่อใช่หรือไม่?",
        ):
            return
        try:
            restore_database(src, self.db.db_path)
            if hasattr(self.db, "invalidate_cache"):
                self.db.invalidate_cache()
            messagebox.showinfo("กู้คืนสำเร็จ", "กู้คืนฐานข้อมูลเรียบร้อยแล้ว!")
        except Exception as e:
            messagebox.showerror("กู้คืนไม่สำเร็จ", str(e))

    def test_audit_integrity(self):
        ok, count, msg = self.db.verify_audit_integrity()
        if ok:
            messagebox.showinfo(
                "การตรวจสอบความสมบูรณ์สำเร็จ",
                f"สถานะ: {msg}\nจำนวนรายการ Audit Logs: {count:,} รายการ\nลูกโซ่ SHA-256 Hash Chain สมบูรณ์ 100% ไม่พบการปลอมแปลง",
            )
        else:
            messagebox.showerror("ตรวจพบความผิดปกติ", f"ตรวจพบปัญหาใน Audit Chain:\n{msg}")

    def save_webhook_config(self):
        self.config["webhook_enabled"] = bool(self.webhook_enabled_var.get())
        self.config["webhook_url"] = self.ent_webhook_url.get().strip()
        try:
            self.config["webhook_alert_threshold"] = float(self.ent_webhook_thresh.get().strip().replace(",", "") or "50000")
        except ValueError:
            self.config["webhook_alert_threshold"] = 50000.0

        save_config(self.config)
        self.webhook.webhook_url = self.config["webhook_url"]
        self.webhook.enabled = self.config["webhook_enabled"]
        self.webhook.alert_threshold = self.config["webhook_alert_threshold"]
        messagebox.showinfo("บันทึกสำเร็จ", "บันทึกการตั้งค่า Webhook เรียบร้อยแล้ว!")

    def test_webhook_connection(self):
        url = self.ent_webhook_url.get().strip()
        if not url:
            messagebox.showwarning("ระบุ URL", "กรุณาระบุ Webhook URL ก่อนทดสอบ")
            return
        dispatcher = FinanceWebhookDispatcher(webhook_url=url, enabled=True, alert_threshold=0.0)

        def cb(success: bool, detail: str):
            if success:
                messagebox.showinfo("Webhook สำเร็จ", f"ส่งสัญญาณ Webhook สำเร็จ!\nตอบกลับ: {detail}")
            else:
                messagebox.showerror("ส่ง Webhook ไม่สำเร็จ", f"ไม่สามารถส่งสัญญาณได้:\n{detail}")

        dispatcher.notify_event_async(
            event_type="TEST_PING",
            title="ทดสอบส่งสัญญาณ Webhook จากระบบ อศ.สุพรรณบุรี",
            amount=950.0,
            actor=self.session.username,
            ref_code="PING-001",
            callback=cb,
        )



# ==============================================================================
# VIEW 5: COURSE REGISTRATION CARD VIEW (บัตรลงทะเบียนรายวิชา อศท.28 ตาม EXofBill)
# ==============================================================================
class CourseRegistrationView(ctk.CTkFrame):
    """
    Course Registration & Official 'บัตรลงทะเบียนรายวิชา (อศท.28)' Generator matching EXofBill:
    - Select Student, Room, Academic Year, Major, Subject Type, Career Group
    - Add / Edit / Remove up to 15 registered courses (รหัสวิชา, ชื่อวิชา, ท., ป., นก.)
    - Displays Total ท. | ป. | นก. and Tuition Fee Summary (1,130 บาท)
    - Exports Official 'บัตรลงทะเบียนรายวิชา (อศท.28)' (.pdf / .docx) with Krungthai Bank tear-off slip
    """

    def __init__(self, parent, db: DatabaseManager, session: UserSession, config: Dict[str, Any]):
        super().__init__(parent, fg_color="transparent")
        self.db = db
        self.session = session
        self.config = config

        self.STUDENT_DROPDOWN_LIMIT = 200
        self._stu_search_job: Optional[str] = None
        self.students: List[Dict[str, Any]] = []
        self.student_map: Dict[str, Dict[str, Any]] = {}
        self.course_catalog = self.db.get_course_catalog()
        self.course_map: Dict[str, Dict[str, Any]] = {}
        self.registered_courses: List[Dict[str, Any]] = []

        header = SectionHeader(
            self,
            "บัตรลงทะเบียนรายวิชา (อศท.28)",
            "จัดการรายวิชาลงทะเบียน (ท./ป./นก.) และพิมพ์บัตรลงทะเบียนรายวิชาพร้อมใบนำฝากธนาคารกรุงไทยตาม EXofBill",
        )
        header.pack(fill="x", pady=(0, 10))
        header.add_action_button(
            "โหลดวิชาตามหลักสูตรต้นฉบับ (11 วิชา)",
            self.load_default_ex_courses,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
        )

        # ------------------------------------------------------------------
        # BOTTOM BAR: Total Credits + Print อศท.28 Button
        # ------------------------------------------------------------------
        btm_bar = ctk.CTkFrame(
            self,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER_STRONG,
        )
        btm_bar.pack(side="bottom", fill="x", pady=(12, 0))

        btm_in = ctk.CTkFrame(btm_bar, fg_color="transparent")
        btm_in.pack(fill="x", padx=20, pady=12)

        self.lbl_totals = ctk.CTkLabel(
            btm_in,
            text="รวม: ท. 10 | ป. 20 | นก. 19  •  ค่าธรรมเนียม ฿1,130.00",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        )
        self.lbl_totals.pack(side="left")

        btn_box = ctk.CTkFrame(btm_in, fg_color="transparent")
        btn_box.pack(side="right")

        self.reg_style_var = ctk.StringVar(value=_style_label_from_code(self.config.get("bill_template_style", "CLASSIC")))
        make_dropdown(
            btn_box,
            values=TEMPLATE_STYLE_OPTIONS,
            variable=self.reg_style_var,
            width=200,
            height=AppTheme.CTRL_H_MD,
            command=self._on_change_reg_style,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_box,
            text="ลบวิชาที่เลือก",
            width=110,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.remove_selected_course,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_box,
            text="บันทึก & พิมพ์ อศท.28 (.pdf)",
            width=240,
            height=AppTheme.CTRL_H_LG,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.save_and_export_registration_card,
        ).pack(side="left")

        # ------------------------------------------------------------------
        # STEP 1: Student & Curriculum Metadata Card (จัดโครงสร้างเป็นระเบียบ เรียบร้อย สบายตา)
        # ------------------------------------------------------------------
        stu_card = ctk.CTkFrame(
            self,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        stu_card.pack(side="top", fill="x", pady=(0, 8))
        self.registration_form = stu_card

        # แถบค้นหาและเลือกนักเรียน
        r_search = ctk.CTkFrame(stu_card, fg_color="transparent")
        r_search.pack(fill="x", padx=16, pady=(10, 6))

        ctk.CTkLabel(
            r_search,
            text="🔍 ค้นหาด่วน:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left", padx=(0, 6))

        self.ent_stu_search = make_entry(
            r_search,
            width=180,
            height=38,
            bold=False,
            placeholder_text="พิมพ์รหัส หรือ ชื่อ...",
        )
        self.ent_stu_search.pack(side="left", padx=(0, 14))
        self.ent_stu_search.bind("<KeyRelease>", self._schedule_student_search)

        ctk.CTkLabel(
            r_search,
            text="เลือกนักศึกษา:",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        ).pack(side="left", padx=(0, 8))

        self.stu_var = ctk.StringVar(value="-")
        self.opt_stu = make_dropdown(
            r_search,
            values=["-"],
            variable=self.stu_var,
            height=38,
            font_size=14,
            command=self.on_select_student,
            on_search=self._search_students_dropdown,
        )
        self.opt_stu.pack(side="left", fill="x", expand=True)

        # ข้อมูลนักศึกษาและหลักสูตร (ไม่มีกรอบซ้อนกรอบ จัดวางตรงบนการ์ดอย่างโปร่งตา)
        info_panel = ctk.CTkFrame(
            stu_card,
            fg_color="transparent",
            border_width=0,
        )
        info_panel.pack(fill="x", padx=16, pady=(0, 12))

        # แถว 1: รหัสนักศึกษา | ชื่อ-นามสกุล | ระดับชั้น
        r1_grid = ctk.CTkFrame(info_panel, fg_color="transparent")
        r1_grid.pack(fill="x", padx=0, pady=(4, 4))
        r1_grid.grid_columnconfigure(1, weight=0)
        r1_grid.grid_columnconfigure(3, weight=1)
        r1_grid.grid_columnconfigure(5, weight=0)

        ctk.CTkLabel(r1_grid, text="รหัสนักศึกษา:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT).grid(row=0, column=0, sticky="w", padx=(0, 6), pady=2)
        self.ent_code = make_entry(r1_grid, width=130, height=36, font_size=14, bold=True)
        self.ent_code.grid(row=0, column=1, sticky="w", padx=(0, 16), pady=2)
        self.ent_code.bind("<Return>", self._lookup_student_by_code_entry)
        self.ent_code.bind("<FocusOut>", self._lookup_student_by_code_entry)

        ctk.CTkLabel(r1_grid, text="ชื่อ-นามสกุล:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT).grid(row=0, column=2, sticky="w", padx=(0, 6), pady=2)
        self.ent_name = make_entry(r1_grid, height=36, font_size=14, bold=True)
        self.ent_name.grid(row=0, column=3, sticky="ew", padx=(0, 16), pady=2)

        ctk.CTkLabel(r1_grid, text="ระดับชั้น/ห้อง:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT).grid(row=0, column=4, sticky="w", padx=(0, 6), pady=2)
        self.ent_cls = make_entry(r1_grid, width=110, height=36)
        self.ent_cls.grid(row=0, column=5, sticky="e", pady=2)

        # แถว 2: สาขาวิชา | ประเภทวิชา | กลุ่มอาชีพ | ปีการศึกษา
        r2_grid = ctk.CTkFrame(info_panel, fg_color="transparent")
        r2_grid.pack(fill="x", padx=0, pady=(2, 6))
        r2_grid.grid_columnconfigure(1, weight=1)
        r2_grid.grid_columnconfigure(3, weight=1)
        r2_grid.grid_columnconfigure(5, weight=1)
        r2_grid.grid_columnconfigure(7, weight=0)

        ctk.CTkLabel(r2_grid, text="สาขาวิชา:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT).grid(row=0, column=0, sticky="w", padx=(0, 6), pady=2)
        self.ent_major = make_entry(r2_grid, height=36)
        self.ent_major.grid(row=0, column=1, sticky="ew", padx=(0, 14), pady=2)

        ctk.CTkLabel(r2_grid, text="ประเภทวิชา:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT).grid(row=0, column=2, sticky="w", padx=(0, 6), pady=2)
        self.ent_subj_type = make_entry(r2_grid, height=36, bold=False)
        self.ent_subj_type.grid(row=0, column=3, sticky="ew", padx=(0, 14), pady=2)
        self.ent_subj_type.insert(0, "อุตสาหกรรมดิจิทัลและเทคโนโลยีสารสนเทศ")

        ctk.CTkLabel(r2_grid, text="กลุ่มอาชีพ:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT).grid(row=0, column=4, sticky="w", padx=(0, 6), pady=2)
        self.ent_career = make_entry(r2_grid, height=36, bold=False)
        self.ent_career.grid(row=0, column=5, sticky="ew", padx=(0, 14), pady=2)
        self.ent_career.insert(0, "ธุรกิจดิจิทัลและพาณิชย์อิเล็กทรอนิกส์")

        ctk.CTkLabel(r2_grid, text="ปีการศึกษา:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_TEXT).grid(row=0, column=6, sticky="w", padx=(0, 6), pady=2)
        self.ent_year = make_entry(r2_grid, width=75, height=36)
        self.ent_year.grid(row=0, column=7, sticky="e", pady=2)
        self.ent_year.insert(0, self.config.get("fiscal_year", "2569"))

        # ------------------------------------------------------------------
        # STEP 2: Add Course Row Bar
        # ------------------------------------------------------------------
        add_bar = ctk.CTkFrame(
            self,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        add_bar.pack(side="top", fill="x", pady=(0, 8))

        ab_in = ctk.CTkFrame(add_bar, fg_color="transparent")
        ab_in.pack(fill="x", padx=16, pady=8)

        course_labels = []
        for c in self.course_catalog:
            clbl = f"{c['course_code']} — {c['course_name']} ({c['theory_hours']}-{c['practice_hours']}-{c['credits']})"
            self.course_map[clbl] = c
            course_labels.append(clbl)
        if not course_labels:
            course_labels = ["-"]

        ctk.CTkLabel(ab_in, text="เลือกรายวิชา:", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_ACCENT).pack(side="left", padx=(0, 6))
        self.course_pick_var = ctk.StringVar(value=course_labels[0])
        make_dropdown(
            ab_in,
            values=course_labels,
            variable=self.course_pick_var,
            width=260,
            height=36,
            command=self.on_pick_catalog_course,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkLabel(ab_in, text="รหัสวิชา:", font=(AppTheme.FONT_FAMILY, 13), text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 4))
        self.ent_c_code = make_entry(ab_in, width=105, height=36)
        self.ent_c_code.pack(side="left", padx=(0, 8))

        ctk.CTkLabel(ab_in, text="ชื่อวิชา:", font=(AppTheme.FONT_FAMILY, 13), text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 4))
        self.ent_c_name = make_entry(ab_in, width=190, height=36)
        self.ent_c_name.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkLabel(ab_in, text="ท.:", font=(AppTheme.FONT_FAMILY, 13), text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 2))
        self.ent_c_t = make_entry(ab_in, width=42, height=36)
        self.ent_c_t.pack(side="left", padx=(0, 6))

        ctk.CTkLabel(ab_in, text="ป.:", font=(AppTheme.FONT_FAMILY, 13), text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 2))
        self.ent_c_p = make_entry(ab_in, width=42, height=36)
        self.ent_c_p.pack(side="left", padx=(0, 6))

        ctk.CTkLabel(ab_in, text="นก.:", font=(AppTheme.FONT_FAMILY, 13), text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 2))
        self.ent_c_cr = make_entry(ab_in, width=42, height=36)
        self.ent_c_cr.pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            ab_in,
            text="➕ เพิ่มรายวิชา",
            width=125,
            height=36,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.add_course_to_table,
        ).pack(side="left")

        cols = [
            ("line_no", "ลำดับที่", 70, "center"),
            ("course_code", "รหัสวิชา", 140, "center"),
            ("course_name", "ชื่อวิชา", 450, "w"),
            ("theory_hours", "ท. (ทฤษฎี)", 100, "center"),
            ("practice_hours", "ป. (ปฏิบัติ)", 100, "center"),
            ("credits", "นก. (หน่วยกิต)", 110, "center"),
        ]
        self.table = DataTable(self, columns=cols, height=7, show_inspector=False)
        self.table.pack(side="top", fill="both", expand=True)

        self.on_pick_catalog_course(course_labels[0])
        self.refresh_student_dropdown(prefer_code="69219100049")
        self._last_data_version = getattr(self.db, "data_version", -1)

    def on_show(self):
        """Reloads student dropdown only if database data has changed."""
        if getattr(self.db, "data_version", None) == getattr(self, "_last_data_version", -1):
            return
        self.refresh_student_dropdown()
        self._last_data_version = getattr(self.db, "data_version", -1)

    def _schedule_student_search(self, _event=None):
        if self._stu_search_job is not None:
            try:
                self.after_cancel(self._stu_search_job)
            except Exception:
                pass
        self._stu_search_job = self.after(250, self.refresh_student_dropdown)

    def _lookup_student_by_code_entry(self, _event=None):
        code = self.ent_code.get().strip()
        if not code:
            return
        stu = self.db.get_student_by_code(code)
        if stu:
            self._populate_student_fields(stu)

    def _search_students_dropdown(self, query: str) -> List[str]:
        q = (query or "").strip()
        stus = self.db.get_students(search=q, limit=40, offset=0)
        labels = []
        for s in stus:
            lbl = f"{s['student_code']} — {s['full_name']} ({s.get('class_group', '')})"
            self.student_map[lbl] = s
            labels.append(lbl)
        return labels or ["(ไม่พบรายชื่อนักศึกษาที่ค้นหา)"]

    def refresh_student_dropdown(self, prefer_code: str = ""):
        self._stu_search_job = None
        search_txt = self.ent_stu_search.get().strip() if hasattr(self, "ent_stu_search") else ""
        total_matching = self.db.count_students(search=search_txt)
        self.students = self.db.get_students(search=search_txt, limit=self.STUDENT_DROPDOWN_LIMIT, offset=0)

        if prefer_code and not any(s["student_code"] == prefer_code for s in self.students):
            pref_stu = self.db.get_student_by_code(prefer_code)
            if pref_stu:
                self.students.insert(0, pref_stu)

        self.student_map.clear()
        stu_labels = []
        default_lbl = ""
        for s in self.students:
            lbl = f"{s['student_code']} — {s['full_name']} ({s['class_group']})"
            self.student_map[lbl] = s
            stu_labels.append(lbl)
            if prefer_code and s["student_code"] == prefer_code:
                default_lbl = lbl

        if not stu_labels:
            stu_labels = ["ไม่พบรายชื่อนักศึกษาที่ค้นหา"]
            self.opt_stu.configure(values=stu_labels)
            self.stu_var.set(stu_labels[0])
            return

        if total_matching > len(self.students):
            stu_labels.append(f"--- แสดง {len(self.students):,} จาก {total_matching:,} คน (พิมพ์ค้นหาเพื่อกรองเพิ่ม) ---")

        if not default_lbl:
            default_lbl = stu_labels[0]
        self.opt_stu.configure(values=stu_labels)
        self.stu_var.set(default_lbl)
        self.on_select_student(default_lbl)

    def on_pick_catalog_course(self, label: str):
        c = self.course_map.get(label)
        if not c:
            return
        self.ent_c_code.delete(0, "end")
        self.ent_c_code.insert(0, c["course_code"])
        self.ent_c_name.delete(0, "end")
        self.ent_c_name.insert(0, c["course_name"])
        self.ent_c_t.delete(0, "end")
        self.ent_c_t.insert(0, str(c["theory_hours"]))
        self.ent_c_p.delete(0, "end")
        self.ent_c_p.insert(0, str(c["practice_hours"]))
        self.ent_c_cr.delete(0, "end")
        self.ent_c_cr.insert(0, str(c["credits"]))

    def _populate_student_fields(self, stu: Dict[str, Any]):
        self.ent_code.delete(0, "end")
        self.ent_code.insert(0, stu.get("student_code", ""))
        self.ent_name.delete(0, "end")
        self.ent_name.insert(0, stu.get("full_name", ""))
        cls_val = stu.get("class_group", "ปวช.1/2")
        if cls_val == "ปวช.1" and stu.get("student_code") == "69219100049":
            cls_val = "ปวช. 1/2"
        self.ent_cls.delete(0, "end")
        self.ent_cls.insert(0, cls_val)
        self.ent_major.delete(0, "end")
        self.ent_major.insert(0, stu.get("major", "เทคโนโลยีธุรกิจดิจิทัล"))

        self.registered_courses = self.db.get_student_registered_courses(stu["student_code"])
        self.refresh_table_ui()

    def on_select_student(self, label: str):
        stu = self.student_map.get(label)
        if not stu:
            return
        self._populate_student_fields(stu)

    def load_default_ex_courses(self):
        self.registered_courses = [dict(c) for c in self.db.get_course_catalog()]
        self.refresh_table_ui()

    def add_course_to_table(self):
        ccode = self.ent_c_code.get().strip()
        cname = self.ent_c_name.get().strip()
        if not ccode or not cname:
            messagebox.showwarning("ระบุข้อมูลวิชา", "กรุณากรอกรหัสวิชาและชื่อวิชา")
            return
        try:
            th = int(float(self.ent_c_t.get().strip() or "0"))
            pr = int(float(self.ent_c_p.get().strip() or "0"))
            cr = int(float(self.ent_c_cr.get().strip() or "0"))
        except ValueError:
            messagebox.showwarning("ตัวเลขไม่ถูกต้อง", "กรุณากรอก ท., ป., นก. เป็นตัวเลข")
            return

        self.registered_courses.append(
            {
                "course_code": ccode,
                "course_name": cname,
                "theory_hours": th,
                "practice_hours": pr,
                "credits": cr,
            }
        )
        self.db.upsert_course_catalog(ccode, cname, th, pr, cr, self.ent_major.get().strip(), self.ent_cls.get().strip())
        self.refresh_table_ui()

    def remove_selected_course(self):
        vals = self.table.get_selected_values()
        if not vals:
            if self.registered_courses:
                self.registered_courses.pop()
                self.refresh_table_ui()
            return
        try:
            idx = int(vals[0]) - 1
            if 0 <= idx < len(self.registered_courses):
                self.registered_courses.pop(idx)
                self.refresh_table_ui()
        except Exception:
            pass

    def refresh_table_ui(self):
        rows = []
        tot_t, tot_p, tot_c = 0, 0, 0
        for idx, c in enumerate(self.registered_courses, start=1):
            t_v = int(float(c.get("theory_hours", 0)))
            p_v = int(float(c.get("practice_hours", 0)))
            c_v = int(float(c.get("credits", 0)))
            tot_t += t_v
            tot_p += p_v
            tot_c += c_v
            rows.append(
                (
                    [
                        str(idx),
                        str(c.get("course_code", "")),
                        str(c.get("course_name", "")),
                        str(t_v),
                        str(p_v),
                        str(c_v),
                    ],
                    "PAID",
                )
            )
        self.table.set_rows(rows)
        self.lbl_totals.configure(
            text=f"รวม ({len(self.registered_courses)} วิชา): ท. {tot_t} | ป. {tot_p} | นก. {tot_c}  •  ค่าธรรมเนียม ฿1,130.00"
        )

    def _on_change_reg_style(self, selected_label: str):
        self.config["bill_template_style"] = normalize_template_style(selected_label)
        save_config(self.config)

    def save_and_export_registration_card(self):
        stu_code = self.ent_code.get().strip()
        stu_name = self.ent_name.get().strip()
        if not stu_code or not stu_name:
            highlight_missing_field(self.ent_code, "กรุณาเลือกหรือกรอกรหัสนักศึกษา", show_toast_msg=True)
            return

        self.db.save_student_registered_courses(
            student_code=stu_code,
            term_code=self.config.get("current_semester", "2/2569"),
            courses_list=self.registered_courses,
            actor=self.session.username,
            role=self.session.role,
        )

        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)
        default_fn = f"CourseReg_OST28_{stu_code}.pdf"

        save_path = filedialog.asksaveasfilename(
            title="เลือกตำแหน่งบันทึกบัตรลงทะเบียนรายวิชา อศท.28 (.pdf / .docx)",
            initialdir=out_dir,
            initialfile=default_fn,
            defaultextension=".pdf",
            filetypes=[
                ("PDF Document (*.pdf)", "*.pdf"),
                ("Word Document (*.docx)", "*.docx"),
            ],
        )
        if not save_path:
            messagebox.showinfo("ยกเลิกการออกบัตรลงทะเบียน", "คุณได้ยกเลิกการเลือกตำแหน่งบันทึกไฟล์", parent=self.winfo_toplevel())
            return

        chosen_style = normalize_template_style(self.reg_style_var.get())
        self.config["bill_template_style"] = chosen_style
        save_config(self.config)

        payload = {
            "report_type": "COURSE_REGISTRATION",
            "template_style": chosen_style,
            "student_code": stu_code,
            "full_name": stu_name,
            "class_group": self.ent_cls.get().strip() or "ปวช. 1/2",
            "fiscal_year": self.ent_year.get().strip() or "2569",
            "major": self.ent_major.get().strip() or "เทคโนโลยีธุรกิจดิจิทัล",
            "subject_type": self.ent_subj_type.get().strip() or "อุตสาหกรรมดิจิทัลและเทคโนโลยีสารสนเทศ",
            "career_group": self.ent_career.get().strip() or "ธุรกิจดิจิทัลและพาณิชย์อิเล็กทรอนิกส์",
            "courses": self.registered_courses,
            "fee_items": [
                {"item_name": "ค่าใช้บริการอินเตอร์เน็ต", "line_total": 300.0},
                {"item_name": "ค่าจ้างครูหรือผู้ที่มีความเชี่ยวชาญในสาขาเฉพาะ", "line_total": 500.0},
                {"item_name": "ค่ารักษาสภาพแวดล้อม", "line_total": 330.0},
            ],
        }

        try:
            export_document_auto(
                save_path,
                payload,
                college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
                metadata=self.config,
                is_batch=False,
            )
            open_ok = open_document_file(save_path, parent=self.winfo_toplevel())
            messagebox.showinfo(
                "พิมพ์บัตรลงทะเบียนรายวิชาสำเร็จ",
                f"บันทึกไฟล์บัตรลงทะเบียนรายวิชา (อศท.28) ของ {stu_name} เรียบร้อยที่:\n{save_path}\n\n"
                + ("(ระบบได้เปิดไฟล์เอกสารขึ้นมาให้พร้อมพิมพ์เรียบร้อยแล้ว)" if open_ok else ""),
                parent=self.winfo_toplevel(),
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการออกบัตรลงทะเบียน", f"ไม่สามารถสร้างหรือพิมพ์บัตรลงทะเบียนได้:\n{e}", parent=self.winfo_toplevel())

