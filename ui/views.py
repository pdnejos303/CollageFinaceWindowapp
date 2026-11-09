import datetime
import os
from tkinter import filedialog, messagebox
from typing import Any, Callable, Dict, List, Optional
import customtkinter as ctk

from core.config import (
    backup_database,
    export_full_zip_backup,
    format_receipt_date_thai,
    format_thai_date_short,
    restore_database,
    save_config,
    thai_baht_text,
)
from core.database import DatabaseManager
from core.models import UserSession
from exporters.docx_exporter import DocxReceiptExporter
from exporters.excel_exporter import ExcelFinanceExporter
from exporters.pdf_exporter import PdfReceiptExporter
from exporters.webhook_exporter import FinanceWebhookDispatcher
from ui.components import DataTable, SectionHeader
from ui.theme import AppTheme


def export_document_auto(
    save_path: str,
    payload: Any,
    college_name: str,
    metadata: Dict[str, Any],
    is_batch: bool = False,
) -> str:
    """Exports receipt/report to PDF (.pdf) by default, or Word (.docx) if selected."""
    if save_path.lower().endswith(".docx"):
        exporter = DocxReceiptExporter(save_path, college_name=college_name)
    else:
        exporter = PdfReceiptExporter(save_path, college_name=college_name)
    if is_batch:
        return exporter.export_batch(payload, metadata=metadata)
    return exporter.export(payload, metadata=metadata)


def make_dropdown(
    parent,
    values: List[str],
    variable: ctk.StringVar,
    width: int = 160,
    height: int = 38,
    font_size: int = 13,
    bold: bool = True,
    command: Optional[Callable] = None,
) -> ctk.CTkOptionMenu:
    """Unified, clean corporate dropdown control that blends seamlessly on light cards."""
    weight = "bold" if bold else "normal"
    kwargs: Dict[str, Any] = {
        "values": values,
        "variable": variable,
        "width": width,
        "height": height,
        "corner_radius": 8,
        "fg_color": AppTheme.COLOR_DROPDOWN_FG,
        "button_color": AppTheme.COLOR_DROPDOWN_BTN,
        "button_hover_color": AppTheme.COLOR_DROPDOWN_HOVER,
        "text_color": AppTheme.COLOR_DROPDOWN_TEXT,
        "dropdown_fg_color": AppTheme.COLOR_DROPDOWN_MENU,
        "dropdown_hover_color": AppTheme.COLOR_SECONDARY_BG,
        "dropdown_text_color": AppTheme.COLOR_TEXT,
        "font": (AppTheme.FONT_FAMILY, font_size, weight),
        "dropdown_font": (AppTheme.FONT_FAMILY, 13),
    }
    if command is not None:
        kwargs["command"] = command
    return ctk.CTkOptionMenu(parent, **kwargs)


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
        "corner_radius": 8,
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
    Step-by-Step Minimal UI with Unified Corporate Navy & Slate Palette:
    - ขั้นตอนที่ 1: เลือกผู้ชำระเงิน (หรือสลับไปโหมด 'ออกใบเสร็จยกห้อง')
    - ขั้นตอนที่ 2: เลือกรายการค่าธรรมเนียม (หลักสูตร -> รายการ -> กดปุ่ม '+ เพิ่มลงใบเสร็จ')
    - ขั้นตอนที่ 3: ตรวจสอบตารางใบเสร็จ & กดปุ่มหลัก 'บันทึก & พิมพ์ใบเสร็จ (.docx)'
    """

    def __init__(
        self,
        parent,
        db: DatabaseManager,
        session: UserSession,
        config: Dict[str, Any],
        webhook: FinanceWebhookDispatcher,
    ):
        super().__init__(parent, fg_color="transparent")
        self.db = db
        self.session = session
        self.config = config
        self.webhook = webhook

        self.students: List[Dict[str, Any]] = []
        self.student_map: Dict[str, Dict[str, Any]] = {}
        self.cart_items: List[Dict[str, Any]] = []
        self.current_catalog: List[Dict[str, Any]] = []
        self.catalog_map: Dict[str, Dict[str, Any]] = {}

        # Load master options
        self.curricula = self.db.get_options("CURRICULUM") or ["หลักสูตรในระบบ-ปวช"]
        self.faculties = self.db.get_options("FACULTY") or ["บริหารธุรกิจ / พาณิชยกรรม"]
        self.departments = self.db.get_options("DEPARTMENT") or ["การบัญชี"]
        self.class_levels = ["ทั้งหมด"] + (self.db.get_options("CLASS_LEVEL") or ["ปวช.1/1"])
        self.terms = self.db.get_options("TERM") or ["1/2569"]

        # ------------------------------------------------------------------
        # TOP BAR: Title + Mode Switcher (ออกรายบุคคล vs ออกใบเสร็จยกห้อง)
        # ------------------------------------------------------------------
        top_header = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        top_header.pack(side="top", fill="x", pady=(0, 10))

        th_in = ctk.CTkFrame(top_header, fg_color="transparent")
        th_in.pack(fill="x", padx=18, pady=10)

        ctk.CTkLabel(
            th_in,
            text="ออกใบเสร็จรับเงินค่าเทอม",
            font=(AppTheme.FONT_FAMILY, 20, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left")

        ctk.CTkButton(
            th_in,
            text="🏫 โหมดออกใบเสร็จยกห้อง / นำเข้าไฟล์ Excel",
            height=38,
            corner_radius=8,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.open_batch_class_modal,
        ).pack(side="right")

        # ------------------------------------------------------------------
        # STEP 3 (PACKED BOTTOM FIRST): สรุปยอดเงิน & ปุ่มหลักโทนเดียวกัน
        # ------------------------------------------------------------------
        bottom_bar = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER_STRONG,
        )
        bottom_bar.pack(side="bottom", fill="x", pady=(10, 0))

        btm_inner = ctk.CTkFrame(bottom_bar, fg_color="transparent")
        btm_inner.pack(fill="x", padx=18, pady=12)

        # Left: Total & Thai Baht Text
        tot_box = ctk.CTkFrame(btm_inner, fg_color="transparent")
        tot_box.pack(side="left", fill="x", expand=True)

        self.lbl_net_total = ctk.CTkLabel(
            tot_box,
            text="ยอดรวมสุทธิ: ฿0.00 บาท",
            font=(AppTheme.FONT_FAMILY, 22, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
            anchor="w",
        )
        self.lbl_net_total.pack(anchor="w")

        self.lbl_baht_text = ctk.CTkLabel(
            tot_box,
            text="ตัวอักษร: (ศูนย์บาทถ้วน)",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_MUTED,
            anchor="w",
        )
        self.lbl_baht_text.pack(anchor="w")

        # Right: Subtle table tools + 1 Authoritative Primary CTA Button
        btn_box = ctk.CTkFrame(btm_inner, fg_color="transparent")
        btn_box.pack(side="right")

        self.pay_method_var = ctk.StringVar(value="เงินสด (Cash)")
        make_dropdown(
            btn_box,
            values=["เงินสด (Cash)", "สแกน QR PromptPay", "โอนเข้าบัญชีวิทยาลัย"],
            variable=self.pay_method_var,
            width=170,
            height=42,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_box,
            text="ลบแถวที่เลือก",
            width=110,
            height=42,
            corner_radius=8,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.remove_selected_item,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            btn_box,
            text="ล้างตาราง",
            width=95,
            height=42,
            corner_radius=8,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_MUTED,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.reset_pos_form,
        ).pack(side="left", padx=(0, 12))

        ctk.CTkButton(
            btn_box,
            text="🖨️ ขั้นตอนที่ 3: บันทึก & พิมพ์ใบเสร็จ (.pdf)",
            width=290,
            height=44,
            corner_radius=8,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            command=self.issue_receipt_and_export,
        ).pack(side="left")

        # ------------------------------------------------------------------
        # STEP 1: ขั้นตอนที่ 1 — เลือกนักศึกษาและข้อมูลหัวใบเสร็จ
        # ------------------------------------------------------------------
        step1_card = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        step1_card.pack(side="top", fill="x", pady=(0, 10))

        s1_title = ctk.CTkFrame(step1_card, fg_color="transparent")
        s1_title.pack(fill="x", padx=16, pady=(10, 4))

        ctk.CTkLabel(
            s1_title,
            text="ขั้นตอนที่ 1 : เลือกนักศึกษาที่จะออกใบเสร็จ",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        ).pack(side="left")

        # Receipt number & Date tucked cleanly on the right of Step 1 header
        ctk.CTkLabel(
            s1_title,
            text="ภาคเรียน:",
            font=(AppTheme.FONT_FAMILY, 13),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="right", padx=(10, 4))
        self.term_var = ctk.StringVar(value=self.config.get("current_semester", self.terms[0]))
        self.opt_term = make_dropdown(
            s1_title,
            values=self.terms,
            variable=self.term_var,
            width=95,
            height=32,
        )
        self.opt_term.pack(side="right")

        self.entry_date = make_entry(s1_title, width=105, height=32, bold=False)
        self.entry_date.pack(side="right", padx=(4, 10))
        self.entry_date.insert(0, format_thai_date_short())
        ctk.CTkLabel(
            s1_title,
            text="วันที่:",
            font=(AppTheme.FONT_FAMILY, 13),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="right")

        self.entry_rcp_no = make_entry(s1_title, width=120, height=32, bold=True)
        self.entry_rcp_no.pack(side="right", padx=(4, 10))
        ctk.CTkLabel(
            s1_title,
            text="เลขที่ใบเสร็จ:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="right")

        # Main Student Selector Row
        self.STUDENT_DROPDOWN_LIMIT = 200
        self._stu_search_job: Optional[str] = None

        r1 = ctk.CTkFrame(step1_card, fg_color="transparent")
        r1.pack(fill="x", padx=16, pady=(4, 6))

        ctk.CTkLabel(
            r1,
            text="กรองห้องเรียน:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))
        self.filter_class_var = ctk.StringVar(value="ทั้งหมด")
        self.opt_filter_class = make_dropdown(
            r1,
            values=self.class_levels,
            variable=self.filter_class_var,
            width=125,
            height=38,
            command=lambda _v: self.refresh_student_list(),
        )
        self.opt_filter_class.pack(side="left", padx=(0, 10))

        self.entry_stu_search = make_entry(
            r1,
            width=185,
            height=38,
            bold=False,
            placeholder_text="🔍 ค้นหารหัส/ชื่อ...",
        )
        self.entry_stu_search.pack(side="left", padx=(0, 10))
        self.entry_stu_search.bind("<KeyRelease>", self._schedule_student_search)

        ctk.CTkLabel(
            r1,
            text="รายชื่อนักศึกษา:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))
        self.stu_var = ctk.StringVar(value="")
        self.opt_student = make_dropdown(
            r1,
            values=["-"],
            variable=self.stu_var,
            height=38,
            font_size=14,
            command=self.on_select_student,
        )
        self.opt_student.pack(side="left", fill="x", expand=True)

        # Student Details Row (Auto-filled, editable if needed)
        r2 = ctk.CTkFrame(step1_card, fg_color="transparent")
        r2.pack(fill="x", padx=16, pady=(2, 12))

        ctk.CTkLabel(
            r2,
            text="ได้รับเงินจาก:",
            font=(AppTheme.FONT_FAMILY, 13),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.entry_stu_name = make_entry(r2, width=185, height=34)
        self.entry_stu_name.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            r2,
            text="รหัสประจำตัว:",
            font=(AppTheme.FONT_FAMILY, 13),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.entry_stu_code = make_entry(r2, width=110, height=34)
        self.entry_stu_code.pack(side="left", padx=(0, 10))
        self.entry_stu_code.bind("<Return>", self._lookup_student_by_code_entry)
        self.entry_stu_code.bind("<FocusOut>", self._lookup_student_by_code_entry)

        ctk.CTkLabel(
            r2,
            text="คณะ:",
            font=(AppTheme.FONT_FAMILY, 13),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.fac_var = ctk.StringVar(value=self.faculties[0])
        self.opt_faculty = make_dropdown(
            r2,
            values=self.faculties,
            variable=self.fac_var,
            width=150,
            height=34,
            bold=False,
        )
        self.opt_faculty.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            r2,
            text="แผนกวิชา:",
            font=(AppTheme.FONT_FAMILY, 13),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.dept_var = ctk.StringVar(value=self.departments[0])
        self.opt_dept = make_dropdown(
            r2,
            values=self.departments,
            variable=self.dept_var,
            width=165,
            height=34,
            bold=False,
        )
        self.opt_dept.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            r2,
            text="ชั้น:",
            font=(AppTheme.FONT_FAMILY, 13),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        cls_only = self.db.get_options("CLASS_LEVEL") or ["ปวช.1/1"]
        self.cls_var = ctk.StringVar(value=cls_only[0])
        self.opt_cls = make_dropdown(
            r2,
            values=cls_only,
            variable=self.cls_var,
            width=110,
            height=34,
            bold=False,
        )
        self.opt_cls.pack(side="left")

        # ------------------------------------------------------------------
        # STEP 2: ขั้นตอนที่ 2 — เลือกรายการค่าธรรมเนียมใส่ใบเสร็จ
        # ------------------------------------------------------------------
        step2_card = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        step2_card.pack(side="top", fill="x", pady=(0, 10))

        ib_inner = ctk.CTkFrame(step2_card, fg_color="transparent")
        ib_inner.pack(fill="x", padx=16, pady=10)

        ctk.CTkLabel(
            ib_inner,
            text="ขั้นตอนที่ 2 : เพิ่มรายการ",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        ).pack(side="left", padx=(0, 12))

        self.curriculum_var = ctk.StringVar(value=self.curricula[0])
        self.opt_curriculum = make_dropdown(
            ib_inner,
            values=self.curricula,
            variable=self.curriculum_var,
            width=185,
            height=38,
            command=self.on_change_curriculum,
        )
        self.opt_curriculum.pack(side="left", padx=(0, 8))

        self.fee_item_var = ctk.StringVar(value="")
        self.opt_fee_item = make_dropdown(
            ib_inner,
            values=["-"],
            variable=self.fee_item_var,
            width=260,
            height=38,
            command=self.on_select_fee_item,
        )
        self.opt_fee_item.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkLabel(
            ib_inner,
            text="จำนวน:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.entry_qty = make_entry(ib_inner, width=65, height=38, font_size=14)
        self.entry_qty.pack(side="left", padx=(0, 10))
        self.entry_qty.insert(0, "1")

        ctk.CTkLabel(
            ib_inner,
            text="ราคา:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.entry_unit_price = make_entry(ib_inner, width=105, height=38, font_size=14)
        self.entry_unit_price.pack(side="left", padx=(0, 10))
        self.entry_unit_price.insert(0, "0.00")

        ctk.CTkButton(
            ib_inner,
            text="➕ เพิ่มลงใบเสร็จ",
            width=155,
            height=38,
            corner_radius=8,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.add_selected_fee_to_cart,
        ).pack(side="left")

        # ------------------------------------------------------------------
        # RECEIPT ITEMS TABLE
        # ------------------------------------------------------------------
        cols = [
            ("line_no", "ลำดับ", 70, "center"),
            ("item_name", "รายการค่าธรรมเนียมในใบเสร็จ", 520, "w"),
            ("quantity", "จำนวน/หน่วยกิต", 140, "center"),
            ("unit_price", "หน่วยละ (บาท)", 160, "e"),
            ("line_total", "จำนวนเงินรวม (บาท)", 180, "e"),
        ]
        self.table = DataTable(self, columns=cols, height=6, show_inspector=False)
        self.table.pack(side="top", fill="both", expand=True)

        self.refresh_next_receipt_no()
        self.refresh_student_list()
        self.on_change_curriculum(self.curricula[0])

    def refresh_next_receipt_no(self):
        prefix = self.config.get("receipt_prefix", "")
        start_no = self.config.get("receipt_no_start", "700605433")
        end_no = self.config.get("receipt_no_end", "700999999")
        nxt = self.db.get_next_receipt_no(prefix, start_no=start_no, end_no=end_no)
        self.entry_rcp_no.delete(0, "end")
        self.entry_rcp_no.insert(0, nxt)

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
            self.fac_var.set(stu.get("faculty") or self.faculties[0])
            self.dept_var.set(stu.get("major") or self.departments[0])
            self.cls_var.set(stu.get("class_group") or "ปวช.1/1")

    def refresh_student_list(self):
        self._stu_search_job = None
        cls_filter = self.filter_class_var.get()
        lvl = cls_filter if cls_filter != "ทั้งหมด" else "ALL"
        search_txt = self.entry_stu_search.get().strip() if hasattr(self, "entry_stu_search") else ""

        total_matching = self.db.count_students(search=search_txt, level_filter=lvl)
        self.students = self.db.get_students(
            search=search_txt,
            level_filter=lvl,
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
            self.opt_student.configure(values=labels)
            self.stu_var.set(labels[0])
            return

        if total_matching > len(self.students):
            labels.append(
                f"--- แสดง {len(self.students):,} จาก {total_matching:,} คน (พิมพ์ค้นหาหรือเลือกห้องเรียนเพื่อกรองเพิ่ม) ---"
            )

        self.opt_student.configure(values=labels)
        self.stu_var.set(labels[0])
        self.on_select_student(labels[0])

    def on_select_student(self, label: str):
        stu = self.student_map.get(label)
        if not stu:
            return
        self.entry_stu_name.delete(0, "end")
        self.entry_stu_name.insert(0, stu.get("full_name", ""))
        self.entry_stu_code.delete(0, "end")
        self.entry_stu_code.insert(0, stu.get("student_code", ""))

        self.fac_var.set(stu.get("faculty") or self.faculties[0])
        self.dept_var.set(stu.get("major") or self.departments[0])
        self.cls_var.set(stu.get("class_group") or "ปวช.1/1")

    def on_change_curriculum(self, curriculum_name: str):
        self.current_catalog = self.db.get_fee_catalog(curriculum_filter=curriculum_name)
        self.catalog_map.clear()
        labels = []
        for f in self.current_catalog:
            lbl = f"{f['fee_name']} ({float(f['default_price']):,.2f} บ.)"
            self.catalog_map[lbl] = f
            labels.append(lbl)

        if not labels:
            labels = ["ไม่มีรายการในหลักสูตรนี้"]
            self.opt_fee_item.configure(values=labels)
            self.fee_item_var.set(labels[0])
            return

        self.opt_fee_item.configure(values=labels)
        self.fee_item_var.set(labels[0])
        self.on_select_fee_item(labels[0])

    def on_select_fee_item(self, label: str):
        fee = self.catalog_map.get(label)
        if not fee:
            return
        cred = float(fee.get("default_credits", 0))
        self.entry_qty.delete(0, "end")
        self.entry_qty.insert(0, f"{cred:.0f}" if cred > 0 else "1")
        self.entry_unit_price.delete(0, "end")
        self.entry_unit_price.insert(0, f"{float(fee.get('default_price', 0)):.2f}")

    def add_selected_fee_to_cart(self):
        label = self.fee_item_var.get()
        fee = self.catalog_map.get(label)
        item_name = fee["fee_name"] if fee else label
        fee_code = fee["fee_code"] if fee else "CUSTOM"
        curr_name = fee.get("curriculum_name", self.curriculum_var.get()) if fee else self.curriculum_var.get()

        try:
            qty = float(self.entry_qty.get().strip().replace(",", "") or "1")
            unit_price = float(self.entry_unit_price.get().strip().replace(",", "") or "0")
        except ValueError:
            messagebox.showwarning("ตัวเลขไม่ถูกต้อง", "กรุณาระบุจำนวนและราคาเป็นตัวเลข")
            return

        if qty <= 0 or unit_price < 0:
            messagebox.showwarning("ข้อมูลไม่ถูกต้อง", "จำนวนต้องมากกว่า 0")
            return

        self.cart_items.append(
            {
                "fee_code": fee_code,
                "curriculum_name": curr_name,
                "item_name": item_name,
                "quantity": qty,
                "unit_price": unit_price,
                "discount": 0.0,
                "line_total": round(qty * unit_price, 2),
            }
        )
        self.refresh_cart_ui()

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

    def refresh_cart_ui(self):
        rows = []
        net_total = 0.0
        for idx, it in enumerate(self.cart_items, start=1):
            it["line_no"] = idx
            line_tot = round(float(it["quantity"]) * float(it["unit_price"]), 2)
            it["line_total"] = line_tot
            net_total += line_tot
            rows.append(
                (
                    [
                        str(idx),
                        it["item_name"],
                        f"{float(it['quantity']):,.0f}",
                        f"฿{float(it['unit_price']):,.2f}",
                        f"฿{line_tot:,.2f}",
                    ],
                    "PAID",
                )
            )
        self.table.set_rows(rows)
        self.lbl_net_total.configure(text=f"ยอดรวมสุทธิ: ฿{net_total:,.2f} บาท")
        self.lbl_baht_text.configure(text=f"ตัวอักษร: ({thai_baht_text(net_total)})")

    def issue_receipt_and_export(self):
        stu_code = self.entry_stu_code.get().strip()
        stu_name = self.entry_stu_name.get().strip()
        if not stu_code or not stu_name:
            messagebox.showwarning("เลือกนักศึกษา", "กรุณาเลือกนักศึกษาในขั้นตอนที่ 1 ก่อนบันทึกใบเสร็จ")
            return
        if not self.cart_items:
            messagebox.showwarning("ยังไม่มีรายการในใบเสร็จ", "กรุณากดปุ่ม '➕ เพิ่มลงใบเสร็จ' ในขั้นตอนที่ 2 อย่างน้อย 1 รายการ")
            return

        rcp_no = self.entry_rcp_no.get().strip()
        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)
        default_filename = f"Receipt_{rcp_no}_{stu_code}.pdf"

        save_path = filedialog.asksaveasfilename(
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
            return

        self.config["default_export_folder"] = os.path.dirname(save_path)
        save_config(self.config)

        try:
            saved = self.db.issue_flexible_receipt(
                student_code=stu_code,
                line_items=self.cart_items,
                payment_method=self.pay_method_var.get(),
                actor=self.session.username,
                cashier_name=self.session.full_name,
                role=self.session.role,
                book_no=self.config.get("receipt_book_no", "7006"),
                term_code=self.term_var.get(),
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

            export_document_auto(
                save_path,
                saved,
                college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
                metadata=self.config,
                is_batch=False,
            )

            messagebox.showinfo(
                "ออกใบเสร็จสำเร็จ",
                f"เลขที่ใบเสร็จ: {saved['receipt_no']}\n"
                f"ผู้ชำระ: {saved['full_name']}\n"
                f"ยอดสุทธิ: ฿{float(saved['net_total_amount']):,.2f} บาท\n\n"
                f"ไฟล์ใบเสร็จพร้อมพิมพ์บันทึกอยู่ที่:\n{save_path}",
            )
            self.reset_pos_form()
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถออกใบเสร็จได้:\n{e}")

    # ------------------------------------------------------------------
    # BATCH CLASSROOM RECEIPT ISSUING (ออกใบเสร็จแบบเดียวยกห้อง / Import รายชื่อเด็ก)
    # ------------------------------------------------------------------
    def open_batch_class_modal(self):
        if not self.cart_items:
            messagebox.showwarning(
                "กรุณาเพิ่มรายการค่าเทอมในขั้นตอนที่ 2 ก่อน",
                "กรุณาเลือกรายการค่าธรรมเนียมใส่ตารางใน 'ขั้นตอนที่ 2' ก่อนครับ\n"
                "แล้วกดปุ่มนี้อีกครั้งเพื่อนำรายการในตารางไปออกใบเสร็จให้เด็กทั้งห้องพร้อมกัน!",
            )
            return

        modal = ctk.CTkToplevel(self)
        modal.title("ออกใบเสร็จแบบเดียวยกห้อง / นำเข้าไฟล์รายชื่อเด็ก")
        modal.geometry("640x440+250+140")
        modal.configure(fg_color=AppTheme.COLOR_BG)
        modal.grab_set()
        modal.focus_force()

        net_total = sum(float(i["line_total"]) for i in self.cart_items)

        ctk.CTkLabel(
            modal,
            text="🏫 ออกใบเสร็จแบบเดียวยกห้อง",
            font=(AppTheme.FONT_FAMILY, 20, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(anchor="w", padx=24, pady=(20, 4))

        ctk.CTkLabel(
            modal,
            text=f"ยอดที่จะออกใบเสร็จคนละ: ฿{net_total:,.2f} บาท ({thai_baht_text(net_total)})",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        ).pack(anchor="w", padx=24, pady=(0, 16))

        # Option 1: Select Classroom Dropdown
        box1 = ctk.CTkFrame(modal, corner_radius=10, fg_color=AppTheme.COLOR_SURFACE, border_width=1, border_color=AppTheme.COLOR_BORDER)
        box1.pack(fill="x", padx=24, pady=8)

        ctk.CTkLabel(
            box1,
            text="วิธีที่ 1 : เลือกห้องเรียนที่มีรายชื่ออยู่ในระบบ",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(anchor="w", padx=16, pady=(12, 8))

        r_cls = ctk.CTkFrame(box1, fg_color="transparent")
        r_cls.pack(fill="x", padx=16, pady=(0, 14))

        cls_options = self.db.get_options("CLASS_LEVEL") or ["ปวช.1/1"]
        batch_cls_var = ctk.StringVar(value=cls_options[0])
        make_dropdown(
            r_cls,
            values=cls_options,
            variable=batch_cls_var,
            width=180,
            height=42,
            font_size=14,
        ).pack(side="left", padx=(0, 12))

        def run_batch_for_room():
            room = batch_cls_var.get()
            students_in_room = self.db.get_students(level_filter=room)
            if not students_in_room:
                messagebox.showwarning("ไม่มีรายชื่อในห้องนี้", f"ไม่พบรายชื่อนักศึกษาในห้อง {room}")
                return
            if not messagebox.askyesno(
                "ยืนยันออกใบเสร็จยกห้อง",
                f"ออกใบเสร็จให้ห้อง '{room}' จำนวน {len(students_in_room)} คน\nยอดคนละ ฿{net_total:,.2f} บาท ใช่หรือไม่?",
            ):
                return
            self._execute_batch_receipts(modal, students_in_room, label_tag=room)

        ctk.CTkButton(
            r_cls,
            text="🖨️ ออกใบเสร็จทั้งห้อง & สร้างไฟล์ PDF (.pdf)",
            height=42,
            corner_radius=8,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=run_batch_for_room,
        ).pack(side="left", fill="x", expand=True)

        # Option 2: Import Excel/CSV Student List & Issue Receipts Immediately
        box2 = ctk.CTkFrame(modal, corner_radius=10, fg_color=AppTheme.COLOR_SURFACE, border_width=1, border_color=AppTheme.COLOR_BORDER)
        box2.pack(fill="x", padx=24, pady=8)

        ctk.CTkLabel(
            box2,
            text="วิธีที่ 2 : นำเข้าไฟล์รายชื่อเด็ก (Excel .xlsx / .csv) แล้วออกใบเสร็จให้ทั้งไฟล์ทันที",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(anchor="w", padx=16, pady=(12, 8))

        def import_and_issue():
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
            if not messagebox.askyesno(
                "นำเข้ารายชื่อสำเร็จ",
                f"นำเข้ารายชื่อสำเร็จ {len(imported_students)} คน!\nต้องการออกใบเสร็จให้ทั้ง {len(imported_students)} คนทันทีหรือไม่?",
            ):
                return
            self._execute_batch_receipts(modal, imported_students, label_tag="ImportedClass")

        ctk.CTkButton(
            box2,
            text="📥 เลือกไฟล์ Excel/CSV รายชื่อเด็ก ➔ ออกใบเสร็จยกห้องทันที",
            height=42,
            corner_radius=8,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=import_and_issue,
        ).pack(fill="x", padx=16, pady=(0, 14))

    def _execute_batch_receipts(self, modal, students_list: List[Dict[str, Any]], label_tag: str):
        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)
        safe_tag = label_tag.replace("/", "-").replace(" ", "_")
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"BatchReceipts_{safe_tag}_{len(students_list)}students_{stamp}.pdf"

        save_path = filedialog.asksaveasfilename(
            parent=modal,
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
            return

        self.config["default_export_folder"] = os.path.dirname(save_path)
        save_config(self.config)

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
                start_no=self.config.get("receipt_no_start", "700605433"),
                end_no=self.config.get("receipt_no_end", "700999999"),
                receipt_date_override=self.entry_date.get().strip(),
            )

            export_document_auto(
                save_path,
                issued_list,
                college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
                metadata=self.config,
                is_batch=True,
            )

            first_no = issued_list[0]["receipt_no"]
            last_no = issued_list[-1]["receipt_no"]
            modal.destroy()
            self.reset_pos_form()
            messagebox.showinfo(
                "ออกใบเสร็จยกห้องสำเร็จ!",
                f"ออกใบเสร็จครบทั้งห้องจำนวน: {len(issued_list)} ฉบับ\n"
                f"เลขที่ใบเสร็จ: {first_no} ถึง {last_no}\n\n"
                f"ไฟล์ใบเสร็จรวมทั้งห้องพร้อมพิมพ์บันทึกอยู่ที่:\n{save_path}",
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถออกใบเสร็จยกห้องได้:\n{e}")


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
            "📥 นำเข้ารายชื่อจากไฟล์ Excel / CSV",
            self.import_students_dialog,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
        )

        # ------------------------------------------------------------------
        # BOTTOM EDITOR BAR
        # ------------------------------------------------------------------
        form_card = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        form_card.pack(side="bottom", fill="x", pady=(10, 0))

        f_row1 = ctk.CTkFrame(form_card, fg_color="transparent")
        f_row1.pack(fill="x", padx=14, pady=(12, 6))

        ctk.CTkLabel(
            f_row1,
            text="รหัสนักศึกษา:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.ent_code = make_entry(f_row1, width=120, height=38)
        self.ent_code.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            f_row1,
            text="ชื่อ-สกุล:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.ent_name = make_entry(f_row1, width=190, height=38)
        self.ent_name.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            f_row1,
            text="คณะ:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.form_fac_var = ctk.StringVar(value=self.faculties[0])
        make_dropdown(
            f_row1,
            values=self.faculties,
            variable=self.form_fac_var,
            width=150,
            height=38,
        ).pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            f_row1,
            text="สาขา:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.form_dept_var = ctk.StringVar(value=self.departments[0])
        make_dropdown(
            f_row1,
            values=self.departments,
            variable=self.form_dept_var,
            width=160,
            height=38,
        ).pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            f_row1,
            text="ห้อง:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.form_cls_var = ctk.StringVar(value=self.class_levels[0])
        make_dropdown(
            f_row1,
            values=self.class_levels,
            variable=self.form_cls_var,
            width=110,
            height=38,
        ).pack(side="left")

        f_row2 = ctk.CTkFrame(form_card, fg_color="transparent")
        f_row2.pack(fill="x", padx=14, pady=(4, 12))

        ctk.CTkButton(
            f_row2,
            text="💾 บันทึกข้อมูลนักศึกษา (เพิ่มใหม่ / อัปเดต)",
            width=260,
            height=40,
            corner_radius=8,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.save_or_update_student,
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            f_row2,
            text="ล้างช่องกรอก",
            width=115,
            height=40,
            corner_radius=8,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.clear_form,
        ).pack(side="left")

        ctk.CTkButton(
            f_row2,
            text="🗑 ลบรายชื่อนี้",
            width=130,
            height=40,
            corner_radius=8,
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
            corner_radius=10,
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
            width=180,
            height=38,
            command=lambda _v: self._reset_and_refresh(),
        ).pack(side="left", padx=(0, 14))

        ctk.CTkLabel(
            fb_in,
            text="กรองห้องเรียน:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))
        self.flt_cls_var = ctk.StringVar(value="ทั้งหมด")
        make_dropdown(
            fb_in,
            values=["ทั้งหมด"] + self.class_levels,
            variable=self.flt_cls_var,
            width=130,
            height=38,
            command=lambda _v: self._reset_and_refresh(),
        ).pack(side="left", padx=(0, 14))

        self.ent_search = make_entry(
            fb_in,
            height=38,
            bold=False,
            placeholder_text="🔍 ค้นหารหัสนักศึกษา หรือ ชื่อ-นามสกุล...",
        )
        self.ent_search.pack(side="left", fill="x", expand=True)
        self.ent_search.bind("<KeyRelease>", self._schedule_search)

        cols = [
            ("student_code", "รหัสนักศึกษา", 140, "center"),
            ("full_name", "ชื่อ-นามสกุล", 260, "w"),
            ("faculty", "คณะ", 180, "w"),
            ("major", "แผนกวิชา / สาขา", 220, "w"),
            ("class_group", "ระดับชั้น / ห้อง", 130, "center"),
            ("paid_amount", "ยอดชำระสะสม (บาท)", 160, "e"),
        ]
        self.table = DataTable(
            self,
            columns=cols,
            height=6,
            on_select=self.on_select_row,
            show_inspector=False,
            page_size=self.page_size,
            on_page_change=self.on_page_change,
            show_pagination=True,
        )
        self.table.pack(side="top", fill="both", expand=True)

        self.refresh_table()

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
        f_dept = self.flt_dept_var.get()
        f_cls = self.flt_cls_var.get()
        maj_flt = f_dept if f_dept != "ทั้งหมด" else "ALL"
        lvl_flt = f_cls if f_cls != "ทั้งหมด" else "ALL"

        total_count = self.db.count_students(
            search=s_txt,
            major_filter=maj_flt,
            level_filter=lvl_flt,
        )
        max_pages = max(1, (total_count + self.page_size - 1) // self.page_size)
        if self.current_page > max_pages:
            self.current_page = max_pages

        offset = (self.current_page - 1) * self.page_size
        students = self.db.get_students(
            search=s_txt,
            major_filter=maj_flt,
            level_filter=lvl_flt,
            limit=self.page_size,
            offset=offset,
        )
        rows = []
        for s in students:
            rows.append(
                (
                    [
                        s["student_code"],
                        s["full_name"],
                        s.get("faculty") or "บริหารธุรกิจ / พาณิชยกรรม",
                        s.get("major") or "-",
                        s.get("class_group") or "-",
                        f"฿{float(s.get('paid_amount', 0)):,.2f}",
                    ],
                    "PAID",
                )
            )
        self.table.set_rows(rows, total_count=total_count, page=self.current_page)

    def on_select_row(self, vals: List[str]):
        if len(vals) < 5:
            return
        self.ent_code.delete(0, "end")
        self.ent_code.insert(0, vals[0])
        self.ent_name.delete(0, "end")
        self.ent_name.insert(0, vals[1])
        self.form_fac_var.set(vals[2])
        self.form_dept_var.set(vals[3])
        self.form_cls_var.set(vals[4])

    def clear_form(self):
        self.ent_code.delete(0, "end")
        self.ent_name.delete(0, "end")

    def save_or_update_student(self):
        code = self.ent_code.get().strip()
        name = self.ent_name.get().strip()
        if not code or not name:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอกรหัสนักศึกษาและชื่อ-นามสกุล")
            return
        existing = self.db.get_student_by_code(code)
        is_exact = existing is not None
        try:
            if is_exact:
                self.db.update_student(
                    student_code=code,
                    full_name=name,
                    faculty=self.form_fac_var.get(),
                    major=self.form_dept_var.get(),
                    class_group=self.form_cls_var.get(),
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
                    actor=self.session.username,
                    role=self.session.role,
                )
                msg = f"เพิ่มข้อมูลนักศึกษาใหม่ {code} ({name}) เรียบร้อยแล้ว"
            self.refresh_table()
            self.clear_form()
            messagebox.showinfo("บันทึกสำเร็จ", msg)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def delete_selected_student(self):
        code = self.ent_code.get().strip()
        if not code:
            messagebox.showwarning("เลือกรายชื่อ", "กรุณาคลิกเลือกนักศึกษาในตารางก่อนลบ")
            return
        if not messagebox.askyesno("ยืนยันการลบ", f"ต้องการลบข้อมูลนักศึกษารหัส {code} ใช่หรือไม่?"):
            return
        self.db.delete_student(code, actor=self.session.username, role=self.session.role)
        self.refresh_table()
        self.clear_form()

    def import_students_dialog(self):
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์รายชื่อนักศึกษา (Excel .xlsx หรือ .csv)",
            filetypes=[("Excel / CSV Files", "*.xlsx *.csv")],
        )
        if not file_path:
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
            self.refresh_table()
            messagebox.showinfo(
                "นำเข้ารายชื่อสำเร็จ",
                f"นำเข้าข้อมูลนักศึกษาสำเร็จทั้งหมด {res['total']} คน\n"
                f"• เพิ่มใหม่: {res['inserted']} คน\n"
                f"• อัปเดตข้อมูลเดิม: {res['updated']} คน",
            )
        except Exception as e:
            messagebox.showerror("Import Error", f"ไม่สามารถนำเข้าไฟล์ได้:\n{e}")


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
            "📊 ส่งออกรายงาน Excel",
            self.export_excel_report,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        )
        header.add_action_button(
            "🖨️ พิมพ์รายงานประจำวัน (.pdf)",
            self.export_daily_closing_docx,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
        )

        # ------------------------------------------------------------------
        # BOTTOM ACTION BAR: Reprint or Void Selected Receipt
        # ------------------------------------------------------------------
        action_bar = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        action_bar.pack(side="bottom", fill="x", pady=(10, 0))

        ab_in = ctk.CTkFrame(action_bar, fg_color="transparent")
        ab_in.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(
            ab_in,
            text="เลขที่ใบเสร็จ:",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))

        self.ent_void_no = make_entry(
            ab_in,
            width=180,
            height=40,
            font_size=14,
            placeholder_text="คลิกเลือกในตาราง...",
        )
        self.ent_void_no.pack(side="left", padx=(0, 12))

        ctk.CTkButton(
            ab_in,
            text="🖨️ พิมพ์ใบเสร็จใบนี้ซ้ำ (.pdf)",
            width=210,
            height=40,
            corner_radius=8,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.reprint_selected_receipt,
        ).pack(side="left", padx=(0, 16))

        ctk.CTkLabel(
            ab_in,
            text="สาเหตุยกเลิก:",
            font=(AppTheme.FONT_FAMILY, 13),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))
        self.ent_void_reason = make_entry(ab_in, height=40, bold=False)
        self.ent_void_reason.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.ent_void_reason.insert(0, "รายการคลาดเคลื่อน / ออกใบเสร็จใหม่")

        ctk.CTkButton(
            ab_in,
            text="🚫 ยกเลิกใบเสร็จนี้ (Void)",
            width=185,
            height=40,
            corner_radius=8,
            fg_color=AppTheme.COLOR_DANGER_BG,
            hover_color="#FEE2E2",
            border_width=1,
            border_color=AppTheme.COLOR_DANGER_BORDER,
            text_color=AppTheme.COLOR_DANGER,
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.void_by_number,
        ).pack(side="right")

        # ------------------------------------------------------------------
        # TOP DATE & SEARCH BAR
        # ------------------------------------------------------------------
        top_bar = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        top_bar.pack(side="top", fill="x", pady=(0, 10))

        tb_in = ctk.CTkFrame(top_bar, fg_color="transparent")
        tb_in.pack(fill="x", padx=16, pady=10)

        ctk.CTkLabel(
            tb_in,
            text="📅 วันที่ออกใบเสร็จ:",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))

        dates = ["ทั้งหมด"] + self.db.get_distinct_receipt_dates()
        today_str = datetime.date.today().isoformat()
        if today_str not in dates:
            dates.insert(1, today_str)

        self.page_size = 100
        self.current_page = 1
        self._search_job: Optional[str] = None

        self.date_var = ctk.StringVar(value="ทั้งหมด")
        self.opt_date = make_dropdown(
            tb_in,
            values=dates,
            variable=self.date_var,
            width=155,
            height=38,
            command=lambda _v: self._reset_and_refresh(),
        )
        self.opt_date.pack(side="left", padx=(0, 14))

        self.ent_search = make_entry(
            tb_in,
            height=38,
            bold=False,
            placeholder_text="🔍 ค้นหาเลขที่ใบเสร็จ / รหัส / ชื่อนักศึกษา...",
        )
        self.ent_search.pack(side="left", fill="x", expand=True, padx=(0, 16))
        self.ent_search.bind("<KeyRelease>", self._schedule_search)

        self.lbl_summary = ctk.CTkLabel(
            tb_in,
            text="ยอดรับสุทธิ: ฿0.00 บาท (0 ฉบับ)",
            font=(AppTheme.FONT_FAMILY, 16, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        )
        self.lbl_summary.pack(side="right")

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
            self,
            columns=cols,
            height=6,
            on_select=self.on_select_row,
            show_inspector=False,
            page_size=self.page_size,
            on_page_change=self.on_page_change,
            show_pagination=True,
        )
        self.table.pack(side="top", fill="both", expand=True)

        self.refresh_receipts()

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
        d_val = self.date_var.get()
        search_txt = self.ent_search.get().strip()
        date_pref = d_val if d_val != "ทั้งหมด" else ""

        stats = self.db.get_receipts_stats(
            status_filter="ALL",
            search=search_txt,
            date_prefix=date_pref,
        )
        total_count = stats["total_count"]
        valid_count = stats["valid_count"]
        total_amt = stats["total_amount"]

        max_pages = max(1, (total_count + self.page_size - 1) // self.page_size)
        if self.current_page > max_pages:
            self.current_page = max_pages

        offset = (self.current_page - 1) * self.page_size
        receipts = self.db.get_receipts(
            search=search_txt,
            status_filter="ALL",
            date_prefix=date_pref,
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
        self.lbl_summary.configure(text=f"ยอดรับสุทธิ: ฿{total_amt:,.2f} บาท ({valid_count:,} ฉบับ)")

    def on_select_row(self, vals: List[str]):
        if vals:
            self.ent_void_no.delete(0, "end")
            self.ent_void_no.insert(0, vals[0])

    def void_by_number(self):
        rcp_no = self.ent_void_no.get().strip()
        reason = self.ent_void_reason.get().strip() or "ยกเลิกใบเสร็จ"
        if not rcp_no:
            messagebox.showwarning("ระบุเลขที่ใบเสร็จ", "กรุณาคลิกเลือกใบเสร็จในตาราง หรือพิมพ์เลขใบเสร็จก่อนยกเลิก")
            return
        if not messagebox.askyesno("ยืนยันยกเลิกใบเสร็จ", f"ต้องการยกเลิกใบเสร็จเลขที่ '{rcp_no}' ใช่หรือไม่?\nสาเหตุ: {reason}"):
            return
        try:
            self.db.void_receipt(rcp_no, reason, self.session.username, self.session.role)
            self.refresh_receipts()
            messagebox.showinfo("ยกเลิกสำเร็จ", f"ยกเลิกใบเสร็จเลขที่ {rcp_no} เรียบร้อยแล้ว")
        except Exception as e:
            messagebox.showerror("ไม่สามารถยกเลิกได้", str(e))

    def reprint_selected_receipt(self):
        rcp_no = self.ent_void_no.get().strip()
        if not rcp_no:
            messagebox.showwarning("เลือกใบเสร็จ", "กรุณาคลิกเลือกใบเสร็จในตารางก่อนสั่งพิมพ์ซ้ำ")
            return
        data = self.db.get_receipt_by_no(rcp_no)
        if not data:
            messagebox.showerror("ไม่พบข้อมูล", f"ไม่พบใบเสร็จเลขที่ {rcp_no}")
            return
        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)
        path = filedialog.asksaveasfilename(
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
            return
        self.config["default_export_folder"] = os.path.dirname(path)
        save_config(self.config)
        export_document_auto(
            path,
            data,
            college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
            metadata=self.config,
            is_batch=False,
        )
        messagebox.showinfo("พิมพ์ใบเสร็จซ้ำสำเร็จ", f"ไฟล์ใบเสร็จพร้อมพิมพ์บันทึกอยู่ที่:\n{path}")

    def export_daily_closing_docx(self):
        d_val = self.date_var.get()
        target_date = d_val if d_val != "ทั้งหมด" else datetime.date.today().isoformat()
        closing = self.db.get_daily_closing_summary(target_date)
        closing["report_type"] = "DAILY_CLOSING"
        closing["doc_title"] = f"รายงานสรุปการรับเงินค่าเทอมประจำวัน ({target_date})"
        closing["recorded_by"] = self.session.full_name

        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)
        path = filedialog.asksaveasfilename(
            title="เลือกตำแหน่งบันทึกรายงานประจำวัน (.pdf)",
            initialdir=out_dir,
            initialfile=f"SPVC_DailyReport_{target_date}.pdf",
            defaultextension=".pdf",
            filetypes=[
                ("PDF Document (*.pdf)", "*.pdf"),
                ("Word Document (*.docx)", "*.docx"),
            ],
        )
        if not path:
            return
        self.config["default_export_folder"] = os.path.dirname(path)
        save_config(self.config)
        export_document_auto(
            path,
            closing,
            college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
            metadata=self.config,
            is_batch=False,
        )
        messagebox.showinfo("พิมพ์รายงานประจำวันสำเร็จ", f"สร้างไฟล์รายงานสรุปยอดประจำวัน ({target_date}) เรียบร้อยที่:\n{path}")

    def export_excel_report(self):
        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = filedialog.asksaveasfilename(
            title="เลือกตำแหน่งบันทึกรายงาน Excel (.xlsx)",
            initialdir=out_dir,
            initialfile=f"SPVC_Receipt_Report_{stamp}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook (*.xlsx)", "*.xlsx")],
        )
        if not path:
            return
        self.config["default_export_folder"] = os.path.dirname(path)
        save_config(self.config)
        payload = {
            "summary": self.db.get_dashboard_summary(),
            "transactions": self.db.get_receipts(),
            "students": self.db.get_students(),
            "budgets": [],
        }
        ExcelFinanceExporter(path, college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")).export(
            payload, metadata=self.config
        )
        messagebox.showinfo("ส่งออก Excel สำเร็จ", f"บันทึกไฟล์รายงาน Excel เรียบร้อยที่:\n{path}")


# ==============================================================================
# VIEW 4: SETTINGS, DROPDOWN MASTER & FEE CATALOG (หน้า 4: ตั้งค่าระบบ)
# ==============================================================================
class SettingsAndBackupView(ctk.CTkFrame):
    """
    Clean 3-Tab Settings View with Unified Corporate Palette:
    - แท็บ 1: รายการค่าธรรมเนียม (ตามหลักสูตร)
    - แท็บ 2: ตัวเลือกใน Dropdown (คณะ / สาขา / ห้องเรียน / หลักสูตร)
    - แท็บ 3: เลขที่ใบเสร็จ & สำรองข้อมูล
    """

    def __init__(self, parent, db: DatabaseManager, session: UserSession, config: Dict[str, Any]):
        super().__init__(parent, fg_color="transparent")
        self.db = db
        self.session = session
        self.config = config

        header = SectionHeader(
            self,
            "ตั้งค่าระบบ",
            "เพิ่ม/ลบ รายการค่าธรรมเนียม, ตัวเลือก Dropdown และกำหนดช่วงเลขที่ใบเสร็จ",
        )
        header.pack(fill="x", pady=(0, 8))

        self.tabs = ctk.CTkTabview(
            self,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
            segmented_button_fg_color="#334155",
            segmented_button_selected_color=AppTheme.COLOR_PRIMARY_BLUE,
            segmented_button_selected_hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            segmented_button_unselected_color="#475569",
            segmented_button_unselected_hover_color="#334155",
            text_color="#FFFFFF",
        )
        self.tabs.pack(fill="both", expand=True)
        self.tabs._segmented_button.configure(font=(AppTheme.FONT_FAMILY, 14, "bold"), height=40)

        self.tab_fees = self.tabs.add("1. รายการค่าธรรมเนียม (ตามหลักสูตร)")
        self.tab_dropdowns = self.tabs.add("2. เพิ่ม/ลบตัวเลือก Dropdown")
        self.tab_system = self.tabs.add("3. เลขที่ใบเสร็จ & สำรองข้อมูล")

        self._build_tab_fees()
        self._build_tab_dropdowns()
        self._build_tab_system()

    def _build_tab_fees(self):
        curricula = self.db.get_options("CURRICULUM") or ["หลักสูตรในระบบ-ปวช"]

        edit_bar = ctk.CTkFrame(
            self.tab_fees,
            corner_radius=8,
            fg_color=AppTheme.COLOR_SURFACE_ALT,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        edit_bar.pack(side="bottom", fill="x", padx=10, pady=8)

        eb_in = ctk.CTkFrame(edit_bar, fg_color="transparent")
        eb_in.pack(fill="x", padx=12, pady=10)

        ctk.CTkLabel(
            eb_in,
            text="รหัส:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.ent_fee_code = make_entry(eb_in, width=90, height=38)
        self.ent_fee_code.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            eb_in,
            text="ชื่อรายการค่าธรรมเนียม:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.ent_fee_name = make_entry(eb_in, height=38)
        self.ent_fee_name.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkLabel(
            eb_in,
            text="ราคา (บาท):",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 4))
        self.ent_fee_price = make_entry(eb_in, width=110, height=38)
        self.ent_fee_price.pack(side="left", padx=(0, 12))

        ctk.CTkButton(
            eb_in,
            text="💾 บันทึกรายการ",
            width=150,
            height=38,
            corner_radius=8,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.save_fee_item,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            eb_in,
            text="🗑 ลบ",
            width=90,
            height=38,
            corner_radius=8,
            fg_color=AppTheme.COLOR_DANGER_BG,
            hover_color="#FEE2E2",
            border_width=1,
            border_color=AppTheme.COLOR_DANGER_BORDER,
            text_color=AppTheme.COLOR_DANGER,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.delete_fee_item,
        ).pack(side="left")

        top_f = ctk.CTkFrame(self.tab_fees, fg_color="transparent")
        top_f.pack(side="top", fill="x", padx=10, pady=(4, 8))

        ctk.CTkLabel(
            top_f,
            text="เลือกหลักสูตร:",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 8))
        self.fee_curr_var = ctk.StringVar(value=curricula[0])
        self.opt_fee_curr = make_dropdown(
            top_f,
            values=curricula,
            variable=self.fee_curr_var,
            width=240,
            height=38,
            font_size=14,
            command=lambda _v: self.refresh_fee_table(),
        )
        self.opt_fee_curr.pack(side="left")

        cols = [
            ("fee_code", "รหัสรายการ", 130, "center"),
            ("curriculum_name", "หลักสูตร", 220, "w"),
            ("fee_name", "ชื่อรายการค่าธรรมเนียม", 420, "w"),
            ("default_price", "จำนวนเงินต่อหน่วย (บาท)", 180, "e"),
        ]
        self.fee_table = DataTable(self.tab_fees, columns=cols, height=5, on_select=self.on_select_fee_row, show_inspector=False)
        self.fee_table.pack(side="top", fill="both", expand=True, padx=10)
        self.refresh_fee_table()

    def refresh_fee_table(self):
        curr = self.fee_curr_var.get()
        items = self.db.get_fee_catalog(curriculum_filter=curr)
        rows = []
        for it in items:
            rows.append(
                (
                    [
                        it["fee_code"],
                        it["curriculum_name"],
                        it["fee_name"],
                        f"{float(it['default_price']):,.2f}",
                    ],
                    "PAID",
                )
            )
        self.fee_table.set_rows(rows)

    def on_select_fee_row(self, vals: List[str]):
        if len(vals) < 4:
            return
        self.ent_fee_code.delete(0, "end")
        self.ent_fee_code.insert(0, vals[0])
        self.ent_fee_name.delete(0, "end")
        self.ent_fee_name.insert(0, vals[2])
        self.ent_fee_price.delete(0, "end")
        self.ent_fee_price.insert(0, vals[3].replace(",", ""))

    def save_fee_item(self):
        code = self.ent_fee_code.get().strip()
        name = self.ent_fee_name.get().strip()
        curr = self.fee_curr_var.get()
        if not name:
            messagebox.showwarning("ระบุชื่อรายการ", "กรุณาระบุชื่อรายการค่าธรรมเนียม")
            return
        try:
            price = float(self.ent_fee_price.get().strip().replace(",", "") or "0")
        except ValueError:
            messagebox.showwarning("ราคาไม่ถูกต้อง", "กรุณากรอกจำนวนเงินเป็นตัวเลข")
            return

        self.db.upsert_fee_item(
            fee_code=code,
            fee_name=name,
            default_price=price,
            curriculum_name=curr,
            actor=self.session.username,
            role=self.session.role,
        )
        self.refresh_fee_table()
        messagebox.showinfo("บันทึกสำเร็จ", f"บันทึกรายการ '{name}' ({price:,.2f} บาท) เรียบร้อยแล้ว")

    def delete_fee_item(self):
        code = self.ent_fee_code.get().strip()
        if not code:
            messagebox.showwarning("เลือกรายการ", "กรุณาคลิกเลือกรายการในตารางก่อนลบ")
            return
        if not messagebox.askyesno("ยืนยันการลบ", f"ต้องการลบรายการรหัส {code} ใช่หรือไม่?"):
            return
        self.db.delete_fee_item(code, actor=self.session.username, role=self.session.role)
        self.refresh_fee_table()

    def _build_tab_dropdowns(self):
        self.category_labels = {
            "หลักสูตร (CURRICULUM)": "CURRICULUM",
            "คณะ (FACULTY)": "FACULTY",
            "แผนกวิชา / สาขา (DEPARTMENT)": "DEPARTMENT",
            "ชั้นปี / ห้องเรียน (CLASS_LEVEL)": "CLASS_LEVEL",
            "ภาค / ปีการศึกษา (TERM)": "TERM",
            "ตำแหน่งผู้รับเงิน (CASHIER_POSITION)": "CASHIER_POSITION",
        }

        bottom_d = ctk.CTkFrame(
            self.tab_dropdowns,
            corner_radius=8,
            fg_color=AppTheme.COLOR_SURFACE_ALT,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        bottom_d.pack(side="bottom", fill="x", padx=10, pady=8)

        bd_in = ctk.CTkFrame(bottom_d, fg_color="transparent")
        bd_in.pack(fill="x", padx=12, pady=10)

        ctk.CTkLabel(
            bd_in,
            text="ชื่อตัวเลือก:",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 8))
        self.ent_new_opt = make_entry(
            bd_in,
            height=38,
            font_size=14,
            placeholder_text="พิมพ์ตัวเลือกที่ต้องการเพิ่ม เช่น ปวช.1/5 หรือ คอมพิวเตอร์กราฟิก...",
        )
        self.ent_new_opt.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkButton(
            bd_in,
            text="➕ เพิ่มตัวเลือก",
            width=160,
            height=38,
            corner_radius=8,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.add_dropdown_option,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            bd_in,
            text="🗑 ลบตัวเลือกนี้",
            width=140,
            height=38,
            corner_radius=8,
            fg_color=AppTheme.COLOR_DANGER_BG,
            hover_color="#FEE2E2",
            border_width=1,
            border_color=AppTheme.COLOR_DANGER_BORDER,
            text_color=AppTheme.COLOR_DANGER,
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.delete_dropdown_option,
        ).pack(side="left")

        top_d = ctk.CTkFrame(self.tab_dropdowns, fg_color="transparent")
        top_d.pack(side="top", fill="x", padx=10, pady=(4, 8))

        ctk.CTkLabel(
            top_d,
            text="เลือกหมวดหมู่ Dropdown:",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 8))
        cat_keys = list(self.category_labels.keys())
        self.opt_cat_var = ctk.StringVar(value=cat_keys[0])
        make_dropdown(
            top_d,
            values=cat_keys,
            variable=self.opt_cat_var,
            width=300,
            height=38,
            font_size=14,
            command=lambda _v: self.refresh_dropdown_table(),
        ).pack(side="left")

        cols = [
            ("idx", "ลำดับ", 90, "center"),
            ("cat", "หมวดหมู่ Dropdown", 260, "w"),
            ("val", "ค่าตัวเลือกที่แสดงใน Dropdown", 520, "w"),
        ]
        self.dd_table = DataTable(self.tab_dropdowns, columns=cols, height=5, on_select=self.on_select_dd_row, show_inspector=False)
        self.dd_table.pack(side="top", fill="both", expand=True, padx=10)
        self.refresh_dropdown_table()

    def refresh_dropdown_table(self):
        lbl = self.opt_cat_var.get()
        cat_code = self.category_labels[lbl]
        opts = self.db.get_options(cat_code)
        rows = []
        for idx, val in enumerate(opts, start=1):
            rows.append(([str(idx), lbl, val], "PAID"))
        self.dd_table.set_rows(rows)

    def on_select_dd_row(self, vals: List[str]):
        if len(vals) >= 3:
            self.ent_new_opt.delete(0, "end")
            self.ent_new_opt.insert(0, vals[2])

    def add_dropdown_option(self):
        val = self.ent_new_opt.get().strip()
        if not val:
            messagebox.showwarning("ระบุข้อมูล", "กรุณาพิมพ์ชื่อตัวเลือกที่ต้องการเพิ่ม")
            return
        cat_code = self.category_labels[self.opt_cat_var.get()]
        self.db.add_option(cat_code, val, actor=self.session.username, role=self.session.role)
        self.ent_new_opt.delete(0, "end")
        self.refresh_dropdown_table()
        if cat_code == "CURRICULUM":
            currs = self.db.get_options("CURRICULUM")
            self.opt_fee_curr.configure(values=currs)
        messagebox.showinfo("เพิ่มสำเร็จ", f"เพิ่ม '{val}' ลงใน Dropdown เรียบร้อยแล้ว!")

    def delete_dropdown_option(self):
        val = self.ent_new_opt.get().strip()
        if not val:
            messagebox.showwarning("เลือกรายการ", "กรุณาคลิกเลือกแถวในตารางก่อนลบ")
            return
        cat_code = self.category_labels[self.opt_cat_var.get()]
        if not messagebox.askyesno("ยืนยันการลบ", f"ต้องการลบตัวเลือก '{val}' ออกจาก Dropdown ใช่หรือไม่?"):
            return
        self.db.delete_option(cat_code, val, actor=self.session.username, role=self.session.role)
        self.ent_new_opt.delete(0, "end")
        self.refresh_dropdown_table()

    def _build_tab_system(self):
        grid = ctk.CTkFrame(self.tab_system, fg_color="transparent")
        grid.pack(fill="both", expand=True, padx=14, pady=10)
        grid.grid_columnconfigure((0, 1), weight=1)

        c_left = ctk.CTkFrame(
            grid,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE_ALT,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        c_left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        ctk.CTkLabel(
            c_left,
            text="🏛️ กำหนดข้อมูลใบเสร็จเบื้องต้น",
            font=(AppTheme.FONT_FAMILY, 16, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(anchor="w", padx=18, pady=(16, 10))

        self.ent_college = self._labeled_entry(c_left, "ที่ทำการ (ชื่อสถานศึกษา):", self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"))

        pos_row = ctk.CTkFrame(c_left, fg_color="transparent")
        pos_row.pack(fill="x", padx=18, pady=5)
        ctk.CTkLabel(
            pos_row,
            text="ตำแหน่งผู้รับเงิน:",
            width=170,
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
            height=36,
        ).pack(side="left", fill="x", expand=True)

        self.ent_book = self._labeled_entry(c_left, "เล่มที่ใบเสร็จ:", self.config.get("receipt_book_no", "7006"))
        self.ent_start_no = self._labeled_entry(c_left, "เลขที่ใบเสร็จตั้งแต่:", self.config.get("receipt_no_start", "700605433"))
        self.ent_end_no = self._labeled_entry(c_left, "ถึงใบเสร็จเลขที่:", self.config.get("receipt_no_end", "700999999"))

        ctk.CTkButton(
            c_left,
            text="💾 บันทึกการตั้งค่า",
            height=42,
            corner_radius=8,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.save_system_config,
        ).pack(fill="x", padx=18, pady=(16, 16))

        c_right = ctk.CTkFrame(
            grid,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE_ALT,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        c_right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        ctk.CTkLabel(
            c_right,
            text="🛡️ สำรองและกู้คืนฐานข้อมูล",
            font=(AppTheme.FONT_FAMILY, 16, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(anchor="w", padx=18, pady=(16, 8))

        summary = self.db.get_dashboard_summary()
        ctk.CTkLabel(
            c_right,
            text=(
                f"• นักศึกษาในระบบ: {summary['student_count']:,} คน\n"
                f"• ใบเสร็จที่ออกแล้ว: {summary['receipt_count']:,} ฉบับ\n"
                f"• ยอดรับชำระสะสม: ฿{summary['total_income']:,.2f} บาท"
            ),
            font=(AppTheme.FONT_FAMILY, 13),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            justify="left",
        ).pack(anchor="w", padx=18, pady=(0, 16))

        ctk.CTkButton(
            c_right,
            text="📦 สำรองข้อมูลเก็บไว้ (.zip Full Backup)",
            height=42,
            corner_radius=8,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.run_zip_backup,
        ).pack(fill="x", padx=18, pady=6)

        ctk.CTkButton(
            c_right,
            text="♻️ กู้คืนฐานข้อมูลจากไฟล์สำรอง (.db)",
            height=42,
            corner_radius=8,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.run_restore_db,
        ).pack(fill="x", padx=18, pady=6)

    def _labeled_entry(self, parent, label_text: str, initial_val: str) -> ctk.CTkEntry:
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=5)
        ctk.CTkLabel(
            row,
            text=label_text,
            width=170,
            anchor="w",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left")
        ent = make_entry(row, height=36)
        ent.pack(side="left", fill="x", expand=True)
        ent.insert(0, str(initial_val))
        return ent

    def save_system_config(self):
        self.config["college_name"] = self.ent_college.get().strip()
        self.config["cashier_position"] = self.pos_var.get().strip()
        self.config["receipt_book_no"] = self.ent_book.get().strip()
        self.config["receipt_no_start"] = self.ent_start_no.get().strip()
        self.config["receipt_no_end"] = self.ent_end_no.get().strip()
        save_config(self.config)
        messagebox.showinfo("บันทึกสำเร็จ", "บันทึกการตั้งค่าข้อมูลเบื้องต้นเรียบร้อยแล้ว")

    def run_zip_backup(self):
        out_folder = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับเก็บไฟล์สำรองข้อมูล (.zip)")
        if not out_folder:
            return
        zip_path = export_full_zip_backup(
            self.db.db_path,
            self.config.get("evidence_folder", ""),
            out_folder,
        )
        messagebox.showinfo("สำรองข้อมูลสำเร็จ", f"บันทึกไฟล์สำรองข้อมูล (.zip) เรียบร้อยที่:\n{zip_path}")

    def run_restore_db(self):
        src = filedialog.askopenfilename(
            title="เลือกไฟล์ฐานข้อมูลสำรอง (.db) เพื่อกู้คืน",
            filetypes=[("SQLite Database", "*.db")],
        )
        if not src:
            return
        if not messagebox.askyesno("ยืนยันการกู้คืนฐานข้อมูล", "ต้องการกู้คืนฐานข้อมูลจากไฟล์ที่เลือกใช่หรือไม่?"):
            return
        restore_database(src, self.db.db_path)
        messagebox.showinfo("กู้คืนสำเร็จ", "กู้คืนฐานข้อมูลเรียบร้อยแล้ว!")


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

        self.students = self.db.get_students()
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
            "🔄 โหลดวิชาตามหลักสูตรต้นฉบับ (11 วิชา)",
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
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER_STRONG,
        )
        btm_bar.pack(side="bottom", fill="x", pady=(10, 0))

        btm_in = ctk.CTkFrame(btm_bar, fg_color="transparent")
        btm_in.pack(fill="x", padx=18, pady=12)

        self.lbl_totals = ctk.CTkLabel(
            btm_in,
            text="รวมชั่วโมง/หน่วยกิต: ทฤษฎี (ท.) 10  |  ปฏิบัติ (ป.) 20  |  หน่วยกิต (นก.) 19   •   ค่าธรรมเนียมการศึกษา: ฿1,130.00 บาท",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
        )
        self.lbl_totals.pack(side="left")

        btn_box = ctk.CTkFrame(btm_in, fg_color="transparent")
        btn_box.pack(side="right")

        ctk.CTkButton(
            btn_box,
            text="ลบวิชาที่เลือก",
            width=115,
            height=40,
            corner_radius=8,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.remove_selected_course,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_box,
            text="🖨️ บันทึก & พิมพ์บัตรลงทะเบียนรายวิชา อศท.28 (.pdf)",
            width=340,
            height=42,
            corner_radius=8,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=self.save_and_export_registration_card,
        ).pack(side="left")

        # ------------------------------------------------------------------
        # STEP 1: Student & Curriculum Metadata Card
        # ------------------------------------------------------------------
        stu_card = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        stu_card.pack(side="top", fill="x", pady=(0, 8))

        r1 = ctk.CTkFrame(stu_card, fg_color="transparent")
        r1.pack(fill="x", padx=16, pady=(10, 4))

        ctk.CTkLabel(
            r1,
            text="เลือกนักศึกษา:",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))

        stu_labels = []
        default_lbl = ""
        for s in self.students:
            lbl = f"{s['student_code']} — {s['full_name']} ({s['class_group']})"
            self.student_map[lbl] = s
            stu_labels.append(lbl)
            if s["student_code"] == "69219100049":
                default_lbl = lbl
        if not stu_labels:
            stu_labels = ["-"]
        if not default_lbl:
            default_lbl = stu_labels[0]

        self.stu_var = ctk.StringVar(value=default_lbl)
        self.opt_stu = make_dropdown(
            r1,
            values=stu_labels,
            variable=self.stu_var,
            width=340,
            height=36,
            command=self.on_select_student,
        )
        self.opt_stu.pack(side="left", padx=(0, 12))

        ctk.CTkLabel(r1, text="รหัสนักศึกษา:", font=(AppTheme.FONT_FAMILY, 13), text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 4))
        self.ent_code = make_entry(r1, width=115, height=34)
        self.ent_code.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(r1, text="ชื่อ-นามสกุล:", font=(AppTheme.FONT_FAMILY, 13), text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 4))
        self.ent_name = make_entry(r1, width=190, height=34)
        self.ent_name.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(r1, text="ระดับชั้น:", font=(AppTheme.FONT_FAMILY, 13), text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 4))
        self.ent_cls = make_entry(r1, width=85, height=34)
        self.ent_cls.pack(side="left")

        r2 = ctk.CTkFrame(stu_card, fg_color="transparent")
        r2.pack(fill="x", padx=16, pady=(2, 10))

        ctk.CTkLabel(r2, text="สาขาวิชา:", font=(AppTheme.FONT_FAMILY, 13), text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 4))
        self.ent_major = make_entry(r2, width=170, height=34)
        self.ent_major.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(r2, text="ประเภทวิชา:", font=(AppTheme.FONT_FAMILY, 13), text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 4))
        self.ent_subj_type = make_entry(r2, width=230, height=34, bold=False)
        self.ent_subj_type.pack(side="left", padx=(0, 10))
        self.ent_subj_type.insert(0, "อุตสาหกรรมดิจิทัลและเทคโนโลยีสารสนเทศ")

        ctk.CTkLabel(r2, text="กลุ่มอาชีพ:", font=(AppTheme.FONT_FAMILY, 13), text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 4))
        self.ent_career = make_entry(r2, width=220, height=34, bold=False)
        self.ent_career.pack(side="left", padx=(0, 10))
        self.ent_career.insert(0, "ธุรกิจดิจิทัลและพาณิชย์อิเล็กทรอนิกส์")

        ctk.CTkLabel(r2, text="ปีการศึกษา:", font=(AppTheme.FONT_FAMILY, 13), text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 4))
        self.ent_year = make_entry(r2, width=70, height=34)
        self.ent_year.pack(side="left")
        self.ent_year.insert(0, self.config.get("fiscal_year", "2569"))

        # ------------------------------------------------------------------
        # STEP 2: Add Course Row Bar
        # ------------------------------------------------------------------
        add_bar = ctk.CTkFrame(
            self,
            corner_radius=10,
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
            corner_radius=8,
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

    def on_select_student(self, label: str):
        stu = self.student_map.get(label)
        if not stu:
            return
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
            text=f"รวมชั่วโมง/หน่วยกิต ({len(self.registered_courses)} วิชา): ท. {tot_t}  |  ป. {tot_p}  |  นก. {tot_c}   •   ค่าธรรมเนียมการศึกษา: ฿1,130.00 บาท"
        )

    def save_and_export_registration_card(self):
        stu_code = self.ent_code.get().strip()
        stu_name = self.ent_name.get().strip()
        if not stu_code or not stu_name:
            messagebox.showwarning("เลือกนักศึกษา", "กรุณาเลือกนักศึกษาก่อนพิมพ์บัตรลงทะเบียนรายวิชา")
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
            return

        payload = {
            "report_type": "COURSE_REGISTRATION",
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

        export_document_auto(
            save_path,
            payload,
            college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
            metadata=self.config,
            is_batch=False,
        )
        messagebox.showinfo(
            "พิมพ์บัตรลงทะเบียนรายวิชาสำเร็จ",
            f"บันทึกไฟล์บัตรลงทะเบียนรายวิชา (อศท.28) ของ {stu_name} เรียบร้อยที่:\n{save_path}",
        )

