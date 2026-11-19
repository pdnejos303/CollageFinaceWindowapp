import datetime
import os
from tkinter import filedialog, messagebox
from typing import Any, Callable, Dict, List, Optional
import customtkinter as ctk

from core.config import (
    format_next_receipt_number,
    format_receipt_date_thai,
    normalize_template_style,
    save_config,
    split_receipt_prefix_and_number,
    thai_baht_text,
)
from core.database import DatabaseManager
from core.models import UserSession
from exporters.docx_exporter import DocxReceiptExporter
from exporters.pdf_exporter import PdfReceiptExporter
from ui.components import (
    DataTable,
    SearchableDropdown,
    SectionHeader,
    choose_export_folder,
    format_export_folder_button_text,
    open_document_file,
)
from ui.notifications import highlight_missing_field
from ui.theme import AppTheme

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
except ImportError:
    openpyxl = None

TEMPLATE_STYLE_OPTIONS = [
    "CLASSIC - ต้นฉบับทางการ (ตามแบบ EX 100%)",
    "MODERN - รูปแบบทันสมัย (อัปเกรด)",
]


class BatchClassReceiptView(ctk.CTkFrame):
    """
    Classroom Batch Receipt Studio (ระบบจัดใบเสร็จยกห้องเรียน).
    Clean, Flat Rectangular Design (No Nested Frames, Sharp Square Corners).
    """

    def __init__(
        self,
        parent,
        db: DatabaseManager,
        session: UserSession,
        config: Dict[str, Any],
        initial_cart: Optional[List[Dict[str, Any]]] = None,
        on_navigate: Optional[Callable[[str], None]] = None,
    ):
        super().__init__(parent, fg_color="transparent")
        self.db = db
        self.session = session
        self.config = config
        self.on_navigate = on_navigate

        # State data
        self.classroom_students: List[Dict[str, Any]] = []
        self.classroom_fees: List[Dict[str, Any]] = []
        self.current_catalog: List[Dict[str, Any]] = []
        self.catalog_map: Dict[str, Dict[str, Any]] = {}
        self.selected_file_path: str = ""
        self._last_data_version: int = -1

        # Master dropdown options
        self.curricula = self.db.get_options("CURRICULUM") or ["หลักสูตรในระบบ-ปวช"]
        self.faculties = self.db.get_options("FACULTY") or ["บริหารธุรกิจ / พาณิชยกรรม"]
        self.departments = self.db.get_options("DEPARTMENT") or ["การบัญชี"]
        self.class_levels = self.db.get_options("CLASS_LEVEL") or ["ปวช.1/1"]
        self.terms = self.db.get_options("TERM") or ["1/2569"]

        # Variables
        self.curriculum_var = ctk.StringVar(value=self.curricula[0] if self.curricula else "หลักสูตรในระบบ-ปวช")
        self.fee_item_var = ctk.StringVar(value="")
        self.unit_price_var = ctk.StringVar(value="0.00")
        self.qty_var = ctk.StringVar(value="1")
        self.fallback_cls_var = ctk.StringVar(value=self.class_levels[0] if self.class_levels else "ปวช.1/1")
        self.fallback_dept_var = ctk.StringVar(value=self.departments[0] if self.departments else "การบัญชี")
        self.fallback_fac_var = ctk.StringVar(value=self.faculties[0] if self.faculties else "บริหารธุรกิจ / พาณิชยกรรม")
        self.system_room_var = ctk.StringVar(value=self.class_levels[0] if self.class_levels else "ปวช.1/1")

        curr_term = self.config.get("current_semester", "1/2569")
        self.term_var = ctk.StringVar(value=curr_term if curr_term in self.terms else (self.terms[0] if self.terms else "1/2569"))
        self.pay_method_var = ctk.StringVar(value="เงินสด (Cash)")
        self.tpl_style_var = ctk.StringVar(
            value=TEMPLATE_STYLE_OPTIONS[1] if normalize_template_style(self.config.get("bill_template_style", "CLASSIC")) == "MODERN" else TEMPLATE_STYLE_OPTIONS[0]
        )

        # ------------------------------------------------------------------
        # 1. TOP HEADER BAR (Spacious, Flat)
        # ------------------------------------------------------------------
        header = SectionHeader(
            self,
            "จัดใบเสร็จทั้งห้อง",
            "นำเข้าไฟล์รายชื่อนักศึกษา (Excel/CSV) หรือเลือกห้องเรียน และตั้งค่าธรรมเนียมเดียวกันทั้งห้อง",
            corner_radius=AppTheme.RADIUS_CARD,
        )
        header.pack(fill="x", pady=(0, AppTheme.SPACE_MD))

        header.add_action_button(
            "ดาวน์โหลดตัวอย่าง Excel",
            self.download_excel_template,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            corner_radius=AppTheme.RADIUS_BTN,
        )

        # ------------------------------------------------------------------
        # 2. BOTTOM EXECUTION BAR (Flat, Single Border, Packed First)
        # ------------------------------------------------------------------
        bottom_bar = ctk.CTkFrame(
            self,
            corner_radius=AppTheme.RADIUS_CARD,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER_STRONG,
        )
        bottom_bar.pack(side="bottom", fill="x", pady=(AppTheme.SPACE_MD, 0))

        btm_inner = ctk.CTkFrame(bottom_bar, fg_color="transparent")
        btm_inner.pack(fill="x", padx=20, pady=14)

        # Left: Financial Summary for Classroom
        tot_box = ctk.CTkFrame(btm_inner, fg_color="transparent")
        tot_box.pack(side="left", fill="x", expand=True)

        self.lbl_grand_total = ctk.CTkLabel(
            tot_box,
            text="ยอดรวมทั้งห้อง: ฿0.00 บาท",
            font=(AppTheme.FONT_FAMILY, 20, "bold"),
            text_color=AppTheme.COLOR_ACCENT,
            anchor="w",
        )
        self.lbl_grand_total.pack(anchor="w")

        self.lbl_per_student_info = ctk.CTkLabel(
            tot_box,
            text="ค่าใช้จ่ายคนละ: ฿0.00 บาท (0 คน)  •  (ศูนย์บาทถ้วน)",
            font=AppTheme.FONT_LABEL,
            text_color=AppTheme.COLOR_TEXT_MUTED,
            anchor="w",
        )
        self.lbl_per_student_info.pack(anchor="w", pady=(3, 0))

        # Right: Batch Controls & Action Button
        ctrl_box = ctk.CTkFrame(btm_inner, fg_color="transparent")
        ctrl_box.pack(side="right")

        ctk.CTkLabel(
            ctrl_box,
            text="ภาคเรียน:",
            font=AppTheme.FONT_LABEL,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))

        self.opt_term = SearchableDropdown(
            ctrl_box,
            values=self.terms,
            variable=self.term_var,
            width=120,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS_INPUT,
            font=AppTheme.FONT_LABEL,
            allow_custom=True,
        )
        self.opt_term.pack(side="left", padx=(0, 12))

        ctk.CTkLabel(
            ctrl_box,
            text="ช่องทาง:",
            font=AppTheme.FONT_LABEL,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))

        self.opt_pay_method = SearchableDropdown(
            ctrl_box,
            values=["เงินสด (Cash)", "สแกน QR PromptPay", "โอนเข้าบัญชีวิทยาลัย"],
            variable=self.pay_method_var,
            width=150,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS_INPUT,
            font=AppTheme.FONT_LABEL,
        )
        self.opt_pay_method.pack(side="left", padx=(0, 14))

        self.btn_export_folder = ctk.CTkButton(
            ctrl_box,
            text=format_export_folder_button_text(self.config),
            width=155,
            height=AppTheme.CTRL_H_LG,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self._pick_export_folder,
        )
        self.btn_export_folder.pack(side="left", padx=(0, 8))

        self.btn_issue_batch = ctk.CTkButton(
            ctrl_box,
            text="ออกใบเสร็จทั้งห้อง (.pdf / .docx)",
            height=AppTheme.CTRL_H_LG,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.issue_batch_classroom_receipts,
        )
        self.btn_issue_batch.pack(side="left")

        # ------------------------------------------------------------------
        # 3. MIDDLE WORKSPACE (Left-Right Split for Full Table Height)
        # ------------------------------------------------------------------
        workspace = ctk.CTkFrame(self, fg_color="transparent")
        workspace.pack(fill="both", expand=True)
        workspace.grid_columnconfigure(0, weight=11)
        workspace.grid_columnconfigure(1, weight=9)
        workspace.grid_rowconfigure(0, weight=1)

        # ==================================================================
        # LEFT PANEL: 1. Classroom Student Roster
        # ==================================================================
        p_top = ctk.CTkFrame(
            workspace,
            corner_radius=AppTheme.RADIUS_CARD,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        p_top.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=0)

        pl_top = ctk.CTkFrame(p_top, fg_color="transparent")
        pl_top.pack(fill="x", padx=16, pady=(14, 8))

        ctk.CTkLabel(
            pl_top,
            text="1. รายชื่อนักศึกษาในห้อง",
            font=AppTheme.FONT_SECTION,
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left")

        ctk.CTkButton(
            pl_top,
            text="ล้างรายชื่อ",
            height=AppTheme.CTRL_H_SM,
            width=85,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 11, "bold"),
            command=self.clear_students,
        ).pack(side="right")

        self.lbl_student_count = ctk.CTkLabel(
            pl_top,
            text="👥 พบ 0 คน",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_PRIMARY_BLUE,
        )
        self.lbl_student_count.pack(side="right", padx=(0, 10))

        # Source mode tabs (Clean segmented button on its own row)
        seg_row = ctk.CTkFrame(p_top, fg_color="transparent")
        seg_row.pack(fill="x", padx=16, pady=(0, 8))

        self.seg_source = ctk.CTkSegmentedButton(
            seg_row,
            values=["นำเข้าไฟล์ (Excel / CSV)", "เลือกห้องที่มีในระบบ"],
            font=AppTheme.FONT_LABEL,
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_BTN,
            command=self._on_change_source_mode,
        )
        self.seg_source.pack(side="left")
        self.seg_source.set("นำเข้าไฟล์ (Excel / CSV)")

        # Dedicated source container for options
        self.f_source_container = ctk.CTkFrame(p_top, fg_color="transparent")
        self.f_source_container.pack(fill="x", padx=16, pady=(0, 8))

        # Mode A: File Import Options (Clean 2-row layout for left-right split)
        self.f_file_options = ctk.CTkFrame(self.f_source_container, fg_color="transparent")
        self.f_file_options.pack(fill="x")

        f_file_r1 = ctk.CTkFrame(self.f_file_options, fg_color="transparent")
        f_file_r1.pack(fill="x", pady=(0, 6))

        self.ent_filepath = ctk.CTkEntry(
            f_file_r1,
            placeholder_text="เลือกไฟล์รายชื่อนักศึกษา (.xlsx, .csv)...",
            font=(AppTheme.FONT_FAMILY, 12),
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS_INPUT,
        )
        self.ent_filepath.pack(side="left", fill="x", expand=True, padx=(0, 6))

        ctk.CTkButton(
            f_file_r1,
            text="เลือกไฟล์...",
            width=90,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT,
            font=AppTheme.FONT_LABEL,
            command=self.browse_student_file,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            f_file_r1,
            text="โหลดรายชื่อ",
            width=100,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=AppTheme.FONT_LABEL,
            command=self.load_students_from_file,
        ).pack(side="left")

        f_file_r2 = ctk.CTkFrame(self.f_file_options, fg_color="transparent")
        f_file_r2.pack(fill="x")

        ctk.CTkLabel(f_file_r2, text="ห้อง:", font=AppTheme.FONT_LABEL, text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 6))
        self.opt_fallback_cls = SearchableDropdown(f_file_r2, values=self.class_levels, variable=self.fallback_cls_var, width=110, height=AppTheme.CTRL_H_SM, corner_radius=AppTheme.RADIUS_INPUT, font=AppTheme.FONT_LABEL)
        self.opt_fallback_cls.pack(side="left", padx=(0, 12))

        ctk.CTkLabel(f_file_r2, text="สาขา:", font=AppTheme.FONT_LABEL, text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 6))
        self.opt_fallback_dept = SearchableDropdown(f_file_r2, values=self.departments, variable=self.fallback_dept_var, width=180, height=AppTheme.CTRL_H_SM, corner_radius=AppTheme.RADIUS_INPUT, font=AppTheme.FONT_LABEL)
        self.opt_fallback_dept.pack(side="left", fill="x", expand=True)

        # Mode B: System Room Options (Initially not packed)
        self.f_system_options = ctk.CTkFrame(self.f_source_container, fg_color="transparent")

        ctk.CTkLabel(self.f_system_options, text="เลือกห้องเรียน:", font=AppTheme.FONT_LABEL, text_color=AppTheme.COLOR_TEXT_SECONDARY).pack(side="left", padx=(0, 6))
        self.opt_system_room = SearchableDropdown(self.f_system_options, values=self.class_levels, variable=self.system_room_var, width=200, height=AppTheme.CTRL_H_MD, corner_radius=AppTheme.RADIUS_INPUT, font=AppTheme.FONT_LABEL)
        self.opt_system_room.pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            self.f_system_options,
            text="ดึงรายชื่อห้องนี้",
            width=125,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=AppTheme.FONT_LABEL,
            command=self.load_students_from_system,
        ).pack(side="left")

        # Student Table
        stu_columns = [
            ("no", "#", 40, "center"),
            ("code", "รหัสนักศึกษา", 115, "center"),
            ("name", "ชื่อ-นามสกุล", 200, "w"),
            ("class", "ห้องเรียน", 95, "center"),
            ("major", "สาขาวิชา", 170, "w"),
        ]
        self.tbl_students = DataTable(
            p_top,
            columns=stu_columns,
            height=12,
            show_pagination=False,
            border_width=0,
            corner_radius=AppTheme.RADIUS_CARD,
        )
        self.tbl_students.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        # ==================================================================
        # RIGHT PANEL: 2. Uniform Classroom Fee Package
        # ==================================================================
        p_bottom = ctk.CTkFrame(
            workspace,
            corner_radius=AppTheme.RADIUS_CARD,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        p_bottom.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=0)

        pr_top = ctk.CTkFrame(p_bottom, fg_color="transparent")
        pr_top.pack(fill="x", padx=16, pady=(14, 8))

        ctk.CTkLabel(
            pr_top,
            text="2. ค่าใช้จ่ายเดียวกันทั้งห้อง",
            font=AppTheme.FONT_SECTION,
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left")

        ctk.CTkButton(
            pr_top,
            text="ล้างทั้งหมด",
            height=AppTheme.CTRL_H_SM,
            width=80,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_DANGER,
            font=(AppTheme.FONT_FAMILY, 11, "bold"),
            command=self.clear_all_fees,
        ).pack(side="right")

        ctk.CTkButton(
            pr_top,
            text="ลบรายการ",
            height=AppTheme.CTRL_H_SM,
            width=80,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 11, "bold"),
            command=self.remove_selected_fee,
        ).pack(side="right", padx=(0, 6))

        self.lbl_fee_count = ctk.CTkLabel(
            pr_top,
            text="0 รายการ",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_PRIMARY_BLUE,
        )
        self.lbl_fee_count.pack(side="right", padx=(0, 10))

        # Row 1: Curriculum & Preset Button
        f_line1 = ctk.CTkFrame(p_bottom, fg_color="transparent")
        f_line1.pack(fill="x", padx=16, pady=(0, 6))

        ctk.CTkLabel(
            f_line1,
            text="หลักสูตร:",
            font=AppTheme.FONT_LABEL,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 6))

        self.opt_curriculum = SearchableDropdown(
            f_line1,
            values=self.curricula,
            variable=self.curriculum_var,
            width=210,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS_INPUT,
            font=AppTheme.FONT_LABEL,
            command=self.on_change_curriculum,
        )
        self.opt_curriculum.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkButton(
            f_line1,
            text="ดึงค่าเทอมมาตรฐาน",
            height=AppTheme.CTRL_H_MD,
            width=140,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=AppTheme.FONT_LABEL,
            command=self.load_preset_curriculum_fees,
        ).pack(side="left")

        # Row 2: Fee Item Selector, Price, Qty, and Add Button
        f_line2 = ctk.CTkFrame(p_bottom, fg_color="transparent")
        f_line2.pack(fill="x", padx=16, pady=(0, 8))

        self.opt_fee_item = SearchableDropdown(
            f_line2,
            values=["เลือกรายการ..."],
            variable=self.fee_item_var,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS_INPUT,
            font=AppTheme.FONT_LABEL,
            command=self.on_select_fee_item,
        )
        self.opt_fee_item.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.ent_price = ctk.CTkEntry(
            f_line2,
            textvariable=self.unit_price_var,
            width=78,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS_INPUT,
            font=AppTheme.FONT_LABEL,
            justify="right",
        )
        self.ent_price.pack(side="left", padx=(0, 6))

        self.ent_qty = ctk.CTkEntry(
            f_line2,
            textvariable=self.qty_var,
            width=46,
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS_INPUT,
            font=AppTheme.FONT_LABEL,
            justify="center",
        )
        self.ent_qty.pack(side="left", padx=(0, 6))

        self.btn_add_fee = ctk.CTkButton(
            f_line2,
            text="เพิ่ม",
            height=AppTheme.CTRL_H_MD,
            width=70,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=AppTheme.FONT_LABEL,
            command=self.add_fee_item,
        )
        self.btn_add_fee.pack(side="left")

        # Fee Items Table
        fee_columns = [
            ("no", "#", 40, "center"),
            ("item_name", "รายการค่าธรรมเนียม", 230, "w"),
            ("qty", "จำนวน", 60, "center"),
            ("price", "ราคา/หน่วย", 105, "e"),
            ("total", "รวมคนละ", 110, "e"),
        ]
        self.tbl_fees = DataTable(
            p_bottom,
            columns=fee_columns,
            height=12,
            show_pagination=False,
            border_width=0,
            corner_radius=AppTheme.RADIUS_CARD,
        )
        self.tbl_fees.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        # Initialize fee catalog dropdown and any pre-loaded cart
        self.on_change_curriculum(self.curriculum_var.get())
        self._last_data_version = getattr(self.db, "data_version", -1)
        if initial_cart:
            self.classroom_fees = list(initial_cart)
            self.refresh_fee_table()

    def on_show(self):
        """Refresh dropdown options and fee catalog only if DB data_version changed."""
        current_version = getattr(self.db, "data_version", None)
        if current_version is not None and current_version == getattr(self, "_last_data_version", -1):
            return
        self.curricula = self.db.get_options("CURRICULUM") or ["หลักสูตรในระบบ-ปวช"]
        self.faculties = self.db.get_options("FACULTY") or ["บริหารธุรกิจ / พาณิชยกรรม"]
        self.departments = self.db.get_options("DEPARTMENT") or ["การบัญชี"]
        self.class_levels = self.db.get_options("CLASS_LEVEL") or ["ปวช.1/1"]
        self.terms = self.db.get_options("TERM") or ["1/2569"]
        self.opt_term.configure(values=self.terms)
        self.opt_fallback_cls.configure(values=self.class_levels)
        self.opt_fallback_dept.configure(values=self.departments)
        self.opt_system_room.configure(values=self.class_levels)
        self.opt_curriculum.configure(values=self.curricula)
        curr = self.curriculum_var.get()
        if curr not in self.curricula and self.curricula:
            curr = self.curricula[0]
            self.curriculum_var.set(curr)
        self.on_change_curriculum(curr)
        self._last_data_version = getattr(self.db, "data_version", -1)

    # ------------------------------------------------------------------
    # EVENT HANDLERS & LOGIC
    # ------------------------------------------------------------------
    def _on_change_source_mode(self, mode_text: str):
        if "ไฟล์" in mode_text:
            self.f_system_options.pack_forget()
            self.f_file_options.pack(fill="x")
        else:
            self.f_file_options.pack_forget()
            self.f_system_options.pack(fill="x")

    def on_change_curriculum(self, curriculum: str):
        """Update fee catalog dropdown options based on selected curriculum."""
        self.current_catalog = self.db.get_fee_catalog(curriculum_filter=curriculum, active_only=True)
        self.catalog_map = {item["fee_name"]: item for item in self.current_catalog}
        names = list(self.catalog_map.keys())
        if not names:
            names = ["ไม่มีรายการในหลักสูตรนี้"]
            self.fee_item_var.set(names[0])
            self.unit_price_var.set("0.00")
        else:
            self.fee_item_var.set(names[0])
            self.on_select_fee_item(names[0])
        self.opt_fee_item.configure(values=names)

    def on_select_fee_item(self, fee_name: str):
        item = self.catalog_map.get(fee_name)
        if item:
            self.unit_price_var.set(f"{float(item.get('default_price', 0.0)):.2f}")
            self.qty_var.set("1")

    def add_fee_item(self):
        name = self.fee_item_var.get().strip()
        if not name or name == "ไม่มีรายการในหลักสูตรนี้":
            highlight_missing_field(self.opt_fee_item, "กรุณาเลือกรายการค่าธรรมเนียม", show_toast_msg=True)
            return
        try:
            price = max(0.0, float(self.unit_price_var.get().replace(",", "").strip()))
            qty = max(0.5, float(self.qty_var.get().replace(",", "").strip()))
        except ValueError:
            highlight_missing_field(self.ent_price, "กรุณาระบุราคาและจำนวนให้ถูกต้อง", show_toast_msg=True)
            return

        cat_item = self.catalog_map.get(name, {})
        fee_code = cat_item.get("fee_code", "CUSTOM")
        curriculum = self.curriculum_var.get()
        line_tot = round(qty * price, 2)

        self.classroom_fees.append({
            "fee_code": fee_code,
            "curriculum_name": curriculum,
            "item_name": name,
            "quantity": qty,
            "unit_price": price,
            "discount": 0.0,
            "line_total": line_tot,
        })
        self.refresh_fee_table()

    def load_preset_curriculum_fees(self):
        """Load all active fee items in the selected curriculum into the classroom package."""
        curriculum = self.curriculum_var.get()
        catalog_items = self.db.get_fee_catalog(curriculum_filter=curriculum, active_only=True)
        if not catalog_items:
            messagebox.showinfo("แจ้งเตือน", f"ไม่พบรายการค่าธรรมเนียมสำหรับหลักสูตร '{curriculum}'", parent=self.winfo_toplevel())
            return

        for item in catalog_items:
            price = float(item.get("default_price", 0.0))
            if price > 0:
                self.classroom_fees.append({
                    "fee_code": item.get("fee_code", "CUSTOM"),
                    "curriculum_name": curriculum,
                    "item_name": item["fee_name"],
                    "quantity": 1.0,
                    "unit_price": price,
                    "discount": 0.0,
                    "line_total": price,
                })
        self.refresh_fee_table()

    def remove_selected_fee(self):
        selected = self.tbl_fees.get_selected_values()
        if not selected:
            messagebox.showinfo("แจ้งเตือน", "กรุณาคลิกเลือกแถวรายการค่าธรรมเนียมที่ต้องการลบในตาราง", parent=self.winfo_toplevel())
            return
        try:
            idx = int(selected[0]) - 1
            if 0 <= idx < len(self.classroom_fees):
                self.classroom_fees.pop(idx)
                self.refresh_fee_table()
        except Exception:
            pass

    def clear_all_fees(self):
        self.classroom_fees.clear()
        self.refresh_fee_table()

    def refresh_fee_table(self):
        table_rows = []
        for i, item in enumerate(self.classroom_fees, 1):
            table_rows.append((
                [
                    str(i),
                    item["item_name"],
                    f"{float(item['quantity']):g}",
                    f"฿{float(item['unit_price']):,.2f}",
                    f"฿{float(item['line_total']):,.2f}",
                ],
                str(i),
            ))
        self.tbl_fees.set_rows(table_rows)
        self.lbl_fee_count.configure(text=f"{len(self.classroom_fees)} รายการ")
        self.recalculate_totals()

    # ------------------------------------------------------------------
    # STUDENT ROSTER LOADING
    # ------------------------------------------------------------------
    def browse_student_file(self):
        file_path = filedialog.askopenfilename(
            parent=self.winfo_toplevel(),
            title="เลือกไฟล์รายชื่อนักศึกษา (Excel .xlsx หรือ .csv)",
            filetypes=[("Excel / CSV Files", "*.xlsx *.csv"), ("All Files", "*.*")],
        )
        if file_path:
            self.selected_file_path = file_path
            self.ent_filepath.delete(0, "end")
            self.ent_filepath.insert(0, file_path)

    def load_students_from_file(self, show_prompt: bool = True):
        path = self.ent_filepath.get().strip()
        if not path or not os.path.exists(path):
            if show_prompt:
                messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกไฟล์รายชื่อนักศึกษา (.xlsx หรือ .csv) ก่อนครับ", parent=self.winfo_toplevel())
            return

        try:
            res = self.db.bulk_import_students(
                file_path=path,
                actor=self.session.username,
                role=self.session.role,
                default_faculty=self.fallback_fac_var.get(),
                default_major=self.fallback_dept_var.get(),
                default_class_group=self.fallback_cls_var.get(),
            )
            codes = res.get("student_codes", [])
            if not codes:
                if show_prompt:
                    messagebox.showwarning("ไม่พบรายชื่อ", "ไม่พบข้อมูลรายชื่อนักศึกษาในไฟล์ที่เลือก", parent=self.winfo_toplevel())
                return

            self.classroom_students = self.db.get_students_by_codes(codes)
            self.refresh_student_table()
            if show_prompt:
                messagebox.showinfo(
                    "นำเข้าสำเร็จ",
                    f"นำเข้ารายชื่อนักศึกษาสำเร็จ {len(self.classroom_students)} คน จากไฟล์:\n{os.path.basename(path)}",
                    parent=self.winfo_toplevel(),
                )
        except Exception as e:
            if show_prompt:
                messagebox.showerror("เกิดข้อผิดพลาดในการนำเข้า", f"ไม่สามารถนำเข้าไฟล์ได้:\n{str(e)}", parent=self.winfo_toplevel())

    def load_students_from_system(self):
        room = self.system_room_var.get().strip()
        students = self.db.get_students(level_filter=room)
        if not students:
            messagebox.showwarning("ไม่มีรายชื่อในห้องนี้", f"ไม่พบรายชื่อนักศึกษาในห้อง '{room}' ในระบบ", parent=self.winfo_toplevel())
            return
        self.classroom_students = students
        self.refresh_student_table()

    def clear_students(self):
        self.classroom_students.clear()
        self.refresh_student_table()

    def refresh_student_table(self):
        table_rows = []
        for i, s in enumerate(self.classroom_students, 1):
            table_rows.append((
                [
                    str(i),
                    s.get("student_code", ""),
                    s.get("full_name", ""),
                    s.get("class_group", "") or s.get("education_level", "ปวช.1"),
                    s.get("major", "") or s.get("faculty", ""),
                ],
                s.get("student_code", ""),
            ))
        self.tbl_students.set_rows(table_rows)
        self.lbl_student_count.configure(text=f"👥 พบ {len(self.classroom_students)} คน")
        self.recalculate_totals()

    def recalculate_totals(self):
        net_per_student = sum(float(it.get("line_total", 0.0)) for it in self.classroom_fees)
        stu_cnt = len(self.classroom_students)
        grand_total = round(net_per_student * stu_cnt, 2)
        baht_text = thai_baht_text(net_per_student)

        self.lbl_grand_total.configure(text=f"ยอดรวมทั้งห้อง: ฿{grand_total:,.2f} บาท")
        self.lbl_per_student_info.configure(
            text=f"ค่าใช้จ่ายคนละ: ฿{net_per_student:,.2f} บาท ({stu_cnt} คน)  •  ({baht_text})"
        )

    # ------------------------------------------------------------------
    # BATCH RECEIPT ISSUING
    # ------------------------------------------------------------------
    def _pick_export_folder(self):
        """Lets the cashier pick the default folder for saving exported receipts."""
        choose_export_folder(
            self.winfo_toplevel(),
            self.config,
            on_updated=lambda _: self.btn_export_folder.configure(
                text=format_export_folder_button_text(self.config)
            ),
        )

    def issue_batch_classroom_receipts(self):
        """Atomically issue identical receipts for the whole classroom and generate documents."""
        if not self.classroom_students:
            messagebox.showwarning("แจ้งเตือน", "กรุณานำเข้าหรือเลือกรายชื่อนักศึกษาในห้องก่อนครับ (ขั้นตอนที่ 1)", parent=self.winfo_toplevel())
            return
        if not self.classroom_fees:
            messagebox.showwarning("แจ้งเตือน", "กรุณาเพิ่มรายการค่าใช้จ่ายเดียวกันทั้งห้องอย่างน้อย 1 รายการ (ขั้นตอนที่ 2)", parent=self.winfo_toplevel())
            return

        stu_cnt = len(self.classroom_students)
        net_per_student = sum(float(it.get("line_total", 0.0)) for it in self.classroom_fees)
        grand_total = round(net_per_student * stu_cnt, 2)

        confirm = messagebox.askyesno(
            "ยืนยันการออกใบเสร็จทั้งห้อง",
            f"ยืนยันการออกใบเสร็จยกห้องสำหรับนักศึกษา {stu_cnt} คน\n"
            f"ยอดค่าใช้จ่ายคนละ: ฿{net_per_student:,.2f} บาท\n"
            f"ยอดเงินรวมทั้งสิ้น: ฿{grand_total:,.2f} บาท\n\n"
            "ต้องการดำเนินการและสร้างเอกสารทันทีหรือไม่?",
            parent=self.winfo_toplevel(),
        )
        if not confirm:
            return

        out_dir = self.config.get("default_export_folder", os.getcwd())
        os.makedirs(out_dir, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_file = f"BatchReceipts_{stu_cnt}students_{stamp}.pdf"

        save_path = filedialog.asksaveasfilename(
            parent=self.winfo_toplevel(),
            title="เลือกตำแหน่งบันทึกไฟล์ใบเสร็จยกห้อง (.pdf / .docx)",
            initialdir=out_dir,
            initialfile=default_file,
            defaultextension=".pdf",
            filetypes=[("PDF Document (*.pdf)", "*.pdf"), ("Word Document (*.docx)", "*.docx")],
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

        try:
            chosen_term = self.term_var.get().strip() or "1/2569"
            if chosen_term not in self.db.get_options("TERM"):
                try:
                    self.db.add_option("TERM", chosen_term, actor=self.session.username, role=self.session.role)
                except Exception:
                    pass

            issued_list = self.db.issue_batch_class_receipts(
                students_list=self.classroom_students,
                line_items=self.classroom_fees,
                payment_method=self.pay_method_var.get(),
                actor=self.session.username,
                cashier_name=self.session.full_name,
                role=self.session.role,
                term_code=chosen_term,
                book_no=self.config.get("receipt_book_no", "7006"),
                cashier_position=self.config.get("cashier_position", "ครู"),
                receipt_prefix=self.config.get("receipt_prefix", ""),
                start_no=self.config.get("receipt_no_start", "700605433"),
                end_no=self.config.get("receipt_no_end", "700999999"),
            )

            # Export multi-page receipt document (PDF or Word)
            college_name = self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")
            if save_path.lower().endswith(".docx"):
                exporter = DocxReceiptExporter(save_path, college_name=college_name)
            else:
                exporter = PdfReceiptExporter(save_path, college_name=college_name)
            exporter.export_batch(issued_list, metadata=self.config)

            # Export companion summary Excel
            excel_summary_path = os.path.splitext(save_path)[0] + "_Summary.xlsx"
            self._export_batch_summary_excel(excel_summary_path, issued_list, grand_total)

            first_no = issued_list[0]["receipt_no"]
            last_no = issued_list[-1]["receipt_no"]

            # Auto-advance config start_no so the next batch or single receipt continues seamlessly
            used_pfx, used_num = split_receipt_prefix_and_number(last_no)
            next_num = format_next_receipt_number(used_num)
            self.config["receipt_prefix"] = used_pfx
            self.config["receipt_no_start"] = next_num
            save_config(self.config)

            open_ok = open_document_file(save_path, parent=self.winfo_toplevel())

            messagebox.showinfo(
                "ออกใบเสร็จทั้งห้องสำเร็จ",
                f"ออกใบเสร็จรับเงินยกห้องสำเร็จเรียบร้อยแล้ว!\n\n"
                f"• จำนวนนักศึกษา: {len(issued_list)} คน\n"
                f"• เลขที่ใบเสร็จ: {first_no} ถึง {last_no}\n"
                f"• ยอดเงินรวมทั้งสิ้น: ฿{grand_total:,.2f} บาท\n\n"
                f"บันทึกไฟล์ใบเสร็จเรียบร้อยที่:\n{save_path}\n"
                f"และไฟล์สรุป Excel ที่:\n{excel_summary_path}\n\n"
                + ("(ระบบได้เปิดไฟล์เอกสารขึ้นมาให้พร้อมพิมพ์เรียบร้อยแล้ว)" if open_ok else ""),
                parent=self.winfo_toplevel(),
            )
        except Exception as e:
            messagebox.showerror(
                "เกิดข้อผิดพลาดในการออกใบเสร็จ",
                f"ไม่สามารถออกใบเสร็จยกห้องได้:\n{str(e)}",
                parent=self.winfo_toplevel(),
            )

    def _export_batch_summary_excel(self, save_path: str, issued_list: List[Dict[str, Any]], grand_total: float):
        if openpyxl is None:
            return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "สรุปใบเสร็จยกห้อง"

            font_title = Font(name="Segoe UI", size=14, bold=True, color="0F172A")
            font_sub = Font(name="Segoe UI", size=10, color="475569")
            font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            font_data = Font(name="Segoe UI", size=10, color="0F172A")
            fill_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")

            ws.cell(row=1, column=1, value=f"{self.config.get('college_name', 'วิทยาลัยอาชีวศึกษาสุพรรณบุรี')} - รายงานสรุปการออกใบเสร็จรับเงินยกห้อง").font = font_title
            ws.cell(
                row=2,
                column=1,
                value=f"ภาค/ปีการศึกษา: {self.term_var.get()}  |  จำนวน: {len(issued_list)} ฉบับ  |  ยอดรวม: ฿{grand_total:,.2f} บาท  |  ผู้รับเงิน: {self.session.full_name}",
            ).font = font_sub

            headers = ["#", "เลขที่ใบเสร็จ", "วันที่ออก", "รหัสนักศึกษา", "ชื่อ-นามสกุล", "ระดับชั้น", "ห้อง", "แผนกวิชา", "ยอดเงิน (บาท)", "ช่องทาง", "สถานะ"]
            for col_i, h in enumerate(headers, 1):
                c = ws.cell(row=4, column=col_i, value=h)
                c.font = font_header
                c.fill = fill_header
                c.alignment = Alignment(horizontal="center", vertical="center")

            for r_i, r in enumerate(issued_list, 1):
                row_idx = 4 + r_i
                vals = [
                    r_i,
                    r.get("receipt_no", ""),
                    format_receipt_date_thai(r.get("receipt_date", "")),
                    r.get("student_code", ""),
                    r.get("snapshot_student_name", ""),
                    r.get("snapshot_education_level", ""),
                    r.get("snapshot_class_group", ""),
                    r.get("snapshot_major", "") or r.get("snapshot_department", ""),
                    float(r.get("net_total_amount", 0.0)),
                    r.get("payment_method", ""),
                    "ปกติ",
                ]
                for col_i, val in enumerate(vals, 1):
                    cell = ws.cell(row=row_idx, column=col_i, value=val)
                    cell.font = font_data
                    if col_i == 9:
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif col_i in (1, 2, 3, 4, 6, 7, 10, 11):
                        cell.alignment = Alignment(horizontal="center", vertical="center")

            wb.save(save_path)
        except Exception as e:
            messagebox.showwarning("แจ้งเตือน", f"สร้างไฟล์สรุป Excel ไม่สำเร็จ:\n{e}", parent=self.winfo_toplevel())

    def download_excel_template(self):
        """Generate and save an example student Excel template file."""
        if openpyxl is None:
            messagebox.showwarning("แจ้งเตือน", "ต้องการไลบรารี openpyxl สำหรับสร้างไฟล์ Excel", parent=self.winfo_toplevel())
            return

        default_name = "เทมเพลตรายชื่อนักศึกษา_สำหรับจัดใบเสร็จยกห้อง.xlsx"
        save_path = filedialog.asksaveasfilename(
            parent=self.winfo_toplevel(),
            title="บันทึกไฟล์ตัวอย่างรายชื่อนักศึกษา (Excel)",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Files (*.xlsx)", "*.xlsx")],
        )
        if not save_path:
            messagebox.showinfo("ยกเลิกการบันทึก", "คุณได้ยกเลิกการเลือกตำแหน่งบันทึกไฟล์", parent=self.winfo_toplevel())
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "รายชื่อนักศึกษา"

            headers = [
                "รหัสนักศึกษา", "ชื่อ-นามสกุล", "เลขบัตรประชาชน", "คณะ/หมวดวิชา",
                "สาขาวิชา/แผนก", "ชั้นปี", "ประเภททุน", "ค่าเทอมประเมิน",
                "เบอร์โทรศัพท์", "ระดับชั้น", "กลุ่มเรียน/ห้อง"
            ]
            ws.append(headers)

            sample_rows = [
                ["692020101", "นายธนภัทร ศรีสุพรรณ", "1729900123456", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี", 1, "ปกติ (Regular)", 1500.0, "081-234-5678", "ปวช.1", "ปวช.1/1"],
                ["692020102", "นางสาวพิมชนก วงศ์สุวรรณ", "1729900234567", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี", 1, "ปกติ (Regular)", 1500.0, "082-345-6789", "ปวช.1", "ปวช.1/1"],
                ["692020103", "นางสาวกมลชนก สุพรรณภูมิ", "1729900345678", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี", 1, "ปกติ (Regular)", 1500.0, "083-456-7890", "ปวช.1", "ปวช.1/1"],
                ["692020104", "นายกิตติพงศ์ เจริญยิ่ง", "1729900456789", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี", 1, "ปกติ (Regular)", 1500.0, "084-567-8901", "ปวช.1", "ปวช.1/1"],
                ["692020105", "นางสาวชนากานต์ อัศวเดช", "1729900567890", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี", 1, "ปกติ (Regular)", 1500.0, "085-678-9012", "ปวช.1", "ปวช.1/1"],
            ]
            for row in sample_rows:
                ws.append(row)

            # Styling header
            fill_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
            font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            for col_i in range(1, len(headers) + 1):
                c = ws.cell(row=1, column=col_i)
                c.fill = fill_header
                c.font = font_header
                c.alignment = Alignment(horizontal="center", vertical="center")

            col_widths = [16, 26, 18, 26, 18, 10, 16, 16, 16, 12, 16]
            for col_i, w in enumerate(col_widths, 1):
                col_letter = openpyxl.utils.get_column_letter(col_i)
                ws.column_dimensions[col_letter].width = w

            wb.save(save_path)
            open_ok = open_document_file(save_path, parent=self.winfo_toplevel())
            messagebox.showinfo(
                "บันทึกเทมเพลตสำเร็จ",
                f"บันทึกไฟล์ตัวอย่างเรียบร้อยแล้ว:\n{save_path}\n\n"
                + ("(ระบบได้เปิดไฟล์เอกสารขึ้นมาเรียบร้อยแล้ว)" if open_ok else ""),
                parent=self.winfo_toplevel(),
            )
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถสร้างเทมเพลตได้: {str(e)}", parent=self.winfo_toplevel())
