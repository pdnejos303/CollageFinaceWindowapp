import datetime
import os
from tkinter import filedialog, messagebox
from typing import Any, Callable, Dict, Optional
import customtkinter as ctk

from core.config import format_receipt_date_thai
from core.database import DatabaseManager
from core.models import UserSession
from exporters.excel_exporter import ExcelFinanceExporter
from ui.components import DataTable, SearchableDropdown, SectionHeader, StatCard, open_document_file
from ui.theme import AppTheme


class DashboardView(ctk.CTkFrame):
    """
    SPVC Executive Financial & Operations Dashboard.
    Clean, Flat Rectangular Design (No Nested Clutter, Sharp Square Corners).
    Supports Today, Month, and All-Time revenue monitoring with dynamic Level and Department filters.
    """

    PERIOD_MAP = {
        "ทั้งหมด (All Time)": "ALL",
        "วันนี้ (Today)": "TODAY",
        "7 วันล่าสุด (Last 7 Days)": "THIS_WEEK",
        "เดือนนี้ (This Month)": "THIS_MONTH",
    }

    def __init__(
        self,
        parent,
        db: DatabaseManager,
        session: UserSession,
        config: Dict[str, Any],
        on_navigate: Optional[Callable[[str], None]] = None,
    ):
        super().__init__(parent, fg_color="transparent")
        self.db = db
        self.session = session
        self.config = config
        self.on_navigate = on_navigate
        self.analytics_data: Dict[str, Any] = {}
        self._last_data_version: int = -1

        # Master filter options from database
        filter_meta = self.db.get_dashboard_filter_metadata()
        master_terms = filter_meta.get("terms", []) or ["1/2569", "2/2569"]
        master_levels = filter_meta.get("levels", []) or ["ปวช.1", "ปวช.2", "ปวช.3", "ปวส.1", "ปวส.2"]
        master_depts = filter_meta.get("departments", []) or []

        self.terms = ["ทั้งหมด"] + [t for t in master_terms if t != "ทั้งหมด"]
        self.levels = ["ทั้งหมด"] + [l for l in master_levels if l != "ทั้งหมด"]
        self.departments = ["ทั้งหมด"] + [d for d in master_depts if d != "ทั้งหมด"]
        self.periods = list(self.PERIOD_MAP.keys())

        curr_term = self.config.get("current_semester", "1/2569")
        init_term = curr_term if curr_term in self.terms else "ทั้งหมด"

        self.selected_term_var = ctk.StringVar(value=init_term)
        self.selected_level_var = ctk.StringVar(value="ทั้งหมด")
        self.selected_dept_var = ctk.StringVar(value="ทั้งหมด")
        self.selected_period_var = ctk.StringVar(value="ทั้งหมด (All Time)")

        # ------------------------------------------------------------------
        # 0. SCROLLABLE BODY AREA (All dashboard content scrolls together)
        # ------------------------------------------------------------------
        self.scroll_body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_body.pack(fill="both", expand=True)

        # ------------------------------------------------------------------
        # 1. TOP HEADER: Single Flat Header Bar
        # ------------------------------------------------------------------
        header = SectionHeader(
            self.scroll_body,
            "แดชบอร์ดภาพรวมการเงิน",
            "วิทยาลัยอาชีวศึกษาสุพรรณบุรี • ติดตามยอดวันนี้ ยอดรายเดือน ยอดรวม และวิเคราะห์ตามชั้นปี/แผนกวิชา",
            corner_radius=AppTheme.RADIUS_CARD,
        )
        header.pack(fill="x", pady=(0, 12))
        self.header = header

        header.add_action_button(
            "รีเฟรชข้อมูล",
            self.refresh_dashboard,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            corner_radius=AppTheme.RADIUS_BTN,
        )
        header.add_action_button(
            "ส่งออก Excel",
            self.export_excel,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            corner_radius=AppTheme.RADIUS_BTN,
        )

        # ------------------------------------------------------------------
        # 2. FILTER BAR: Clean, Airy Single-Row Filter Scope
        # ------------------------------------------------------------------
        filter_bar = ctk.CTkFrame(
            self.scroll_body,
            corner_radius=AppTheme.RADIUS_CARD,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        filter_bar.pack(fill="x", pady=(0, 12))
        self.filter_bar = filter_bar

        filters_grid = ctk.CTkFrame(filter_bar, fg_color="transparent")
        filters_grid.pack(fill="x", padx=16, pady=(12, 6))

        filters_grid.grid_columnconfigure(0, weight=2)  # Term (ภาคเรียน)
        filters_grid.grid_columnconfigure(1, weight=2)  # Level (ระดับชั้น)
        filters_grid.grid_columnconfigure(2, weight=3)  # Department (แผนกวิชา)
        filters_grid.grid_columnconfigure(3, weight=2)  # Period (ช่วงเวลา)
        filters_grid.grid_columnconfigure(4, weight=1)  # Action (ล้างตัวกรอง)

        # Col 0: Term
        ctk.CTkLabel(
            filters_grid,
            text="ภาค/ปีการศึกษา",
            font=AppTheme.FONT_CAPTION,
            text_color=AppTheme.COLOR_TEXT_MUTED,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=4, pady=(0, 3))

        self.opt_term = SearchableDropdown(
            filters_grid,
            values=self.terms,
            variable=self.selected_term_var,
            width=135,
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_INPUT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            dropdown_font=(AppTheme.FONT_FAMILY, 12),
            command=lambda _: self.refresh_dashboard(),
        )
        self.opt_term.grid(row=1, column=0, sticky="ew", padx=4)

        # Col 1: Level
        ctk.CTkLabel(
            filters_grid,
            text="ระดับชั้น / ชั้นปี",
            font=AppTheme.FONT_CAPTION,
            text_color=AppTheme.COLOR_TEXT_MUTED,
            anchor="w",
        ).grid(row=0, column=1, sticky="w", padx=4, pady=(0, 3))

        self.opt_level = SearchableDropdown(
            filters_grid,
            values=self.levels,
            variable=self.selected_level_var,
            width=90,
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_INPUT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            dropdown_font=(AppTheme.FONT_FAMILY, 12),
            command=lambda _: self.refresh_dashboard(),
        )
        self.opt_level.grid(row=1, column=1, sticky="ew", padx=4)

        # Col 2: Department
        ctk.CTkLabel(
            filters_grid,
            text="แผนกวิชา / สาขา",
            font=AppTheme.FONT_CAPTION,
            text_color=AppTheme.COLOR_TEXT_MUTED,
            anchor="w",
        ).grid(row=0, column=2, sticky="w", padx=4, pady=(0, 3))

        self.opt_dept = SearchableDropdown(
            filters_grid,
            values=self.departments,
            variable=self.selected_dept_var,
            width=120,
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_INPUT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            dropdown_font=(AppTheme.FONT_FAMILY, 12),
            command=lambda _: self.refresh_dashboard(),
        )
        self.opt_dept.grid(row=1, column=2, sticky="ew", padx=4)

        # Col 3: Period
        ctk.CTkLabel(
            filters_grid,
            text="ช่วงเวลาที่บันทึก",
            font=AppTheme.FONT_CAPTION,
            text_color=AppTheme.COLOR_TEXT_MUTED,
            anchor="w",
        ).grid(row=0, column=3, sticky="w", padx=4, pady=(0, 3))

        self.opt_period = SearchableDropdown(
            filters_grid,
            values=self.periods,
            variable=self.selected_period_var,
            width=110,
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_INPUT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            dropdown_font=(AppTheme.FONT_FAMILY, 12),
            command=lambda _: self.refresh_dashboard(),
        )
        self.opt_period.grid(row=1, column=3, sticky="ew", padx=4)

        # Col 4: Action / Reset
        self.btn_reset_filters = ctk.CTkButton(
            filters_grid,
            text="ล้างตัวกรอง",
            width=110,
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.reset_filters,
        )
        self.btn_reset_filters.grid(row=1, column=4, sticky="ew", padx=(4, 0))

        # Compact Muted Scope Caption
        self.lbl_filter_scope = ctk.CTkLabel(
            filter_bar,
            text="กำลังแสดง: ทุกภาคเรียน • ทุกระดับชั้น • ทุกแผนก • ตลอดเวลา",
            font=AppTheme.FONT_CAPTION,
            text_color=AppTheme.COLOR_TEXT_MUTED,
            anchor="w",
        )
        self.lbl_filter_scope.pack(fill="x", padx=20, pady=(0, 10))

        # ------------------------------------------------------------------
        # ROW 1: 5 Executive KPI StatCards (Flat Rectangles, 1 Border)
        # 1: ยอดวันนี้ | 2: ยอดรวมรายเดือน | 3: ยอดทั้งหมด | 4: ยอดค้าง | 5: สถานะนักเรียน
        # ------------------------------------------------------------------
        kpi_row = ctk.CTkFrame(self.scroll_body, fg_color="transparent")
        kpi_row.pack(fill="x", pady=(0, 12))
        kpi_row.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        # Card 0: ยอดรับชำระวันนี้
        self.card_today_income = StatCard(
            kpi_row,
            title="ยอดรับชำระวันนี้",
            value="฿0.00",
            subtitle="0 ฉบับวันนี้",
            accent_color=AppTheme.COLOR_SUCCESS,
            corner_radius=AppTheme.RADIUS_CARD,
        )
        self.card_today_income.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        # Card 1: ยอดรวมประจำเดือนนี้
        self.card_month_income = StatCard(
            kpi_row,
            title="ยอดรวมประจำเดือนนี้",
            value="฿0.00",
            subtitle="เดือนนี้ • 0 ฉบับ",
            accent_color=AppTheme.COLOR_ACCENT,
            corner_radius=AppTheme.RADIUS_CARD,
        )
        self.card_month_income.grid(row=0, column=1, sticky="nsew", padx=6)

        # Card 2: ยอดจัดเก็บเงินรวมทั้งหมด
        self.card_total_income = StatCard(
            kpi_row,
            title="ยอดจัดเก็บเงินรวมทั้งหมด",
            value="฿0.00",
            subtitle="ปกติ 0 • ยกเลิก 0 ฉบับ",
            accent_color=AppTheme.COLOR_PRIMARY_BLUE,
            corner_radius=AppTheme.RADIUS_CARD,
        )
        self.card_total_income.grid(row=0, column=2, sticky="nsew", padx=6)

        # Card 3: ยอดค้างชำระค่าเทอม
        self.card_outstanding = StatCard(
            kpi_row,
            title="ยอดค้างชำระค่าเทอม",
            value="฿0.00",
            subtitle="จากยอดประเมิน ฿0.00",
            accent_color=AppTheme.COLOR_DANGER,
            corner_radius=AppTheme.RADIUS_CARD,
        )
        self.card_outstanding.grid(row=0, column=3, sticky="nsew", padx=6)

        # Card 4: สถานะจัดเก็บค่าเทอม
        self.card_students = StatCard(
            kpi_row,
            title="สถานะจัดเก็บค่าเทอม",
            value="0 / 0 คน",
            subtitle="อัตราจัดเก็บ 0.0%",
            accent_color=AppTheme.COLOR_SLATE,
            corner_radius=AppTheme.RADIUS_CARD,
        )
        self.card_students.grid(row=0, column=4, sticky="nsew", padx=(6, 0))

        # ------------------------------------------------------------------
        # ROW 2: 3 Visual Analytics Panels (Clean Single-Bordered Cards)
        # ------------------------------------------------------------------
        charts_row = ctk.CTkFrame(self.scroll_body, fg_color="transparent")
        charts_row.pack(fill="x", pady=(0, 12))
        charts_row.grid_columnconfigure((0, 1, 2), weight=1)

        # Panel 1: Payment Methods Breakdown
        card_pm = ctk.CTkFrame(
            charts_row,
            corner_radius=AppTheme.RADIUS_CARD,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        card_pm.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        pm_header = ctk.CTkFrame(card_pm, fg_color="transparent")
        pm_header.pack(fill="x", padx=18, pady=(14, 10))
        ctk.CTkLabel(
            pm_header,
            text="ช่องทางการชำระเงิน",
            font=AppTheme.FONT_SECTION,
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left")

        self.pm_container = ctk.CTkFrame(card_pm, fg_color="transparent")
        self.pm_container.pack(fill="both", expand=True, padx=18, pady=(0, 14))

        # Panel 2: Tuition Collection Status
        card_prog = ctk.CTkFrame(
            charts_row,
            corner_radius=AppTheme.RADIUS_CARD,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        card_prog.grid(row=0, column=1, sticky="nsew", padx=6)

        prog_header = ctk.CTkFrame(card_prog, fg_color="transparent")
        prog_header.pack(fill="x", padx=18, pady=(14, 10))
        ctk.CTkLabel(
            prog_header,
            text="ความคืบหน้าจัดเก็บค่าเทอม",
            font=AppTheme.FONT_SECTION,
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left")

        self.prog_container = ctk.CTkFrame(card_prog, fg_color="transparent")
        self.prog_container.pack(fill="both", expand=True, padx=18, pady=(0, 14))

        # Panel 3: Collections by Education Level or Department (Dynamic)
        card_lvl = ctk.CTkFrame(
            charts_row,
            corner_radius=AppTheme.RADIUS_CARD,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        card_lvl.grid(row=0, column=2, sticky="nsew", padx=(6, 0))

        lvl_header = ctk.CTkFrame(card_lvl, fg_color="transparent")
        lvl_header.pack(fill="x", padx=18, pady=(14, 10))
        self.lbl_panel3_title = ctk.CTkLabel(
            lvl_header,
            text="ยอดจัดเก็บแยกตามระดับชั้น",
            font=AppTheme.FONT_SECTION,
            text_color=AppTheme.COLOR_TEXT,
        )
        self.lbl_panel3_title.pack(side="left")

        self.lvl_container = ctk.CTkFrame(card_lvl, fg_color="transparent")
        self.lvl_container.pack(fill="both", expand=True, padx=18, pady=(0, 14))

        # ------------------------------------------------------------------
        # ROW 3: Top Fee Items & Recent Receipts Activity (Zero Nested Boxes)
        # ------------------------------------------------------------------
        row3 = ctk.CTkFrame(self.scroll_body, fg_color="transparent")
        row3.pack(fill="both", expand=True, pady=(0, 4))
        row3.grid_columnconfigure(0, weight=4)
        row3.grid_columnconfigure(1, weight=6)

        # Left: Top 5 Fee Items
        card_fees = ctk.CTkFrame(
            row3,
            corner_radius=AppTheme.RADIUS_CARD,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        card_fees.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        fees_top = ctk.CTkFrame(card_fees, fg_color="transparent")
        fees_top.pack(fill="x", padx=18, pady=(14, 10))
        ctk.CTkLabel(
            fees_top,
            text="5 อันดับค่าธรรมเนียมยอดนิยม",
            font=AppTheme.FONT_SECTION,
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left")

        self.fees_container = ctk.CTkFrame(card_fees, fg_color="transparent")
        self.fees_container.pack(fill="both", expand=True, padx=18, pady=(0, 14))

        # Right: Recent Receipts Activity Table (Single Clean Border via Card)
        card_rec = ctk.CTkFrame(
            row3,
            corner_radius=AppTheme.RADIUS_CARD,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        card_rec.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        rec_top = ctk.CTkFrame(card_rec, fg_color="transparent")
        rec_top.pack(fill="x", padx=18, pady=(14, 10))
        ctk.CTkLabel(
            rec_top,
            text="รายการใบเสร็จล่าสุด",
            font=AppTheme.FONT_SECTION,
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left")

        ctk.CTkButton(
            rec_top,
            text="ดูทั้งหมด",
            width=90,
            height=28,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=lambda: self._navigate("receipts"),
        ).pack(side="right")

        columns = [
            ("receipt_no", "เลขที่ใบเสร็จ", 110, "center"),
            ("date", "วันที่ออก", 95, "center"),
            ("student", "รหัส-ชื่อนักศึกษา", 180, "w"),
            ("level", "ระดับชั้น", 80, "center"),
            ("method", "ช่องทาง", 120, "center"),
            ("amount", "ยอดสุทธิ (บาท)", 110, "e"),
            ("status", "สถานะ", 85, "center"),
        ]
        self.tbl_recent = DataTable(
            card_rec,
            columns=columns,
            height=7,
            show_pagination=False,
            border_width=0,
            corner_radius=AppTheme.RADIUS_CARD,
        )
        self.tbl_recent.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.recent_table = self.tbl_recent

        # Initial data load
        self.refresh_dashboard()

    def _navigate(self, tab_key: str):
        if self.on_navigate:
            self.on_navigate(tab_key)

    def on_show(self):
        """Refresh dashboard data only if the underlying DB data version changed."""
        current_version = getattr(self.db, "data_version", None)
        if current_version is None or current_version != getattr(self, "_last_data_version", -1):
            self.refresh_dashboard()

    def reset_filters(self):
        """Reset all filter dropdowns back to initial/unfiltered state."""
        curr_term = self.config.get("current_semester", "1/2569")
        self.selected_term_var.set(curr_term if curr_term in self.terms else "ทั้งหมด")
        self.selected_level_var.set("ทั้งหมด")
        self.selected_dept_var.set("ทั้งหมด")
        self.selected_period_var.set("ทั้งหมด (All Time)")
        self.refresh_dashboard()

    def refresh_dashboard(self):
        """Fetch latest analytics and refresh all cards, progress meters, and table."""
        term_val = self.selected_term_var.get().strip()
        level_val = self.selected_level_var.get().strip()
        dept_val = self.selected_dept_var.get().strip()
        period_key = self.selected_period_var.get().strip()
        period_val = self.PERIOD_MAP.get(period_key, "ALL")

        # Update Scope Label Badge
        if hasattr(self, "lbl_filter_scope"):
            t_txt = f"ภาค {term_val}" if term_val != "ทั้งหมด" else "ทุกภาคเรียน"
            l_txt = level_val if level_val != "ทั้งหมด" else "ทุกระดับชั้น"
            d_short = (dept_val[:16] + "..") if len(dept_val) > 18 else dept_val
            d_txt = d_short if dept_val != "ทั้งหมด" else "ทุกแผนก"
            p_txt = period_key if period_key != "ทั้งหมด (All Time)" else "ตลอดเวลา"
            self.lbl_filter_scope.configure(
                text=f"กำลังแสดง: {t_txt} • {l_txt} • {d_txt} • {p_txt}"
            )

        self.analytics_data = self.db.get_dashboard_analytics(
            term_filter=term_val,
            period_filter=period_val,
            level_filter=level_val,
            department_filter=dept_val,
        )
        data = self.analytics_data

        # 1. Update Top 5 Stat Cards: Today, Month, Total, Outstanding, Students
        today_inc = data.get("today_income", 0.0)
        today_cnt = data.get("today_receipt_count", 0)
        self.card_today_income.set_value(
            f"฿{today_inc:,.2f}",
            subtitle=f"ออก {today_cnt:,} ฉบับในวันนี้",
        )

        month_inc = data.get("month_income", 0.0)
        month_cnt = data.get("month_receipt_count", 0)
        month_lbl = data.get("month_label_thai", "เดือนนี้")
        self.card_month_income.set_value(
            f"฿{month_inc:,.2f}",
            subtitle=f"{month_lbl} ({month_cnt:,} ฉบับ)",
        )

        tot_inc = data.get("total_income", 0.0)
        r_cnt = data.get("receipt_count", 0)
        v_cnt = data.get("void_receipt_count", 0)
        self.card_total_income.set_value(
            f"฿{tot_inc:,.2f}",
            subtitle=f"ปกติ {r_cnt:,} • ยกเลิก {v_cnt:,} ฉบับ",
        )

        ar = data.get("outstanding_receivables", 0.0)
        tf = data.get("total_tuition_expected", 0.0)
        self.card_outstanding.set_value(
            f"฿{ar:,.2f}",
            subtitle=f"จากยอดประเมิน ฿{tf:,.2f}",
        )

        stu_tot = data.get("student_count", 0)
        stu_paid = data.get("paid_student_count", 0)
        rate = data.get("collection_rate_percent", 0.0)
        self.card_students.set_value(
            f"ชำระแล้ว {stu_paid}/{stu_tot} คน",
            subtitle=f"อัตราจัดเก็บค่าเทอม {rate:.1f}%",
        )

        # 2. Render Payment Methods Breakdown (Clean Flat Progress Bars)
        for widget in self.pm_container.winfo_children():
            widget.destroy()

        methods = data.get("by_payment_method", [])
        if not methods:
            ctk.CTkLabel(
                self.pm_container,
                text="ยังไม่มีข้อมูลการรับชำระเงิน",
                font=(AppTheme.FONT_FAMILY, 12),
                text_color=AppTheme.COLOR_TEXT_MUTED,
            ).pack(pady=16)
        else:
            for m in methods:
                row_f = ctk.CTkFrame(self.pm_container, fg_color="transparent")
                row_f.pack(fill="x", pady=3)

                top_txt = ctk.CTkFrame(row_f, fg_color="transparent")
                top_txt.pack(fill="x")

                ctk.CTkLabel(
                    top_txt,
                    text=f"{m['name']} ({m['count']} ฉบับ)",
                    font=(AppTheme.FONT_FAMILY, 12, "bold"),
                    text_color=AppTheme.COLOR_TEXT_SECONDARY,
                ).pack(side="left")

                ctk.CTkLabel(
                    top_txt,
                    text=f"฿{m['amount']:,.2f} ({m['percent']:.1f}%)",
                    font=(AppTheme.FONT_FAMILY, 12, "bold"),
                    text_color=AppTheme.COLOR_PRIMARY_BLUE,
                ).pack(side="right")

                bar_val = min(1.0, max(0.0, m["percent"] / 100.0))
                pbar = ctk.CTkProgressBar(
                    row_f,
                    height=6,
                    corner_radius=AppTheme.RADIUS,
                    progress_color=AppTheme.COLOR_PRIMARY_BLUE,
                    fg_color="#E2E8F0",
                )
                pbar.pack(fill="x", pady=(2, 0))
                pbar.set(bar_val)

        # 3. Render Tuition Collection Status (Flat Metrics, No Nested Box Bordering)
        for widget in self.prog_container.winfo_children():
            widget.destroy()

        overall_pct = data.get("collection_rate_percent", 0.0)
        collected_amt = data.get("total_tuition_collected", 0.0)

        ctk.CTkLabel(
            self.prog_container,
            text=f"จัดเก็บได้ ฿{collected_amt:,.2f} / ฿{tf:,.2f}",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        ).pack(fill="x")

        prog_bar = ctk.CTkProgressBar(
            self.prog_container,
            height=10,
            corner_radius=AppTheme.RADIUS,
            progress_color=AppTheme.COLOR_SUCCESS,
            fg_color="#E2E8F0",
        )
        prog_bar.pack(fill="x", pady=(4, 10))
        prog_bar.set(min(1.0, max(0.0, overall_pct / 100.0)))

        # Clean Flat 3 Columns
        flat_status_row = ctk.CTkFrame(self.prog_container, fg_color="transparent")
        flat_status_row.pack(fill="x")
        flat_status_row.grid_columnconfigure((0, 1, 2), weight=1)

        # Paid
        col_paid = ctk.CTkFrame(flat_status_row, fg_color="#F0FDF4", corner_radius=AppTheme.RADIUS)
        col_paid.grid(row=0, column=0, sticky="nsew", padx=(0, 3), pady=2)
        ctk.CTkLabel(col_paid, text="ชำระครบ", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_SUCCESS).pack(pady=(6, 0))
        ctk.CTkLabel(col_paid, text=f"{data.get('paid_student_count', 0)} คน", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_SUCCESS).pack(pady=(0, 6))

        # Partial
        col_part = ctk.CTkFrame(flat_status_row, fg_color="#FFFBEB", corner_radius=AppTheme.RADIUS)
        col_part.grid(row=0, column=1, sticky="nsew", padx=3, pady=2)
        ctk.CTkLabel(col_part, text="ค้างบางส่วน", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_WARNING).pack(pady=(6, 0))
        ctk.CTkLabel(col_part, text=f"{data.get('partial_student_count', 0)} คน", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_WARNING).pack(pady=(0, 6))

        # Unpaid
        col_unp = ctk.CTkFrame(flat_status_row, fg_color="#FEF2F2", corner_radius=AppTheme.RADIUS)
        col_unp.grid(row=0, column=2, sticky="nsew", padx=(3, 0), pady=2)
        ctk.CTkLabel(col_unp, text="ยังไม่ชำระ", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_DANGER).pack(pady=(6, 0))
        ctk.CTkLabel(col_unp, text=f"{data.get('unpaid_student_count', 0)} คน", font=(AppTheme.FONT_FAMILY, 13, "bold"), text_color=AppTheme.COLOR_DANGER).pack(pady=(0, 6))

        # 4. Render Dynamic Breakdown (By Department if Level Filtered, else By Level)
        for widget in self.lvl_container.winfo_children():
            widget.destroy()

        if level_val != "ทั้งหมด" and dept_val == "ทั้งหมด":
            self.lbl_panel3_title.configure(text=f"ยอดจัดเก็บแยกตามแผนกวิชา ({level_val})")
            display_items = data.get("by_department", [])
            item_key = "name"
            empty_msg = f"ยังไม่มีข้อมูลแผนกวิชาสำหรับ {level_val}"
        elif dept_val != "ทั้งหมด":
            self.lbl_panel3_title.configure(text=f"ยอดจัดเก็บแยกตามระดับชั้น ({dept_val})")
            display_items = data.get("by_education_level", [])
            item_key = "level"
            empty_msg = f"ยังไม่มีข้อมูลระดับชั้นสำหรับแผนก {dept_val}"
        else:
            self.lbl_panel3_title.configure(text="ยอดจัดเก็บแยกตามระดับชั้น")
            display_items = data.get("by_education_level", [])
            item_key = "level"
            empty_msg = "ยังไม่มีข้อมูลระดับชั้น"

        if not display_items:
            ctk.CTkLabel(
                self.lvl_container,
                text=empty_msg,
                font=(AppTheme.FONT_FAMILY, 12),
                text_color=AppTheme.COLOR_TEXT_MUTED,
            ).pack(pady=16)
        else:
            for item in display_items:
                row_f = ctk.CTkFrame(self.lvl_container, fg_color="transparent")
                row_f.pack(fill="x", pady=3)

                top_txt = ctk.CTkFrame(row_f, fg_color="transparent")
                top_txt.pack(fill="x")

                ctk.CTkLabel(
                    top_txt,
                    text=f"{item[item_key]} ({item['count']} ฉบับ)",
                    font=(AppTheme.FONT_FAMILY, 12, "bold"),
                    text_color=AppTheme.COLOR_TEXT_SECONDARY,
                ).pack(side="left")

                ctk.CTkLabel(
                    top_txt,
                    text=f"฿{item['amount']:,.2f} ({item['percent']:.1f}%)",
                    font=(AppTheme.FONT_FAMILY, 12, "bold"),
                    text_color=AppTheme.COLOR_PRIMARY_BLUE,
                ).pack(side="right")

                bar_val = min(1.0, max(0.0, item["percent"] / 100.0))
                pbar = ctk.CTkProgressBar(
                    row_f,
                    height=6,
                    corner_radius=AppTheme.RADIUS,
                    progress_color="#0284C7",
                    fg_color="#E2E8F0",
                )
                pbar.pack(fill="x", pady=(2, 0))
                pbar.set(bar_val)

        # 5. Render Top 5 Fee Items (Clean Flat Rows, No Nested Frames)
        for widget in self.fees_container.winfo_children():
            widget.destroy()

        top_fees = data.get("top_fee_items", [])
        if not top_fees:
            ctk.CTkLabel(
                self.fees_container,
                text="ยังไม่มีข้อมูลค่าธรรมเนียม",
                font=(AppTheme.FONT_FAMILY, 12),
                text_color=AppTheme.COLOR_TEXT_MUTED,
            ).pack(pady=16)
        else:
            rank_colors = ["#1E40AF", "#0F766E", "#0284C7", "#475569", "#64748B"]
            for idx, fee in enumerate(top_fees, 1):
                item_row = ctk.CTkFrame(self.fees_container, fg_color="transparent")
                item_row.pack(fill="x", pady=3)

                pill_color = rank_colors[idx - 1] if idx <= len(rank_colors) else "#64748B"
                ctk.CTkLabel(
                    item_row,
                    text=str(idx),
                    width=20,
                    height=20,
                    corner_radius=AppTheme.RADIUS,
                    fg_color=pill_color,
                    text_color="#FFFFFF",
                    font=(AppTheme.FONT_FAMILY, 12, "bold"),
                ).pack(side="left", padx=(0, 8))

                name_lbl = ctk.CTkLabel(
                    item_row,
                    text=fee["item_name"],
                    font=(AppTheme.FONT_FAMILY, 12),
                    text_color=AppTheme.COLOR_TEXT,
                    anchor="w",
                )
                name_lbl.pack(side="left", fill="x", expand=True)

                ctk.CTkLabel(
                    item_row,
                    text=f"฿{fee['total']:,.2f}",
                    font=(AppTheme.FONT_FAMILY, 12, "bold"),
                    text_color=AppTheme.COLOR_PRIMARY_BLUE,
                ).pack(side="right")

        # 6. Render Recent Receipts Table
        table_rows = []
        recent = data.get("recent_receipts", [])
        for r in recent:
            st = r.get("status", "ISSUED")
            status_text = "ปกติ" if st == "ISSUED" else "ยกเลิก"
            student_display = f"{r.get('student_code', '')} {r.get('snapshot_student_name', '')}".strip()
            date_display = format_receipt_date_thai(r.get("receipt_date", ""))

            table_rows.append((
                [
                    r.get("receipt_no", ""),
                    date_display,
                    student_display,
                    r.get("snapshot_education_level", "ปวช.1"),
                    r.get("payment_method", "เงินสด"),
                    f"฿{float(r.get('net_total_amount', 0.0)):,.2f}",
                    status_text,
                ],
                r.get("receipt_no", ""),
            ))

        self.tbl_recent.set_rows(table_rows)
        self._last_data_version = getattr(self.db, "data_version", -1)

    def export_excel(self):
        """Export executive dashboard analytics to formatted multi-sheet Excel file."""
        if not self.analytics_data:
            self.refresh_dashboard()

        default_dir = self.config.get("default_export_folder", "") or os.getcwd()
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_file = f"SPVC_Dashboard_Summary_{timestamp}.xlsx"

        save_path = filedialog.asksaveasfilename(
            parent=self.winfo_toplevel(),
            title="ส่งออกรายงานสรุปภาพรวมแดชบอร์ด (Excel)",
            initialfile=default_file,
            initialdir=default_dir,
            defaultextension=".xlsx",
            filetypes=[("Excel Files (*.xlsx)", "*.xlsx"), ("All Files (*.*)", "*.*")],
        )

        if not save_path:
            messagebox.showinfo("ยกเลิกการส่งออก", "คุณได้ยกเลิกการเลือกตำแหน่งบันทึกไฟล์", parent=self.winfo_toplevel())
            return

        try:
            exporter = ExcelFinanceExporter(save_path, college_name=self.config.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"))
            metadata = {
                "cashier_name": self.session.full_name,
                "exported_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            exporter.export_dashboard(self.analytics_data, metadata=metadata)
            open_ok = open_document_file(save_path, parent=self.winfo_toplevel())
            messagebox.showinfo(
                "ส่งออกรายงานสำเร็จ",
                f"บันทึกไฟล์สรุปแดชบอร์ดเรียบร้อยแล้ว:\n{save_path}\n\n"
                + ("(ระบบได้เปิดไฟล์เอกสารขึ้นมาเรียบร้อยแล้ว)" if open_ok else ""),
                parent=self.winfo_toplevel(),
            )
        except Exception as e:
            messagebox.showerror(
                "เกิดข้อผิดพลาดในการส่งออก",
                f"ไม่สามารถส่งออกไฟล์ Excel ได้:\n{str(e)}",
                parent=self.winfo_toplevel(),
            )
