import datetime
import os
import sys
import traceback
from tkinter import messagebox
from typing import Any, Callable, Dict, Optional
import customtkinter as ctk

from core.config import backup_database
from core.database import DatabaseManager
from core.models import UserSession
from exporters.webhook_exporter import FinanceWebhookDispatcher
from ui.components import SearchableDropdown
from ui.icon_manager import IconManager
from ui.theme import AppTheme
from ui.batch_class_view import BatchClassReceiptView
from ui.dashboard_view import DashboardView
from ui.receipt_gallery_view import ReceiptGalleryView
from ui.views import (
    CourseRegistrationView,
    ReceiptRegisterView,
    SettingsAndBackupView,
    SmartReceiptPOSView,
    StudentAndFeeCatalogView,
)


class MainWindow:
    """
    Minimal, Ultra-Clean Single-Root-Window for Suphanburi Vocational College (SPVC).
    Unified Corporate Navy & Slate palette with zero clashing colors.
    Modern vector PNG icons for navigation items (No cheap emojis).
    """

    def __init__(
        self,
        db: DatabaseManager,
        session: UserSession,
        config: Dict[str, Any],
        on_logout: Optional[Callable[[], None]] = None,
    ):
        self.db = db
        self.session = session
        self.config = config
        self.on_logout = on_logout
        self.webhook = FinanceWebhookDispatcher(
            webhook_url=config.get("webhook_url", ""),
            enabled=config.get("webhook_enabled", False),
            alert_threshold=config.get("webhook_alert_threshold", 50000.0),
        )
        self.nav_buttons: Dict[str, ctk.CTkButton] = {}
        self._view_cache: Dict[str, ctk.CTkFrame] = {}
        self.current_view = None
        self.current_tab_key = "dashboard"
        self.top_header = None
        self.lbl_header_title = None
        self.lbl_profile_avatar = None
        self.lbl_user_name = None
        self.lbl_user_role = None
        self.btn_logout = None
        self.sidebar_is_collapsed = False
        self.SIDEBAR_EXPANDED_WIDTH = 250
        self.SIDEBAR_COLLAPSED_WIDTH = 72

        self.nav_items_config = [
            ("dashboard", "ภาพรวมการเงิน"),
            ("pos", "ออกใบเสร็จรับเงิน (POS)"),
            ("batch_pos", "จัดใบเสร็จทั้งห้อง"),
            ("registration", "ทะเบียนวิชา (อศท.28)"),
            ("students", "ฐานข้อมูลนักศึกษา"),
            ("receipts", "ประวัติใบเสร็จ"),
            ("gallery", "คลังภาพใบเสร็จ"),
            ("settings", "ตั้งค่าระบบ"),
        ]

    def show(self):
        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk()
        # Global Error Guard: Never fail silently anywhere in the application
        self.root.report_callback_exception = self._on_tkinter_callback_exception
        sys.excepthook = self._on_tkinter_callback_exception
        IconManager.set_window_icon(self.root)
        self.root.title(
            f"ระบบออกใบเสร็จรับเงินค่าเทอม - {self.config.get('college_name', 'วิทยาลัยอาชีวศึกษาสุพรรณบุรี')}"
        )
        self.root.geometry("1200x680+20+20")
        self.root.minsize(1024, 620)
        self.root.configure(fg_color=AppTheme.COLOR_BG)

        try:
            self.root.state("zoomed")
        except Exception:
            pass

        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(400, lambda: self.root.attributes("-topmost", False))
        self.root.focus_force()

        # Left Navigation Sidebar (Executive Deep Navy)
        self.sidebar = ctk.CTkFrame(
            self.root,
            width=self.SIDEBAR_EXPANDED_WIDTH,
            corner_radius=0,
            fg_color=AppTheme.COLOR_SIDEBAR,
        )
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # 1. Top Section (Brand Header & Cashier Selector)
        self.sidebar_top = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.sidebar_top.pack(side="top", fill="x", padx=12, pady=(18, 10))

        brand_row = ctk.CTkFrame(self.sidebar_top, fg_color="transparent")
        brand_row.pack(fill="x")

        self.brand_text_box = ctk.CTkFrame(brand_row, fg_color="transparent")
        self.brand_text_box.pack(side="left", fill="x", expand=True)

        brand_icon = IconManager.get_icon("brand", state="accent", size=(28, 28))
        self.lbl_brand_title = ctk.CTkLabel(
            self.brand_text_box,
            text="  อศ.สุพรรณบุรี",
            image=brand_icon,
            compound="left",
            font=(AppTheme.FONT_FAMILY, 18, "bold"),
            text_color="#F8FAFC",
            anchor="w",
        )
        self.lbl_brand_title.pack(fill="x")

        self.lbl_brand_sub = ctk.CTkLabel(
            self.brand_text_box,
            text="ระบบการเงิน & ออกใบเสร็จ",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color="#94A3B8",
            anchor="w",
        )
        self.lbl_brand_sub.pack(fill="x", pady=(1, 0))

        chevron_left = IconManager.get_icon("chevron_left", state="normal", size=(14, 14))
        self.btn_toggle_sidebar = ctk.CTkButton(
            brand_row,
            text="",
            image=chevron_left,
            width=32,
            height=32,
            corner_radius=0,
            fg_color="#1E293B",
            hover_color="#334155",
            command=self.toggle_sidebar,
        )
        self.btn_toggle_sidebar.pack(side="right")

        # Cashier Selector (Cohesive dark slate card inside sidebar_top)
        self.user_box = ctk.CTkFrame(self.sidebar_top, fg_color="transparent")
        self.user_box.pack(fill="x", pady=(12, 0))

        ctk.CTkLabel(
            self.user_box,
            text="ผู้รับเงิน:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color="#94A3B8",
            anchor="w",
        ).pack(fill="x", padx=2, pady=(0, 4))

        self.cashier_var = ctk.StringVar(master=self.root, value=self.session.full_name)
        staff_options = [
            "นางสาวสุพรรณี ศรีอาชีวะ",
            "คุณกนกวรรณ รับชำระค่าเทอม",
            "เจ้าหน้าที่การเงิน อศ.สุพรรณบุรี",
        ]
        self.opt_cashier = SearchableDropdown(
            self.user_box,
            values=staff_options,
            variable=self.cashier_var,
            height=36,
            corner_radius=AppTheme.RADIUS_INPUT,
            fg_color="#1E293B",
            button_color="#334155",
            button_hover_color="#475569",
            text_color="#F8FAFC",
            dropdown_fg_color="#FFFFFF",
            dropdown_hover_color="#EFF6FF",
            dropdown_text_color="#0F172A",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            dropdown_font=(AppTheme.FONT_FAMILY, 13),
            command=self.on_change_cashier,
        )
        self.opt_cashier.pack(fill="x")

        # Subtle divider between brand/cashier header and navigation
        ctk.CTkFrame(self.sidebar, height=1, fg_color="#1E293B").pack(fill="x", padx=12, pady=(4, 4))

        # 2. Middle Navigation Section (Buttons with Vector CTkImage Icons)
        self.sidebar_nav = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.sidebar_nav.pack(side="top", fill="both", expand=True, pady=(6, 4))

        for key, full_label in self.nav_items_config:
            icon_img = IconManager.get_icon(key, state="normal", size=(20, 20))
            btn = ctk.CTkButton(
                self.sidebar_nav,
                text=f"  {full_label}",
                image=icon_img,
                compound="left",
                anchor="w",
                height=42,
                corner_radius=AppTheme.RADIUS_BTN,
                fg_color="transparent",
                hover_color=AppTheme.COLOR_SIDEBAR_HOVER,
                text_color="#CBD5E1",
                font=(AppTheme.FONT_FAMILY, 13, "bold"),
                command=lambda k=key: self.switch_tab(k),
            )
            self.nav_buttons[key] = btn
            if key == "registration" and not bool(self.config.get("show_course_registration", False)):
                continue
            btn.pack(fill="x", padx=12, pady=3)

        self.refresh_nav_visibility()

        # 3. Bottom Section (Exit & System Semester Info)
        self.sidebar_bottom = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.sidebar_bottom.pack(side="bottom", fill="x", padx=12, pady=(0, 16))

        ctk.CTkFrame(self.sidebar_bottom, height=1, fg_color="#1E293B").pack(fill="x", pady=(0, 12))

        self.lbl_bottom_meta = ctk.CTkLabel(
            self.sidebar_bottom,
            text=f"เล่มที่ {self.config.get('receipt_book_no', '7006')} • ภาคเรียน {self.config.get('current_semester', '1/2569')}",
            font=(AppTheme.FONT_FAMILY, 11),
            text_color="#64748B",
        )
        self.lbl_bottom_meta.pack(side="top", pady=(0, 8))

        exit_icon = IconManager.get_icon("exit", state="normal", size=(18, 18))
        self.btn_exit = ctk.CTkButton(
            self.sidebar_bottom,
            text="  ออกจากโปรแกรม",
            image=exit_icon,
            compound="left",
            height=AppTheme.CTRL_H_MD,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color="#1E293B",
            hover_color="#334155",
            text_color="#CBD5E1",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.on_close,
        )
        self.btn_exit.pack(side="top", fill="x")

        # Right Workspace Shell (Global Top Header + Content Area)
        self.right_shell = ctk.CTkFrame(self.root, corner_radius=0, fg_color=AppTheme.COLOR_BG)
        self.right_shell.pack(side="right", fill="both", expand=True)

        self._build_top_header(self.right_shell)

        # Main Workspace Area (generous breathing room)
        self.content_area = ctk.CTkFrame(self.right_shell, corner_radius=0, fg_color=AppTheme.COLOR_BG)
        self.content_area.pack(side="top", fill="both", expand=True, padx=24, pady=20)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.switch_tab("dashboard")
        self.root.mainloop()

    def is_button_packed(self, btn: Optional[ctk.CTkButton]) -> bool:
        """Safely checks if a widget is currently packed into the layout."""
        if not btn:
            return False
        try:
            return bool(btn.pack_info())
        except Exception:
            return False

    def refresh_nav_visibility(self):
        """Shows or hides optional navigation items like course registration based on config."""
        show_reg = bool(self.config.get("show_course_registration", False))
        btn_reg = self.nav_buttons.get("registration")
        if not btn_reg:
            return

        is_packed = self.is_button_packed(btn_reg)
        if show_reg:
            if not is_packed:
                btn_stu = self.nav_buttons.get("students")
                pad = 8 if self.sidebar_is_collapsed else 12
                if btn_stu and self.is_button_packed(btn_stu):
                    btn_reg.pack(fill="x", padx=pad, pady=3, before=btn_stu)
                else:
                    btn_reg.pack(fill="x", padx=pad, pady=3)
        else:
            if is_packed:
                btn_reg.pack_forget()
            if self.current_tab_key == "registration":
                self.switch_tab("dashboard")
            self._discard_cached_view("registration")

    def toggle_sidebar(self):
        """Toggles the navigation sidebar between Expanded (275px) and Collapsed (72px) modes."""
        self.sidebar_is_collapsed = not self.sidebar_is_collapsed
        if self.sidebar_is_collapsed:
            # Collapse Sidebar
            self.sidebar.configure(width=self.SIDEBAR_COLLAPSED_WIDTH)
            chevron_right = IconManager.get_icon("chevron_right", state="normal", size=(14, 14))
            self.btn_toggle_sidebar.configure(image=chevron_right, text="")
            self.brand_text_box.pack_forget()
            self.user_box.pack_forget()
            self.lbl_bottom_meta.pack_forget()
            exit_icon = IconManager.get_icon("exit", state="normal", size=(20, 20))
            self.btn_exit.configure(image=exit_icon, text="")

            for key, _full_label in self.nav_items_config:
                btn = self.nav_buttons.get(key)
                if btn:
                    state = "active" if key == self.current_tab_key else "normal"
                    icon = IconManager.get_icon(key, state=state, size=(22, 22))
                    btn.configure(
                        text="",
                        image=icon,
                        anchor="center",
                    )
                    if self.is_button_packed(btn):
                        btn.pack_configure(padx=8)
        else:
            # Expand Sidebar
            self.sidebar.configure(width=self.SIDEBAR_EXPANDED_WIDTH)
            chevron_left = IconManager.get_icon("chevron_left", state="normal", size=(14, 14))
            self.btn_toggle_sidebar.configure(image=chevron_left, text="")
            self.brand_text_box.pack(side="left", fill="x", expand=True)
            self.user_box.pack(fill="x", pady=(12, 0))
            self.lbl_bottom_meta.pack(side="top", pady=(0, 6), before=self.btn_exit)
            exit_icon = IconManager.get_icon("exit", state="normal", size=(18, 18))
            self.btn_exit.configure(image=exit_icon, text="  ออกจากโปรแกรม", compound="left")

            for key, full_label in self.nav_items_config:
                btn = self.nav_buttons.get(key)
                if btn:
                    state = "active" if key == self.current_tab_key else "normal"
                    icon = IconManager.get_icon(key, state=state, size=(20, 20))
                    btn.configure(
                        text=f"  {full_label}",
                        image=icon,
                        compound="left",
                        anchor="w",
                        font=(AppTheme.FONT_FAMILY, 13, "bold"),
                    )
                    if self.is_button_packed(btn):
                        btn.pack_configure(padx=12)

        self.refresh_nav_visibility()

    def _get_tab_title(self, tab_key: str) -> str:
        """Returns the Thai display title for a navigation tab key."""
        for key, label in self.nav_items_config:
            if key == tab_key:
                return label
        return "ระบบการเงินและออกใบเสร็จ"

    def _build_top_header(self, parent) -> ctk.CTkFrame:
        """Builds the global top app header with current view title, user profile, and logout button."""
        self.top_header = ctk.CTkFrame(
            parent,
            height=56,
            corner_radius=0,
            fg_color=AppTheme.COLOR_SURFACE,
        )
        self.top_header.pack(side="top", fill="x")
        self.top_header.pack_propagate(False)

        ctk.CTkFrame(parent, height=1, fg_color=AppTheme.COLOR_BORDER).pack(side="top", fill="x")

        # Left side: active tab title
        left_box = ctk.CTkFrame(self.top_header, fg_color="transparent")
        left_box.pack(side="left", fill="y", padx=24)

        self.lbl_header_title = ctk.CTkLabel(
            left_box,
            text=self._get_tab_title(self.current_tab_key),
            font=(AppTheme.FONT_FAMILY, 16, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        )
        self.lbl_header_title.pack(side="left", pady=12)

        # Right side: profile avatar, user name, role badge, logout button
        right_box = ctk.CTkFrame(self.top_header, fg_color="transparent")
        right_box.pack(side="right", fill="y", padx=24)

        display_name = (getattr(self.session, "full_name", "") or getattr(self.session, "username", "") or "ผู้ใช้งาน").strip()
        role_text = (getattr(self.session, "role", "") or "STAFF").strip().upper()
        initial_char = display_name[0].upper() if display_name else "U"

        self.lbl_profile_avatar = ctk.CTkLabel(
            right_box,
            text=initial_char,
            width=32,
            height=32,
            corner_radius=16,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
        )
        self.lbl_profile_avatar.pack(side="left", padx=(0, 8), pady=12)

        self.lbl_user_name = ctk.CTkLabel(
            right_box,
            text=display_name,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        )
        self.lbl_user_name.pack(side="left", padx=(0, 8), pady=12)

        self.lbl_user_role = ctk.CTkLabel(
            right_box,
            text=role_text,
            height=24,
            corner_radius=6,
            fg_color="#EFF6FF",
            text_color=AppTheme.COLOR_PRIMARY_BLUE,
            font=(AppTheme.FONT_FAMILY, 11, "bold"),
            padx=10,
        )
        self.lbl_user_role.pack(side="left", padx=(0, 14), pady=12)

        self.btn_logout = ctk.CTkButton(
            right_box,
            text="ออกจากระบบ",
            width=104,
            height=32,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color="#FEF2F2",
            hover_color="#FEE2E2",
            text_color=AppTheme.COLOR_DANGER,
            border_width=1,
            border_color="#FECACA",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.logout,
        )
        self.btn_logout.pack(side="left", pady=12)
        return self.top_header

    def on_change_cashier(self, new_name: str):
        self.session.full_name = new_name
        if self.lbl_user_name is not None:
            try:
                if self.lbl_user_name.winfo_exists():
                    self.lbl_user_name.configure(text=new_name)
            except Exception:
                pass
        if self.lbl_profile_avatar is not None and new_name:
            try:
                if self.lbl_profile_avatar.winfo_exists():
                    self.lbl_profile_avatar.configure(text=new_name.strip()[0].upper())
            except Exception:
                pass

    def _discard_cached_view(self, tab_key: str):
        """Destroys and removes a cached view if it is not currently displayed."""
        cached = self._view_cache.get(tab_key)
        if cached is not None and cached is not self.current_view:
            self._view_cache.pop(tab_key, None)
            try:
                if cached.winfo_exists():
                    cached.destroy()
            except Exception:
                pass

    def _create_view(self, tab_key: str) -> Optional[ctk.CTkFrame]:
        """Instantiates the view widget for the given tab key."""
        if tab_key == "dashboard":
            return DashboardView(
                self.content_area, self.db, self.session, self.config, on_navigate=self.switch_tab
            )
        if tab_key == "pos":
            return SmartReceiptPOSView(
                self.content_area, self.db, self.session, self.config, self.webhook, on_navigate=self.switch_tab
            )
        if tab_key == "batch_pos":
            return BatchClassReceiptView(
                self.content_area, self.db, self.session, self.config, on_navigate=self.switch_tab
            )
        if tab_key == "registration":
            return CourseRegistrationView(self.content_area, self.db, self.session, self.config)
        if tab_key == "students":
            return StudentAndFeeCatalogView(self.content_area, self.db, self.session, self.config)
        if tab_key == "receipts":
            return ReceiptRegisterView(self.content_area, self.db, self.session, self.config)
        if tab_key == "gallery":
            return ReceiptGalleryView(
                self.content_area, self.db, self.session, self.config, on_navigate=self.switch_tab
            )
        if tab_key == "settings":
            return SettingsAndBackupView(
                self.content_area,
                self.db,
                self.session,
                self.config,
                webhook=self.webhook,
                on_config_updated=self.refresh_nav_visibility,
            )
        return None

    def _set_nav_button_active(self, tab_key: str, active: bool):
        btn = self.nav_buttons.get(tab_key)
        if not btn:
            return
        icon_size = (22, 22) if self.sidebar_is_collapsed else (20, 20)
        if active:
            btn.configure(
                fg_color=AppTheme.COLOR_SIDEBAR_ACTIVE,
                text_color="#FFFFFF",
                image=IconManager.get_icon(tab_key, state="active", size=icon_size),
            )
        else:
            btn.configure(
                fg_color="transparent",
                text_color="#CBD5E1",
                image=IconManager.get_icon(tab_key, state="normal", size=icon_size),
            )

    def switch_tab(self, tab_key: str):
        if tab_key == "registration" and not self.config.get("show_course_registration", False):
            self._discard_cached_view("registration")
            if self.current_tab_key != "dashboard" or self.current_view is None:
                self.switch_tab("dashboard")
            return

        if tab_key == self.current_tab_key and self.current_view is not None:
            return

        prev_key = self.current_tab_key
        self.current_tab_key = tab_key
        if prev_key != tab_key:
            self._set_nav_button_active(prev_key, active=False)
        self._set_nav_button_active(tab_key, active=True)

        if self.current_view is not None:
            self.current_view.pack_forget()

        view = self._view_cache.get(tab_key)
        is_cached = False
        if view is not None:
            try:
                is_cached = bool(view.winfo_exists())
            except Exception:
                is_cached = False

        if is_cached and view is not None:
            if hasattr(view, "on_show"):
                view.on_show()
        else:
            view = self._create_view(tab_key)
            if view is not None:
                self._view_cache[tab_key] = view

        self.current_view = view
        if self.current_view is not None:
            self.current_view.pack(fill="both", expand=True)

        if self.lbl_header_title is not None:
            try:
                if self.lbl_header_title.winfo_exists():
                    self.lbl_header_title.configure(text=self._get_tab_title(tab_key))
            except Exception:
                pass

    def _on_tkinter_callback_exception(self, exc_type, exc_val, exc_tb):
        """Global error catcher: Logs full stacktrace and displays informative popup to ensure no silent crashes."""
        tb_text = "".join(traceback.format_exception(exc_type, exc_val, exc_tb))
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Write to persistent error log file in data/error_log.txt
        log_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, "error_log.txt")
        try:
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"\n{'='*75}\n[{now_str}] UNHANDLED APPLICATION ERROR:\n{tb_text}\n{'='*75}\n")
        except Exception:
            pass

        # Print to stderr if available
        try:
            sys.stderr.write(f"[{now_str}] {tb_text}\n")
        except Exception:
            pass

        err_type_name = getattr(exc_type, "__name__", "Error")
        err_msg = str(exc_val) or err_type_name
        if "Permission denied" in err_msg or exc_type is PermissionError:
            friendly = "ไฟล์หรือข้อมูลกำลังถูกเปิดใช้งานอยู่ในโปรแกรมอื่น กรุณาปิดโปรแกรมดังกล่าวก่อนแล้วลองใหม่อีกครั้ง"
        elif "database is locked" in err_msg.lower():
            friendly = "ฐานข้อมูลกำลังถูกใช้งานจากอีกกระบวนการหนึ่ง กรุณารอสักครู่แล้วลองใหม่อีกครั้ง"
        elif "no application is associated" in err_msg.lower():
            friendly = "ไม่พบโปรแกรมที่ถูกตั้งค่าให้เปิดไฟล์ประเภทนี้ กรุณาติดตั้งหรือกำหนด Default App"
        else:
            friendly = "ระบบพบข้อผิดพลาดที่ไม่คาดคิดในการทำงาน"

        user_message = (
            f"❌ เกิดข้อผิดพลาดในการทำงานของระบบ (System Error)\n\n"
            f"• ข้อความแจ้งเตือน: {err_msg}\n"
            f"• ประเภท: {err_type_name}\n"
            f"• คำแนะนำ: {friendly}\n\n"
            f"ระบบได้บันทึกรายละเอียดข้อผิดพลาดไว้ที่:\n{log_file}\n"
            f"หากต้องการความช่วยเหลือ กรุณาติดต่อผู้ดูแลระบบ"
        )
        try:
            messagebox.showerror("เกิดข้อผิดพลาดของระบบ", user_message, parent=getattr(self, "root", None))
        except Exception:
            try:
                messagebox.showerror("เกิดข้อผิดพลาดของระบบ", user_message)
            except Exception:
                pass

    def logout(self):
        """Logs out the current user and triggers the re-login callback if configured."""
        if self.on_logout is not None:
            try:
                self.on_logout()
            except Exception:
                pass
        self.on_close()

    def on_close(self):
        if self.config.get("auto_backup_on_exit", True):
            try:
                backup_database(self.db.db_path)
            except Exception:
                pass
        self.root.destroy()
