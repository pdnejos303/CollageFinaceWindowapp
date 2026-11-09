from typing import Any, Dict
import customtkinter as ctk

from core.config import backup_database
from core.database import DatabaseManager
from core.models import UserSession
from exporters.webhook_exporter import FinanceWebhookDispatcher
from ui.theme import AppTheme
from ui.views import (
    ReceiptRegisterView,
    SettingsAndBackupView,
    SmartReceiptPOSView,
    StudentAndFeeCatalogView,
)


class MainWindow:
    """
    Minimal, Ultra-Clean Single-Root-Window for Suphanburi Vocational College (SPVC).
    Unified Corporate Navy & Slate palette with zero clashing colors.
    """

    def __init__(self, db: DatabaseManager, session: UserSession, config: Dict[str, Any]):
        self.db = db
        self.session = session
        self.config = config
        self.webhook = FinanceWebhookDispatcher(
            webhook_url=config.get("webhook_url", ""),
            enabled=config.get("webhook_enabled", False),
            alert_threshold=config.get("webhook_alert_threshold", 50000.0),
        )
        self.nav_buttons: Dict[str, ctk.CTkButton] = {}
        self.current_view = None
        self.current_tab_key = "pos"

    def show(self):
        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk()
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
        self.sidebar = ctk.CTkFrame(self.root, width=250, corner_radius=0, fg_color=AppTheme.COLOR_SIDEBAR)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        brand_box = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        brand_box.pack(fill="x", padx=20, pady=(24, 18))

        ctk.CTkLabel(
            brand_box,
            text="🧾 อศ.สุพรรณบุรี",
            font=(AppTheme.FONT_FAMILY, 20, "bold"),
            text_color="#F8FAFC",
            anchor="w",
        ).pack(fill="x")

        ctk.CTkLabel(
            brand_box,
            text="ระบบออกใบเสร็จรับเงินค่าเทอม",
            font=(AppTheme.FONT_FAMILY, 13),
            text_color="#94A3B8",
            anchor="w",
        ).pack(fill="x", pady=(2, 0))

        # Cashier Selector (Cohesive dark slate card inside sidebar)
        user_box = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        user_box.pack(fill="x", padx=16, pady=(0, 20))

        ctk.CTkLabel(
            user_box,
            text="ผู้รับเงิน:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color="#94A3B8",
            anchor="w",
        ).pack(fill="x", padx=4, pady=(0, 4))

        self.cashier_var = ctk.StringVar(master=self.root, value=self.session.full_name)
        staff_options = [
            "นางสาวสุพรรณี ศรีอาชีวะ",
            "คุณกนกวรรณ รับชำระค่าเทอม",
            "เจ้าหน้าที่การเงิน อศ.สุพรรณบุรี",
        ]
        self.opt_cashier = ctk.CTkOptionMenu(
            user_box,
            values=staff_options,
            variable=self.cashier_var,
            height=36,
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

        # 4 Main Navigation Menus
        nav_items = [
            ("pos", "1. ออกใบเสร็จรับเงิน"),
            ("students", "2. รายชื่อนักศึกษา"),
            ("receipts", "3. ประวัติ & ยกเลิกใบเสร็จ"),
            ("settings", "4. ตั้งค่าระบบ"),
        ]

        for key, label in nav_items:
            btn = ctk.CTkButton(
                self.sidebar,
                text=label,
                anchor="w",
                height=48,
                corner_radius=8,
                fg_color="transparent",
                hover_color=AppTheme.COLOR_SIDEBAR_HOVER,
                text_color="#CBD5E1",
                font=(AppTheme.FONT_FAMILY, 15, "bold"),
                command=lambda k=key: self.switch_tab(k),
            )
            btn.pack(fill="x", padx=14, pady=4)
            self.nav_buttons[key] = btn

        # Bottom Exit Info
        bottom_box = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        bottom_box.pack(side="bottom", fill="x", padx=16, pady=20)

        ctk.CTkLabel(
            bottom_box,
            text=f"เล่มที่ {self.config.get('receipt_book_no', '7006')} • ภาคเรียน {self.config.get('current_semester', '1/2569')}",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color="#64748B",
        ).pack(pady=(0, 8))

        ctk.CTkButton(
            bottom_box,
            text="ออกจากโปรแกรม",
            height=40,
            corner_radius=8,
            fg_color="#1E293B",
            hover_color="#334155",
            text_color="#CBD5E1",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self.on_close,
        ).pack(fill="x")

        # Main Workspace Area
        self.content_area = ctk.CTkFrame(self.root, corner_radius=0, fg_color=AppTheme.COLOR_BG)
        self.content_area.pack(side="right", fill="both", expand=True, padx=20, pady=16)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.switch_tab("pos")
        self.root.mainloop()

    def on_change_cashier(self, new_name: str):
        self.session.full_name = new_name

    def switch_tab(self, tab_key: str):
        self.current_tab_key = tab_key
        for k, btn in self.nav_buttons.items():
            if k == tab_key:
                btn.configure(fg_color=AppTheme.COLOR_SIDEBAR_ACTIVE, text_color="#FFFFFF")
            else:
                btn.configure(fg_color="transparent", text_color="#CBD5E1")

        if self.current_view is not None:
            self.current_view.destroy()

        if tab_key == "pos":
            self.current_view = SmartReceiptPOSView(self.content_area, self.db, self.session, self.config, self.webhook)
        elif tab_key == "students":
            self.current_view = StudentAndFeeCatalogView(self.content_area, self.db, self.session, self.config)
        elif tab_key == "receipts":
            self.current_view = ReceiptRegisterView(self.content_area, self.db, self.session, self.config)
        elif tab_key == "settings":
            self.current_view = SettingsAndBackupView(self.content_area, self.db, self.session, self.config)

        if self.current_view:
            self.current_view.pack(fill="both", expand=True)

    def on_close(self):
        if self.config.get("auto_backup_on_exit", True):
            try:
                backup_database(self.db.db_path)
            except Exception:
                pass
        self.root.destroy()
