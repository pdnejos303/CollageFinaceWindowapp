from tkinter import messagebox
from typing import Optional
import customtkinter as ctk

from core.database import DatabaseManager
from core.models import UserSession
from ui.theme import AppTheme


class LoginWindow:
    """Clean, Fast Login Window for Suphanburi Vocational College (SPVC) Receipt System."""

    def __init__(self, db: DatabaseManager, college_name: str):
        self.db = db
        self.college_name = college_name
        self.session: Optional[UserSession] = None

    def show(self) -> Optional[UserSession]:
        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("blue")

        app = ctk.CTk()
        app.title("SPVC Smart Receipt - ระบบออกใบเสร็จรับเงิน วิทยาลัยอาชีวศึกษาสุพรรณบุรี")
        app.geometry("560x620")
        app.resizable(False, False)
        app.configure(fg_color=AppTheme.COLOR_BG)

        app.update_idletasks()
        x = (app.winfo_screenwidth() // 2) - (560 // 2)
        y = (app.winfo_screenheight() // 2) - (620 // 2)
        app.geometry(f"+{x}+{y}")

        # Top Brand Banner
        banner = ctk.CTkFrame(app, fg_color=AppTheme.COLOR_SIDEBAR, corner_radius=0, height=145)
        banner.pack(fill="x", side="top")
        banner.pack_propagate(False)

        ctk.CTkLabel(
            banner,
            text="🧾 ระบบออกใบเสร็จและรับชำระค่าเทอม",
            font=(AppTheme.FONT_FAMILY, 23, "bold"),
            text_color="#F8FAFC",
        ).pack(pady=(24, 4))

        ctk.CTkLabel(
            banner,
            text=f"{self.college_name}",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color="#93C5FD",
        ).pack()

        ctk.CTkLabel(
            banner,
            text="SPVC SMART TUITION & FLEXIBLE RECEIPT SYSTEM",
            font=(AppTheme.FONT_FAMILY, 11),
            text_color="#94A3B8",
        ).pack(pady=(4, 0))

        # Login Card Container
        card = ctk.CTkFrame(
            app,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        card.pack(fill="both", expand=True, padx=36, pady=24)

        ctk.CTkLabel(
            card,
            text="🔐 เข้าสู่ระบบสำหรับเจ้าหน้าที่การเงิน",
            font=(AppTheme.FONT_FAMILY, 17, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(anchor="w", padx=28, pady=(20, 12))

        ctk.CTkLabel(
            card,
            text="👤 ชื่อผู้ใช้งาน (Username):",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(anchor="w", padx=28)

        entry_user = ctk.CTkEntry(
            card,
            height=42,
            corner_radius=8,
            border_width=1,
            fg_color=AppTheme.COLOR_INPUT_BG,
            border_color=AppTheme.COLOR_INPUT_BORDER,
            text_color=AppTheme.COLOR_TEXT,
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
        )
        entry_user.pack(fill="x", padx=28, pady=(4, 14))
        entry_user.insert(0, "admin")

        ctk.CTkLabel(
            card,
            text="🔑 รหัสผ่าน (Password):",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(anchor="w", padx=28)

        entry_pass = ctk.CTkEntry(
            card,
            height=42,
            corner_radius=8,
            border_width=1,
            show="•",
            fg_color=AppTheme.COLOR_INPUT_BG,
            border_color=AppTheme.COLOR_INPUT_BORDER,
            text_color=AppTheme.COLOR_TEXT,
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
        )
        entry_pass.pack(fill="x", padx=28, pady=(4, 18))
        entry_pass.insert(0, "admin123")

        def do_login(_event=None):
            u = entry_user.get().strip()
            p = entry_pass.get().strip()
            user_row = self.db.authenticate(u, p)
            if not user_row:
                messagebox.showerror("เข้าสู่ระบบไม่สำเร็จ", "ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง กรุณาตรวจสอบอีกครั้ง!")
                return
            self.session = UserSession(
                user_id=user_row["id"],
                username=user_row["username"],
                full_name=user_row["full_name"],
                role=user_row["role"],
                department=user_row["department"],
            )
            app.destroy()

        entry_pass.bind("<Return>", do_login)

        btn_login = ctk.CTkButton(
            card,
            text="เข้าสู่ระบบออกใบเสร็จ (START POS)",
            height=46,
            corner_radius=8,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 15, "bold"),
            command=do_login,
        )
        btn_login.pack(fill="x", padx=28, pady=(0, 14))

        hint_frame = ctk.CTkFrame(
            card,
            fg_color=AppTheme.COLOR_SURFACE_ALT,
            corner_radius=8,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        hint_frame.pack(fill="x", padx=28, pady=(4, 14))

        ctk.CTkLabel(
            hint_frame,
            text="ปุ่มสลับบัญชีเจ้าหน้าที่:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(anchor="w", padx=12, pady=(8, 6))

        roles_row = ctk.CTkFrame(hint_frame, fg_color="transparent")
        roles_row.pack(fill="x", padx=10, pady=(0, 10))

        def fill_account(u: str, p: str):
            entry_user.delete(0, "end")
            entry_user.insert(0, u)
            entry_pass.delete(0, "end")
            entry_pass.insert(0, p)

        for lbl, u, p in [
            ("หัวหน้าการเงิน (Admin)", "admin", "admin123"),
            ("เจ้าหน้าที่รับชำระ (Cashier)", "finance", "fin123"),
            ("ตรวจสอบภายใน (Auditor)", "auditor", "audit123"),
        ]:
            ctk.CTkButton(
                roles_row,
                text=lbl,
                height=34,
                corner_radius=6,
                fg_color=AppTheme.COLOR_NEUTRAL_BTN,
                text_color=AppTheme.COLOR_TEXT_SECONDARY,
                hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
                font=(AppTheme.FONT_FAMILY, 12, "bold"),
                command=lambda u=u, p=p: fill_account(u, p),
            ).pack(side="left", expand=True, fill="x", padx=3)

        app.mainloop()
        return self.session
