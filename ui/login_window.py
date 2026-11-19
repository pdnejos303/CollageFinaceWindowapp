from tkinter import messagebox
from typing import Optional
import customtkinter as ctk

from core.database import DatabaseManager
from core.models import UserSession
from ui.icon_manager import IconManager
from ui.notifications import highlight_missing_field
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
        IconManager.set_window_icon(app)
        app.title("SPVC Smart Receipt - ระบบออกใบเสร็จรับเงิน วิทยาลัยอาชีวศึกษาสุพรรณบุรี")
        app.geometry("500x560")
        app.resizable(False, False)
        app.configure(fg_color=AppTheme.COLOR_BG)

        app.update_idletasks()
        x = (app.winfo_screenwidth() // 2) - (500 // 2)
        y = (app.winfo_screenheight() // 2) - (560 // 2)
        app.geometry(f"+{x}+{y}")

        # Top Brand Banner (Compact 110px with Vector Brand Icon)
        banner = ctk.CTkFrame(app, fg_color=AppTheme.COLOR_SIDEBAR, corner_radius=0, height=110)
        banner.pack(fill="x", side="top")
        banner.pack_propagate(False)

        banner_inner = ctk.CTkFrame(banner, fg_color="transparent")
        banner_inner.pack(expand=True, padx=28)

        brand_icon = IconManager.get_icon("brand", state="accent", size=(32, 32))
        ctk.CTkLabel(
            banner_inner,
            text="  ระบบออกใบเสร็จและรับชำระค่าเทอม",
            image=brand_icon,
            compound="left",
            font=AppTheme.FONT_TITLE,
            text_color="#F8FAFC",
        ).pack(pady=(0, 4))

        ctk.CTkLabel(
            banner_inner,
            text=f"{self.college_name}",
            font=(AppTheme.FONT_FAMILY, 13),
            text_color="#93C5FD",
        ).pack()

        # Login Card Container
        card = ctk.CTkFrame(
            app,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        card.pack(fill="both", expand=True, padx=32, pady=24)

        ctk.CTkLabel(
            card,
            text="เข้าสู่ระบบสำหรับเจ้าหน้าที่การเงิน",
            font=(AppTheme.FONT_FAMILY, 16, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(anchor="w", padx=28, pady=(24, 16))

        ctk.CTkLabel(
            card,
            text="ชื่อผู้ใช้งาน",
            font=AppTheme.FONT_LABEL,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(anchor="w", padx=28)

        entry_user = ctk.CTkEntry(
            card,
            height=AppTheme.CTRL_H_LG,
            corner_radius=AppTheme.RADIUS,
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
            text="รหัสผ่าน",
            font=AppTheme.FONT_LABEL,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(anchor="w", padx=28)

        entry_pass = ctk.CTkEntry(
            card,
            height=AppTheme.CTRL_H_LG,
            corner_radius=AppTheme.RADIUS,
            border_width=1,
            show="•",
            fg_color=AppTheme.COLOR_INPUT_BG,
            border_color=AppTheme.COLOR_INPUT_BORDER,
            text_color=AppTheme.COLOR_TEXT,
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
        )
        entry_pass.pack(fill="x", padx=28, pady=(4, 20))
        entry_pass.insert(0, "admin123")

        def do_login(_event=None):
            u = entry_user.get().strip()
            p = entry_pass.get().strip()
            if not u:
                highlight_missing_field(entry_user, "กรุณากรอกชื่อผู้ใช้งาน", show_toast_msg=True)
                return
            if not p:
                highlight_missing_field(entry_pass, "กรุณากรอกรหัสผ่าน", show_toast_msg=True)
                return
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
            text="เข้าสู่ระบบออกใบเสร็จ",
            height=44,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 14, "bold"),
            command=do_login,
        )
        btn_login.pack(fill="x", padx=28, pady=(0, 18))

        # Subtle divider & lightweight inline role-switcher
        ctk.CTkFrame(card, height=1, fg_color=AppTheme.COLOR_BORDER).pack(fill="x", padx=28, pady=(0, 12))

        ctk.CTkLabel(
            card,
            text="สลับบัญชีเจ้าหน้าที่ (ทดสอบระบบ)",
            font=AppTheme.FONT_CAPTION,
            text_color=AppTheme.COLOR_TEXT_MUTED,
        ).pack(anchor="w", padx=28, pady=(0, 6))

        roles_row = ctk.CTkFrame(card, fg_color="transparent")
        roles_row.pack(fill="x", padx=28, pady=(0, 20))

        def fill_account(u: str, p: str):
            entry_user.delete(0, "end")
            entry_user.insert(0, u)
            entry_pass.delete(0, "end")
            entry_pass.insert(0, p)

        for idx, (lbl, u, p) in enumerate([
            ("หัวหน้าการเงิน", "admin", "admin123"),
            ("เจ้าหน้าที่รับชำระ", "finance", "fin123"),
            ("ตรวจสอบภายใน", "auditor", "audit123"),
        ]):
            padx_val = (0, 4) if idx == 0 else ((4, 0) if idx == 2 else 4)
            ctk.CTkButton(
                roles_row,
                text=lbl,
                height=AppTheme.CTRL_H_SM,
                corner_radius=AppTheme.RADIUS,
                fg_color=AppTheme.COLOR_SURFACE_ALT,
                text_color=AppTheme.COLOR_TEXT_SECONDARY,
                hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
                font=(AppTheme.FONT_FAMILY, 12),
                command=lambda u=u, p=p: fill_account(u, p),
            ).pack(side="left", expand=True, fill="x", padx=padx_val)

        app.mainloop()
        return self.session
