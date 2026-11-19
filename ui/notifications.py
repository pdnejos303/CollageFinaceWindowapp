"""
Modern, Premium Alert Dialogs & Toast Notifications for SPVC College Finance GodApp.
Replaces outdated Windows 95/XP-style Win32 message boxes with sleek, 
high-definition CustomTkinter modal cards and floating toast notifications.
"""

import os
from typing import Any, Dict, List, Optional, Tuple
import customtkinter as ctk
from PIL import Image

from ui.icon_manager import IconManager
from ui.theme import AppTheme


def _get_active_parent(parent: Optional[Any] = None) -> Optional[Any]:
    """Finds the most suitable active parent window for centering and transient modality."""
    if parent is not None:
        try:
            if hasattr(parent, "winfo_exists") and parent.winfo_exists():
                return parent.winfo_toplevel()
        except Exception:
            pass

    # Try default root window
    try:
        import tkinter as tk
        if hasattr(tk, "_default_root") and tk._default_root is not None:
            if tk._default_root.winfo_exists():
                return tk._default_root
    except Exception:
        pass
    return None


class ModernDialog(ctk.CTkToplevel):
    """
    Sleek, Modern Modal Dialog matching the Corporate Navy & Slate Design System.
    Features:
    - High-DPI anti-aliased vector badges (Success, Warning, Error, Question, Info)
    - Distinct top accent band matching status
    - Multi-line selectable / scrollable message support
    - Enter / Escape keyboard shortcuts
    - Auto-centering over parent window with modal grab
    - Sharp window icon matching application branding
    """

    ICONS_DIR = os.path.join(os.path.dirname(__file__), "assets", "dialog_icons")

    ACCENT_COLORS = {
        "success": AppTheme.COLOR_SUCCESS,
        "warning": AppTheme.COLOR_WARNING,
        "error": AppTheme.COLOR_DANGER,
        "question": AppTheme.COLOR_PRIMARY_BLUE,
        "info": AppTheme.COLOR_INFO,
    }

    BUTTON_COLORS = {
        "success": (AppTheme.COLOR_SUCCESS, AppTheme.COLOR_SUCCESS_HOVER),
        "warning": (AppTheme.COLOR_PRIMARY_BLUE, AppTheme.COLOR_PRIMARY_BLUE_HOVER),
        "error": (AppTheme.COLOR_DANGER, AppTheme.COLOR_DANGER_HOVER),
        "question": (AppTheme.COLOR_PRIMARY_BLUE, AppTheme.COLOR_PRIMARY_BLUE_HOVER),
        "info": (AppTheme.COLOR_PRIMARY_BLUE, AppTheme.COLOR_PRIMARY_BLUE_HOVER),
    }

    _icon_cache: Dict[str, ctk.CTkImage] = {}

    def __init__(
        self,
        title: str = "แจ้งเตือน",
        message: str = "",
        kind: str = "info",
        parent: Optional[Any] = None,
        buttons: Optional[List[Tuple[str, Any, str]]] = None,
        default_return: Any = None,
        danger: bool = False,
        width: int = 490,
    ):
        actual_parent = _get_active_parent(parent)
        super().__init__(actual_parent)

        self.title(title)
        self.result = default_return
        self.kind = kind
        self.danger = danger
        self._buttons = buttons or [("ตกลง", "ok", "primary")]
        self.target_width = width

        # Hide initially while computing dimensions and layout
        self.withdraw()
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")

        if actual_parent:
            try:
                self.transient(actual_parent)
            except Exception:
                pass

        # Apply high-definition application icon
        try:
            IconManager.set_window_icon(self)
        except Exception:
            pass

        self._build_ui(title, message)
        self._center_and_show(actual_parent)

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Escape>", lambda e: self._on_close())
        self.bind("<Return>", lambda e: self._on_confirm_key())
        self.bind("<KP_Enter>", lambda e: self._on_confirm_key())

        # Grab focus as modal dialog
        try:
            if self.winfo_exists():
                self.grab_set()
                self.focus_set()
        except Exception:
            pass

        self.wait_window()

    def _get_badge_icon(self, kind: str) -> Optional[ctk.CTkImage]:
        if kind in self._icon_cache:
            return self._icon_cache[kind]

        icon_path = os.path.join(self.ICONS_DIR, f"{kind}.png")
        if os.path.exists(icon_path):
            try:
                pil_img = Image.open(icon_path)
                badge = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(48, 48))
                self._icon_cache[kind] = badge
                return badge
            except Exception:
                pass
        return None

    def _build_ui(self, title: str, message: str):
        # 1. Top Colorful Accent Band (4px)
        accent_color = self.ACCENT_COLORS.get(self.kind, "#1E40AF")
        accent_bar = ctk.CTkFrame(self, height=4, fg_color=accent_color, corner_radius=0)
        accent_bar.pack(fill="x", side="top")

        # 2. Card Content Area
        card_content = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=0)
        card_content.pack(fill="both", expand=True, padx=24, pady=(20, 16))

        # Horizontal Row: Badge Icon (Left) + Text Container (Right)
        header_row = ctk.CTkFrame(card_content, fg_color="transparent")
        header_row.pack(fill="x", anchor="n")

        badge_img = self._get_badge_icon(self.kind)
        if badge_img:
            lbl_icon = ctk.CTkLabel(header_row, text="", image=badge_img)
            lbl_icon.pack(side="left", anchor="n", padx=(0, 16), pady=(2, 0))

        text_col = ctk.CTkFrame(header_row, fg_color="transparent")
        text_col.pack(side="left", fill="both", expand=True)

        # Title Label
        lbl_title = ctk.CTkLabel(
            text_col,
            text=title,
            font=(AppTheme.FONT_FAMILY, 16, "bold"),
            text_color="#0F172A",
            anchor="w",
            justify="left",
        )
        lbl_title.pack(fill="x", anchor="w", pady=(0, 6))

        # Message Section: Auto-detect length and choose clean label or scrollable box
        clean_msg = (message or "").strip()
        lines = clean_msg.split("\n")
        est_lines = sum(max(1, len(line) // 42 + 1) for line in lines)

        if est_lines > 5 or len(clean_msg) > 280:
            # Long message / error trace: use selectable styled scrollable textbox
            txt_box = ctk.CTkTextbox(
                text_col,
                height=min(180, max(90, est_lines * 22)),
                fg_color=AppTheme.COLOR_BG,
                border_width=1,
                border_color=AppTheme.COLOR_BORDER,
                corner_radius=AppTheme.RADIUS,
                font=(AppTheme.FONT_FAMILY, 12),
                text_color=AppTheme.COLOR_TEXT_SECONDARY,
                wrap="word",
            )
            txt_box.pack(fill="both", expand=True, pady=(2, 0))
            txt_box.insert("1.0", clean_msg)
            txt_box.configure(state="disabled")
        else:
            # Short / standard message: clean typography label
            lbl_msg = ctk.CTkLabel(
                text_col,
                text=clean_msg,
                font=(AppTheme.FONT_FAMILY, 13),
                text_color="#334155",
                anchor="w",
                justify="left",
                wraplength=self.target_width - 120,
            )
            lbl_msg.pack(fill="x", anchor="w", pady=(2, 0))

        # 3. Bottom Button Bar (Warm Slate footer with pill buttons)
        btn_bar = ctk.CTkFrame(self, fg_color="#F8FAFC", height=58, corner_radius=0)
        btn_bar.pack(fill="x", side="bottom")
        btn_bar.pack_propagate(False)

        # Separator line
        sep = ctk.CTkFrame(self, height=1, fg_color="#E2E8F0", corner_radius=0)
        sep.pack(fill="x", side="bottom")

        btn_inner = ctk.CTkFrame(btn_bar, fg_color="transparent")
        btn_inner.pack(side="right", padx=20, pady=10)

        self._action_buttons: List[ctk.CTkButton] = []
        for label, ret_val, style in self._buttons:
            if style == "primary":
                fg, hover = self.BUTTON_COLORS.get(self.kind, (AppTheme.COLOR_PRIMARY_BLUE, AppTheme.COLOR_PRIMARY_BLUE_HOVER))
                if self.danger:
                    fg, hover = (AppTheme.COLOR_DANGER, AppTheme.COLOR_DANGER_HOVER)
                btn = ctk.CTkButton(
                    btn_inner,
                    text=label,
                    width=115,
                    height=AppTheme.CTRL_H_MD,
                    corner_radius=AppTheme.RADIUS_BTN,
                    fg_color=fg,
                    hover_color=hover,
                    text_color=AppTheme.COLOR_TEXT_LIGHT,
                    font=(AppTheme.FONT_FAMILY, 13, "bold"),
                    command=lambda v=ret_val: self._select_and_close(v),
                )
            else:
                # Secondary / Cancel button
                btn = ctk.CTkButton(
                    btn_inner,
                    text=label,
                    width=100,
                    height=AppTheme.CTRL_H_MD,
                    corner_radius=AppTheme.RADIUS_BTN,
                    fg_color=AppTheme.COLOR_SURFACE,
                    hover_color=AppTheme.COLOR_SURFACE_ALT,
                    border_color=AppTheme.COLOR_BORDER_STRONG,
                    border_width=1,
                    text_color=AppTheme.COLOR_SLATE,
                    font=(AppTheme.FONT_FAMILY, 13, "bold"),
                    command=lambda v=ret_val: self._select_and_close(v),
                )
            btn.pack(side="right", padx=(12, 0))
            self._action_buttons.append(btn)

    def _center_and_show(self, parent: Optional[Any]):
        self.update_idletasks()
        req_w = max(self.target_width, self.winfo_reqwidth())
        req_h = max(165, self.winfo_reqheight())

        if parent is not None:
            try:
                px = parent.winfo_rootx()
                py = parent.winfo_rooty()
                pw = parent.winfo_width()
                ph = parent.winfo_height()
                x = px + max(10, (pw - req_w) // 2)
                y = py + max(10, (ph - req_h) // 2)
                self.geometry(f"{req_w}x{req_h}+{x}+{y}")
                self.deiconify()
                return
            except Exception:
                pass

        # Fallback to screen center
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = max(10, (sw - req_w) // 2)
        y = max(10, (sh - req_h) // 2)
        self.geometry(f"{req_w}x{req_h}+{x}+{y}")
        self.deiconify()

    def _select_and_close(self, val: Any):
        self.result = val
        self.destroy()

    def _on_confirm_key(self):
        # Trigger primary action button on Enter key
        for label, ret_val, style in self._buttons:
            if style == "primary":
                self._select_and_close(ret_val)
                return
        self._on_close()

    def _revert_withdraw_after_windows_set_titlebar_color(self):
        try:
            if self.winfo_exists():
                super()._revert_withdraw_after_windows_set_titlebar_color()
        except Exception:
            pass

    def _on_close(self):
        self.destroy()

    def focus_set(self):
        try:
            if hasattr(self, "winfo_exists") and self.winfo_exists():
                super().focus_set()
        except Exception:
            pass

    def destroy(self):
        try:
            if hasattr(self, "winfo_exists") and self.winfo_exists():
                self.grab_release()
        except Exception:
            pass
        try:
            super().destroy()
        except Exception:
            pass


class ModernToast(ctk.CTkToplevel):
    """
    Non-blocking, sleek floating toast notification that slides in at the bottom-right
    of the active window and automatically fades out without freezing the UI.
    """

    ACCENTS = {
        "success": (AppTheme.COLOR_SUCCESS, AppTheme.COLOR_SUCCESS_BG),
        "warning": (AppTheme.COLOR_WARNING, AppTheme.COLOR_WARNING_BG),
        "error": (AppTheme.COLOR_DANGER, AppTheme.COLOR_DANGER_BG),
        "info": (AppTheme.COLOR_INFO, AppTheme.COLOR_INFO_BG),
    }

    def __init__(
        self,
        parent: Optional[Any] = None,
        title: str = "แจ้งเตือน",
        message: str = "",
        kind: str = "success",
        duration_ms: int = 2800,
    ):
        actual_parent = _get_active_parent(parent)
        super().__init__(actual_parent)

        self.withdraw()
        self.overrideredirect(True)
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass

        accent_color, bg_tint = self.ACCENTS.get(kind, (AppTheme.COLOR_INFO, AppTheme.COLOR_INFO_BG))

        card = ctk.CTkFrame(
            self,
            fg_color=AppTheme.COLOR_SURFACE,
            corner_radius=AppTheme.RADIUS,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER_STRONG,
        )
        card.pack(fill="both", expand=True, padx=1, pady=1)

        # Left accent vertical indicator
        strip = ctk.CTkFrame(card, width=6, fg_color=accent_color, corner_radius=0)
        strip.pack(side="left", fill="y")

        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(side="left", fill="both", expand=True, padx=14, pady=10)

        ctk.CTkLabel(
            content,
            text=title,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color="#0F172A",
            anchor="w",
        ).pack(fill="x")

        if message:
            ctk.CTkLabel(
                content,
                text=message,
                font=(AppTheme.FONT_FAMILY, 11),
                text_color="#64748B",
                anchor="w",
            ).pack(fill="x", pady=(2, 0))

        # Close button '✕'
        btn_close = ctk.CTkButton(
            card,
            text="✕",
            width=24,
            height=24,
            corner_radius=AppTheme.RADIUS,
            fg_color="transparent",
            hover_color=AppTheme.COLOR_SURFACE_ALT,
            text_color=AppTheme.COLOR_TEXT_MUTED,
            font=(AppTheme.FONT_FAMILY, 12),
            command=self.destroy,
        )
        btn_close.pack(side="right", padx=(0, 10), pady=10)

        # Position at bottom-right of parent window
        tw, th = 350, 70 if message else 54
        if actual_parent:
            try:
                actual_parent.update_idletasks()
                px = actual_parent.winfo_rootx()
                py = actual_parent.winfo_rooty()
                pw = actual_parent.winfo_width()
                ph = actual_parent.winfo_height()
                tx = px + pw - tw - 24
                ty = py + ph - th - 24
                self.geometry(f"{tw}x{th}+{tx}+{ty}")
                self.deiconify()
                self.after(duration_ms, self.destroy)
                return
            except Exception:
                pass

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        tx = sw - tw - 24
        ty = sh - th - 60
        self.geometry(f"{tw}x{th}+{tx}+{ty}")
        self.deiconify()
        self.after(duration_ms, self._safe_destroy)

    def _revert_withdraw_after_windows_set_titlebar_color(self):
        try:
            if self.winfo_exists():
                super()._revert_withdraw_after_windows_set_titlebar_color()
        except Exception:
            pass

    def _safe_destroy(self):
        try:
            if self.winfo_exists():
                self.destroy()
        except Exception:
            pass

    def focus_set(self):
        try:
            if hasattr(self, "winfo_exists") and self.winfo_exists():
                super().focus_set()
        except Exception:
            pass

    def destroy(self):
        try:
            if hasattr(self, "winfo_exists") and self.winfo_exists():
                super().destroy()
        except Exception:
            pass


# ==============================================================================
# DROP-IN REPLACEMENT FUNCTIONS (Matches tkinter.messagebox signature)
# ==============================================================================

def showinfo(title: str = "แจ้งเตือน", message: str = "", parent: Optional[Any] = None, **kwargs) -> str:
    """Displays a modern informative dialog. Auto-detects success messages for emerald styling."""
    t_lower = str(title).lower()
    m_lower = str(message).lower()
    is_success = any(kw in t_lower or kw in m_lower for kw in ["สำเร็จ", "เรียบร้อย", "success", "บันทึกข้อมูล", "สร้างไฟล์"])
    kind = "success" if is_success else "info"

    ModernDialog(
        title=title,
        message=message,
        kind=kind,
        parent=parent or kwargs.get("parent"),
        buttons=[("ตกลง", "ok", "primary")],
        default_return="ok",
    )
    return "ok"


def showwarning(title: str = "แจ้งเตือน", message: str = "", parent: Optional[Any] = None, **kwargs) -> str:
    """Displays a modern amber warning dialog."""
    ModernDialog(
        title=title,
        message=message,
        kind="warning",
        parent=parent or kwargs.get("parent"),
        buttons=[("ตกลง", "ok", "primary")],
        default_return="ok",
    )
    return "ok"


def showerror(title: str = "เกิดข้อผิดพลาด", message: str = "", parent: Optional[Any] = None, **kwargs) -> str:
    """Displays a modern crimson error dialog with selectable / copyable text support."""
    ModernDialog(
        title=title,
        message=message,
        kind="error",
        parent=parent or kwargs.get("parent"),
        buttons=[("ตกลง", "ok", "primary")],
        default_return="ok",
    )
    return "ok"


def askyesno(title: str = "ยืนยันการทำรายการ", message: str = "", parent: Optional[Any] = None, **kwargs) -> bool:
    """Displays a modern confirmation dialog with Yes / No actions."""
    t_lower = str(title).lower()
    m_lower = str(message).lower()
    is_danger = any(kw in t_lower or kw in m_lower for kw in ["ลบ", "ล้าง", "ยกเลิกข้อมูล", "delete", "remove", "กู้คืน"])

    dlg = ModernDialog(
        title=title,
        message=message,
        kind="question",
        parent=parent or kwargs.get("parent"),
        buttons=[
            ("ใช่ / ยืนยัน", True, "primary"),
            ("ไม่ / ยกเลิก", False, "secondary"),
        ],
        default_return=False,
        danger=is_danger,
    )
    return bool(dlg.result)


def askokcancel(title: str = "ยืนยันการทำรายการ", message: str = "", parent: Optional[Any] = None, **kwargs) -> bool:
    """Displays a modern confirmation dialog with OK / Cancel actions."""
    dlg = ModernDialog(
        title=title,
        message=message,
        kind="question",
        parent=parent or kwargs.get("parent"),
        buttons=[
            ("ตกลง", True, "primary"),
            ("ยกเลิก", False, "secondary"),
        ],
        default_return=False,
    )
    return bool(dlg.result)


def askquestion(title: str = "คำถาม", message: str = "", parent: Optional[Any] = None, **kwargs) -> str:
    """Returns 'yes' or 'no' string matching tkinter.messagebox.askquestion."""
    res = askyesno(title=title, message=message, parent=parent, **kwargs)
    return "yes" if res else "no"


def askretrycancel(title: str = "ลองใหม่อีกครั้ง", message: str = "", parent: Optional[Any] = None, **kwargs) -> bool:
    """Displays a retry / cancel dialog."""
    dlg = ModernDialog(
        title=title,
        message=message,
        kind="warning",
        parent=parent or kwargs.get("parent"),
        buttons=[
            ("ลองใหม่", True, "primary"),
            ("ยกเลิก", False, "secondary"),
        ],
        default_return=False,
    )
    return bool(dlg.result)


def show_toast(parent: Optional[Any] = None, title: str = "แจ้งเตือน", message: str = "", kind: str = "success", duration_ms: int = 2800):
    """Spawns a sleek, non-blocking floating toast notification."""
    try:
        ModernToast(parent=parent, title=title, message=message, kind=kind, duration_ms=duration_ms)
    except Exception:
        pass


def highlight_missing_field(widget: Any, message: Optional[str] = None, show_toast_msg: bool = False):
    """
    Elegantly pulses a form field (glowing coral/crimson focus ring with subtle soft tint)
    and jumps focus to it when a required input is missing.
    Automatically resets back to normal styling as soon as the user interacts or types.
    """
    if not widget or not hasattr(widget, "configure"):
        return

    # Cancel any active pulse on this widget
    if hasattr(widget, "_pulse_job") and widget._pulse_job:
        try:
            widget.after_cancel(widget._pulse_job)
        except Exception:
            pass

    # Save original style attributes
    if getattr(widget, "_orig_border_color", None) is None:
        try:
            widget._orig_border_color = widget.cget("border_color")
        except Exception:
            widget._orig_border_color = AppTheme.COLOR_INPUT_BORDER

    if getattr(widget, "_orig_border_width", None) is None:
        try:
            widget._orig_border_width = widget.cget("border_width")
        except Exception:
            widget._orig_border_width = 1

    if getattr(widget, "_orig_fg_color", None) is None:
        try:
            widget._orig_fg_color = widget.cget("fg_color")
        except Exception:
            widget._orig_fg_color = None

    # Focus the missing field
    try:
        widget.focus_set()
    except Exception:
        pass

    # Restoration callback
    def reset_styling(*_):
        try:
            if hasattr(widget, "_pulse_job") and widget._pulse_job:
                widget.after_cancel(widget._pulse_job)
                widget._pulse_job = None
            cfg = {
                "border_color": widget._orig_border_color,
                "border_width": widget._orig_border_width,
            }
            if widget._orig_fg_color is not None:
                cfg["fg_color"] = widget._orig_fg_color
            widget.configure(**cfg)
        except Exception:
            pass

    # Attach programmatic reset function
    widget._reset_missing_highlight = reset_styling

    # Attach auto-restore triggers to widget and inner entry if present
    targets = [widget]
    if hasattr(widget, "_entry") and widget._entry:
        targets.append(widget._entry)
    if hasattr(widget, "entry") and widget.entry:
        targets.append(widget.entry)

    for tgt in targets:
        try:
            tgt.bind("<KeyRelease>", reset_styling, add="+")
            tgt.bind("<FocusIn>", reset_styling, add="+")
            tgt.bind("<Button-1>", reset_styling, add="+")
        except Exception:
            pass

    # Smooth 3-pulse animation timeline (Gentle Crimson / Coral Ring)
    pulse_frames = [
        ("#EF4444", 2, "#FEF2F2"),
        ("#FCA5A5", 2, None),
        ("#EF4444", 2, "#FEF2F2"),
        ("#FCA5A5", 2, None),
        ("#EF4444", 2, "#FEF2F2"),
    ]

    def run_frame(frame_idx: int):
        if frame_idx >= len(pulse_frames):
            widget._pulse_job = None
            return
        b_color, b_w, fg_tint = pulse_frames[frame_idx]
        try:
            cfg = {"border_color": b_color, "border_width": b_w}
            if fg_tint is not None and widget._orig_fg_color is not None:
                cfg["fg_color"] = fg_tint
            widget.configure(**cfg)
            widget._pulse_job = widget.after(160, lambda: run_frame(frame_idx + 1))
        except Exception:
            widget._pulse_job = None

    run_frame(0)

    # Optional non-blocking toast
    if message and show_toast_msg:
        show_toast(
            parent=getattr(widget, "winfo_toplevel", lambda: None)(),
            title="กรุณากรอกข้อมูล",
            message=message,
            kind="warning",
            duration_ms=2500,
        )


class ModernMessageBox:
    """Class container matching tkinter.messagebox static API."""
    showinfo = staticmethod(showinfo)
    showwarning = staticmethod(showwarning)
    showerror = staticmethod(showerror)
    askyesno = staticmethod(askyesno)
    askokcancel = staticmethod(askokcancel)
    askquestion = staticmethod(askquestion)
    askretrycancel = staticmethod(askretrycancel)
    show_toast = staticmethod(show_toast)
    highlight_missing_field = staticmethod(highlight_missing_field)


def patch_tkinter_messagebox():
    """
    Globally patches tkinter.messagebox so that all existing dialog calls across
    the entire application automatically render as modern CustomTkinter dialogs.
    """
    try:
        import tkinter.messagebox as mb
        mb.showinfo = showinfo
        mb.showwarning = showwarning
        mb.showerror = showerror
        mb.askyesno = askyesno
        mb.askokcancel = askokcancel
        mb.askquestion = askquestion
        mb.askretrycancel = askretrycancel
        mb.highlight_missing_field = highlight_missing_field
    except Exception:
        pass
