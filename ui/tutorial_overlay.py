"""
Interactive Thai Onboarding & Guided Tutorial System for SPVC Smart Receipt System.

Provides:
- TutorialWelcomeDialog: First-launch friendly Thai modal asking whether to start the guided tour.
- TutorialStep / TUTORIAL_STEPS: Structured step definitions covering every page in the app.
- TutorialOverlay: Floating AppTheme-styled guide card with target highlight ring, progress
  indicator, clickable step dots, section jump menu, and navigation controls.
- InteractiveTutorialManager: Orchestrates role/config filtering, page navigation callbacks,
  widget highlighting, and completion/skip state.
"""

from dataclasses import dataclass
import inspect
import tkinter as tk
from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple
import customtkinter as ctk

from ui.icon_manager import IconManager
from ui.theme import AppTheme


@dataclass(frozen=True)
class TutorialStep:
    """Represents a single guided step in the interactive Thai tutorial."""

    step_id: str
    page_key: str
    page_title: str
    title: str
    description: str
    tips: Tuple[str, ...] = ()
    target_attr: Optional[str] = None
    allowed_roles: Optional[Tuple[str, ...]] = None


TUTORIAL_STEPS: Tuple[TutorialStep, ...] = (
    TutorialStep(
        step_id="dashboard_overview",
        page_key="dashboard",
        page_title="ภาพรวมการเงิน",
        title="1. แดชบอร์ดภาพรวมการเงิน (Dashboard)",
        description=(
            "หน้าแรกสำหรับผู้บริหารและเจ้าหน้าที่การเงิน เพื่อติดตามยอดรับชำระวันนี้ "
            "ยอดรวมประจำเดือน ยอดจัดเก็บสะสม และยอดค้างชำระค่าเทอมแบบเรียลไทม์"
        ),
        tips=(
            "ใช้แถบตัวกรองด้านบนเพื่อดูข้อมูลแยกตามภาค/ปีการศึกษา ระดับชั้น แผนกวิชา หรือช่วงเวลา",
            "กดปุ่ม 'รีเฟรชข้อมูล' เพื่ออัปเดตตัวเลขล่าสุด หรือ 'ส่งออก Excel' เพื่อทำรายงาน",
        ),
        target_attr="filter_bar",
    ),
    TutorialStep(
        step_id="dashboard_analytics",
        page_key="dashboard",
        page_title="ภาพรวมการเงิน",
        title="2. ช่องทางชำระเงิน & ใบเสร็จล่าสุด",
        description=(
            "ส่วนล่างของแดชบอร์ดแสดงสัดส่วนช่องทางการชำระเงิน ความคืบหน้าการเก็บค่าเทอม "
            "5 อันดับค่าธรรมเนียมยอดนิยม และรายการใบเสร็จล่าสุด"
        ),
        tips=(
            "สามารถกดปุ่ม 'ดูทั้งหมด' ในตารางใบเสร็จล่าสุดเพื่อกระโดดไปยังหน้าประวัติใบเสร็จได้ทันที",
        ),
        target_attr="recent_table",
    ),
    TutorialStep(
        step_id="pos_normal",
        page_key="pos",
        page_title="ออกใบเสร็จรับเงิน (POS)",
        title="3. ออกใบเสร็จรับเงินรายคน (POS)",
        description=(
            "หน้าหลักสำหรับรับชำระเงินและออกใบเสร็จรายบุคคล ค้นหานักศึกษา เลือกรายการค่าธรรมเนียม "
            "กำหนดจำนวนและวิธีชำระเงิน จากนั้นบันทึกและพิมพ์ใบเสร็จ PDF ได้ในคลิกเดียว"
        ),
        tips=(
            "เลือกวิธีชำระเงิน (เงินสด / โอนเงิน / QR PromptPay) และตรวจสอบโฟลเดอร์ปลายทางก่อนกดพิมพ์",
            "ปุ่ม 'บันทึกข้อมูล' จะบันทึกโดยไม่เปิดไฟล์ ส่วน 'บันทึก & พิมพ์ใบเสร็จ (.pdf)' จะสร้างไฟล์พร้อมเปิดทันที",
        ),
        target_attr="pos_form",
    ),
    TutorialStep(
        step_id="pos_advanced",
        page_key="pos",
        page_title="ออกใบเสร็จรับเงิน (POS)",
        title="4. โหมดขั้นสูงใน POS (ออกแบบ & แอดวานซ์)",
        description=(
            "ด้านบนของหน้า POS สามารถสลับโหมดเป็น 'ออกใบเสร็จยกห้อง' หรือ 'ออกแบบ & แอดวานซ์' "
            "เพื่อใช้งาน QR PromptPay, ชุดค่าธรรมเนียมสำเร็จรูป, ออกแบบเทมเพลต และ Webhook แจ้งเตือน"
        ),
        tips=(
            "เหมาะสำหรับจัดชุดค่าธรรมเนียมที่ใช้บ่อย เพื่อลดเวลาในการกรอกรายการซ้ำๆ",
        ),
        target_attr="mode_switcher",
    ),
    TutorialStep(
        step_id="batch_pos",
        page_key="batch_pos",
        page_title="จัดใบเสร็จทั้งห้อง",
        title="5. จัดใบเสร็จทั้งห้อง (Batch Classroom)",
        description=(
            "ออกใบเสร็จพร้อมกันทั้งห้องเรียนอย่างรวดเร็ว ด้านซ้ายเลือกนำเข้าไฟล์ Excel/CSV "
            "หรือดึงรายชื่อห้องที่มีในระบบ ด้านขวากำหนดรายการค่าใช้จ่ายเดียวกันทั้งห้อง"
        ),
        tips=(
            "สามารถกด 'ดาวน์โหลดตัวอย่าง Excel' เพื่อดูรูปแบบไฟล์นำเข้าที่ถูกต้อง",
            "กด 'ดึงค่าเทอมมาตรฐาน' เพื่อโหลดรายการค่าธรรมเนียมตามหลักสูตร แล้วกด 'ออกใบเสร็จทั้งห้อง (.pdf / .docx)'",
        ),
        target_attr="batch_workspace",
    ),
    TutorialStep(
        step_id="gallery",
        page_key="gallery",
        page_title="คลังภาพใบเสร็จ",
        title="6. คลังภาพใบเสร็จ (Receipt Gallery)",
        description=(
            "ศูนย์รวมไฟล์ภาพและใบเสร็จที่บันทึกไว้ ค้นหาตามเลขที่ใบเสร็จ ชื่อ ห้อง ชั้นปี หรือวันที่ "
            "พร้อมหน้าต่างแสดงตัวอย่างภาพที่ซูมเข้า-ออกและเลื่อนดูรายละเอียดได้"
        ),
        tips=(
            "สลับมุมมองแบบตารางภาพย่อหรือแบบรายการ (☰ รายการ) และติ๊ก 'รวมโฟลเดอร์ย่อย' เมื่อต้องการค้นหาทั้งหมด",
            "ด้านขวาสามารถกด 'บันทึกเป็น...', 'ย้ายโฟลเดอร์', 'เปลี่ยนชื่อ' หรือเปิดโฟลเดอร์ไฟล์ต้นฉบับได้",
        ),
        target_attr="gallery_split",
    ),
    TutorialStep(
        step_id="receipts",
        page_key="receipts",
        page_title="ประวัติใบเสร็จ",
        title="7. ประวัติใบเสร็จ & การยกเลิก (Void)",
        description=(
            "ตรวจสอบประวัติใบเสร็จทั้งหมด กรองตามวันที่หรือคำค้นหา พิมพ์ใบเสร็จซ้ำ (.pdf) "
            "ส่งออกรายงานประจำวัน และยกเลิกใบเสร็จที่ออกผิดพลาดพร้อมระบุเหตุผล"
        ),
        tips=(
            "ระบบบันทึก Audit Log สำหรับทุกการแก้ไข/ยกเลิก เพื่อความโปร่งใสในการตรวจสอบบัญชี",
            "สามารถกด 'ส่งออกรายงาน Excel' หรือ 'พิมพ์รายงานประจำวัน (.pdf)' จากแถบด้านบน",
        ),
        target_attr="receipts_table",
    ),
    TutorialStep(
        step_id="students",
        page_key="students",
        page_title="ฐานข้อมูลนักศึกษา",
        title="8. ฐานข้อมูลนักศึกษา",
        description=(
            "จัดการรายชื่อนักศึกษาทั้งหมดในวิทยาลัย รองรับการนำเข้าข้อมูลจำนวนมากจากไฟล์ Excel / CSV "
            "การค้นหาแยกตามรุ่นปี ระดับชั้น ห้อง และแผนกวิชา"
        ),
        tips=(
            "คลิกที่แถวรายชื่อนักศึกษาในตารางเพื่อดึงข้อมูลลงฟอร์มด้านล่างสำหรับแก้ไขหรือลบรายการ",
        ),
        target_attr="students_table",
    ),
    TutorialStep(
        step_id="registration",
        page_key="registration",
        page_title="ทะเบียนวิชา (อศท.28)",
        title="9. ลงทะเบียนเรียน (อศท.28)",
        description=(
            "จัดทำบัตรลงทะเบียนรายวิชา (อศท.28) โหลดวิชาตามหลักสูตรต้นฉบับ 11 วิชา "
            "คำนวณชั่วโมงทฤษฎี/ปฏิบัติ/หน่วยกิต และค่าธรรมเนียมรวมให้อัตโนมัติ"
        ),
        tips=(
            "เพิ่มรายวิชาได้สูงสุด 15 วิชา เลือกรูปแบบเทมเพลต แล้วกด 'บันทึก & พิมพ์ อศท.28 (.pdf)'",
            "เมนูนี้สามารถเปิด/ปิดการแสดงผลได้ในหน้า ตั้งค่าระบบ > ข้อมูลสถานศึกษา & ใบเสร็จ",
        ),
        target_attr="registration_form",
    ),
    TutorialStep(
        step_id="settings",
        page_key="settings",
        page_title="ตั้งค่าระบบ",
        title="10. ตั้งค่าระบบ & สำรองข้อมูล",
        description=(
            "กำหนดข้อมูลสถานศึกษา เลขเล่ม/คำนำหน้าใบเสร็จ จัดการรายการค่าธรรมเนียม "
            "ตัวเลือก Dropdown และสำรองข้อมูลฐานข้อมูลเป็นไฟล์ .zip เพื่อความปลอดภัย"
        ),
        tips=(
            "อย่าลืมกด 'บันทึกการตั้งค่าทั้งหมด' หลังแก้ไขข้อมูล และกด 'สำรองข้อมูลด่วน (.zip)' เป็นประจำ",
            "หากต้องการเปิดดูคู่มือสอนใช้งานอีกครั้ง กดปุ่ม '❓ สอนใช้งาน' ที่แถบเมนูด้านซ้ายได้ทุกเมื่อ",
        ),
        target_attr="settings_tabs",
    ),
)


class TutorialWelcomeDialog(ctk.CTkToplevel):
    """
    Friendly Thai modal shown on first launch asking if the user wants
    an interactive step-by-step walkthrough.
    """

    def __init__(
        self,
        parent: Optional[Any] = None,
        on_start: Optional[Callable[[], None]] = None,
        on_skip: Optional[Callable[..., None]] = None,
        config: Optional[Dict[str, Any]] = None,
        title: str = "ยินดีต้อนรับสู่ระบบออกใบเสร็จรับเงิน",
    ):
        super().__init__(parent)
        self._parent = parent
        self._on_start = on_start
        self._on_skip = on_skip
        self._config = config
        self.result: Optional[str] = None
        self.dont_ask_var = ctk.BooleanVar(value=True)
        self.dont_ask_again: bool = True

        self.withdraw()
        self.title(title)
        self.resizable(False, False)
        self.configure(fg_color=AppTheme.COLOR_SURFACE)
        IconManager.set_window_icon(self)

        if parent is not None:
            try:
                self.transient(parent)
            except Exception:
                pass

        self._build_ui()
        self._center_over_parent()

        self.protocol("WM_DELETE_WINDOW", self._handle_skip)
        self.bind("<Return>", lambda _e: self._handle_start())
        self.bind("<Escape>", lambda _e: self._handle_skip())

        try:
            self.deiconify()
            self.lift()
            self.focus_force()
            self.grab_set()
        except Exception:
            pass

    def _build_ui(self) -> None:
        # Top executive navy accent bar
        accent_bar = ctk.CTkFrame(
            self,
            fg_color=AppTheme.COLOR_ACCENT,
            height=6,
            corner_radius=AppTheme.RADIUS,
        )
        accent_bar.pack(fill="x", side="top")

        body = ctk.CTkFrame(
            self,
            fg_color=AppTheme.COLOR_SURFACE,
            corner_radius=AppTheme.RADIUS,
        )
        body.pack(fill="both", expand=True, padx=AppTheme.SPACE_LG, pady=AppTheme.SPACE_LG)

        badge = ctk.CTkLabel(
            body,
            text="  คู่มือเริ่มต้นใช้งาน (Interactive Guide)  ",
            font=AppTheme.FONT_CAPTION,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            height=26,
            corner_radius=AppTheme.RADIUS,
        )
        badge.pack(anchor="w", pady=(0, AppTheme.SPACE_SM))

        self.lbl_question = ctk.CTkLabel(
            body,
            text="ต้องการให้ระบบสอนการใช้งานแบบทีละขั้นตอนไหมครับ?",
            font=AppTheme.FONT_TITLE,
            text_color=AppTheme.COLOR_TEXT,
            wraplength=460,
            justify="left",
        )
        self.lbl_question.pack(anchor="w", pady=(0, AppTheme.SPACE_SM))

        self.lbl_subtitle = ctk.CTkLabel(
            body,
            text=(
                "ระบบจะพาคุณแนะนำทุกหน้าจอสำคัญทีละขั้นตอน ตั้งแต่ภาพรวมการเงิน (Dashboard), "
                "การออกใบเสร็จรายคนและยกห้อง, คลังภาพใบเสร็จ, ประวัติ/ยกเลิกใบเสร็จ, "
                "ฐานข้อมูลนักศึกษา ไปจนถึงการตั้งค่าระบบและสำรองข้อมูล"
            ),
            font=AppTheme.FONT_BODY,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            wraplength=460,
            justify="left",
        )
        self.lbl_subtitle.pack(anchor="w", pady=(0, AppTheme.SPACE_MD))

        hint_box = ctk.CTkFrame(
            body,
            fg_color=AppTheme.COLOR_SURFACE_ALT,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
            corner_radius=AppTheme.RADIUS,
        )
        hint_box.pack(fill="x", pady=(0, AppTheme.SPACE_MD))

        self.lbl_reopen_hint = ctk.CTkLabel(
            hint_box,
            text='สามารถเปิดคู่มือสอนใช้งานได้ทุกเมื่อที่ปุ่ม "❓ สอนใช้งาน" ในแถบเมนูด้านซ้าย',
            font=AppTheme.FONT_CAPTION,
            text_color=AppTheme.COLOR_TEXT_MUTED,
            wraplength=440,
            justify="left",
        )
        self.lbl_reopen_hint.pack(anchor="w", padx=AppTheme.SPACE_SM, pady=AppTheme.SPACE_SM)

        self.chk_dont_ask = ctk.CTkCheckBox(
            body,
            text="ไม่ต้องถามอีกตอนเปิดโปรแกรม",
            variable=self.dont_ask_var,
            font=AppTheme.FONT_BODY,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            border_color=AppTheme.COLOR_BORDER_STRONG,
            corner_radius=AppTheme.RADIUS,
        )
        self.chk_dont_ask.pack(anchor="w", pady=(0, AppTheme.SPACE_MD))

        btn_row = ctk.CTkFrame(body, fg_color="transparent")
        btn_row.pack(fill="x", pady=(AppTheme.SPACE_XS, 0))

        self.btn_skip = ctk.CTkButton(
            btn_row,
            text="ข้ามไปก่อน",
            font=AppTheme.FONT_LABEL,
            height=AppTheme.CTRL_H_MD,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER_STRONG,
            corner_radius=AppTheme.RADIUS_BTN,
            command=self._handle_skip,
        )
        self.btn_skip.pack(side="left")

        self.btn_start = ctk.CTkButton(
            btn_row,
            text="เริ่มสอนใช้งานเลย (แนะนำ)",
            font=AppTheme.FONT_LABEL,
            height=AppTheme.CTRL_H_MD,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color=AppTheme.COLOR_TEXT_LIGHT,
            corner_radius=AppTheme.RADIUS_BTN,
            command=self._handle_start,
        )
        self.btn_start.pack(side="right")

    def _center_over_parent(self) -> None:
        try:
            self.update_idletasks()
            w = max(520, self.winfo_reqwidth())
            h = max(310, self.winfo_reqheight())
            if self._parent is not None and self._parent.winfo_exists():
                px = self._parent.winfo_rootx()
                py = self._parent.winfo_rooty()
                pw = self._parent.winfo_width()
                ph = self._parent.winfo_height()
                if pw > 1 and ph > 1:
                    x = px + (pw - w) // 2
                    y = py + (ph - h) // 2
                    self.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")
                    return
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            x = max(0, (sw - w) // 2)
            y = max(0, (sh - h) // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

    def _revert_withdraw_after_windows_set_titlebar_color(self) -> None:
        try:
            if self.winfo_exists():
                super()._revert_withdraw_after_windows_set_titlebar_color()
        except Exception:
            pass

    def focus_set(self) -> None:
        try:
            if hasattr(self, "winfo_exists") and self.winfo_exists():
                super().focus_set()
        except Exception:
            pass

    def destroy(self) -> None:
        try:
            if hasattr(self, "winfo_exists") and self.winfo_exists():
                super().destroy()
        except Exception:
            pass

    def _safe_destroy(self) -> None:
        try:
            self.grab_release()
        except Exception:
            pass
        try:
            if self.winfo_exists():
                self.withdraw()
                self.destroy()
        except Exception:
            pass

    def _handle_start(self) -> None:
        self.result = "start"
        self.dont_ask_again = bool(self.dont_ask_var.get())
        self._safe_destroy()
        if self._on_start is not None:
            self._on_start()

    def _handle_skip(self) -> None:
        self.result = "skip"
        self.dont_ask_again = bool(self.dont_ask_var.get())
        if isinstance(self._config, dict) and self.dont_ask_again:
            self._config["tutorial_completed"] = True
        self._safe_destroy()
        if self._on_skip is not None:
            self._invoke_skip_callback(self._on_skip, self.dont_ask_again)

    @staticmethod
    def _invoke_skip_callback(cb: Callable[..., None], dont_ask_again: bool) -> None:
        try:
            sig = inspect.signature(cb)
            params = [
                p
                for p in sig.parameters.values()
                if p.kind
                in (
                    inspect.Parameter.POSITIONAL_ONLY,
                    inspect.Parameter.POSITIONAL_OR_KEYWORD,
                )
            ]
            has_varargs = any(
                p.kind == inspect.Parameter.VAR_POSITIONAL
                for p in sig.parameters.values()
            )
            if has_varargs or len(params) >= 1:
                cb(dont_ask_again)
            else:
                cb()
        except (TypeError, ValueError):
            cb()


class TutorialOverlay(ctk.CTkFrame):
    """
    Floating AppTheme-styled guided tour card with a 4-panel dimmed backdrop,
    spotlight cutout + highlight ring around the target widget, smart anchoring
    next to the target, readable typography, click-anywhere-to-continue,
    clickable step dots, section jump menu, and navigation buttons.
    """

    HIGHLIGHT_THICKNESS = 4
    SPOTLIGHT_PADDING = 10
    SPOTLIGHT_RADIUS = 18
    DIM_COLOR = "#0B1120"
    SCRIM_ALPHA = 0.60
    TRANSPARENT_KEY = "#FF00FE"
    SPOTLIGHT_RING_COLOR = "#38BDF8"
    CARD_WIDTH = 500

    FONT_CARD_BADGE = (AppTheme.FONT_FAMILY, 13, "bold")
    FONT_CARD_PROGRESS = (AppTheme.FONT_FAMILY, 13)
    FONT_CARD_TITLE = (AppTheme.FONT_FAMILY, 19, "bold")
    FONT_CARD_BODY = (AppTheme.FONT_FAMILY, 15)
    FONT_CARD_TIPS = (AppTheme.FONT_FAMILY, 14)
    FONT_CARD_HINT = (AppTheme.FONT_FAMILY, 13, "bold")

    def __init__(
        self,
        root: Any,
        on_prev: Optional[Callable[[], None]] = None,
        on_next: Optional[Callable[[], None]] = None,
        on_skip: Optional[Callable[[], None]] = None,
        on_complete: Optional[Callable[[], None]] = None,
        on_jump: Optional[Callable[[int], None]] = None,
    ):
        super().__init__(
            root,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=2,
            border_color=AppTheme.COLOR_ACCENT,
            corner_radius=AppTheme.RADIUS_CARD,
        )
        self._host_root = root
        self._on_prev = on_prev
        self._on_next = on_next
        self._on_skip = on_skip
        self._on_complete = on_complete
        self._on_jump = on_jump
        self._is_last_step = False
        self._step_dots: List[ctk.CTkButton] = []
        self._jump_labels: List[str] = []
        self._dim_panels: List[ctk.CTkFrame] = []
        self._highlight_bars: List[ctk.CTkFrame] = []
        self._scrim: Optional[tk.Toplevel] = None
        self._scrim_canvas: Optional[tk.Canvas] = None
        self._scrim_supported: bool = False
        self._last_target_widget: Optional[Any] = None
        self._last_root_size: Tuple[int, int] = (0, 0)
        self._configure_bind_id: Optional[str] = None
        self._map_bind_id: Optional[str] = None
        self._unmap_bind_id: Optional[str] = None

        self._build_scrim()
        if not self._scrim_supported:
            self._build_dim_panels()
            self._build_highlight_bars()
        self._build_card_ui()
        self._bind_root_events()

    def _build_scrim(self) -> None:
        try:
            scrim = tk.Toplevel(self._host_root)
            scrim.withdraw()
            scrim.overrideredirect(True)
            scrim.transient(self._host_root)
            scrim.configure(bg=self.TRANSPARENT_KEY)
            scrim.attributes("-alpha", self.SCRIM_ALPHA)
            scrim.attributes("-transparentcolor", self.TRANSPARENT_KEY)

            canvas = tk.Canvas(
                scrim,
                bg=self.TRANSPARENT_KEY,
                highlightthickness=0,
                bd=0,
                cursor="hand2",
            )
            canvas.pack(fill="both", expand=True)
            canvas.bind("<Button-1>", self._handle_backdrop_click)

            self._scrim = scrim
            self._scrim_canvas = canvas
            self._scrim_supported = True
        except Exception:
            if self._scrim is not None:
                try:
                    self._scrim.destroy()
                except Exception:
                    pass
            self._scrim = None
            self._scrim_canvas = None
            self._scrim_supported = False

    def _bind_root_events(self) -> None:
        try:
            self._configure_bind_id = self._host_root.bind(
                "<Configure>", self._on_root_configure, add="+"
            )
            self._map_bind_id = self._host_root.bind(
                "<Map>", self._on_root_map, add="+"
            )
            self._unmap_bind_id = self._host_root.bind(
                "<Unmap>", self._on_root_unmap, add="+"
            )
        except Exception:
            pass

    def _on_root_map(self, _event: Any = None) -> None:
        try:
            if self.winfo_exists() and self.winfo_ismapped():
                self._layout_spotlight_and_card(self._last_target_widget)
        except Exception:
            pass

    def _on_root_unmap(self, _event: Any = None) -> None:
        try:
            if self._scrim is not None and self._scrim.winfo_exists():
                self._scrim.withdraw()
        except Exception:
            pass

    def _on_root_configure(self, _event: Any = None) -> None:
        try:
            if not self.winfo_exists() or not self.winfo_ismapped():
                return
            rw = self._host_root.winfo_width()
            rh = self._host_root.winfo_height()
            rx = self._host_root.winfo_rootx()
            ry = self._host_root.winfo_rooty()
            if (rw, rh) != self._last_root_size and rw > 1 and rh > 1:
                self._last_root_size = (rw, rh)
                self._layout_spotlight_and_card(self._last_target_widget)
            elif self._scrim is not None and self._scrim.winfo_exists() and self._host_root.winfo_viewable():
                self._scrim.geometry(f"{rw}x{rh}+{rx}+{ry}")
        except Exception:
            pass

    def _build_dim_panels(self) -> None:
        for _ in range(4):
            panel = ctk.CTkFrame(
                self._host_root,
                fg_color=self.DIM_COLOR,
                corner_radius=0,
                border_width=0,
            )
            try:
                panel.configure(cursor="hand2")
            except Exception:
                pass
            panel.bind("<Button-1>", self._handle_backdrop_click)
            self._dim_panels.append(panel)

    def _build_highlight_bars(self) -> None:
        for _ in range(4):
            bar = ctk.CTkFrame(
                self._host_root,
                fg_color=self.SPOTLIGHT_RING_COLOR,
                corner_radius=0,
                border_width=0,
            )
            self._highlight_bars.append(bar)

    def _bind_click_to_advance(self, widget: Any) -> None:
        try:
            widget.bind("<Button-1>", self._handle_backdrop_click)
        except Exception:
            pass

    def _build_card_ui(self) -> None:
        self._bind_click_to_advance(self)

        # Top Navy header strip
        header = ctk.CTkFrame(
            self,
            fg_color=AppTheme.COLOR_ACCENT,
            corner_radius=AppTheme.RADIUS,
            height=42,
        )
        header.pack(fill="x", side="top")
        self._bind_click_to_advance(header)

        self.lbl_page_badge = ctk.CTkLabel(
            header,
            text="หน้า: -",
            font=self.FONT_CARD_BADGE,
            text_color=AppTheme.COLOR_TEXT_LIGHT,
        )
        self.lbl_page_badge.pack(side="left", padx=AppTheme.SPACE_MD, pady=AppTheme.SPACE_SM)
        self._bind_click_to_advance(self.lbl_page_badge)

        self.lbl_progress = ctk.CTkLabel(
            header,
            text="ขั้นตอนที่ 1 / 1",
            font=self.FONT_CARD_PROGRESS,
            text_color=AppTheme.COLOR_SECONDARY_BG,
        )
        self.lbl_progress.pack(side="right", padx=AppTheme.SPACE_MD, pady=AppTheme.SPACE_SM)
        self._bind_click_to_advance(self.lbl_progress)

        content = ctk.CTkFrame(self, fg_color=AppTheme.COLOR_SURFACE, corner_radius=0)
        content.pack(fill="both", expand=True, padx=AppTheme.SPACE_MD, pady=AppTheme.SPACE_MD)
        self._bind_click_to_advance(content)

        # Jump menu row
        jump_row = ctk.CTkFrame(content, fg_color="transparent")
        jump_row.pack(fill="x", pady=(0, AppTheme.SPACE_SM))

        lbl_jump = ctk.CTkLabel(
            jump_row,
            text="ข้ามไปยังหัวข้อ:",
            font=AppTheme.FONT_CAPTION,
            text_color=AppTheme.COLOR_TEXT_MUTED,
        )
        lbl_jump.pack(side="left", padx=(0, AppTheme.SPACE_SM))

        self.jump_var = ctk.StringVar(value="")
        self.jump_menu = ctk.CTkOptionMenu(
            jump_row,
            variable=self.jump_var,
            values=["-"],
            height=30,
            font=AppTheme.FONT_CAPTION,
            fg_color=AppTheme.COLOR_DROPDOWN_FG,
            button_color=AppTheme.COLOR_DROPDOWN_BTN,
            button_hover_color=AppTheme.COLOR_DROPDOWN_HOVER,
            dropdown_fg_color=AppTheme.COLOR_DROPDOWN_MENU,
            text_color=AppTheme.COLOR_DROPDOWN_TEXT,
            corner_radius=AppTheme.RADIUS,
            command=self._handle_jump_menu,
        )
        self.jump_menu.pack(side="left", fill="x", expand=True)

        self.lbl_title = ctk.CTkLabel(
            content,
            text="",
            font=self.FONT_CARD_TITLE,
            text_color=AppTheme.COLOR_TEXT,
            wraplength=452,
            justify="left",
        )
        self.lbl_title.pack(anchor="w", pady=(0, AppTheme.SPACE_XS))
        self._bind_click_to_advance(self.lbl_title)

        self.lbl_description = ctk.CTkLabel(
            content,
            text="",
            font=self.FONT_CARD_BODY,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            wraplength=452,
            justify="left",
        )
        self.lbl_description.pack(anchor="w", pady=(0, AppTheme.SPACE_SM))
        self._bind_click_to_advance(self.lbl_description)

        self.tips_frame = ctk.CTkFrame(
            content,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            corner_radius=AppTheme.RADIUS,
        )
        self._bind_click_to_advance(self.tips_frame)
        self.lbl_tips = ctk.CTkLabel(
            self.tips_frame,
            text="",
            font=self.FONT_CARD_TIPS,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            wraplength=432,
            justify="left",
        )
        self.lbl_tips.pack(anchor="w", padx=AppTheme.SPACE_SM, pady=AppTheme.SPACE_SM)
        self._bind_click_to_advance(self.lbl_tips)

        # Clickable step dots row
        self.dots_frame = ctk.CTkFrame(content, fg_color="transparent")
        self.dots_frame.pack(fill="x", pady=(AppTheme.SPACE_SM, AppTheme.SPACE_XS))
        self._bind_click_to_advance(self.dots_frame)

        # Click-anywhere hint
        self.lbl_click_hint = ctk.CTkLabel(
            content,
            text="คลิกที่ใดก็ได้เพื่อไปต่อ",
            font=self.FONT_CARD_HINT,
            text_color=AppTheme.COLOR_PRIMARY_BLUE,
        )
        self.lbl_click_hint.pack(anchor="w", pady=(0, AppTheme.SPACE_SM))
        self._bind_click_to_advance(self.lbl_click_hint)

        # Action controls row
        controls = ctk.CTkFrame(content, fg_color="transparent")
        controls.pack(fill="x", pady=(AppTheme.SPACE_XS, 0))
        self._bind_click_to_advance(controls)

        self.btn_skip = ctk.CTkButton(
            controls,
            text="ข้าม",
            font=AppTheme.FONT_LABEL,
            height=AppTheme.CTRL_H_SM,
            width=72,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_MUTED,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER_STRONG,
            corner_radius=AppTheme.RADIUS_BTN,
            command=self._handle_skip,
        )
        self.btn_skip.pack(side="left")

        nav_right = ctk.CTkFrame(controls, fg_color="transparent")
        nav_right.pack(side="right")

        self.btn_prev = ctk.CTkButton(
            nav_right,
            text="← ย้อนกลับ",
            font=AppTheme.FONT_LABEL,
            height=AppTheme.CTRL_H_SM,
            width=104,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER_STRONG,
            corner_radius=AppTheme.RADIUS_BTN,
            command=self._handle_prev,
        )
        self.btn_prev.pack(side="left", padx=(0, AppTheme.SPACE_SM))

        self.btn_next = ctk.CTkButton(
            nav_right,
            text="ถัดไป →",
            font=AppTheme.FONT_LABEL,
            height=AppTheme.CTRL_H_SM,
            width=110,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color=AppTheme.COLOR_TEXT_LIGHT,
            corner_radius=AppTheme.RADIUS_BTN,
            command=self._handle_next,
        )
        self.btn_next.pack(side="left")

    def show_step(
        self,
        step: TutorialStep,
        index: int,
        total: int,
        all_steps: List[TutorialStep],
        target_widget: Optional[Any] = None,
    ) -> None:
        """Renders the given tutorial step, updates dots/menu, and highlights target."""
        self._is_last_step = index >= total - 1
        self.lbl_page_badge.configure(text=f"หน้า: {step.page_title}")
        self.lbl_progress.configure(text=f"ขั้นตอนที่ {index + 1} / {total}")
        self.lbl_title.configure(text=step.title)
        self.lbl_description.configure(text=step.description)

        if step.tips:
            tips_text = "\n".join(f"• {tip}" for tip in step.tips)
            self.lbl_tips.configure(text=tips_text)
            if not self.tips_frame.winfo_manager():
                self.tips_frame.pack(
                    fill="x",
                    before=self.dots_frame,
                    pady=(0, AppTheme.SPACE_SM),
                )
        else:
            if self.tips_frame.winfo_manager():
                self.tips_frame.pack_forget()

        # Update jump menu
        self._jump_labels = [
            f"{i + 1}. {s.page_title} — {s.title.split('. ', 1)[-1]}"
            for i, s in enumerate(all_steps)
        ]
        try:
            self.jump_menu.configure(values=self._jump_labels or ["-"])
            if 0 <= index < len(self._jump_labels):
                self.jump_var.set(self._jump_labels[index])
        except Exception:
            pass

        # Rebuild clickable step dots
        self._render_step_dots(index, total)

        # Prev / Next / Finish button states
        self.btn_prev.configure(state="disabled" if index <= 0 else "normal")
        if self._is_last_step:
            self.btn_next.configure(
                text="เสร็จสิ้น ✔",
                fg_color=AppTheme.COLOR_SUCCESS,
                hover_color=AppTheme.COLOR_SUCCESS_HOVER,
            )
            self.lbl_click_hint.configure(text="คลิกที่ใดก็ได้เพื่อเสร็จสิ้น")
        else:
            self.btn_next.configure(
                text="ถัดไป →",
                fg_color=AppTheme.COLOR_PRIMARY_BLUE,
                hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            )
            self.lbl_click_hint.configure(text="คลิกที่ใดก็ได้เพื่อไปต่อ")

        self._last_target_widget = target_widget
        self._layout_spotlight_and_card(target_widget)

    def _render_step_dots(self, current_index: int, total: int) -> None:
        if len(self._step_dots) == total:
            for i, dot in enumerate(self._step_dots):
                if i == current_index:
                    color = AppTheme.COLOR_PRIMARY_BLUE
                    hover = AppTheme.COLOR_PRIMARY_BLUE_HOVER
                    text_color = AppTheme.COLOR_TEXT_LIGHT
                elif i < current_index:
                    color = AppTheme.COLOR_SUCCESS
                    hover = AppTheme.COLOR_SUCCESS_HOVER
                    text_color = AppTheme.COLOR_TEXT_LIGHT
                else:
                    color = AppTheme.COLOR_BORDER_STRONG
                    hover = AppTheme.COLOR_SLATE
                    text_color = AppTheme.COLOR_TEXT
                dot.configure(fg_color=color, hover_color=hover, text_color=text_color)
            return

        for dot in self._step_dots:
            try:
                dot.destroy()
            except Exception:
                pass
        self._step_dots.clear()

        for i in range(total):
            if i == current_index:
                color = AppTheme.COLOR_PRIMARY_BLUE
                hover = AppTheme.COLOR_PRIMARY_BLUE_HOVER
                text_color = AppTheme.COLOR_TEXT_LIGHT
            elif i < current_index:
                color = AppTheme.COLOR_SUCCESS
                hover = AppTheme.COLOR_SUCCESS_HOVER
                text_color = AppTheme.COLOR_TEXT_LIGHT
            else:
                color = AppTheme.COLOR_BORDER_STRONG
                hover = AppTheme.COLOR_SLATE
                text_color = AppTheme.COLOR_TEXT

            dot = ctk.CTkButton(
                self.dots_frame,
                text=str(i + 1),
                width=24,
                height=22,
                font=AppTheme.FONT_CAPTION,
                fg_color=color,
                hover_color=hover,
                text_color=text_color,
                corner_radius=AppTheme.RADIUS,
                command=lambda idx=i: self._handle_dot_click(idx),
            )
            dot.pack(side="left", padx=(0, 4))
            self._step_dots.append(dot)

    def _layout_spotlight_and_card(self, target_widget: Optional[Any] = None) -> None:
        """Atomically measures geometry once, then updates scrim/spotlight and card in one pass."""
        try:
            self._host_root.update_idletasks()
            self.update_idletasks()
        except Exception:
            pass

        try:
            rw = self._host_root.winfo_width()
            rh = self._host_root.winfo_height()
            rx = self._host_root.winfo_rootx()
            ry = self._host_root.winfo_rooty()
        except Exception:
            return

        self._last_root_size = (rw, rh)
        card_w = min(self.CARD_WIDTH, max(320, rw - 24)) if rw > 40 else self.CARD_WIDTH
        card_h = max(self.winfo_reqheight(), 280)
        spotlight = self._get_spotlight_box(target_widget) if (rw > 1 and rh > 1) else None

        if rw > 320 and rh > 280:
            card_x, card_y = self._compute_card_position(rw, rh, card_w, card_h, spotlight)
            self.configure(width=card_w)
            self.place(x=card_x, y=card_y)
        else:
            card_x = max(0, (rw - card_w) // 2)
            card_y = max(0, (rh - card_h) // 2)
            self.place(relx=0.5, rely=0.5, anchor="center")

        if self._scrim_supported and self._scrim is not None and self._scrim_canvas is not None:
            self._render_scrim(rw, rh, rx, ry, spotlight, (card_x, card_y, card_w, card_h))
        else:
            self._render_fallback_panels(rw, rh, spotlight)

        try:
            self.lift()
        except Exception:
            pass

    @staticmethod
    def _rounded_rect_points(x1: int, y1: int, x2: int, y2: int, r: int) -> List[int]:
        r = max(4, min(r, max(4, (x2 - x1) // 2), max(4, (y2 - y1) // 2)))
        return [
            x1 + r, y1,
            x2 - r, y1,
            x2, y1,
            x2, y1 + r,
            x2, y2 - r,
            x2, y2,
            x2 - r, y2,
            x1 + r, y2,
            x1, y2,
            x1, y2 - r,
            x1, y1 + r,
            x1, y1,
        ]

    def _render_scrim(
        self,
        rw: int,
        rh: int,
        rx: int,
        ry: int,
        spotlight: Optional[Tuple[int, int, int, int]],
        card_rect: Tuple[int, int, int, int],
    ) -> None:
        if self._scrim is None or self._scrim_canvas is None or rw <= 1 or rh <= 1:
            return
        try:
            canvas = self._scrim_canvas
            canvas.delete("all")
            canvas.create_rectangle(0, 0, rw, rh, fill=self.DIM_COLOR, outline="", tags="backdrop")

            if spotlight is not None:
                sx, sy, sw, sh = spotlight
                pts = self._rounded_rect_points(sx, sy, sx + sw, sy + sh, self.SPOTLIGHT_RADIUS)
                canvas.create_polygon(
                    pts,
                    smooth=True,
                    fill=self.TRANSPARENT_KEY,
                    outline=self.SPOTLIGHT_RING_COLOR,
                    width=self.HIGHLIGHT_THICKNESS,
                    tags="spotlight",
                )

            cx, cy, cw, ch = card_rect
            canvas.create_rectangle(
                cx, cy, cx + cw, cy + ch,
                fill=self.TRANSPARENT_KEY,
                outline="",
                tags="card_cutout",
            )

            self._scrim.geometry(f"{rw}x{rh}+{rx}+{ry}")
            if self._host_root.winfo_viewable():
                self._scrim.deiconify()
                self._scrim.lift()
            else:
                self._scrim.withdraw()
        except Exception:
            pass

    def _render_fallback_panels(
        self,
        rw: int,
        rh: int,
        spotlight: Optional[Tuple[int, int, int, int]],
    ) -> None:
        if rw <= 1 or rh <= 1 or not self._dim_panels:
            return
        try:
            top_dim, bottom_dim, left_dim, right_dim = self._dim_panels
            if spotlight is None:
                self._place_rect(top_dim, 0, 0, rw, rh)
                for p in (bottom_dim, left_dim, right_dim):
                    p.place_forget()
                for b in self._highlight_bars:
                    b.place_forget()
                top_dim.lift()
                return

            sx, sy, sw, sh = spotlight
            if sy > 0:
                self._place_rect(top_dim, 0, 0, rw, sy)
                top_dim.lift()
            else:
                top_dim.place_forget()

            if sy + sh < rh:
                self._place_rect(bottom_dim, 0, sy + sh, rw, rh - (sy + sh))
                bottom_dim.lift()
            else:
                bottom_dim.place_forget()

            if sx > 0:
                self._place_rect(left_dim, 0, sy, sx, sh)
                left_dim.lift()
            else:
                left_dim.place_forget()

            if sx + sw < rw:
                self._place_rect(right_dim, sx + sw, sy, rw - (sx + sw), sh)
                right_dim.lift()
            else:
                right_dim.place_forget()

            t = self.HIGHLIGHT_THICKNESS
            top_bar, bottom_bar, left_bar, right_bar = self._highlight_bars
            self._place_rect(top_bar, sx - t, sy - t, sw + 2 * t, t)
            self._place_rect(bottom_bar, sx - t, sy + sh, sw + 2 * t, t)
            self._place_rect(left_bar, sx - t, sy, t, sh)
            self._place_rect(right_bar, sx + sw, sy, t, sh)
            for bar in self._highlight_bars:
                bar.lift()
        except Exception:
            self._clear_highlight()

    @staticmethod
    def _rect_intersection_area(
        ax: int, ay: int, aw: int, ah: int,
        bx: int, by: int, bw: int, bh: int,
    ) -> int:
        ix = max(0, min(ax + aw, bx + bw) - max(ax, bx))
        iy = max(0, min(ay + ah, by + bh) - max(ay, by))
        return ix * iy

    def _get_spotlight_box(self, target_widget: Optional[Any]) -> Optional[Tuple[int, int, int, int]]:
        if target_widget is None:
            return None
        try:
            if not target_widget.winfo_exists():
                return None
            rx = self._host_root.winfo_rootx()
            ry = self._host_root.winfo_rooty()
            rw = self._host_root.winfo_width()
            rh = self._host_root.winfo_height()
            tx = target_widget.winfo_rootx() - rx
            ty = target_widget.winfo_rooty() - ry
            tw = target_widget.winfo_width()
            th = target_widget.winfo_height()
            if tw <= 1 or th <= 1:
                return None
            pad = self.SPOTLIGHT_PADDING
            sx = max(0, tx - pad)
            sy = max(0, ty - pad)
            sx2 = min(rw, tx + tw + pad)
            sy2 = min(rh, ty + th + pad)
            if sx2 <= sx or sy2 <= sy:
                return None
            return (sx, sy, sx2 - sx, sy2 - sy)
        except Exception:
            return None

    def _compute_card_position(
        self,
        rw: int,
        rh: int,
        card_w: int,
        card_h: int,
        spotlight: Optional[Tuple[int, int, int, int]],
    ) -> Tuple[int, int]:
        margin = 16
        gap = 16
        max_x = max(0, rw - card_w - margin)
        max_y = max(0, rh - card_h - margin)

        if spotlight is None:
            return (max(0, (rw - card_w) // 2), max(0, (rh - card_h) // 2))

        sx, sy, sw, sh = spotlight
        candidates = [
            # Below target
            (max(margin, min(sx, max_x)), sy + sh + gap),
            # Above target
            (max(margin, min(sx, max_x)), sy - card_h - gap),
            # Right of target
            (sx + sw + gap, max(margin, min(sy, max_y))),
            # Left of target
            (sx - card_w - gap, max(margin, min(sy, max_y))),
        ]

        for cx, cy in candidates:
            if (
                margin <= cx <= rw - card_w - margin
                and margin <= cy <= rh - card_h - margin
                and self._rect_intersection_area(cx, cy, card_w, card_h, sx, sy, sw, sh) == 0
            ):
                return (cx, cy)

        # Fallback: clamp candidates + corners and pick minimal overlap with spotlight
        fallback_points = [
            (max(0, min(cx, max_x)), max(0, min(cy, max_y)))
            for cx, cy in candidates
        ] + [
            (max_x, margin),
            (margin, margin),
            (max_x, max_y),
        ]
        return min(
            fallback_points,
            key=lambda pt: self._rect_intersection_area(pt[0], pt[1], card_w, card_h, sx, sy, sw, sh),
        )

    def _place_card(self, target_widget: Optional[Any] = None) -> None:
        """Positions the floating tutorial card anchored next to the target widget."""
        self._layout_spotlight_and_card(target_widget)

    @staticmethod
    def _place_rect(widget: Any, x: int, y: int, w: int, h: int) -> None:
        widget.configure(width=max(1, int(w)), height=max(1, int(h)))
        widget.place(x=int(x), y=int(y))

    def _highlight_widget(self, target_widget: Optional[Any]) -> None:
        """Updates spotlight backdrop and ring around target_widget without clearing first."""
        try:
            rw = self._host_root.winfo_width()
            rh = self._host_root.winfo_height()
            rx = self._host_root.winfo_rootx()
            ry = self._host_root.winfo_rooty()
            spotlight = self._get_spotlight_box(target_widget)
            if self._scrim_supported and self._scrim is not None and self._scrim_canvas is not None:
                card_w = self.winfo_width() if self.winfo_width() > 1 else self.CARD_WIDTH
                card_h = self.winfo_height() if self.winfo_height() > 1 else max(self.winfo_reqheight(), 280)
                self._render_scrim(rw, rh, rx, ry, spotlight, (self.winfo_x(), self.winfo_y(), card_w, card_h))
            else:
                self._render_fallback_panels(rw, rh, spotlight)
        except Exception:
            self._clear_highlight()

    def _clear_highlight(self) -> None:
        if self._scrim is not None:
            try:
                if self._scrim.winfo_exists():
                    self._scrim.withdraw()
            except Exception:
                pass
        for panel in self._dim_panels:
            try:
                panel.place_forget()
            except Exception:
                pass
        for bar in self._highlight_bars:
            try:
                bar.place_forget()
            except Exception:
                pass

    def hide(self) -> None:
        """Hides the overlay card, dimmed backdrop, and spotlight ring without destroying."""
        self._clear_highlight()
        try:
            self.place_forget()
        except Exception:
            pass

    def destroy_overlay(self) -> None:
        """Destroys dimmed backdrop, spotlight ring, and the overlay card."""
        for seq, bind_id in (
            ("<Configure>", self._configure_bind_id),
            ("<Map>", self._map_bind_id),
            ("<Unmap>", self._unmap_bind_id),
        ):
            if bind_id is not None:
                try:
                    self._host_root.unbind(seq, bind_id)
                except Exception:
                    pass
        self._configure_bind_id = None
        self._map_bind_id = None
        self._unmap_bind_id = None

        if self._scrim is not None:
            try:
                if self._scrim.winfo_exists():
                    self._scrim.destroy()
            except Exception:
                pass
        self._scrim = None
        self._scrim_canvas = None

        for panel in self._dim_panels:
            try:
                panel.destroy()
            except Exception:
                pass
        self._dim_panels.clear()
        for bar in self._highlight_bars:
            try:
                bar.destroy()
            except Exception:
                pass
        self._highlight_bars.clear()
        try:
            if self.winfo_exists():
                self.destroy()
        except Exception:
            pass

    def _handle_backdrop_click(self, _event: Any = None) -> str:
        self._handle_next()
        return "break"

    def _handle_prev(self) -> None:
        if self._on_prev is not None:
            self._on_prev()

    def _handle_next(self) -> None:
        if self._is_last_step:
            if self._on_complete is not None:
                self._on_complete()
        elif self._on_next is not None:
            self._on_next()

    def _handle_skip(self) -> None:
        if self._on_skip is not None:
            self._on_skip()

    def _handle_dot_click(self, index: int) -> None:
        if self._on_jump is not None:
            self._on_jump(index)

    def _handle_jump_menu(self, selected: str) -> None:
        if self._on_jump is not None and selected in self._jump_labels:
            self._on_jump(self._jump_labels.index(selected))


class InteractiveTutorialManager:
    """
    Coordinates the multi-page Thai interactive tutorial:
    - Filters steps by visible pages, user role, and config flags.
    - Invokes navigate_callback(page_key) when switching pages.
    - Resolves target widgets for highlighting.
    - Persists completion/skip state into config when provided.
    """

    def __init__(
        self,
        root: Any,
        navigate_callback: Optional[Callable[[str], None]] = None,
        visible_pages: Optional[Iterable[str]] = None,
        user_role: str = "STAFF",
        config: Optional[Dict[str, Any]] = None,
        on_complete: Optional[Callable[[], None]] = None,
        on_skip: Optional[Callable[[], None]] = None,
        widget_resolver: Optional[Callable[[TutorialStep], Optional[Any]]] = None,
        steps: Optional[Iterable[TutorialStep]] = None,
    ):
        self.root = root
        self.navigate_callback = navigate_callback
        self.visible_pages = list(visible_pages) if visible_pages is not None else None
        self.user_role = user_role or "STAFF"
        self.config = config
        self.on_complete = on_complete
        self.on_skip = on_skip
        self.widget_resolver = widget_resolver

        base_steps = list(steps) if steps is not None else list(TUTORIAL_STEPS)
        self.steps: List[TutorialStep] = self.filter_steps(
            steps=base_steps,
            visible_pages=self.visible_pages,
            user_role=self.user_role,
            config=self.config,
        )
        self.current_index: int = 0
        self.is_active: bool = False
        self.overlay: Optional[TutorialOverlay] = None
        self._last_page_key: Optional[str] = None

    @staticmethod
    def filter_steps(
        steps: Iterable[TutorialStep],
        visible_pages: Optional[Iterable[str]] = None,
        user_role: str = "STAFF",
        config: Optional[Dict[str, Any]] = None,
    ) -> List[TutorialStep]:
        """Filters tutorial steps based on visible page keys, config, and user role."""
        visible_set = set(visible_pages) if visible_pages is not None else None
        role_upper = (user_role or "STAFF").upper()
        filtered: List[TutorialStep] = []

        for step in steps:
            if visible_set is not None and step.page_key not in visible_set:
                continue
            if (
                visible_set is None
                and isinstance(config, dict)
                and step.page_key == "registration"
                and not config.get("show_course_registration", False)
            ):
                continue
            if step.allowed_roles is not None:
                allowed = {r.upper() for r in step.allowed_roles}
                if role_upper not in allowed:
                    continue
            filtered.append(step)

        return filtered

    @property
    def total_steps(self) -> int:
        return len(self.steps)

    @property
    def current_step(self) -> Optional[TutorialStep]:
        if 0 <= self.current_index < len(self.steps):
            return self.steps[self.current_index]
        return None

    def start(self, start_index: int = 0) -> bool:
        """Starts the interactive tutorial at the given step index."""
        if not self.steps:
            return False

        if self.overlay is None:
            self.overlay = TutorialOverlay(
                root=self.root,
                on_prev=self.prev_step,
                on_next=self.next_step,
                on_skip=self.skip,
                on_complete=self.complete,
                on_jump=self.go_to_step,
            )

        self.is_active = True
        self._last_page_key = None
        clamped = max(0, min(start_index, len(self.steps) - 1))
        self.go_to_step(clamped)
        return True

    def go_to_step(self, index: int) -> None:
        """Navigates to the specified step index and updates the overlay."""
        if not self.steps:
            return
        if index < 0 or index >= len(self.steps):
            return

        self.current_index = index
        step = self.steps[index]

        if self.navigate_callback is not None and step.page_key != self._last_page_key:
            self.navigate_callback(step.page_key)
            self._last_page_key = step.page_key

        target_widget = self._resolve_target_widget(step)
        if self.overlay is not None:
            self.overlay.show_step(
                step=step,
                index=self.current_index,
                total=len(self.steps),
                all_steps=self.steps,
                target_widget=target_widget,
            )

    def next_step(self) -> None:
        """Advances to the next step, or completes the tutorial on the last step."""
        if not self.steps:
            return
        if self.current_index >= len(self.steps) - 1:
            self.complete()
        else:
            self.go_to_step(self.current_index + 1)

    def prev_step(self) -> None:
        """Moves back to the previous step."""
        if self.current_index > 0:
            self.go_to_step(self.current_index - 1)

    def skip(self) -> None:
        """Closes the tutorial early, marks completion in config, and invokes on_skip."""
        self.stop()
        if isinstance(self.config, dict):
            self.config["tutorial_completed"] = True
        if self.on_skip is not None:
            self.on_skip()

    def complete(self) -> None:
        """Finishes the tutorial, marks completion in config, and invokes on_complete."""
        self.stop()
        if isinstance(self.config, dict):
            self.config["tutorial_completed"] = True
        if self.on_complete is not None:
            self.on_complete()

    def stop(self) -> None:
        """Tears down the overlay cleanly without triggering completion callbacks."""
        self.is_active = False
        self._last_page_key = None
        if self.overlay is not None:
            self.overlay.destroy_overlay()
            self.overlay = None

    def _resolve_target_widget(self, step: TutorialStep) -> Optional[Any]:
        if self.widget_resolver is not None:
            try:
                return self.widget_resolver(step)
            except Exception:
                return None
        return None
