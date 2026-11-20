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
    Floating AppTheme-styled guided tour card with target widget highlight border,
    progress indicator, clickable step dots, section jump menu, and navigation buttons.
    """

    HIGHLIGHT_THICKNESS = 3

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
        self._highlight_bars: List[ctk.CTkFrame] = []

        self._build_highlight_bars()
        self._build_card_ui()

    def _build_highlight_bars(self) -> None:
        for _ in range(4):
            bar = ctk.CTkFrame(
                self._host_root,
                fg_color=AppTheme.COLOR_PRIMARY_BLUE,
                corner_radius=0,
                border_width=0,
            )
            self._highlight_bars.append(bar)

    def _build_card_ui(self) -> None:
        # Top Navy header strip
        header = ctk.CTkFrame(
            self,
            fg_color=AppTheme.COLOR_ACCENT,
            corner_radius=AppTheme.RADIUS,
            height=40,
        )
        header.pack(fill="x", side="top")

        self.lbl_page_badge = ctk.CTkLabel(
            header,
            text="หน้า: -",
            font=AppTheme.FONT_LABEL,
            text_color=AppTheme.COLOR_TEXT_LIGHT,
        )
        self.lbl_page_badge.pack(side="left", padx=AppTheme.SPACE_MD, pady=AppTheme.SPACE_SM)

        self.lbl_progress = ctk.CTkLabel(
            header,
            text="ขั้นตอนที่ 1 / 1",
            font=AppTheme.FONT_CAPTION,
            text_color=AppTheme.COLOR_SECONDARY_BG,
        )
        self.lbl_progress.pack(side="right", padx=AppTheme.SPACE_MD, pady=AppTheme.SPACE_SM)

        content = ctk.CTkFrame(self, fg_color=AppTheme.COLOR_SURFACE, corner_radius=0)
        content.pack(fill="both", expand=True, padx=AppTheme.SPACE_MD, pady=AppTheme.SPACE_MD)

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
            height=28,
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
            font=AppTheme.FONT_SECTION,
            text_color=AppTheme.COLOR_TEXT,
            wraplength=420,
            justify="left",
        )
        self.lbl_title.pack(anchor="w", pady=(0, AppTheme.SPACE_XS))

        self.lbl_description = ctk.CTkLabel(
            content,
            text="",
            font=AppTheme.FONT_BODY,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            wraplength=420,
            justify="left",
        )
        self.lbl_description.pack(anchor="w", pady=(0, AppTheme.SPACE_SM))

        self.tips_frame = ctk.CTkFrame(
            content,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            border_width=1,
            border_color=AppTheme.COLOR_SECONDARY_BORDER,
            corner_radius=AppTheme.RADIUS,
        )
        self.lbl_tips = ctk.CTkLabel(
            self.tips_frame,
            text="",
            font=AppTheme.FONT_CAPTION,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            wraplength=400,
            justify="left",
        )
        self.lbl_tips.pack(anchor="w", padx=AppTheme.SPACE_SM, pady=AppTheme.SPACE_SM)

        # Clickable step dots row
        self.dots_frame = ctk.CTkFrame(content, fg_color="transparent")
        self.dots_frame.pack(fill="x", pady=(AppTheme.SPACE_SM, AppTheme.SPACE_SM))

        # Action controls row
        controls = ctk.CTkFrame(content, fg_color="transparent")
        controls.pack(fill="x", pady=(AppTheme.SPACE_XS, 0))

        self.btn_skip = ctk.CTkButton(
            controls,
            text="ข้ามบทเรียน (Skip)",
            font=AppTheme.FONT_CAPTION,
            height=AppTheme.CTRL_H_SM,
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
            text="ย้อนกลับ",
            font=AppTheme.FONT_LABEL,
            height=AppTheme.CTRL_H_SM,
            width=84,
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
            text="ถัดไป",
            font=AppTheme.FONT_LABEL,
            height=AppTheme.CTRL_H_SM,
            width=100,
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
                text="เสร็จสิ้น",
                fg_color=AppTheme.COLOR_SUCCESS,
                hover_color=AppTheme.COLOR_SUCCESS_HOVER,
            )
        else:
            self.btn_next.configure(
                text="ถัดไป",
                fg_color=AppTheme.COLOR_PRIMARY_BLUE,
                hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            )

        self._place_card(target_widget)
        self._highlight_widget(target_widget)

    def _render_step_dots(self, current_index: int, total: int) -> None:
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

    def _place_card(self, target_widget: Optional[Any] = None) -> None:
        """Positions the floating tutorial card cleanly inside root."""
        try:
            self._host_root.update_idletasks()
            self.update_idletasks()
            rw = self._host_root.winfo_width()
            rh = self._host_root.winfo_height()
            card_w = 460
            card_h = max(self.winfo_reqheight(), 260)
            if rw > 500 and rh > 380:
                margin = 20
                x = rw - card_w - margin
                y = rh - card_h - margin
                if target_widget is not None and target_widget.winfo_exists():
                    tx = target_widget.winfo_rootx() - self._host_root.winfo_rootx()
                    ty = target_widget.winfo_rooty() - self._host_root.winfo_rooty()
                    tw = target_widget.winfo_width()
                    th = target_widget.winfo_height()
                    # If target sits in bottom-right, move card to bottom-left or top-right
                    if tx + tw > x and ty + th > y:
                        if tx > card_w + margin * 2:
                            x = margin
                        elif ty > card_h + margin * 2:
                            y = margin + 56
                self.place(x=max(margin, x), y=max(margin, y), width=card_w)
            else:
                self.place(relx=0.98, rely=0.96, anchor="se")
            self.lift()
        except Exception:
            try:
                self.place(relx=0.98, rely=0.96, anchor="se")
            except Exception:
                pass

    def _highlight_widget(self, target_widget: Optional[Any]) -> None:
        """Draws a crisp 4-edge highlight border around the target widget."""
        self._clear_highlight()
        if target_widget is None:
            return
        try:
            if not target_widget.winfo_exists():
                return
            self._host_root.update_idletasks()
            rx = self._host_root.winfo_rootx()
            ry = self._host_root.winfo_rooty()
            tx = target_widget.winfo_rootx() - rx
            ty = target_widget.winfo_rooty() - ry
            tw = target_widget.winfo_width()
            th = target_widget.winfo_height()
            if tw <= 1 or th <= 1:
                return

            t = self.HIGHLIGHT_THICKNESS
            top_bar, bottom_bar, left_bar, right_bar = self._highlight_bars
            top_bar.place(x=tx - t, y=ty - t, width=tw + 2 * t, height=t)
            bottom_bar.place(x=tx - t, y=ty + th, width=tw + 2 * t, height=t)
            left_bar.place(x=tx - t, y=ty, width=t, height=th)
            right_bar.place(x=tx + tw, y=ty, width=t, height=th)
            for bar in self._highlight_bars:
                bar.lift()
            self.lift()
        except Exception:
            self._clear_highlight()

    def _clear_highlight(self) -> None:
        for bar in self._highlight_bars:
            try:
                bar.place_forget()
            except Exception:
                pass

    def hide(self) -> None:
        """Hides the overlay card and highlight border without destroying."""
        self._clear_highlight()
        try:
            self.place_forget()
        except Exception:
            pass

    def destroy_overlay(self) -> None:
        """Destroys highlight bars and the overlay card."""
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
