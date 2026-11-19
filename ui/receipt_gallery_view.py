"""
In-App Receipt Image Gallery & Manager View for SPVC Smart Receipt System.
Allows viewing, filtering, zooming, and managing exported/printed receipt images
directly inside the application without opening external programs.
"""

import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Any, Callable, Dict, List, Optional, Tuple

import customtkinter as ctk
from PIL import Image, ImageOps, ImageTk

from core.config import save_config
from core.database import DatabaseManager
from core.models import UserSession
from core.receipt_gallery import (
    ReceiptImageRecord,
    copy_receipt_image,
    delete_receipt_image,
    filter_receipt_images,
    get_gallery_filter_options,
    migrate_receipt_folder,
    move_receipt_image,
    rename_receipt_image,
    scan_receipt_images,
)
from ui.components import SearchableDropdown, SectionHeader
from ui.notifications import show_toast
from ui.theme import AppTheme


def _open_folder_in_explorer(folder_path: str, parent: Any = None) -> bool:
    """Opens a folder in the OS file manager safely across platforms."""
    target = os.path.abspath(folder_path)
    if not os.path.isdir(target):
        messagebox.showwarning("ไม่พบโฟลเดอร์", f"ไม่พบโฟลเดอร์ที่ระบุ:\n{target}", parent=parent)
        return False
    try:
        if sys.platform.startswith("win"):
            os.startfile(target)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", target])
        else:
            subprocess.Popen(["xdg-open", target])
        return True
    except Exception as exc:
        messagebox.showerror("เปิดโฟลเดอร์ไม่สำเร็จ", f"ไม่สามารถเปิดโฟลเดอร์ได้:\n{exc}", parent=parent)
        return False


def _format_file_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024 * 1024):.2f} MB"


class ReceiptGalleryView(ctk.CTkFrame):
    """
    In-App Receipt Image Gallery & Manager:
    - Scans the configured export/gallery folder for receipt images.
    - Enriches metadata from SQLite/Postgres receipts table via core.receipt_gallery.
    - Provides Grid/List toggle, lazy thumbnail loading, and a zoomable/pannable in-app preview.
    - Supports changing default export folder (with optional migration) and file management.
    """

    THUMB_SIZE: Tuple[int, int] = (120, 150)
    LIST_THUMB_SIZE: Tuple[int, int] = (48, 60)
    THUMB_BATCH_SIZE: int = 8

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

        self.all_records: List[ReceiptImageRecord] = []
        self.filtered_records: List[ReceiptImageRecord] = []
        self.selected_record: Optional[ReceiptImageRecord] = None

        # Image / UI state
        self.view_mode: str = "grid"  # "grid" | "list"
        self.zoom_mode: str = "fit"  # "fit" | "manual"
        self.zoom_scale: float = 1.0
        self._preview_pil: Optional[Image.Image] = None
        self._preview_photo: Optional[ImageTk.PhotoImage] = None
        self._preview_image_id: Optional[int] = None
        self._thumb_refs: Dict[str, ctk.CTkImage] = {}
        self._thumb_labels: List[Tuple[ctk.CTkLabel, str, Tuple[int, int]]] = []
        self._thumb_after_id: Optional[str] = None
        self._search_after_id: Optional[str] = None
        self._card_frames: Dict[str, ctk.CTkFrame] = {}
        self._pan_start: Tuple[int, int] = (0, 0)

        # Filter variables
        self.search_var = ctk.StringVar(value="")
        self.class_var = ctk.StringVar(value="ทั้งหมด")
        self.year_var = ctk.StringVar(value="ทั้งหมด")
        self.date_var = ctk.StringVar(value="")
        self.recursive_var = ctk.BooleanVar(value=True)

        self._build_ui()
        self.refresh_gallery()

    def on_show(self):
        """Called by MainWindow when the user switches back to this tab."""
        self.refresh_gallery()

    # ------------------------------------------------------------------
    # UI CONSTRUCTION
    # ------------------------------------------------------------------
    def _build_ui(self):
        # 1. Header
        header = SectionHeader(
            self,
            "คลังภาพใบเสร็จรับเงิน",
            "ดูภาพใบเสร็จในแอป ค้นหาตามห้อง/ชั้นปี/วันที่ และจัดการโฟลเดอร์จัดเก็บ",
        )
        header.pack(fill="x", pady=(0, 8))
        header.add_action_button(
            "🔄 รีเฟรช",
            self.refresh_gallery,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
        )
        header.add_action_button(
            "📂 เปิดโฟลเดอร์",
            self.open_current_folder,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
        )
        header.add_action_button(
            "📁 เปลี่ยนโฟลเดอร์จัดเก็บ",
            self.change_gallery_folder,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
        )

        # 2. Folder Status Strip
        folder_strip = ctk.CTkFrame(
            self,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        folder_strip.pack(fill="x", pady=(0, 8))

        fs_inner = ctk.CTkFrame(folder_strip, fg_color="transparent")
        fs_inner.pack(fill="x", padx=12, pady=8)

        ctk.CTkLabel(
            fs_inner,
            text="โฟลเดอร์จัดเก็บปัจจุบัน:",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 8))

        self.lbl_folder_path = ctk.CTkLabel(
            fs_inner,
            text=self.get_gallery_folder(),
            font=(AppTheme.FONT_MONO, 12),
            text_color=AppTheme.COLOR_PRIMARY_BLUE,
            anchor="w",
        )
        self.lbl_folder_path.pack(side="left", fill="x", expand=True)

        self.lbl_count_badge = ctk.CTkLabel(
            fs_inner,
            text="0 ภาพ",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_MUTED,
        )
        self.lbl_count_badge.pack(side="right", padx=(8, 0))

        # 3. Filter & View Bar
        filter_card = ctk.CTkFrame(
            self,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        filter_card.pack(fill="x", pady=(0, 10))

        flt_row = ctk.CTkFrame(filter_card, fg_color="transparent")
        flt_row.pack(fill="x", padx=12, pady=8)

        self.ent_search = ctk.CTkEntry(
            flt_row,
            textvariable=self.search_var,
            placeholder_text="ค้นหา ชื่อนักศึกษา / เลขที่ใบเสร็จ / ห้อง / ชื่อไฟล์...",
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_INPUT,
            border_width=1,
            border_color=AppTheme.COLOR_INPUT_BORDER,
            fg_color=AppTheme.COLOR_INPUT_BG,
            font=(AppTheme.FONT_FAMILY, 12),
        )
        self.ent_search.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.ent_search.bind("<KeyRelease>", self._schedule_filter)

        ctk.CTkLabel(
            flt_row, text="ห้อง:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY
        ).pack(side="left", padx=(0, 4))
        self.opt_class = SearchableDropdown(
            flt_row,
            values=["ทั้งหมด"],
            variable=self.class_var,
            width=130,
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_INPUT,
            font=(AppTheme.FONT_FAMILY, 12),
            command=lambda _v: self.apply_filters(),
        )
        self.opt_class.pack(side="left", padx=(0, 8))

        ctk.CTkLabel(
            flt_row, text="ชั้นปี:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY
        ).pack(side="left", padx=(0, 4))
        self.opt_year = SearchableDropdown(
            flt_row,
            values=["ทั้งหมด", "1", "2", "3", "4"],
            variable=self.year_var,
            width=95,
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_INPUT,
            font=(AppTheme.FONT_FAMILY, 12),
            command=lambda _v: self.apply_filters(),
        )
        self.opt_year.pack(side="left", padx=(0, 8))

        ctk.CTkLabel(
            flt_row, text="วันที่:", font=(AppTheme.FONT_FAMILY, 12, "bold"), text_color=AppTheme.COLOR_TEXT_SECONDARY
        ).pack(side="left", padx=(0, 4))
        self.ent_date = ctk.CTkEntry(
            flt_row,
            textvariable=self.date_var,
            placeholder_text="YYYY-MM-DD",
            width=110,
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_INPUT,
            border_width=1,
            border_color=AppTheme.COLOR_INPUT_BORDER,
            fg_color=AppTheme.COLOR_INPUT_BG,
            font=(AppTheme.FONT_FAMILY, 12),
        )
        self.ent_date.pack(side="left", padx=(0, 8))
        self.ent_date.bind("<KeyRelease>", self._schedule_filter)

        self.chk_recursive = ctk.CTkCheckBox(
            flt_row,
            text="รวมโฟลเดอร์ย่อย",
            variable=self.recursive_var,
            font=(AppTheme.FONT_FAMILY, 12),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            command=self.refresh_gallery,
        )
        self.chk_recursive.pack(side="left", padx=(0, 8))

        self.btn_view_mode = ctk.CTkButton(
            flt_row,
            text="☰ รายการ",
            width=86,
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.toggle_view_mode,
        )
        self.btn_view_mode.pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            flt_row,
            text="ล้าง",
            width=56,
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.reset_filters,
        ).pack(side="left")

        # 4. Split Workspace: Left Gallery List/Grid + Right In-App Preview
        workspace = ctk.CTkFrame(self, fg_color="transparent")
        workspace.pack(fill="both", expand=True)
        workspace.grid_columnconfigure(0, weight=5, uniform="gallery_split")
        workspace.grid_columnconfigure(1, weight=6, uniform="gallery_split")
        workspace.grid_rowconfigure(0, weight=1)

        # Left Panel: Scrollable Gallery Grid/List
        left_card = ctk.CTkFrame(
            workspace,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        left_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.gallery_scroll = ctk.CTkScrollableFrame(
            left_card,
            fg_color="transparent",
            corner_radius=0,
        )
        self.gallery_scroll.pack(fill="both", expand=True, padx=8, pady=8)

        # Right Panel: Large Zoomable/Pannable In-App Preview + Actions
        right_card = ctk.CTkFrame(
            workspace,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        right_card.grid(row=0, column=1, sticky="nsew")
        right_card.grid_rowconfigure(1, weight=1)
        right_card.grid_columnconfigure(0, weight=1)

        # Preview Toolbar
        pv_bar = ctk.CTkFrame(right_card, fg_color=AppTheme.COLOR_SURFACE_ALT, corner_radius=0)
        pv_bar.grid(row=0, column=0, sticky="ew")

        pv_bar_in = ctk.CTkFrame(pv_bar, fg_color="transparent")
        pv_bar_in.pack(fill="x", padx=10, pady=6)

        ctk.CTkLabel(
            pv_bar_in,
            text="ตัวอย่างใบเสร็จในแอป",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        ).pack(side="left")

        self.lbl_zoom = ctk.CTkLabel(
            pv_bar_in,
            text="Fit",
            width=52,
            font=(AppTheme.FONT_MONO, 12, "bold"),
            text_color=AppTheme.COLOR_PRIMARY_BLUE,
        )
        self.lbl_zoom.pack(side="right", padx=(6, 0))

        for label, cmd in [
            ("100%", self.zoom_actual),
            ("พอดีจอ", self.zoom_fit),
            ("🔍+", self.zoom_in),
            ("🔍-", self.zoom_out),
        ]:
            ctk.CTkButton(
                pv_bar_in,
                text=label,
                width=58,
                height=28,
                corner_radius=AppTheme.RADIUS_BTN,
                fg_color=AppTheme.COLOR_SURFACE,
                hover_color=AppTheme.COLOR_SECONDARY_HOVER,
                text_color=AppTheme.COLOR_TEXT_SECONDARY,
                border_width=1,
                border_color=AppTheme.COLOR_BORDER_STRONG,
                font=(AppTheme.FONT_FAMILY, 11, "bold"),
                command=cmd,
            ).pack(side="right", padx=2)

        # Scrollable / Pannable Preview Canvas
        canvas_wrap = ctk.CTkFrame(right_card, fg_color="#E2E8F0", corner_radius=0)
        canvas_wrap.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        canvas_wrap.grid_rowconfigure(0, weight=1)
        canvas_wrap.grid_columnconfigure(0, weight=1)

        self.preview_canvas = tk.Canvas(
            canvas_wrap,
            bg="#E2E8F0",
            highlightthickness=0,
        )
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")

        self.v_scroll = ctk.CTkScrollbar(canvas_wrap, orientation="vertical", command=self.preview_canvas.yview)
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll = ctk.CTkScrollbar(canvas_wrap, orientation="horizontal", command=self.preview_canvas.xview)
        self.h_scroll.grid(row=1, column=0, sticky="ew")
        self.preview_canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)

        self.preview_canvas.bind("<Configure>", self._on_canvas_configure)
        self.preview_canvas.bind("<ButtonPress-1>", self._on_pan_start)
        self.preview_canvas.bind("<B1-Motion>", self._on_pan_move)
        self.preview_canvas.bind("<MouseWheel>", self._on_mousewheel_zoom)

        # Metadata & Actions Card
        meta_card = ctk.CTkFrame(right_card, fg_color=AppTheme.COLOR_SURFACE_ALT, corner_radius=0)
        meta_card.grid(row=2, column=0, sticky="ew")

        self.lbl_meta_title = ctk.CTkLabel(
            meta_card,
            text="ยังไม่ได้เลือกภาพใบเสร็จ",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        )
        self.lbl_meta_title.pack(fill="x", padx=12, pady=(8, 2))

        self.lbl_meta_detail = ctk.CTkLabel(
            meta_card,
            text="คลิกเลือกภาพใบเสร็จจากรายการด้านซ้ายเพื่อดูตัวอย่างและจัดการไฟล์",
            font=(AppTheme.FONT_FAMILY, 12),
            text_color=AppTheme.COLOR_TEXT_MUTED,
            anchor="w",
            justify="left",
        )
        self.lbl_meta_detail.pack(fill="x", padx=12, pady=(0, 8))

        act_row = ctk.CTkFrame(meta_card, fg_color="transparent")
        act_row.pack(fill="x", padx=12, pady=(0, 10))

        self.btn_save_as = ctk.CTkButton(
            act_row,
            text="💾 บันทึกเป็น...",
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_PRIMARY_BLUE,
            hover_color=AppTheme.COLOR_PRIMARY_BLUE_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.save_as_selected,
            state="disabled",
        )
        self.btn_save_as.pack(side="left", padx=(0, 6))

        self.btn_move = ctk.CTkButton(
            act_row,
            text="📦 ย้ายโฟลเดอร์",
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.move_selected,
            state="disabled",
        )
        self.btn_move.pack(side="left", padx=(0, 6))

        self.btn_rename = ctk.CTkButton(
            act_row,
            text="✏️ เปลี่ยนชื่อ",
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_SECONDARY_BG,
            hover_color=AppTheme.COLOR_SECONDARY_HOVER,
            text_color=AppTheme.COLOR_SECONDARY_TEXT,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.rename_selected,
            state="disabled",
        )
        self.btn_rename.pack(side="left", padx=(0, 6))

        self.btn_open_item_folder = ctk.CTkButton(
            act_row,
            text="📂 โฟลเดอร์ไฟล์",
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.open_selected_folder,
            state="disabled",
        )
        self.btn_open_item_folder.pack(side="left", padx=(0, 6))

        self.btn_delete = ctk.CTkButton(
            act_row,
            text="🗑️ ลบ",
            width=72,
            height=AppTheme.CTRL_H_SM,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_DANGER,
            hover_color=AppTheme.COLOR_DANGER_HOVER,
            text_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self.delete_selected,
            state="disabled",
        )
        self.btn_delete.pack(side="right")

    # ------------------------------------------------------------------
    # FOLDER & SCANNING
    # ------------------------------------------------------------------
    def get_gallery_folder(self) -> str:
        raw = str(self.config.get("default_export_folder", "") or "").strip()
        if raw:
            return os.path.abspath(raw)
        return os.path.abspath(os.path.join(os.getcwd(), "exports_output"))

    def refresh_gallery(self):
        folder = self.get_gallery_folder()
        self.lbl_folder_path.configure(text=folder)
        self.all_records = scan_receipt_images(
            folder,
            db=self.db,
            recursive=bool(self.recursive_var.get()),
        )

        facets = get_gallery_filter_options(self.all_records)
        class_opts = ["ทั้งหมด"] + facets.get("class_groups", [])
        year_opts = ["ทั้งหมด"] + facets.get("year_levels", [])
        self.opt_class.configure(values=class_opts)
        if self.class_var.get() not in class_opts:
            self.class_var.set("ทั้งหมด")
        self.opt_year.configure(values=year_opts)
        if self.year_var.get() not in year_opts:
            self.year_var.set("ทั้งหมด")

        self.apply_filters()

    def _schedule_filter(self, _event=None):
        if self._search_after_id is not None:
            try:
                self.after_cancel(self._search_after_id)
            except Exception:
                pass
        self._search_after_id = self.after(180, self.apply_filters)

    def apply_filters(self):
        cls_val = self.class_var.get().strip()
        yr_val = self.year_var.get().strip()
        self.filtered_records = filter_receipt_images(
            self.all_records,
            query=self.search_var.get().strip(),
            class_group="" if cls_val in ("", "ทั้งหมด") else cls_val,
            year_level="" if yr_val in ("", "ทั้งหมด") else yr_val,
            date_str=self.date_var.get().strip(),
        )
        self.lbl_count_badge.configure(
            text=f"แสดง {len(self.filtered_records)} / {len(self.all_records)} ภาพ"
        )
        self._render_gallery_items()

        # Keep selection if still present, otherwise clear if missing
        if self.selected_record is not None:
            still_here = next(
                (r for r in self.filtered_records if r.file_path == self.selected_record.file_path),
                None,
            )
            if still_here is None:
                self.clear_preview()
            else:
                self.select_record(still_here, reload_image=False)

    def reset_filters(self):
        self.search_var.set("")
        self.class_var.set("ทั้งหมด")
        self.year_var.set("ทั้งหมด")
        self.date_var.set("")
        self.apply_filters()

    def toggle_view_mode(self):
        self.view_mode = "list" if self.view_mode == "grid" else "grid"
        self.btn_view_mode.configure(text="⊞ การ์ด" if self.view_mode == "list" else "☰ รายการ")
        self._render_gallery_items()

    # ------------------------------------------------------------------
    # GALLERY ITEMS & LAZY THUMBNAIL LOADING
    # ------------------------------------------------------------------
    def _cancel_thumb_loader(self):
        if self._thumb_after_id is not None:
            try:
                self.after_cancel(self._thumb_after_id)
            except Exception:
                pass
            self._thumb_after_id = None
        self._thumb_labels.clear()

    def _render_gallery_items(self):
        self._cancel_thumb_loader()
        self._card_frames.clear()
        self._thumb_refs.clear()

        for child in list(self.gallery_scroll.winfo_children()):
            child.destroy()

        if not self.filtered_records:
            empty_box = ctk.CTkFrame(self.gallery_scroll, fg_color="transparent")
            empty_box.pack(fill="both", expand=True, pady=48, padx=16)
            ctk.CTkLabel(
                empty_box,
                text="ไม่พบภาพใบเสร็จในโฟลเดอร์นี้",
                font=(AppTheme.FONT_FAMILY, 15, "bold"),
                text_color=AppTheme.COLOR_TEXT_SECONDARY,
            ).pack(pady=(0, 6))
            ctk.CTkLabel(
                empty_box,
                text="ลองเปลี่ยนคำค้นหา ตัวกรอง หรือกด 'เปลี่ยนโฟลเดอร์จัดเก็บ' เพื่อเลือกโฟลเดอร์ที่มีไฟล์ภาพ (.png, .jpg)",
                font=(AppTheme.FONT_FAMILY, 12),
                text_color=AppTheme.COLOR_TEXT_MUTED,
                wraplength=360,
            ).pack()
            return

        if self.view_mode == "grid":
            self.gallery_scroll.grid_columnconfigure(0, weight=1, uniform="gal_col")
            self.gallery_scroll.grid_columnconfigure(1, weight=1, uniform="gal_col")
            for idx, rec in enumerate(self.filtered_records):
                row_i, col_i = divmod(idx, 2)
                card = self._create_grid_card(self.gallery_scroll, rec)
                card.grid(row=row_i, column=col_i, sticky="nsew", padx=5, pady=5)
                self._card_frames[rec.file_path] = card
        else:
            self.gallery_scroll.grid_columnconfigure(0, weight=1)
            self.gallery_scroll.grid_columnconfigure(1, weight=0)
            for idx, rec in enumerate(self.filtered_records):
                row_card = self._create_list_row(self.gallery_scroll, rec)
                row_card.grid(row=idx, column=0, columnspan=2, sticky="ew", padx=4, pady=3)
                self._card_frames[rec.file_path] = row_card

        self._schedule_next_thumb_batch()

    def _create_grid_card(self, parent, rec: ReceiptImageRecord) -> ctk.CTkFrame:
        is_sel = self.selected_record is not None and self.selected_record.file_path == rec.file_path
        card = ctk.CTkFrame(
            parent,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG if is_sel else AppTheme.COLOR_SURFACE_ALT,
            border_width=2 if is_sel else 1,
            border_color=AppTheme.COLOR_PRIMARY_BLUE if is_sel else AppTheme.COLOR_BORDER,
        )

        lbl_img = ctk.CTkLabel(
            card,
            text="กำลังโหลด...",
            width=self.THUMB_SIZE[0],
            height=self.THUMB_SIZE[1],
            fg_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 11),
            text_color=AppTheme.COLOR_TEXT_MUTED,
        )
        lbl_img.pack(padx=8, pady=(8, 4))
        self._thumb_labels.append((lbl_img, rec.file_path, self.THUMB_SIZE))

        title_text = f"เลขที่ {rec.receipt_no}" if rec.receipt_no else rec.filename
        lbl_title = ctk.CTkLabel(
            card,
            text=title_text,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        )
        lbl_title.pack(fill="x", padx=8)

        sub_parts = [p for p in [rec.student_name, rec.class_group, rec.receipt_date or rec.modified_date_str] if p]
        sub_text = " • ".join(sub_parts) if sub_parts else rec.filename
        lbl_sub = ctk.CTkLabel(
            card,
            text=sub_text,
            font=(AppTheme.FONT_FAMILY, 11),
            text_color=AppTheme.COLOR_TEXT_MUTED,
            anchor="w",
            wraplength=170,
            justify="left",
        )
        lbl_sub.pack(fill="x", padx=8, pady=(0, 8))

        for w in (card, lbl_img, lbl_title, lbl_sub):
            w.bind("<Button-1>", lambda _e, r=rec: self.select_record(r))
        return card

    def _create_list_row(self, parent, rec: ReceiptImageRecord) -> ctk.CTkFrame:
        is_sel = self.selected_record is not None and self.selected_record.file_path == rec.file_path
        row = ctk.CTkFrame(
            parent,
            corner_radius=AppTheme.RADIUS,
            fg_color=AppTheme.COLOR_SECONDARY_BG if is_sel else AppTheme.COLOR_SURFACE_ALT,
            border_width=2 if is_sel else 1,
            border_color=AppTheme.COLOR_PRIMARY_BLUE if is_sel else AppTheme.COLOR_BORDER,
        )

        lbl_img = ctk.CTkLabel(
            row,
            text="...",
            width=self.LIST_THUMB_SIZE[0],
            height=self.LIST_THUMB_SIZE[1],
            fg_color="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 10),
            text_color=AppTheme.COLOR_TEXT_MUTED,
        )
        lbl_img.pack(side="left", padx=8, pady=6)
        self._thumb_labels.append((lbl_img, rec.file_path, self.LIST_THUMB_SIZE))

        info = ctk.CTkFrame(row, fg_color="transparent")
        info.pack(side="left", fill="both", expand=True, padx=(0, 8), pady=6)

        title_text = f"เลขที่ {rec.receipt_no} — {rec.filename}" if rec.receipt_no else rec.filename
        lbl_title = ctk.CTkLabel(
            info,
            text=title_text,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        )
        lbl_title.pack(fill="x")

        sub_parts = [p for p in [rec.student_name, rec.class_group, rec.receipt_date or rec.modified_date_str, _format_file_size(rec.size_bytes)] if p]
        lbl_sub = ctk.CTkLabel(
            info,
            text=" | ".join(sub_parts),
            font=(AppTheme.FONT_FAMILY, 11),
            text_color=AppTheme.COLOR_TEXT_MUTED,
            anchor="w",
        )
        lbl_sub.pack(fill="x")

        for w in (row, lbl_img, info, lbl_title, lbl_sub):
            w.bind("<Button-1>", lambda _e, r=rec: self.select_record(r))
        return row

    def _schedule_next_thumb_batch(self):
        if not self._thumb_labels:
            self._thumb_after_id = None
            return
        self._thumb_after_id = self.after(15, self._load_thumb_batch)

    def _load_thumb_batch(self):
        self._thumb_after_id = None
        batch = self._thumb_labels[: self.THUMB_BATCH_SIZE]
        del self._thumb_labels[: self.THUMB_BATCH_SIZE]

        for lbl, path, size in batch:
            try:
                if not lbl.winfo_exists():
                    continue
            except Exception:
                continue

            ctk_img = self._load_thumbnail_ctk_image(path, size)
            if ctk_img is not None:
                key = f"{path}::{size[0]}x{size[1]}"
                self._thumb_refs[key] = ctk_img
                lbl.configure(image=ctk_img, text="")
            else:
                lbl.configure(text="⚠️ ภาพเสียหาย")

        if self._thumb_labels:
            self._schedule_next_thumb_batch()

    @staticmethod
    def _load_thumbnail_ctk_image(file_path: str, size: Tuple[int, int]) -> Optional[ctk.CTkImage]:
        """Opens, thumbnails, and immediately closes the PIL file handle to avoid Windows locks."""
        try:
            with Image.open(file_path) as raw:
                img = ImageOps.exif_transpose(raw).convert("RGBA")
            img.thumbnail(size, Image.Resampling.LANCZOS)
            return ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
        except Exception:
            return None

    # ------------------------------------------------------------------
    # IN-APP PREVIEW & ZOOM/PAN
    # ------------------------------------------------------------------
    def select_record(self, rec: ReceiptImageRecord, reload_image: bool = True):
        prev_path = self.selected_record.file_path if self.selected_record else None
        self.selected_record = rec

        if prev_path and prev_path in self._card_frames:
            try:
                self._card_frames[prev_path].configure(
                    fg_color=AppTheme.COLOR_SURFACE_ALT,
                    border_width=1,
                    border_color=AppTheme.COLOR_BORDER,
                )
            except Exception:
                pass
        if rec.file_path in self._card_frames:
            try:
                self._card_frames[rec.file_path].configure(
                    fg_color=AppTheme.COLOR_SECONDARY_BG,
                    border_width=2,
                    border_color=AppTheme.COLOR_PRIMARY_BLUE,
                )
            except Exception:
                pass

        for btn in (self.btn_save_as, self.btn_move, self.btn_rename, self.btn_open_item_folder, self.btn_delete):
            btn.configure(state="normal")

        if reload_image:
            self._release_preview_image()
            try:
                with Image.open(rec.file_path) as raw:
                    self._preview_pil = ImageOps.exif_transpose(raw).convert("RGBA")
                self.zoom_mode = "fit"
                self._render_preview_canvas()
            except Exception as exc:
                self._preview_pil = None
                self.preview_canvas.delete("all")
                self.preview_canvas.create_text(
                    max(160, self.preview_canvas.winfo_width() // 2),
                    max(120, self.preview_canvas.winfo_height() // 2),
                    text=f"ไม่สามารถเปิดภาพได้:\n{exc}",
                    fill=AppTheme.COLOR_DANGER,
                    font=(AppTheme.FONT_FAMILY, 13, "bold"),
                )

        dims = f"{self._preview_pil.width}×{self._preview_pil.height} px" if self._preview_pil else "-"
        title = f"{rec.filename}" + (f"  (เลขที่ใบเสร็จ: {rec.receipt_no})" if rec.receipt_no else "")
        self.lbl_meta_title.configure(text=title)

        details = [
            f"ผู้ชำระ: {rec.student_name or '-'}",
            f"ห้อง/ชั้น: {rec.class_group or '-'}",
            f"วันที่: {rec.receipt_date or rec.modified_date_str}",
            f"ขนาด: {_format_file_size(rec.size_bytes)} ({dims})",
        ]
        if rec.net_total_amount is not None:
            details.append(f"ยอดสุทธิ: ฿{rec.net_total_amount:,.2f}")
        self.lbl_meta_detail.configure(text="   |   ".join(details))

    def _release_preview_image(self):
        self._preview_pil = None
        self._preview_photo = None
        self._preview_image_id = None

    def clear_preview(self):
        self.selected_record = None
        self._release_preview_image()
        self.preview_canvas.delete("all")
        self.lbl_meta_title.configure(text="ยังไม่ได้เลือกภาพใบเสร็จ")
        self.lbl_meta_detail.configure(text="คลิกเลือกภาพใบเสร็จจากรายการด้านซ้ายเพื่อดูตัวอย่างและจัดการไฟล์")
        self.lbl_zoom.configure(text="Fit")
        for btn in (self.btn_save_as, self.btn_move, self.btn_rename, self.btn_open_item_folder, self.btn_delete):
            btn.configure(state="disabled")

    def _render_preview_canvas(self):
        if self._preview_pil is None:
            return
        cw = max(100, self.preview_canvas.winfo_width())
        ch = max(100, self.preview_canvas.winfo_height())

        if self.zoom_mode == "fit":
            scale_w = (cw - 24) / max(1, self._preview_pil.width)
            scale_h = (ch - 24) / max(1, self._preview_pil.height)
            scale = max(0.05, min(scale_w, scale_h, 2.0))
            self.zoom_scale = scale
            self.lbl_zoom.configure(text=f"Fit ({int(scale * 100)}%)")
        else:
            scale = max(0.1, min(self.zoom_scale, 4.0))
            self.zoom_scale = scale
            self.lbl_zoom.configure(text=f"{int(scale * 100)}%")

        new_w = max(1, int(self._preview_pil.width * scale))
        new_h = max(1, int(self._preview_pil.height * scale))
        resized = self._preview_pil.resize((new_w, new_h), Image.Resampling.LANCZOS)
        self._preview_photo = ImageTk.PhotoImage(resized)

        self.preview_canvas.delete("all")
        x_pos = max(0, (cw - new_w) // 2)
        y_pos = max(0, (ch - new_h) // 2)
        self._preview_image_id = self.preview_canvas.create_image(x_pos, y_pos, anchor="nw", image=self._preview_photo)
        self.preview_canvas.configure(scrollregion=(0, 0, max(cw, new_w + x_pos), max(ch, new_h + y_pos)))

    def _on_canvas_configure(self, _event=None):
        if self.zoom_mode == "fit" and self._preview_pil is not None:
            self._render_preview_canvas()

    def _on_pan_start(self, event):
        self.preview_canvas.scan_mark(event.x, event.y)

    def _on_pan_move(self, event):
        self.preview_canvas.scan_dragto(event.x, event.y, gain=1)

    def _on_mousewheel_zoom(self, event):
        if self._preview_pil is None:
            return
        if event.delta > 0:
            self.zoom_in()
        elif event.delta < 0:
            self.zoom_out()

    def zoom_in(self):
        if self._preview_pil is None:
            return
        self.zoom_mode = "manual"
        self.zoom_scale = min(4.0, self.zoom_scale * 1.25)
        self._render_preview_canvas()

    def zoom_out(self):
        if self._preview_pil is None:
            return
        self.zoom_mode = "manual"
        self.zoom_scale = max(0.1, self.zoom_scale / 1.25)
        self._render_preview_canvas()

    def zoom_fit(self):
        self.zoom_mode = "fit"
        self._render_preview_canvas()

    def zoom_actual(self):
        if self._preview_pil is None:
            return
        self.zoom_mode = "manual"
        self.zoom_scale = 1.0
        self._render_preview_canvas()

    # ------------------------------------------------------------------
    # FOLDER & FILE MANAGEMENT ACTIONS
    # ------------------------------------------------------------------
    def open_current_folder(self):
        folder = self.get_gallery_folder()
        os.makedirs(folder, exist_ok=True)
        _open_folder_in_explorer(folder, parent=self.winfo_toplevel())

    def open_selected_folder(self):
        if not self.selected_record:
            return
        _open_folder_in_explorer(self.selected_record.folder, parent=self.winfo_toplevel())

    def change_gallery_folder(self):
        old_folder = self.get_gallery_folder()
        init_dir = old_folder if os.path.isdir(old_folder) else os.getcwd()
        new_folder = filedialog.askdirectory(
            parent=self.winfo_toplevel(),
            title="เลือกโฟลเดอร์สำหรับจัดเก็บภาพ/ไฟล์ใบเสร็จ",
            initialdir=init_dir,
        )
        if not new_folder:
            return

        new_abs = os.path.abspath(new_folder)
        if os.path.normcase(new_abs) == os.path.normcase(old_folder):
            return

        self.config["default_export_folder"] = new_abs
        save_config(self.config)

        if self.all_records and os.path.isdir(old_folder):
            ans = messagebox.askyesnocancel(
                "ย้ายหรือคัดลอกภาพใบเสร็จเดิม?",
                f"พบภาพใบเสร็จในโฟลเดอร์เดิม ({len(self.all_records)} ภาพ)\n"
                f"ต้องการ 'ย้าย' ภาพทั้งหมดไปยังโฟลเดอร์ใหม่หรือไม่?\n\n"
                f"• กด Yes = ย้ายไฟล์ทั้งหมดไปยังโฟลเดอร์ใหม่\n"
                f"• กด No = คัดลอกสำเนาไปยังโฟลเดอร์ใหม่ (เก็บไฟล์เดิมไว้)\n"
                f"• กด Cancel = เปลี่ยนเฉพาะโฟลเดอร์จัดเก็บ ไม่ย้ายไฟล์เดิม",
                parent=self.winfo_toplevel(),
            )
            if ans is not None:
                self.clear_preview()
                res = migrate_receipt_folder(old_folder, new_abs, move_files=bool(ans))
                action_th = "ย้าย" if ans else "คัดลอก"
                msg = f"{action_th}สำเร็จ {res['processed']} ไฟล์"
                if res["skipped"]:
                    msg += f" (ข้าม {res['skipped']})"
                if res["errors"]:
                    messagebox.showwarning(
                        "ย้าย/คัดลอกบางไฟล์ไม่สำเร็จ",
                        msg + "\n\nข้อผิดพลาด:\n" + "\n".join(res["errors"][:5]),
                        parent=self.winfo_toplevel(),
                    )
                else:
                    show_toast(self.winfo_toplevel(), "อัปเดตโฟลเดอร์สำเร็จ", msg, kind="success")

        self.refresh_gallery()

    def save_as_selected(self):
        if not self.selected_record:
            return
        src = self.selected_record.file_path
        ext = self.selected_record.ext or ".png"
        dest = filedialog.asksaveasfilename(
            parent=self.winfo_toplevel(),
            title="บันทึกสำเนาภาพใบเสร็จเป็น...",
            initialdir=self.selected_record.folder,
            initialfile=self.selected_record.filename,
            defaultextension=ext,
            filetypes=[("Image Files", "*.png *.jpg *.jpeg *.webp *.bmp *.gif"), ("All Files", "*.*")],
        )
        if not dest:
            return
        try:
            saved_path = copy_receipt_image(src, dest, overwrite=True)
            show_toast(self.winfo_toplevel(), "บันทึกสำเนาสำเร็จ", f"บันทึกไว้ที่: {saved_path}", kind="success")
        except Exception as exc:
            messagebox.showerror("บันทึกไม่สำเร็จ", f"ไม่สามารถบันทึกไฟล์ได้:\n{exc}", parent=self.winfo_toplevel())

    def move_selected(self):
        if not self.selected_record:
            return
        src = self.selected_record.file_path
        dest_dir = filedialog.askdirectory(
            parent=self.winfo_toplevel(),
            title="เลือกโฟลเดอร์ปลายทางสำหรับย้ายภาพ",
            initialdir=self.selected_record.folder,
        )
        if not dest_dir:
            return
        try:
            self.clear_preview()
            new_path = move_receipt_image(src, dest_dir)
            show_toast(self.winfo_toplevel(), "ย้ายไฟล์สำเร็จ", f"ย้ายไปที่: {new_path}", kind="success")
            self.refresh_gallery()
        except Exception as exc:
            messagebox.showerror("ย้ายไฟล์ไม่สำเร็จ", f"ไม่สามารถย้ายไฟล์ได้:\n{exc}", parent=self.winfo_toplevel())

    def rename_selected(self):
        if not self.selected_record:
            return
        rec = self.selected_record
        stem, _ext = os.path.splitext(rec.filename)
        dlg = ctk.CTkInputDialog(
            text=f"กรอกชื่อไฟล์ใหม่ (ไม่ต้องใส่นามสกุล {rec.ext}):",
            title="เปลี่ยนชื่อไฟล์ใบเสร็จ",
        )
        new_stem = dlg.get_input()
        if not new_stem or not new_stem.strip() or new_stem.strip() == stem:
            return
        try:
            src = rec.file_path
            self.clear_preview()
            new_path = rename_receipt_image(src, new_stem.strip(), allowed_root=self.get_gallery_folder())
            self.refresh_gallery()
            matched = next((r for r in self.filtered_records if r.file_path == new_path), None)
            if matched:
                self.select_record(matched)
            show_toast(self.winfo_toplevel(), "เปลี่ยนชื่อสำเร็จ", os.path.basename(new_path), kind="success")
        except Exception as exc:
            messagebox.showerror("เปลี่ยนชื่อไม่สำเร็จ", f"ไม่สามารถเปลี่ยนชื่อไฟล์ได้:\n{exc}", parent=self.winfo_toplevel())

    def delete_selected(self):
        if not self.selected_record:
            return
        rec = self.selected_record
        confirmed = messagebox.askyesno(
            "ยืนยันการลบภาพใบเสร็จ",
            f"คุณต้องการลบไฟล์ภาพนี้ถาวรหรือไม่?\n\n{rec.filename}",
            parent=self.winfo_toplevel(),
        )
        if not confirmed:
            return
        try:
            target = rec.file_path
            self.clear_preview()
            ok = delete_receipt_image(target, allowed_root=self.get_gallery_folder())
            if ok:
                show_toast(self.winfo_toplevel(), "ลบไฟล์แล้ว", rec.filename, kind="info")
            self.refresh_gallery()
        except Exception as exc:
            messagebox.showerror("ลบไฟล์ไม่สำเร็จ", f"ไม่สามารถลบไฟล์ได้:\n{exc}", parent=self.winfo_toplevel())
