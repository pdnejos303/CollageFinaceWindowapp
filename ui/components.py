from tkinter import ttk
from typing import Callable, List, Optional, Tuple
import customtkinter as ctk

from ui.theme import AppTheme


class StatCard(ctk.CTkFrame):
    """Minimal KPI Card with unified corporate styling."""

    def __init__(
        self,
        parent,
        title: str,
        value: str,
        subtitle: str = "",
        accent_color: str = AppTheme.COLOR_PRIMARY_BLUE,
        badge_bg: str = AppTheme.COLOR_INFO_BG,
    ):
        super().__init__(
            parent,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=10)

        self.lbl_title = ctk.CTkLabel(
            body,
            text=title,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        )
        self.lbl_title.pack(fill="x")

        self.lbl_value = ctk.CTkLabel(
            body,
            text=value,
            font=(AppTheme.FONT_FAMILY, 20, "bold"),
            text_color=accent_color,
            anchor="w",
        )
        self.lbl_value.pack(fill="x", pady=(2, 0))


class SectionHeader(ctk.CTkFrame):
    """Clean, single-line title bar with unified corporate styling."""

    def __init__(self, parent, title: str, subtitle: str = ""):
        super().__init__(
            parent,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        inner = ctk.CTkFrame(self, fg_color="transparent")
        inner.pack(fill="x", padx=18, pady=10)

        self.actions_frame = ctk.CTkFrame(inner, fg_color="transparent")
        self.actions_frame.pack(side="right")

        left = ctk.CTkFrame(inner, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            left,
            text=title,
            font=(AppTheme.FONT_FAMILY, 19, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        ).pack(side="left")

        if subtitle:
            ctk.CTkLabel(
                left,
                text=f"   •   {subtitle}",
                font=(AppTheme.FONT_FAMILY, 13),
                text_color=AppTheme.COLOR_TEXT_MUTED,
                anchor="w",
            ).pack(side="left")

    def add_action_button(
        self,
        text: str,
        command: Callable,
        fg_color: str = AppTheme.COLOR_SECONDARY_BG,
        text_color: str = AppTheme.COLOR_SECONDARY_TEXT,
        hover_color: str = AppTheme.COLOR_SECONDARY_HOVER,
    ) -> ctk.CTkButton:
        btn = ctk.CTkButton(
            self.actions_frame,
            text=text,
            command=command,
            corner_radius=8,
            height=36,
            fg_color=fg_color,
            hover_color=hover_color,
            text_color=text_color,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
        )
        btn.pack(side="left", padx=(8, 0))
        return btn


class DataTable(ctk.CTkFrame):
    """
    Harmonious, Enterprise-Grade Table with Built-in Pagination Support.
    Soft Slate-100 header, subtle zebra rows, calm ice-blue selection highlight,
    and automatic/server-side pagination for 100,000+ rows.
    """

    def __init__(
        self,
        parent,
        columns: List[Tuple[str, str, int, str]],
        height: int = 6,
        on_select: Optional[Callable[[List[str]], None]] = None,
        show_inspector: bool = False,
        page_size: int = 100,
        on_page_change: Optional[Callable[[int], None]] = None,
        show_pagination: bool = False,
    ):
        super().__init__(
            parent,
            corner_radius=10,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=1,
            border_color=AppTheme.COLOR_BORDER,
        )
        self.on_select = on_select
        self.columns_meta = columns
        self.page_size = max(10, int(page_size))
        self.on_page_change = on_page_change
        self.show_pagination = show_pagination
        self.current_page = 1
        self.total_count = 0
        self.total_pages = 1
        self._client_rows: List[Tuple[List[str], str]] = []

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "GodApp.Treeview",
            background="#FFFFFF",
            foreground=AppTheme.COLOR_TEXT,
            rowheight=44,
            fieldbackground="#FFFFFF",
            font=(AppTheme.FONT_FAMILY, 13),
            borderwidth=0,
        )
        style.configure(
            "GodApp.Treeview.Heading",
            background="#F1F5F9",
            foreground="#1E293B",
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            padding=(10, 10),
            relief="flat",
            borderwidth=0,
        )
        style.map(
            "GodApp.Treeview",
            background=[("selected", "#DBEAFE")],
            foreground=[("selected", "#1E3A8A")],
        )
        style.map(
            "GodApp.Treeview.Heading",
            background=[("active", "#E2E8F0")],
        )

        # Optional Pagination Footer Bar
        self.pagination_bar = ctk.CTkFrame(self, fg_color="#F8FAFC", corner_radius=6, height=36)
        if self.show_pagination or self.on_page_change is not None:
            self.pagination_bar.pack(side="bottom", fill="x", padx=6, pady=(0, 6))

        self.lbl_page_info = ctk.CTkLabel(
            self.pagination_bar,
            text="แสดง 0 รายการ",
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        )
        self.lbl_page_info.pack(side="left", padx=12, pady=4)

        nav_box = ctk.CTkFrame(self.pagination_bar, fg_color="transparent")
        nav_box.pack(side="right", padx=8, pady=4)

        self.btn_first = ctk.CTkButton(
            nav_box,
            text="⏮ หน้าแรก",
            width=75,
            height=28,
            corner_radius=6,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self._go_first_page,
        )
        self.btn_first.pack(side="left", padx=2)

        self.btn_prev = ctk.CTkButton(
            nav_box,
            text="◀ ก่อนหน้า",
            width=80,
            height=28,
            corner_radius=6,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self._go_prev_page,
        )
        self.btn_prev.pack(side="left", padx=2)

        self.lbl_page_num = ctk.CTkLabel(
            nav_box,
            text="หน้า 1 / 1",
            width=95,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            text_color=AppTheme.COLOR_TEXT,
        )
        self.lbl_page_num.pack(side="left", padx=6)

        self.btn_next = ctk.CTkButton(
            nav_box,
            text="ถัดไป ▶",
            width=80,
            height=28,
            corner_radius=6,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self._go_next_page,
        )
        self.btn_next.pack(side="left", padx=2)

        self.btn_last = ctk.CTkButton(
            nav_box,
            text="สุดท้าย ⏭",
            width=75,
            height=28,
            corner_radius=6,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=(AppTheme.FONT_FAMILY, 12, "bold"),
            command=self._go_last_page,
        )
        self.btn_last.pack(side="left", padx=2)

        hsb = ttk.Scrollbar(self, orient="horizontal")
        hsb.pack(side="bottom", fill="x", padx=6, pady=(0, 4))

        table_box = ctk.CTkFrame(self, fg_color="transparent")
        table_box.pack(side="top", fill="both", expand=True, padx=6, pady=(6, 0))

        col_ids = [c[0] for c in columns]
        self.tree = ttk.Treeview(
            table_box,
            columns=col_ids,
            show="headings",
            height=min(height, 7),
            style="GodApp.Treeview",
            selectmode="browse",
        )

        for col_id, heading, width, anchor in columns:
            self.tree.heading(col_id, text=heading)
            self.tree.column(col_id, width=width, minwidth=max(70, width - 30), anchor=anchor, stretch=True)

        self.tree.tag_configure("PAID", background="#FFFFFF", foreground="#0F172A", font=(AppTheme.FONT_FAMILY, 13))
        self.tree.tag_configure("ODD", background="#F8FAFC", foreground="#0F172A", font=(AppTheme.FONT_FAMILY, 13))
        self.tree.tag_configure("VOIDED", background="#FEF2F2", foreground="#991B1B", font=(AppTheme.FONT_FAMILY, 13))

        vsb = ttk.Scrollbar(table_box, orient="vertical", command=self.tree.yview)
        hsb.configure(command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.bind("<<TreeviewSelect>>", self._handle_select)

    def _handle_select(self, _event):
        vals = self.get_selected_values()
        if vals and self.on_select:
            self.on_select(vals)

    def _go_first_page(self):
        if self.current_page > 1:
            self._change_page(1)

    def _go_prev_page(self):
        if self.current_page > 1:
            self._change_page(self.current_page - 1)

    def _go_next_page(self):
        if self.current_page < self.total_pages:
            self._change_page(self.current_page + 1)

    def _go_last_page(self):
        if self.current_page < self.total_pages:
            self._change_page(self.total_pages)

    def _change_page(self, new_page: int):
        self.current_page = max(1, min(new_page, self.total_pages))
        if self.on_page_change:
            self.on_page_change(self.current_page)
        else:
            self._render_client_page()

    def _render_tree_rows(self, rows: List[Tuple[List[str], str]]):
        self.clear()
        for idx, (vals, tag) in enumerate(rows):
            if tag == "VOIDED":
                row_tag = "VOIDED"
            else:
                row_tag = "PAID" if idx % 2 == 0 else "ODD"
            self.tree.insert("", "end", values=vals, tags=(row_tag,))

    def _update_pagination_ui(self, shown_count: int):
        self.total_pages = max(1, (self.total_count + self.page_size - 1) // self.page_size)
        self.current_page = max(1, min(self.current_page, self.total_pages))
        if self.total_count == 0 or shown_count == 0:
            self.lbl_page_info.configure(text="ไม่พบรายการข้อมูล (0 รายการ)")
        else:
            start_idx = (self.current_page - 1) * self.page_size + 1
            end_idx = start_idx + shown_count - 1
            self.lbl_page_info.configure(
                text=f"แสดง {start_idx:,} – {end_idx:,} จากทั้งหมด {self.total_count:,} รายการ"
            )
        self.lbl_page_num.configure(text=f"หน้า {self.current_page:,} / {self.total_pages:,}")
        prev_state = "normal" if self.current_page > 1 else "disabled"
        next_state = "normal" if self.current_page < self.total_pages else "disabled"
        self.btn_first.configure(state=prev_state)
        self.btn_prev.configure(state=prev_state)
        self.btn_next.configure(state=next_state)
        self.btn_last.configure(state=next_state)

    def _render_client_page(self):
        start = (self.current_page - 1) * self.page_size
        end = start + self.page_size
        page_slice = self._client_rows[start:end]
        self._render_tree_rows(page_slice)
        self._update_pagination_ui(len(page_slice))

    def clear(self):
        items = self.tree.get_children()
        if items:
            self.tree.delete(*items)

    def set_rows(
        self,
        rows: List[Tuple[List[str], str]],
        total_count: Optional[int] = None,
        page: Optional[int] = None,
    ):
        if page is not None:
            self.current_page = max(1, int(page))

        if total_count is not None:
            # Server/DB-side paginated mode: `rows` is already the current page slice
            self.total_count = max(len(rows), int(total_count))
            self._render_tree_rows(rows)
            self._update_pagination_ui(len(rows))
        else:
            # Client-side mode: slice to page_size so Treeview never freezes on 100k+ rows
            self._client_rows = list(rows)
            self.total_count = len(self._client_rows)
            if page is None:
                self.current_page = 1
            if self.total_count > self.page_size and not self.pagination_bar.winfo_ismapped():
                self.pagination_bar.pack(side="bottom", fill="x", padx=6, pady=(0, 6))
            self._render_client_page()

    def get_selected_values(self) -> Optional[List[str]]:
        sel = self.tree.selection()
        if not sel:
            return None
        item = self.tree.item(sel[0])
        return [str(v) for v in item.get("values", [])]
