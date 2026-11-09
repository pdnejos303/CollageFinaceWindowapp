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
    Harmonious, Enterprise-Grade Table.
    Soft Slate-100 header, subtle zebra rows, and calm ice-blue selection highlight.
    """

    def __init__(
        self,
        parent,
        columns: List[Tuple[str, str, int, str]],
        height: int = 6,
        on_select: Optional[Callable[[List[str]], None]] = None,
        show_inspector: bool = False,
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

    def clear(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def set_rows(self, rows: List[Tuple[List[str], str]]):
        self.clear()
        for idx, (vals, tag) in enumerate(rows):
            if tag == "VOIDED":
                row_tag = "VOIDED"
            else:
                row_tag = "PAID" if idx % 2 == 0 else "ODD"
            self.tree.insert("", "end", values=vals, tags=(row_tag,))

    def get_selected_values(self) -> Optional[List[str]]:
        sel = self.tree.selection()
        if not sel:
            return None
        item = self.tree.item(sel[0])
        return [str(v) for v in item.get("values", [])]
