import tkinter as tk
from tkinter import ttk
from typing import Any, Callable, List, Optional, Tuple
import customtkinter as ctk

from ui.theme import AppTheme


class SearchableDropdown(ctk.CTkFrame):
    """
    Ultra-lightweight Searchable Dropdown (Combobox) for CustomTkinter.
    - Uses a single native tk.Entry + shared/lazy tk.Listbox popup for zero lag.
    - Supports instant typing to filter options, keyboard navigation (Up/Down/Enter/Esc),
      mouse wheel scrolling, and full CTkOptionMenu API compatibility (.get, .set, .configure).
    """

    _active_dropdown: Optional["SearchableDropdown"] = None

    def __init__(
        self,
        parent,
        values: Optional[List[str]] = None,
        variable: Optional[ctk.StringVar] = None,
        width: int = 160,
        height: int = 38,
        corner_radius: int = 8,
        fg_color: str = AppTheme.COLOR_DROPDOWN_FG,
        border_color: str = AppTheme.COLOR_INPUT_BORDER,
        border_width: int = 1,
        button_color: str = AppTheme.COLOR_DROPDOWN_BTN,
        button_hover_color: str = AppTheme.COLOR_DROPDOWN_HOVER,
        text_color: str = AppTheme.COLOR_DROPDOWN_TEXT,
        dropdown_fg_color: str = AppTheme.COLOR_DROPDOWN_MENU,
        dropdown_hover_color: str = AppTheme.COLOR_SECONDARY_BG,
        dropdown_text_color: str = AppTheme.COLOR_TEXT,
        font: Tuple[Any, ...] = (AppTheme.FONT_FAMILY, 13, "bold"),
        dropdown_font: Tuple[Any, ...] = (AppTheme.FONT_FAMILY, 13),
        command: Optional[Callable[[str], None]] = None,
        on_search: Optional[Callable[[str], Optional[List[str]]]] = None,
        max_display: int = 120,
        state: str = "normal",
        **kwargs,
    ):
        super().__init__(
            parent,
            width=width,
            height=height,
            corner_radius=corner_radius,
            fg_color=fg_color,
            border_width=border_width,
            border_color=border_color,
        )
        self.pack_propagate(False)
        self.grid_propagate(False)

        self._values: List[str] = list(values) if values else []
        self._filtered_values: List[str] = list(self._values)
        self._variable = variable
        self._command = command
        self._on_search = on_search
        self._max_display = max(20, int(max_display))
        self._state = state

        self._fg_color = fg_color
        self._text_color = text_color
        self._button_color = button_color
        self._button_hover_color = button_hover_color
        self._dropdown_fg_color = dropdown_fg_color
        self._dropdown_hover_color = dropdown_hover_color
        self._dropdown_text_color = dropdown_text_color
        self._font = font
        self._dropdown_font = dropdown_font

        self._popup: Optional[tk.Toplevel] = None
        self._listbox: Optional[tk.Listbox] = None
        self._search_job: Optional[str] = None
        self._updating_entry = False
        self._trace_name: Optional[str] = None

        initial_val = ""
        if self._variable is not None:
            initial_val = self._variable.get()
        elif self._values:
            initial_val = self._values[0]
        self._committed_value: str = initial_val

        # Arrow button on the right
        btn_w = max(26, min(32, height - 4))
        self.btn_arrow = ctk.CTkButton(
            self,
            text="▾",
            width=btn_w,
            height=max(20, height - 4),
            corner_radius=max(4, corner_radius - 2),
            fg_color=button_color,
            hover_color=button_hover_color,
            text_color=text_color,
            font=(AppTheme.FONT_FAMILY, 13, "bold"),
            command=self._toggle_popup,
        )
        self.btn_arrow.pack(side="right", padx=(0, 2), pady=2)

        # Native tk.Entry for ultra-fast, lightweight typing & filtering
        self.entry = tk.Entry(
            self,
            bd=0,
             highlightthickness=0,
            relief="flat",
            bg=fg_color,
            fg=text_color,
            insertbackground=text_color,
            selectbackground="#BFDBFE",
            selectforeground="#1E3A8A",
            disabledbackground=fg_color,
            disabledforeground="#94A3B8",
            font=font,
        )
        self.entry.pack(side="left", fill="both", expand=True, padx=(10, 4), pady=2)
        self._set_entry_text(initial_val)

        if self._state == "disabled":
            self.entry.configure(state="disabled")
            self.btn_arrow.configure(state="disabled")

        if self._variable is not None:
            self._trace_name = self._variable.trace_add("write", self._on_var_changed)

        # Bindings
        self.entry.bind("<Button-1>", self._on_entry_click)
        self.entry.bind("<KeyRelease>", self._on_key_release)
        self.entry.bind("<Down>", self._on_arrow_down)
        self.entry.bind("<Up>", self._on_arrow_up)
        self.entry.bind("<Return>", self._on_enter)
        self.entry.bind("<KP_Enter>", self._on_enter)
        self.entry.bind("<Escape>", self._on_escape)
        self.entry.bind("<Tab>", self._on_tab)
        self.entry.bind("<FocusOut>", self._on_focus_out)
        self.entry.bind("<MouseWheel>", self._on_mousewheel)
        self.bind("<Destroy>", self._on_destroy)

        self._ensure_root_hooks()

    def _ensure_root_hooks(self):
        try:
            top = self.winfo_toplevel()
            if not getattr(top, "_searchable_dd_hooked", False):
                top._searchable_dd_hooked = True
                top.bind("<Button-1>", SearchableDropdown._on_global_root_click, add="+")
                top.bind("<Configure>", SearchableDropdown._on_global_root_configure, add="+")
        except Exception:
            pass

    @classmethod
    def _on_global_root_click(cls, event):
        active = cls._active_dropdown
        if active is None:
            return
        try:
            w = event.widget
            if w is None:
                return
            w_str = str(w)
            if w_str.startswith(str(active)):
                return
            if active._popup is not None and w_str.startswith(str(active._popup)):
                return
            active._commit_or_restore_and_close()
        except Exception:
            active._close_popup()

    @classmethod
    def _on_global_root_configure(cls, event):
        active = cls._active_dropdown
        if active is None or active._popup is None:
            return
        try:
            if event.widget == active.winfo_toplevel():
                active._close_popup()
        except Exception:
            pass

    def _set_entry_text(self, text: str):
        self._updating_entry = True
        try:
            prev_state = str(self.entry.cget("state"))
            if prev_state == "disabled":
                self.entry.configure(state="normal")
            self.entry.delete(0, "end")
            self.entry.insert(0, str(text))
            if prev_state == "disabled":
                self.entry.configure(state="disabled")
        finally:
            self._updating_entry = False

    def _on_var_changed(self, *_args):
        if self._updating_entry or self._variable is None:
            return
        val = self._variable.get()
        self._committed_value = val
        self._set_entry_text(val)

    def get(self) -> str:
        if self._variable is not None:
            return self._variable.get()
        return self._committed_value

    def set(self, value: str):
        val = str(value)
        self._committed_value = val
        self._set_entry_text(val)
        if self._variable is not None and self._variable.get() != val:
            self._updating_entry = True
            try:
                self._variable.set(val)
            finally:
                self._updating_entry = False

    def configure(self, require_redraw: bool = False, **kwargs):
        if "values" in kwargs:
            vals = kwargs.pop("values")
            self._values = list(vals) if vals else []
            self._filtered_values = list(self._values)
            if self._popup is not None and self._popup.winfo_exists():
                self._refresh_listbox(self.entry.get().strip())
        if "variable" in kwargs:
            if self._variable is not None and self._trace_name is not None:
                try:
                    self._variable.trace_remove("write", self._trace_name)
                except Exception:
                    pass
            self._variable = kwargs.pop("variable")
            if self._variable is not None:
                self._trace_name = self._variable.trace_add("write", self._on_var_changed)
                self.set(self._variable.get())
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if "on_search" in kwargs:
            self._on_search = kwargs.pop("on_search")
        if "state" in kwargs:
            self._state = kwargs.pop("state")
            ent_state = "disabled" if self._state == "disabled" else "normal"
            self.entry.configure(state=ent_state)
            self.btn_arrow.configure(state=ent_state)
        if "fg_color" in kwargs:
            self._fg_color = kwargs["fg_color"]
            self.entry.configure(bg=self._fg_color, disabledbackground=self._fg_color)
        if "text_color" in kwargs:
            self._text_color = kwargs.pop("text_color")
            self.entry.configure(fg=self._text_color, insertbackground=self._text_color)
            self.btn_arrow.configure(text_color=self._text_color)
        if "button_color" in kwargs:
            self._button_color = kwargs.pop("button_color")
            self.btn_arrow.configure(fg_color=self._button_color)
        if "button_hover_color" in kwargs:
            self._button_hover_color = kwargs.pop("button_hover_color")
            self.btn_arrow.configure(hover_color=self._button_hover_color)
        for unused_k in ("dropdown_fg_color", "dropdown_hover_color", "dropdown_text_color", "font", "dropdown_font"):
            kwargs.pop(unused_k, None)
        if kwargs:
            super().configure(require_redraw=require_redraw, **kwargs)

    def _on_entry_click(self, _event=None):
        if self._state == "disabled":
            return
        self.entry.focus_set()
        self.after(10, lambda: self.entry.select_range(0, "end"))
        if self._popup is None or not self._popup.winfo_exists():
            self._open_popup(filter_query="")

    def _toggle_popup(self):
        if self._state == "disabled":
            return
        if self._popup is not None and self._popup.winfo_exists():
            self._commit_or_restore_and_close()
        else:
            self.entry.focus_set()
            self.entry.select_range(0, "end")
            self._open_popup(filter_query="")

    def _on_key_release(self, event):
        if self._state == "disabled":
            return
        if event.keysym in ("Up", "Down", "Return", "KP_Enter", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Left", "Right"):
            return
        if self._search_job is not None:
            try:
                self.after_cancel(self._search_job)
            except Exception:
                pass
        delay = 150 if self._on_search is not None else 40
        self._search_job = self.after(delay, self._perform_live_filter)

    def _perform_live_filter(self):
        self._search_job = None
        if not self.winfo_exists():
            return
        query = self.entry.get().strip()
        if self._on_search is not None:
            try:
                new_vals = self._on_search(query)
                if new_vals is not None:
                    self._values = list(new_vals)
            except Exception:
                pass
        if self._popup is None or not self._popup.winfo_exists():
            self._open_popup(filter_query=query)
        else:
            self._refresh_listbox(filter_query=query)

    def _open_popup(self, filter_query: str = ""):
        if SearchableDropdown._active_dropdown is not None and SearchableDropdown._active_dropdown is not self:
            SearchableDropdown._active_dropdown._commit_or_restore_and_close()

        self._close_popup()
        SearchableDropdown._active_dropdown = self

        self._popup = tk.Toplevel(self)
        self._popup.wm_overrideredirect(True)
        try:
            self._popup.wm_attributes("-topmost", True)
        except Exception:
            pass
        self._popup.configure(bg=AppTheme.COLOR_BORDER_STRONG, padx=1, pady=1)

        container = tk.Frame(self._popup, bg=self._dropdown_fg_color, bd=0)
        container.pack(fill="both", expand=True)

        sb = ttk.Scrollbar(container, orient="vertical")
        self._listbox = tk.Listbox(
            container,
            bg=self._dropdown_fg_color,
            fg=self._dropdown_text_color,
            selectbackground="#DBEAFE",
            selectforeground="#1E3A8A",
            font=self._dropdown_font,
            bd=0,
            highlightthickness=0,
            relief="flat",
            activestyle="none",
            exportselection=False,
            yscrollcommand=sb.set,
        )
        sb.configure(command=self._listbox.yview)
        self._listbox.pack(side="left", fill="both", expand=True, padx=2, pady=2)

        self._listbox.bind("<ButtonRelease-1>", self._on_listbox_click)
        self._listbox.bind("<Motion>", self._on_listbox_motion)
        self._listbox.bind("<MouseWheel>", self._on_mousewheel)
        self._listbox.bind("<Return>", self._on_enter)
        self._listbox.bind("<Escape>", self._on_escape)

        self._refresh_listbox(filter_query=filter_query, sb=sb)

    def _refresh_listbox(self, filter_query: str = "", sb: Optional[ttk.Scrollbar] = None):
        if self._listbox is None or self._popup is None or not self._popup.winfo_exists():
            return

        q = filter_query.strip().lower()
        if not q:
            matched = list(self._values)
        else:
            # Prioritize prefix matches first, then substring matches
            starts = [v for v in self._values if v.lower().startswith(q)]
            contains = [v for v in self._values if q in v.lower() and not v.lower().startswith(q)]
            matched = starts + contains

        self._filtered_values = matched[: self._max_display]
        self._listbox.delete(0, "end")

        if not self._filtered_values:
            self._listbox.insert("end", "   (ไม่พบตัวเลือกที่ค้นหา)")
            self._listbox.itemconfig(0, fg=AppTheme.COLOR_TEXT_MUTED)
        else:
            for item in self._filtered_values:
                self._listbox.insert("end", f"  {item}")

            # Highlight current or first match
            sel_idx = 0
            if not q and self._committed_value in self._filtered_values:
                sel_idx = self._filtered_values.index(self._committed_value)
            self._listbox.selection_clear(0, "end")
            self._listbox.selection_set(sel_idx)
            self._listbox.activate(sel_idx)
            self._listbox.see(sel_idx)

        # Position & size popup cleanly
        self.update_idletasks()
        count = max(1, len(self._filtered_values))
        visible_rows = min(count, 8)
        row_h = 26
        popup_h = visible_rows * row_h + 8

        if sb is not None:
            if count > 8:
                sb.pack(side="right", fill="y")
            else:
                sb.pack_forget()

        longest = max((len(str(v)) for v in self._filtered_values), default=10)
        min_w = max(self.winfo_width(), min(520, longest * 9 + 36))
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 2
        screen_h = self.winfo_screenheight()
        screen_w = self.winfo_screenwidth()
        if y + popup_h > screen_h - 40:
            y = max(10, self.winfo_rooty() - popup_h - 2)
        if x + min_w > screen_w - 16:
            x = max(8, screen_w - min_w - 16)

        self._popup.wm_geometry(f"{min_w}x{popup_h}+{x}+{y}")

    def _on_listbox_motion(self, event):
        if self._listbox is None or not self._filtered_values:
            return
        idx = self._listbox.nearest(event.y)
        if 0 <= idx < len(self._filtered_values):
            self._listbox.selection_clear(0, "end")
            self._listbox.selection_set(idx)
            self._listbox.activate(idx)

    def _on_listbox_click(self, event):
        if self._listbox is None or not self._filtered_values:
            self._close_popup()
            return
        idx = self._listbox.nearest(event.y)
        if 0 <= idx < len(self._filtered_values):
            chosen = self._filtered_values[idx]
            if chosen.startswith("---") and chosen.endswith("---"):
                return
            self._select_value(chosen)

    def _select_value(self, chosen: str):
        self._close_popup()
        self.set(chosen)
        self.entry.icursor("end")
        self.entry.select_clear()
        if self._command is not None:
            self._command(chosen)

    def _on_arrow_down(self, _event=None):
        if self._state == "disabled":
            return "break"
        if self._popup is None or not self._popup.winfo_exists():
            self._open_popup(filter_query="")
            return "break"
        if self._listbox is not None and self._filtered_values:
            cur = self._listbox.curselection()
            nxt = min(len(self._filtered_values) - 1, (cur[0] + 1) if cur else 0)
            self._listbox.selection_clear(0, "end")
            self._listbox.selection_set(nxt)
            self._listbox.activate(nxt)
            self._listbox.see(nxt)
        return "break"

    def _on_arrow_up(self, _event=None):
        if self._state == "disabled":
            return "break"
        if self._popup is None or not self._popup.winfo_exists():
            self._open_popup(filter_query="")
            return "break"
        if self._listbox is not None and self._filtered_values:
            cur = self._listbox.curselection()
            prv = max(0, (cur[0] - 1) if cur else 0)
            self._listbox.selection_clear(0, "end")
            self._listbox.selection_set(prv)
            self._listbox.activate(prv)
            self._listbox.see(prv)
        return "break"

    def _on_enter(self, _event=None):
        if self._state == "disabled":
            return "break"
        if self._popup is not None and self._popup.winfo_exists() and self._listbox is not None and self._filtered_values:
            cur = self._listbox.curselection()
            idx = cur[0] if cur else 0
            if 0 <= idx < len(self._filtered_values):
                chosen = self._filtered_values[idx]
                if not (chosen.startswith("---") and chosen.endswith("---")):
                    self._select_value(chosen)
                    return "break"
        self._commit_or_restore_and_close()
        return "break"

    def _on_escape(self, _event=None):
        self._set_entry_text(self._committed_value)
        self.entry.select_clear()
        self._close_popup()
        return "break"

    def _on_tab(self, _event=None):
        self._commit_or_restore_and_close()

    def _on_focus_out(self, _event=None):
        self.after(90, self._check_focus_after_delay)

    def _check_focus_after_delay(self):
        if not self.winfo_exists():
            return
        try:
            focused = self.focus_get()
            if focused is not None:
                f_str = str(focused)
                if f_str.startswith(str(self)):
                    return
                if self._popup is not None and f_str.startswith(str(self._popup)):
                    return
        except Exception:
            pass
        self._commit_or_restore_and_close()

    def _commit_or_restore_and_close(self):
        typed = self.entry.get().strip()
        self._close_popup()
        if not typed:
            self._set_entry_text(self._committed_value)
            return
        if typed == self._committed_value:
            return
        # Exact match (case-insensitive)
        for v in self._values:
            if v.lower() == typed.lower() and not (v.startswith("---") and v.endswith("---")):
                self._select_value(v)
                return
        # Restore committed value if no exact match was selected
        self._set_entry_text(self._committed_value)

    def _on_mousewheel(self, event):
        if self._popup is not None and self._popup.winfo_exists() and self._listbox is not None:
            delta = int(-1 * (event.delta / 120)) if event.delta else 0
            if delta != 0:
                self._listbox.yview_scroll(delta, "units")
            return "break"

    def _close_popup(self):
        if self._search_job is not None:
            try:
                self.after_cancel(self._search_job)
            except Exception:
                pass
            self._search_job = None
        if self._popup is not None:
            try:
                self._popup.destroy()
            except Exception:
                pass
            self._popup = None
            self._listbox = None
        if SearchableDropdown._active_dropdown is self:
            SearchableDropdown._active_dropdown = None

    def _on_destroy(self, event=None):
        if event is not None and event.widget is not self:
            return
        self._close_popup()
        if self._variable is not None and self._trace_name is not None:
            try:
                self._variable.trace_remove("write", self._trace_name)
            except Exception:
                pass
            self._trace_name = None



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
