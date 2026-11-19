import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Dict, List, Optional, Tuple
import weakref
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
        corner_radius: int = AppTheme.RADIUS_INPUT,
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
        allow_custom: bool = True,
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
        self._allow_custom = bool(allow_custom)
        self._arrow_navigated = False
        self._sb: Optional[ttk.Scrollbar] = None

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
            corner_radius=corner_radius,
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
        self.bind("<Button-1>", self._on_entry_click)
        if hasattr(self, "_canvas") and self._canvas is not None:
            try:
                self._canvas.bind("<Button-1>", self._on_entry_click)
            except Exception:
                pass

        self.entry.bind("<Button-1>", self._on_entry_click)
        self.entry.bind("<KeyPress>", self._on_key_press)
        self.entry.bind("<KeyRelease>", self._on_key_release)
        self.entry.bind("<Down>", self._on_arrow_down)
        self.entry.bind("<Up>", self._on_arrow_up)
        self.entry.bind("<Return>", self._on_enter)
        self.entry.bind("<KP_Enter>", self._on_enter)
        self.entry.bind("<Escape>", self._on_escape)
        self.entry.bind("<Tab>", self._on_tab)
        self.entry.bind("<FocusIn>", self._on_focus_in)
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
            self.entry.icursor(0)
            self.entry.xview_moveto(0)
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
        if hasattr(self, "entry") and self.entry.winfo_exists():
            entry_text = self.entry.get().strip()
            if entry_text and entry_text != self._committed_value:
                self._committed_value = entry_text
                if self._variable is not None and self._variable.get() != entry_text:
                    self._updating_entry = True
                    try:
                        self._variable.set(entry_text)
                    finally:
                        self._updating_entry = False
                return entry_text
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
        if "allow_custom" in kwargs:
            self._allow_custom = bool(kwargs.pop("allow_custom"))
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

    def _on_focus_in(self, _event=None):
        if self._state == "disabled":
            return
        self._select_all_entry()

    def _select_all_entry(self):
        try:
            self.entry.select_range(0, "end")
            self.entry.icursor("end")
        except Exception:
            pass

    def _on_entry_click(self, _event=None):
        if self._state == "disabled":
            return
        if self.focus_get() != self.entry:
            self.entry.focus_set()
            self._select_all_entry()
        if self._popup is None or not self._popup.winfo_exists():
            self._open_popup(filter_query="")

    def _on_key_press(self, event):
        if self._state == "disabled":
            return
        # If user presses printable character while text is selected, delete selection so typed character replaces it cleanly
        if event.char and event.char.isprintable() and not (event.state & 4):
            try:
                if self.entry.select_present():
                    self.entry.delete("sel.first", "sel.last")
            except Exception:
                pass
        elif event.keysym in ("BackSpace", "Delete"):
            try:
                if self.entry.select_present():
                    self.entry.delete("sel.first", "sel.last")
                    self._on_key_release(event)
                    return "break"
            except Exception:
                pass

    def _toggle_popup(self):
        if self._state == "disabled":
            return
        if self._popup is not None and self._popup.winfo_exists():
            self._commit_or_restore_and_close()
        else:
            self.entry.focus_set()
            self._select_all_entry()
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
        delay = 120 if self._on_search is not None else 30
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
        self._arrow_navigated = False

        self._popup = tk.Toplevel(self)
        self._popup.withdraw()
        self._popup.wm_overrideredirect(True)
        try:
            self._popup.wm_attributes("-topmost", True)
        except Exception:
            pass
        self._popup.configure(bg=AppTheme.COLOR_BORDER_STRONG, padx=1, pady=1)

        container = tk.Frame(self._popup, bg=self._dropdown_fg_color, bd=0)
        container.pack(fill="both", expand=True)

        self._sb = ttk.Scrollbar(container, orient="vertical")
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
            yscrollcommand=self._sb.set,
        )
        self._sb.configure(command=self._listbox.yview)
        self._listbox.pack(side="left", fill="both", expand=True, padx=2, pady=2)

        self._listbox.bind("<ButtonRelease-1>", self._on_listbox_click)
        self._listbox.bind("<Motion>", self._on_listbox_motion)
        self._listbox.bind("<MouseWheel>", self._on_mousewheel)
        self._listbox.bind("<Return>", self._on_enter)
        self._listbox.bind("<Escape>", self._on_escape)

        self._refresh_listbox(filter_query=filter_query, sb=self._sb)

    def _refresh_listbox(self, filter_query: str = "", sb: Optional[ttk.Scrollbar] = None):
        if self._listbox is None or self._popup is None or not self._popup.winfo_exists():
            return

        q = filter_query.strip().lower()
        if not q or self._on_search is not None:
            matched = list(self._values)
        else:
            tokens = [t for t in q.split() if t]
            starts = [v for v in self._values if v.lower().startswith(q)]
            token_matches = [
                v for v in self._values
                if all(tok in v.lower() for tok in tokens) and not v.lower().startswith(q)
            ]
            matched = starts + token_matches

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

        # Position & size popup cleanly directly beneath the dropdown widget
        self.update_idletasks()
        count = max(1, len(self._filtered_values))
        visible_rows = min(count, 8)
        row_h = 26
        popup_h = visible_rows * row_h + 8

        scrollbar = sb or getattr(self, "_sb", None)
        if scrollbar is not None:
            if count > 8:
                scrollbar.pack(side="right", fill="y")
            else:
                scrollbar.pack_forget()

        # The popup width should match the dropdown widget width, but be capped at a readable max (560px)
        w = self.winfo_width()
        if w <= 1:
            w = getattr(self, "_current_width", 160)
        popup_w = max(160, min(w, 560))

        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 2

        if x <= 0 and y <= 0:
            try:
                self.winfo_toplevel().update_idletasks()
                x = self.winfo_rootx()
                y = self.winfo_rooty() + self.winfo_height() + 2
            except Exception:
                pass

        try:
            scale = ctk.ScalingTracker.get_window_scaling(self.winfo_toplevel())
        except Exception:
            scale = 1.0

        physical_screen_w = int(self.winfo_screenwidth() * scale)
        physical_screen_h = int(self.winfo_screenheight() * scale)

        # If popup exceeds bottom of screen, open upwards above the dropdown
        if y + popup_h > physical_screen_h - 40:
            y = max(10, self.winfo_rooty() - popup_h - 2)

        # Keep popup width within physical screen width without moving x unless near screen edge
        if x + popup_w > physical_screen_w - 10:
            if x < physical_screen_w - 150:
                popup_w = max(140, physical_screen_w - x - 10)
            else:
                x = max(10, physical_screen_w - popup_w - 10)

        self._popup.wm_geometry(f"{popup_w}x{popup_h}+{x}+{y}")
        self._popup.deiconify()
        self._popup.lift()

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
        self.entry.icursor(0)
        self.entry.xview_moveto(0)
        self.entry.select_clear()
        if self._command is not None:
            self._command(chosen)

    def _on_arrow_down(self, _event=None):
        if self._state == "disabled":
            return "break"
        if self._popup is None or not self._popup.winfo_exists():
            self._open_popup(filter_query="")
            return "break"
        self._arrow_navigated = True
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
        self._arrow_navigated = True
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
        typed = self.entry.get().strip()

        # If user navigated the list with arrow keys or if popup has filtered items
        if self._popup is not None and self._popup.winfo_exists() and self._listbox is not None and self._filtered_values:
            # Check exact match first
            for v in self._filtered_values:
                if v.lower() == typed.lower() and not (v.startswith("---") and v.endswith("---")):
                    self._select_value(v)
                    return "break"
            # If navigated via arrow keys
            if getattr(self, "_arrow_navigated", False):
                cur = self._listbox.curselection()
                idx = cur[0] if cur else 0
                if 0 <= idx < len(self._filtered_values):
                    chosen = self._filtered_values[idx]
                    if not (chosen.startswith("---") and chosen.endswith("---")):
                        self._select_value(chosen)
                        return "break"
            # If user typed something and there is at least one match
            if self._filtered_values:
                chosen = self._filtered_values[0]
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

    def cget(self, attribute_name: str) -> Any:
        if attribute_name == "values":
            return list(self._values)
        if attribute_name == "state":
            return self._state
        return super().cget(attribute_name)

    def _check_focus_after_delay(self):
        if not self.winfo_exists():
            return
        try:
            if self._popup is not None and self._popup.winfo_exists():
                x, y = self.winfo_pointerxy()
                px, py = self._popup.winfo_rootx(), self._popup.winfo_rooty()
                pw, ph = self._popup.winfo_width(), self._popup.winfo_height()
                if px <= x <= px + pw and py <= y <= py + ph:
                    return
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
        # Exact match (case-insensitive) against existing values
        for v in self._values:
            if v.lower() == typed.lower() and not (v.startswith("---") and v.endswith("---")):
                self._select_value(v)
                return
        # Accept custom typed text (default True)
        if getattr(self, "_allow_custom", True):
            if typed not in self._values:
                self._values.append(typed)
                if typed not in self._filtered_values:
                    self._filtered_values.append(typed)
            self._select_value(typed)
            return
        # Restore committed value only if allow_custom is False
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
            self._sb = None
        self._arrow_navigated = False
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
    """Minimal KPI Card with unified corporate styling and subtle left accent strip."""

    def __init__(
        self,
        parent,
        title: str,
        value: str,
        subtitle: str = "",
        accent_color: str = AppTheme.COLOR_PRIMARY_BLUE,
        badge_bg: str = AppTheme.COLOR_INFO_BG,
        corner_radius: int = AppTheme.RADIUS_CARD,
        border_width: int = 1,
    ):
        super().__init__(
            parent,
            corner_radius=corner_radius,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=border_width,
            border_color=AppTheme.COLOR_BORDER,
        )
        strip = ctk.CTkFrame(self, width=3, fg_color=accent_color, corner_radius=0)
        strip.pack(side="left", fill="y")

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(side="left", fill="both", expand=True, padx=18, pady=14)

        self.lbl_title = ctk.CTkLabel(
            body,
            text=title,
            font=AppTheme.FONT_LABEL,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            anchor="w",
        )
        self.lbl_title.pack(fill="x")

        self.lbl_value = ctk.CTkLabel(
            body,
            text=value,
            font=(AppTheme.FONT_FAMILY, 22, "bold"),
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        )
        self.lbl_value.pack(fill="x", pady=(4, 0))

        self.lbl_subtitle = None
        if subtitle:
            self.lbl_subtitle = ctk.CTkLabel(
                body,
                text=subtitle,
                font=AppTheme.FONT_CAPTION,
                text_color=AppTheme.COLOR_TEXT_MUTED,
                anchor="w",
            )
            self.lbl_subtitle.pack(fill="x", pady=(3, 0))

    def set_value(self, value: str, subtitle: Optional[str] = None):
        """Dynamically update card value and optional subtitle."""
        self.lbl_value.configure(text=value)
        if subtitle is not None:
            if self.lbl_subtitle is None:
                # Find body frame to add subtitle label
                body = self.lbl_value.master
                self.lbl_subtitle = ctk.CTkLabel(
                    body,
                    text=subtitle,
                    font=AppTheme.FONT_CAPTION,
                    text_color=AppTheme.COLOR_TEXT_MUTED,
                    anchor="w",
                )
                self.lbl_subtitle.pack(fill="x", pady=(3, 0))
            else:
                self.lbl_subtitle.configure(text=subtitle)


class SectionHeader(ctk.CTkFrame):
    """Clean, single-line title bar with unified corporate styling."""

    def __init__(
        self,
        parent,
        title: str,
        subtitle: str = "",
        corner_radius: int = AppTheme.RADIUS_CARD,
        border_width: int = 1,
    ):
        super().__init__(
            parent,
            corner_radius=corner_radius,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=border_width,
            border_color=AppTheme.COLOR_BORDER,
        )
        inner = ctk.CTkFrame(self, fg_color="transparent")
        inner.pack(fill="x", padx=20, pady=14)

        # Pack actions_frame to the RIGHT first, so action buttons are NEVER squeezed or crushed by title expansion
        self.actions_frame = ctk.CTkFrame(inner, fg_color="transparent")
        self.actions_frame.pack(side="right")
        self._actions_packed = True

        left = ctk.CTkFrame(inner, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            left,
            text=title,
            font=AppTheme.FONT_TITLE,
            text_color=AppTheme.COLOR_TEXT,
            anchor="w",
        ).pack(anchor="w")

        if subtitle:
            ctk.CTkLabel(
                left,
                text=subtitle,
                font=(AppTheme.FONT_FAMILY, 12),
                text_color=AppTheme.COLOR_TEXT_MUTED,
                anchor="w",
            ).pack(anchor="w", pady=(2, 0))

    def add_action_button(
        self,
        text: str,
        command: Callable,
        fg_color: str = AppTheme.COLOR_SECONDARY_BG,
        text_color: str = AppTheme.COLOR_SECONDARY_TEXT,
        hover_color: str = AppTheme.COLOR_SECONDARY_HOVER,
        corner_radius: int = AppTheme.RADIUS_BTN,
    ) -> ctk.CTkButton:
        if not self._actions_packed:
            self.actions_frame.pack(side="right")
            self._actions_packed = True
        btn = ctk.CTkButton(
            self.actions_frame,
            text=text,
            command=command,
            corner_radius=corner_radius,
            height=AppTheme.CTRL_H_MD,
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

    _styled_roots: "weakref.WeakSet[Any]" = weakref.WeakSet()

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
        border_width: int = 1,
        corner_radius: int = AppTheme.RADIUS_CARD,
        selectmode: str = "extended",
    ):
        super().__init__(
            parent,
            corner_radius=corner_radius,
            fg_color=AppTheme.COLOR_SURFACE,
            border_width=border_width,
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

        # File-Explorer style Drag-Selection & Multi-selection support
        self._drag_anchor: Optional[str] = None
        self._drag_start_item: Optional[str] = None
        self._pre_drag_selection: Tuple[str, ...] = ()
        self._is_dragging: bool = False

        root = self._root()
        style = ttk.Style()
        if root not in DataTable._styled_roots or style.theme_use() != "clam":
            try:
                style.theme_use("clam")
            except Exception:
                pass

            style.configure(
                "GodApp.Treeview",
                background="#FFFFFF",
                foreground=AppTheme.COLOR_TEXT,
                rowheight=40,
                fieldbackground="#FFFFFF",
                font=AppTheme.FONT_BODY,
                borderwidth=0,
            )
            style.configure(
                "GodApp.Treeview.Heading",
                background="#F8FAFC",
                foreground=AppTheme.COLOR_TEXT_SECONDARY,
                font=AppTheme.FONT_LABEL,
                padding=(10, 8),
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
                background=[("active", "#F1F5F9")],
            )
            if root is not None:
                DataTable._styled_roots.add(root)

        # Optional Pagination Footer Bar
        self.pagination_bar = ctk.CTkFrame(self, fg_color="#F8FAFC", corner_radius=AppTheme.RADIUS, height=36)
        if self.show_pagination or self.on_page_change is not None:
            self.pagination_bar.pack(side="bottom", fill="x", padx=0, pady=0)

        self.lbl_page_info = ctk.CTkLabel(
            self.pagination_bar,
            text="แสดง 0 รายการ",
            font=AppTheme.FONT_LABEL,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
        )
        self.lbl_page_info.pack(side="left", padx=12, pady=4)

        nav_box = ctk.CTkFrame(self.pagination_bar, fg_color="transparent")
        nav_box.pack(side="right", padx=8, pady=4)

        self.btn_first = ctk.CTkButton(
            nav_box,
            text="<<",
            width=36,
            height=28,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=AppTheme.FONT_LABEL,
            command=self._go_first_page,
        )
        self.btn_first.pack(side="left", padx=2)

        self.btn_prev = ctk.CTkButton(
            nav_box,
            text="<",
            width=34,
            height=28,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=AppTheme.FONT_LABEL,
            command=self._go_prev_page,
        )
        self.btn_prev.pack(side="left", padx=2)

        self.lbl_page_num = ctk.CTkLabel(
            nav_box,
            text="หน้า 1 / 1",
            width=95,
            font=AppTheme.FONT_LABEL,
            text_color=AppTheme.COLOR_TEXT,
        )
        self.lbl_page_num.pack(side="left", padx=6)

        self.btn_next = ctk.CTkButton(
            nav_box,
            text=">",
            width=34,
            height=28,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=AppTheme.FONT_LABEL,
            command=self._go_next_page,
        )
        self.btn_next.pack(side="left", padx=2)

        self.btn_last = ctk.CTkButton(
            nav_box,
            text=">>",
            width=36,
            height=28,
            corner_radius=AppTheme.RADIUS_BTN,
            fg_color=AppTheme.COLOR_NEUTRAL_BTN,
            hover_color=AppTheme.COLOR_NEUTRAL_BTN_HOVER,
            text_color=AppTheme.COLOR_TEXT_SECONDARY,
            font=AppTheme.FONT_LABEL,
            command=self._go_last_page,
        )
        self.btn_last.pack(side="left", padx=2)

        hsb = ttk.Scrollbar(self, orient="horizontal")
        hsb.pack(side="bottom", fill="x", padx=0, pady=0)

        table_box = ctk.CTkFrame(self, fg_color="transparent")
        table_box.pack(side="top", fill="both", expand=True, padx=0, pady=0)

        col_ids = [c[0] for c in columns]
        self.tree = ttk.Treeview(
            table_box,
            columns=col_ids,
            show="headings",
            height=height,
            style="GodApp.Treeview",
            selectmode=selectmode,
        )

        for col_id, heading, width, anchor in columns:
            self.tree.heading(col_id, text=heading)
            self.tree.column(col_id, width=width, minwidth=max(35, width - 25), anchor=anchor, stretch=True)

        self.tree.tag_configure("PAID", background="#FFFFFF", foreground="#0F172A", font=(AppTheme.FONT_FAMILY, 13))
        self.tree.tag_configure("ODD", background="#F8FAFC", foreground="#0F172A", font=(AppTheme.FONT_FAMILY, 13))
        self.tree.tag_configure("VOIDED", background="#FEF2F2", foreground="#991B1B", font=(AppTheme.FONT_FAMILY, 13))

        vsb = ttk.Scrollbar(table_box, orient="vertical", command=self.tree.yview)
        hsb.configure(command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.bind("<<TreeviewSelect>>", self._handle_select)
        self.tree.bind("<ButtonPress-1>", self._on_tree_button_press, add=True)
        self.tree.bind("<B1-Motion>", self._on_tree_b1_motion, add=True)
        self.tree.bind("<ButtonRelease-1>", self._on_tree_button_release, add=True)
        self.tree.bind("<Control-a>", self._on_select_all)
        self.tree.bind("<Control-A>", self._on_select_all)

    def _on_tree_button_press(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region in ("heading", "separator"):
            return

        item = self.tree.identify_row(event.y)
        self._is_dragging = False
        self._drag_start_item = item
        self._pre_drag_selection = self.tree.selection()

        ctrl_held = bool(event.state & 0x0004)
        shift_held = bool(event.state & 0x0001)

        if item:
            if not shift_held and not ctrl_held:
                self._drag_anchor = item
            elif shift_held and not self._drag_anchor:
                self._drag_anchor = item
        else:
            if not ctrl_held and not shift_held:
                self.tree.selection_set([])
                self._drag_anchor = None
                self._handle_select(None)

    def _on_tree_b1_motion(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region in ("heading", "separator"):
            return

        cur_item = self.tree.identify_row(event.y)
        all_items = self.tree.get_children("")
        if not all_items:
            return

        if not cur_item:
            if event.y < 10:
                cur_item = all_items[0]
            elif event.y > self.tree.winfo_height() - 10:
                cur_item = all_items[-1]
            else:
                return

        self._is_dragging = True
        anchor = self._drag_anchor or self._drag_start_item
        if not anchor:
            self._drag_anchor = cur_item
            anchor = cur_item

        if anchor in all_items and cur_item in all_items:
            idx1 = all_items.index(anchor)
            idx2 = all_items.index(cur_item)
            start_idx, end_idx = min(idx1, idx2), max(idx1, idx2)
            dragged_slice = all_items[start_idx : end_idx + 1]

            ctrl_held = bool(event.state & 0x0004)
            if ctrl_held:
                new_sel = list(dict.fromkeys(list(self._pre_drag_selection) + list(dragged_slice)))
            else:
                new_sel = list(dragged_slice)

            self.tree.selection_set(new_sel)
            self._handle_select(None)

        if event.y < 10:
            self.tree.yview_scroll(-1, "units")
        elif event.y > self.tree.winfo_height() - 10:
            self.tree.yview_scroll(1, "units")

    def _on_tree_button_release(self, _event):
        if self._is_dragging:
            self._is_dragging = False
            self._handle_select(None)

    def _on_select_all(self, _event=None):
        all_items = self.tree.get_children("")
        if all_items:
            self.tree.selection_set(all_items)
            self._handle_select(None)
        return "break"

    def _handle_select(self, _event):
        if self.on_select:
            vals = self.get_selected_values()
            self.on_select(vals if vals else [])

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
                self.pagination_bar.pack(side="bottom", fill="x", padx=0, pady=0)
            self._render_client_page()

    def get_selected_values(self) -> Optional[List[str]]:
        sel = self.tree.selection()
        if not sel:
            return None
        item = self.tree.item(sel[0])
        return [str(v) for v in item.get("values", [])]

    def get_selected_row(self) -> Optional[List[str]]:
        """Alias for get_selected_values to return values of the first selected row."""
        return self.get_selected_values()

    def get_all_selected_values(self) -> List[List[str]]:
        """Returns values of ALL currently selected rows in the table."""
        res: List[List[str]] = []
        for iid in self.tree.selection():
            item = self.tree.item(iid)
            vals = [str(v) for v in item.get("values", [])]
            if vals:
                res.append(vals)
        return res

    def get_selected_count(self) -> int:
        """Returns the number of selected rows."""
        return len(self.tree.selection())


def open_document_file(file_path: str, parent: Optional[Any] = None) -> bool:
    """
    Opens the exported document file (PDF, Word, Excel) with the OS default program
    so the cashier can immediately preview and print it.
    If opening fails, displays an informative error dialog instead of failing silently.
    """
    if not file_path or not os.path.exists(file_path):
        messagebox.showerror(
            "ไม่พบไฟล์เอกสาร",
            f"ระบบไม่พบไฟล์ที่ต้องการเปิดเพื่อสั่งพิมพ์:\n{file_path}\n\nกรุณาตรวจสอบว่าไฟล์ถูกลบหรือย้ายหรือไม่",
            parent=parent,
        )
        return False

    try:
        if sys.platform == "win32":
            os.startfile(os.path.abspath(file_path))
            return True
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", os.path.abspath(file_path)])
            return True
        else:
            import subprocess
            subprocess.Popen(["xdg-open", os.path.abspath(file_path)])
            return True
    except OSError as oe:
        messagebox.showwarning(
            "ไม่สามารถเปิดไฟล์อัตโนมัติได้",
            f"ไฟล์เอกสารถูกสร้างและบันทึกเรียบร้อยที่:\n{file_path}\n\n"
            f"⚠️ แต่ระบบไม่สามารถเปิดไฟล์ขึ้นมาอัตโนมัติได้:\n{oe}\n\n"
            f"(คำแนะนำ: เครื่องคอมพิวเตอร์นี้อาจยังไม่ได้ตั้งโปรแกรมเริ่มต้นสำหรับเปิดไฟล์ประเภทนี้ เช่น Adobe Acrobat, Microsoft Edge หรือ Word)\n"
            f"คุณสามารถเปิดไฟล์ดังกล่าวได้ด้วยตนเองจากโฟลเดอร์ที่บันทึกไว้ครับ",
            parent=parent,
        )
        return False
    except Exception as e:
        messagebox.showerror(
            "เกิดข้อผิดพลาดในการเปิดไฟล์",
            f"ไม่สามารถเปิดไฟล์เอกสารเพื่อสั่งพิมพ์ได้:\n{e}\n\nไฟล์อยู่ที่:\n{file_path}",
            parent=parent,
        )
        return False


def format_export_folder_button_text(config: Dict[str, Any], max_len: int = 16) -> str:
    """Returns a concise button label showing the current receipt export folder name."""
    raw = str(config.get("default_export_folder", "") or "").strip()
    if not raw:
        return "📁 เลือกที่เก็บไฟล์"
    folder_name = os.path.basename(os.path.normpath(raw)) or raw
    if len(folder_name) > max_len:
        folder_name = folder_name[: max_len - 1] + "…"
    return f"📁 ที่เก็บ: {folder_name}"


def choose_export_folder(
    parent: Any,
    config: Dict[str, Any],
    on_updated: Optional[Callable[[str], None]] = None,
) -> Optional[str]:
    """Prompts user to pick the default receipt export folder, saves config, and notifies."""
    current_dir = str(config.get("default_export_folder", "") or "").strip()
    if not current_dir or not os.path.isdir(current_dir):
        current_dir = os.getcwd()

    folder = filedialog.askdirectory(
        parent=parent,
        title="เลือกโฟลเดอร์สำหรับเก็บไฟล์ใบเสร็จ",
        initialdir=current_dir,
    )
    if not folder:
        return None

    config["default_export_folder"] = folder
    try:
        from core.config import save_config
        save_config(config)
    except Exception:
        pass

    if on_updated is not None:
        on_updated(folder)

    try:
        from ui.notifications import show_toast
        show_toast(
            parent=parent,
            title="ตั้งค่าที่เก็บไฟล์แล้ว",
            message=f"บันทึกใบเสร็จไว้ที่: {folder}",
            kind="success",
        )
    except Exception:
        pass

    return folder
