"""Reusable UI theme, components, views, and modern notifications for College Finance God App."""

from ui.notifications import (
    ModernDialog,
    ModernMessageBox,
    ModernToast,
    askokcancel,
    askquestion,
    askretrycancel,
    askyesno,
    highlight_missing_field,
    patch_tkinter_messagebox,
    show_toast,
    showerror,
    showinfo,
    showwarning,
)

# Apply global modern styling to standard messagebox dialogs
patch_tkinter_messagebox()
