"""
Unified Corporate Financial Design System for SPVC Smart Receipt System.
Cohesive Navy, Slate & Soft Ice-Blue palette with zero clashing colors.
"""


class AppTheme:
    FONT_FAMILY = "Tahoma"
    FONT_MONO = "Consolas"

    # Spacing & Control Sizing Tokens
    SPACE_XS = 4
    SPACE_SM = 8
    SPACE_MD = 16
    SPACE_LG = 24

    CTRL_H_SM = 32
    CTRL_H_MD = 36
    CTRL_H_LG = 42

    # Typography Tokens
    FONT_TITLE = (FONT_FAMILY, 18, "bold")
    FONT_SECTION = (FONT_FAMILY, 14, "bold")
    FONT_BODY = (FONT_FAMILY, 13)
    FONT_LABEL = (FONT_FAMILY, 12, "bold")
    FONT_CAPTION = (FONT_FAMILY, 11)

    # Corner Radius (Pure Flat Rectangular Design - No Rounded Clutter)
    RADIUS = 0
    RADIUS_CARD = 0
    RADIUS_BTN = 0
    RADIUS_INPUT = 0

    # 1. App Surfaces (Warm-Cool Corporate Slate)
    COLOR_BG = "#F8FAFC"            # Main workspace background (Slate 50 - calm & airy)
    COLOR_SURFACE = "#FFFFFF"       # Primary card surface (Pure White)
    COLOR_SURFACE_ALT = "#F1F5F9"   # Secondary panel / subtle strip (Slate 100)
    COLOR_SURFACE_MUTED = "#F1F5F9" # Input / control background

    # 2. Sidebar Navigation (Deep Executive Navy)
    COLOR_SIDEBAR = "#0F172A"        # Slate 900
    COLOR_SIDEBAR_SURFACE = "#1E293B" # Slate 800
    COLOR_SIDEBAR_HOVER = "#1E293B"  # Slate 800
    COLOR_SIDEBAR_ACTIVE = "#1E40AF" # Corporate Royal Navy (Blue 800)

    # 3. Primary & Brand Accents (Cohesive Corporate Navy-Blue)
    COLOR_ACCENT = "#1E3A8A"             # Deep Navy
    COLOR_ACCENT_HOVER = "#172554"       # Darker Navy
    COLOR_PRIMARY_BLUE = "#1E40AF"       # Primary CTA Blue (Blue 800 - authoritative & calm)
    COLOR_PRIMARY_BLUE_HOVER = "#1E3A8A" # Blue 900
    COLOR_SECONDARY_BG = "#EFF6FF"       # Soft Ice Blue tint (Blue 50)
    COLOR_SECONDARY_HOVER = "#DBEAFE"    # Blue 100
    COLOR_SECONDARY_BORDER = "#BFDBFE"   # Blue 200
    COLOR_SECONDARY_TEXT = "#1E40AF"     # Blue 800

    # 4. Borders & Dividers
    COLOR_BORDER = "#E2E8F0"        # Subtle card border (Slate 200)
    COLOR_BORDER_STRONG = "#CBD5E1" # Input border (Slate 300)
    COLOR_BORDER_LIGHT = "#F1F5F9"  # Hairline divider

    # 5. Unified Form Controls (Dropdowns & Entries match seamlessly)
    COLOR_INPUT_BG = "#F8FAFC"
    COLOR_INPUT_BORDER = "#CBD5E1"
    COLOR_DROPDOWN_FG = "#F1F5F9"
    COLOR_DROPDOWN_BTN = "#E2E8F0"
    COLOR_DROPDOWN_HOVER = "#CBD5E1"
    COLOR_DROPDOWN_MENU = "#FFFFFF"
    COLOR_DROPDOWN_TEXT = "#0F172A"

    # 6. Typography Hierarchy
    COLOR_TEXT = "#0F172A"           # Primary heading & body (Slate 900)
    COLOR_TEXT_SECONDARY = "#334155" # Labels & subheadings (Slate 700)
    COLOR_TEXT_MUTED = "#64748B"     # Captions & hints (Slate 500)
    COLOR_TEXT_LIGHT = "#FFFFFF"

    # 7. Semantic Actions (Toned-down, harmonious enterprise shades)
    # Primary Save / Confirm uses deep corporate teal-emerald or navy so it doesn't clash
    COLOR_SUCCESS = "#0F766E"        # Deep Corporate Teal (Teal 700 - blends with Navy/Blue)
    COLOR_SUCCESS_HOVER = "#115E59"  # Teal 800
    COLOR_SUCCESS_BG = "#F0FDFA"     # Teal 50

    COLOR_DANGER = "#B91C1C"         # Muted Crimson (Red 700)
    COLOR_DANGER_HOVER = "#991B1B"   # Red 800
    COLOR_DANGER_BG = "#FEF2F2"      # Soft Rose tint (Red 50)
    COLOR_DANGER_BORDER = "#FECACA"  # Red 200

    COLOR_WARNING = "#B45309"
    COLOR_WARNING_HOVER = "#92400E"
    COLOR_WARNING_BG = "#FFFBEB"

    COLOR_INFO = "#1E40AF"
    COLOR_INFO_BG = "#EFF6FF"

    COLOR_SLATE = "#475569"
    COLOR_SLATE_HOVER = "#334155"
    COLOR_NEUTRAL_BTN = "#F1F5F9"
    COLOR_NEUTRAL_BTN_HOVER = "#E2E8F0"

    STATUS_THAI_MAP = {
        "PAID": "✔ ชำระครบ",
        "PARTIAL": "⏳ ค้างบางส่วน",
        "UNPAID": "✖ ค้างชำระ",
        "ISSUED": "✔ ปกติ",
        "COMPLETED": "✔ ปกติ",
        "VOIDED": "🚫 ยกเลิกแล้ว",
    }

    @classmethod
    def format_status(cls, code: str) -> str:
        return cls.STATUS_THAI_MAP.get((code or "").upper(), code)

    @classmethod
    def apply_sharp_theme(cls):
        """
        Forces all CustomTkinter widgets across the entire system to have
        pure sharp rectangular edges (corner_radius = 0, no rounded corners).
        """
        try:
            import customtkinter as ctk
            for _widget_name, props in ctk.ThemeManager.theme.items():
                if isinstance(props, dict):
                    for k in list(props.keys()):
                        if "corner_radius" in k:
                            props[k] = 0
        except Exception:
            pass


# Automatically hook and apply pure sharp rectangular geometry across CustomTkinter
try:
    import customtkinter as _ctk
    _orig_set_default_color_theme = _ctk.set_default_color_theme

    def _set_default_color_theme_sharp(theme):
        _orig_set_default_color_theme(theme)
        AppTheme.apply_sharp_theme()

    _ctk.set_default_color_theme = _set_default_color_theme_sharp
    AppTheme.apply_sharp_theme()
except Exception:
    pass
