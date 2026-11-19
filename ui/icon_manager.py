import math
import os
from typing import Dict, Tuple
from PIL import Image, ImageDraw
import customtkinter as ctk


class IconManager:
    """
    Manages high-definition vector PNG icons for SPVC Smart Receipt System.
    Replaces cheap emojis with elegant, high-definition monochrome and dual-tone vector icons.
    """

    _cache: Dict[Tuple[str, str, int, int], ctk.CTkImage] = {}
    _icons_ensured: bool = False
    ASSETS_DIR = os.path.join(os.path.dirname(__file__), "assets")
    ICONS_DIR = os.path.join(os.path.dirname(__file__), "assets", "icons")
    APP_ICON_ICO = os.path.join(ASSETS_DIR, "app_icon.ico")
    APP_ICON_PNG = os.path.join(ASSETS_DIR, "app_icon.png")

    COLORS = {
        "normal": "#94A3B8",  # Slate 400 (Inactive)
        "active": "#FFFFFF",  # Pure Crisp White (Active Selected)
        "accent": "#38BDF8",  # Sky Blue (Hover / Highlights)
    }

    @classmethod
    def get_app_icon_ico(cls) -> str:
        if os.path.exists(cls.APP_ICON_ICO):
            return cls.APP_ICON_ICO
        root_ico = os.path.join(os.path.dirname(os.path.dirname(__file__)), "app_icon.ico")
        if os.path.exists(root_ico):
            return root_ico
        return ""

    @classmethod
    def get_app_icon_png(cls) -> str:
        if os.path.exists(cls.APP_ICON_PNG):
            return cls.APP_ICON_PNG
        return ""

    @classmethod
    def set_window_icon(cls, window):
        """
        Sets the high-definition application icon for both titlebar and Windows taskbar.
        Uses native Win32 API WM_SETICON with DPI-aware system metrics, multi-resolution
        iconphoto, and iconbitmap to guarantee ultra-sharp rendering on all Windows scales.
        """
        import sys
        if sys.platform == "win32":
            try:
                import ctypes
                myappid = "spvc.collegefinance.godapp.2"
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
            except Exception:
                pass

        ico_path = cls.get_app_icon_ico()
        png_path = cls.get_app_icon_png()

        # 1. Multi-resolution iconphoto (anti-aliased Lanczos pre-scaled images)
        # Passing multiple sizes ensures Tkinter never forces downsampling from a single photo!
        if png_path and os.path.exists(png_path):
            try:
                from PIL import ImageTk
                base_img = Image.open(png_path).convert("RGBA")
                resolutions = [16, 20, 24, 32, 48, 64, 128, 256]
                photo_images = []
                for res in resolutions:
                    resized = base_img.resize((res, res), Image.Resampling.LANCZOS)
                    photo_images.append(ImageTk.PhotoImage(resized))

                # default=False applies to this window without overriding child window defaults
                window.iconphoto(False, *photo_images)
                window._icon_photo_refs = photo_images
            except Exception:
                pass

        # 2. Standard Tkinter / CustomTkinter iconbitmap
        if ico_path and os.path.exists(ico_path):
            try:
                window.iconbitmap(ico_path)
            except Exception:
                pass

        # 3. Direct Win32 WM_SETICON with exact monitor DPI metrics
        if sys.platform == "win32" and ico_path and os.path.exists(ico_path):
            def _apply_native_win32_icon():
                try:
                    if not hasattr(window, "winfo_exists") or not window.winfo_exists():
                        return
                    import ctypes
                    user32 = ctypes.windll.user32
                    WM_SETICON = 0x0080
                    ICON_SMALL = 0
                    ICON_BIG = 1
                    LR_LOADFROMFILE = 0x0010
                    IMAGE_ICON = 1

                    # Get exact DPI-scaled metrics for the current monitor
                    cx_sm = user32.GetSystemMetrics(49)  # SM_CXSMICON (Titlebar icon width)
                    cy_sm = user32.GetSystemMetrics(50)  # SM_CYSMICON (Titlebar icon height)
                    cx_lg = user32.GetSystemMetrics(11)  # SM_CXICON (Taskbar icon width)
                    cy_lg = user32.GetSystemMetrics(12)  # SM_CYICON (Taskbar icon height)

                    abs_ico = os.path.abspath(ico_path)
                    h_sm = user32.LoadImageW(None, abs_ico, IMAGE_ICON, cx_sm, cy_sm, LR_LOADFROMFILE)
                    h_lg = user32.LoadImageW(None, abs_ico, IMAGE_ICON, cx_lg, cy_lg, LR_LOADFROMFILE)

                    child_hwnd = window.winfo_id()
                    parent_hwnd = user32.GetParent(child_hwnd)
                    hwnd = parent_hwnd if parent_hwnd else child_hwnd

                    if h_sm:
                        user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, h_sm)
                    if h_lg:
                        user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, h_lg)
                except Exception:
                    pass

            _apply_native_win32_icon()
            try:
                window.after(50, _apply_native_win32_icon)
                window.after(250, _apply_native_win32_icon)
            except Exception:
                pass

    @classmethod
    def ensure_icons(cls):
        """Generates all PNG icons to disk if not already present."""
        if cls._icons_ensured:
            return
        os.makedirs(cls.ICONS_DIR, exist_ok=True)
        test_file = os.path.join(cls.ICONS_DIR, "dashboard_normal.png")
        gallery_file = os.path.join(cls.ICONS_DIR, "gallery_normal.png")
        if os.path.exists(test_file) and os.path.exists(gallery_file):
            cls._icons_ensured = True
            return

        for state, col in cls.COLORS.items():
            w = 8

            # 1. Dashboard (Modern 4-tile layout)
            im = Image.new("RGBA", (128, 128), (0, 0, 0, 0))
            d = ImageDraw.Draw(im)
            d.rounded_rectangle([18, 18, 54, 70], radius=10, outline=col, width=w)
            d.rounded_rectangle([74, 18, 110, 48], radius=10, outline=col, width=w)
            d.rounded_rectangle([18, 84, 54, 110], radius=10, outline=col, width=w)
            d.rounded_rectangle([74, 62, 110, 110], radius=10, outline=col, width=w)
            cls._save_icon(im, f"dashboard_{state}.png")

            # 2. POS / Receipt (Vertical invoice with items)
            im = Image.new("RGBA", (128, 128), (0, 0, 0, 0))
            d = ImageDraw.Draw(im)
            pts = [
                (28, 18), (100, 18), (100, 102),
                (88, 110), (76, 102), (64, 110), (52, 102), (40, 110), (28, 102),
                (28, 18)
            ]
            d.line(pts, fill=col, width=w, joint="round")
            d.line([(44, 38), (84, 38)], fill=col, width=w)
            d.line([(44, 58), (84, 58)], fill=col, width=w)
            d.line([(44, 78), (68, 78)], fill=col, width=w)
            cls._save_icon(im, f"pos_{state}.png")

            # 3. Batch POS / Classroom Stack (Dual-sheet layered receipts)
            im = Image.new("RGBA", (128, 128), (0, 0, 0, 0))
            d = ImageDraw.Draw(im)
            d.rounded_rectangle([42, 16, 108, 90], radius=10, outline=col, width=w)
            d.rounded_rectangle([20, 36, 86, 110], radius=10, outline=col, width=w)
            d.line([(36, 56), (70, 56)], fill=col, width=w)
            d.line([(36, 74), (70, 74)], fill=col, width=w)
            d.line([(36, 92), (54, 92)], fill=col, width=w)
            cls._save_icon(im, f"batch_pos_{state}.png")

            # 4. Course Registration (Clipboard with checkmarks - อศท.28)
            im = Image.new("RGBA", (128, 128), (0, 0, 0, 0))
            d = ImageDraw.Draw(im)
            d.rounded_rectangle([24, 28, 104, 112], radius=10, outline=col, width=w)
            d.rounded_rectangle([46, 14, 82, 34], radius=6, outline=col, width=w)
            d.line([(36, 54), (44, 62), (56, 48)], fill=col, width=w, joint="round")
            d.line([(66, 54), (88, 54)], fill=col, width=w)
            d.line([(36, 76), (44, 84), (56, 70)], fill=col, width=w, joint="round")
            d.line([(66, 76), (88, 76)], fill=col, width=w)
            d.line([(38, 96), (88, 96)], fill=col, width=w)
            cls._save_icon(im, f"registration_{state}.png")

            # 5. Students Database (Community / Users silhouettes)
            im = Image.new("RGBA", (128, 128), (0, 0, 0, 0))
            d = ImageDraw.Draw(im)
            d.ellipse([34, 22, 68, 56], outline=col, width=w)
            d.arc([20, 68, 82, 130], start=180, end=0, fill=col, width=w)
            d.ellipse([72, 18, 98, 44], outline=col, width=w)
            d.arc([62, 56, 108, 102], start=210, end=350, fill=col, width=w)
            cls._save_icon(im, f"students_{state}.png")

            # 6. Receipts Archive (Ledger Document + Clock history badge)
            im = Image.new("RGBA", (128, 128), (0, 0, 0, 0))
            d = ImageDraw.Draw(im)
            d.rounded_rectangle([20, 18, 82, 108], radius=8, outline=col, width=w)
            d.line([(36, 38), (66, 38)], fill=col, width=w)
            d.line([(36, 56), (56, 56)], fill=col, width=w)
            d.ellipse([58, 58, 110, 110], outline=col, width=w, fill=(15, 23, 42))
            d.line([(84, 84), (84, 70)], fill=col, width=w)
            d.line([(84, 84), (96, 84)], fill=col, width=w)
            cls._save_icon(im, f"receipts_{state}.png")

            # 6b. Receipt Image Gallery (Framed photo + landscape peak)
            im = Image.new("RGBA", (128, 128), (0, 0, 0, 0))
            d = ImageDraw.Draw(im)
            d.rounded_rectangle([18, 22, 110, 106], radius=10, outline=col, width=w)
            d.ellipse([34, 38, 52, 56], outline=col, width=w - 1)
            d.line([(26, 94), (54, 66), (74, 84), (90, 62), (104, 94)], fill=col, width=w, joint="round")
            cls._save_icon(im, f"gallery_{state}.png")

            # 7. Settings (Modern high-precision 8-tooth gear)
            im = Image.new("RGBA", (128, 128), (0, 0, 0, 0))
            d = ImageDraw.Draw(im)
            cx, cy, r_out, r_in = 64, 64, 46, 20
            for i in range(8):
                angle = i * (math.pi / 4)
                tx = cx + 46 * math.cos(angle)
                ty = cy + 46 * math.sin(angle)
                d.line([(cx, cy), (tx, ty)], fill=col, width=16)
            d.ellipse([cx - r_out, cy - r_out, cx + r_out, cy + r_out], fill=col)
            d.ellipse([cx - 30, cy - 30, cx + 30, cy + 30], fill=(0, 0, 0, 0))
            d.ellipse([cx - r_in, cy - r_in, cx + r_in, cy + r_in], outline=col, width=w)
            cls._save_icon(im, f"settings_{state}.png")

            # 8. Exit / Logout (Door frame + outgoing arrow)
            im = Image.new("RGBA", (128, 128), (0, 0, 0, 0))
            d = ImageDraw.Draw(im)
            d.line([(62, 22), (26, 22), (26, 106), (62, 106)], fill=col, width=w, joint="round")
            d.line([(48, 64), (102, 64)], fill=col, width=w)
            d.line([(84, 46), (102, 64), (84, 82)], fill=col, width=w, joint="round")
            cls._save_icon(im, f"exit_{state}.png")

            # 9. Brand / Institution (Academic college building with pediment & pillars)
            im = Image.new("RGBA", (128, 128), (0, 0, 0, 0))
            d = ImageDraw.Draw(im)
            d.line([(64, 18), (18, 46), (110, 46), (64, 18)], fill=col, width=w, joint="round")
            d.line([(16, 52), (112, 52)], fill=col, width=w)
            d.line([(32, 56), (32, 94)], fill=col, width=w)
            d.line([(53, 56), (53, 94)], fill=col, width=w)
            d.line([(75, 56), (75, 94)], fill=col, width=w)
            d.line([(96, 56), (96, 94)], fill=col, width=w)
            d.line([(14, 100), (114, 100)], fill=col, width=w)
            if os.path.exists(cls.APP_ICON_PNG):
                brand_img = Image.open(cls.APP_ICON_PNG)
                cls._save_icon(brand_img, f"brand_{state}.png")
            else:
                cls._save_icon(im, f"brand_{state}.png")

            # 10. Chevrons for sidebar toggle
            im_left = Image.new("RGBA", (128, 128), (0, 0, 0, 0))
            d_left = ImageDraw.Draw(im_left)
            d_left.line([(76, 32), (44, 64), (76, 96)], fill=col, width=w + 2, joint="round")
            cls._save_icon(im_left, f"chevron_left_{state}.png")

            im_right = Image.new("RGBA", (128, 128), (0, 0, 0, 0))
            d_right = ImageDraw.Draw(im_right)
            d_right.line([(52, 32), (84, 64), (52, 96)], fill=col, width=w + 2, joint="round")
            cls._save_icon(im_right, f"chevron_right_{state}.png")

        cls._icons_ensured = True

    @classmethod
    def _save_icon(cls, image: Image.Image, filename: str):
        # Save at high resolution (128x128) so CTkImage renders crisp on high-DPI screens
        target_size = (128, 128) if (image.width >= 128 and image.height >= 128) else (image.width, image.height)
        if image.size != target_size:
            resized = image.resize(target_size, Image.Resampling.LANCZOS)
        else:
            resized = image
        out_path = os.path.join(cls.ICONS_DIR, filename)
        resized.save(out_path, format="PNG")

    @classmethod
    def get_icon(cls, name: str, state: str = "normal", size: Tuple[int, int] = (20, 20)) -> ctk.CTkImage:
        """
        Retrieves a cached CTkImage for the specified icon, state ('normal', 'active', 'accent'), and size.
        """
        cls.ensure_icons()
        cache_key = (name, state, size[0], size[1])
        if cache_key in cls._cache:
            return cls._cache[cache_key]

        filename = f"{name}_{state}.png"
        file_path = os.path.join(cls.ICONS_DIR, filename)
        if not os.path.exists(file_path):
            file_path = os.path.join(cls.ICONS_DIR, f"{name}_normal.png")

        if os.path.exists(file_path):
            pil_img = Image.open(file_path)
        else:
            # Fallback blank transparent image
            pil_img = Image.new("RGBA", (48, 48), (0, 0, 0, 0))

        ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=size)
        cls._cache[cache_key] = ctk_img
        return ctk_img
