import os
import sys

if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("spvc.collegefinance.godapp.2")
    except Exception:
        pass

from core.config import get_base_dir, load_config
from core.database import DatabaseManager
from core.models import UserSession
from ui.notifications import patch_tkinter_messagebox
from ui.main_window import MainWindow

# Apply modern dialog theme to all popups
patch_tkinter_messagebox()


class DummyStream:
    def write(self, *args, **kwargs): pass
    def flush(self, *args, **kwargs): pass
    def reconfigure(self, *args, **kwargs): pass


if sys.stdout is None:
    sys.stdout = DummyStream()
    sys.stderr = DummyStream()
else:
    try:
        sys.stdout.reconfigure(line_buffering=True)
    except Exception:
        pass


class CollegeFinanceGodApp:
    """Single-Window Application Controller for SPVC Smart Tuition & Receipt System."""

    def __init__(self):
        self.base_dir = get_base_dir()
        self.config = load_config(self.base_dir)
        self.db_path = os.path.join(self.base_dir, "data", "college_finance.db")
        self.db = DatabaseManager(self.db_path, self.config.get("database_url", ""))

    def run(self):
        default_session = UserSession(
            user_id=1,
            username="admin",
            full_name="นางสาวสุพรรณี ศรีอาชีวะ",
            role="ADMIN",
            department="งานการเงิน วิทยาลัยอาชีวศึกษาสุพรรณบุรี",
        )
        main_win = MainWindow(self.db, default_session, self.config)
        main_win.show()


def main():
    try:
        app = CollegeFinanceGodApp()
        app.run()
    except Exception as e:
        import traceback
        import tkinter.messagebox as mb
        tb = traceback.format_exc()
        try:
            with open(os.path.join(get_base_dir(), "data", "startup_error.log"), "a", encoding="utf-8") as f:
                f.write(f"\nSTARTUP ERROR:\n{tb}\n")
        except Exception:
            pass
        mb.showerror("เกิดข้อผิดพลาดในการเริ่มต้นโปรแกรม", f"ไม่สามารถเปิดโปรแกรมได้:\n{e}\n\n{tb}")


if __name__ == "__main__":
    main()
