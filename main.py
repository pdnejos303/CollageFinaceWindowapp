import os
import sys

from core.config import get_base_dir, load_config
from core.database import DatabaseManager
from core.models import UserSession
from ui.main_window import MainWindow


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
    app = CollegeFinanceGodApp()
    app.run()


if __name__ == "__main__":
    main()
