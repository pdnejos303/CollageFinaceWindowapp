"""
Automated tests for the Thai interactive onboarding & tutorial system:
- core/config.py (tutorial_completed default & persistence)
- ui/tutorial_overlay.py (TutorialWelcomeDialog, TutorialOverlay, InteractiveTutorialManager)
- ui/main_window.py (first-launch trigger, sidebar button, navigation wiring)
"""

import os
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

tcl_dir = os.path.join(os.path.dirname(sys.executable), "tcl")
if os.path.exists(tcl_dir):
    tcl_lib = os.path.join(tcl_dir, "tcl8.6").replace("\\", "/")
    tk_lib = os.path.join(tcl_dir, "tk8.6").replace("\\", "/")
    if os.path.exists(tcl_lib):
        os.environ["TCL_LIBRARY"] = tcl_lib
    if os.path.exists(tk_lib):
        os.environ["TK_LIBRARY"] = tk_lib

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import customtkinter as ctk

from core.config import DEFAULT_CONFIG, load_config, save_config
from core.database import DatabaseManager
from core.models import UserSession
from ui.main_window import MainWindow
from ui.tutorial_overlay import (
    TUTORIAL_STEPS,
    InteractiveTutorialManager,
    TutorialOverlay,
    TutorialStep,
    TutorialWelcomeDialog,
)

_MODULE_ROOT = None


def get_module_root():
    global _MODULE_ROOT
    if _MODULE_ROOT is None or not _MODULE_ROOT.winfo_exists():
        _MODULE_ROOT = ctk.CTk()
        _MODULE_ROOT.withdraw()
    return _MODULE_ROOT


class BaseTutorialTestCase(unittest.TestCase):
    @property
    def root(self):
        return get_module_root()

    def tearDown(self):
        r = get_module_root()
        for child in list(r.winfo_children()):
            try:
                child.destroy()
            except Exception:
                pass


class TestTutorialConfig(unittest.TestCase):
    def test_default_config_contains_tutorial_completed_false(self):
        self.assertIn("tutorial_completed", DEFAULT_CONFIG)
        self.assertIs(DEFAULT_CONFIG["tutorial_completed"], False)

    def test_save_and_load_tutorial_completed(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            cfg = load_config(tmpdir)
            self.assertFalse(cfg["tutorial_completed"])

            cfg["tutorial_completed"] = True
            self.assertTrue(save_config(cfg, tmpdir))

            reloaded = load_config(tmpdir)
            self.assertTrue(reloaded["tutorial_completed"])


class TestTutorialStepsAndFiltering(unittest.TestCase):
    def test_all_tutorial_steps_match_main_window_nav_keys(self):
        db = MagicMock()
        session = UserSession(1, "staff", "เจ้าหน้าที่", "STAFF", "การเงิน")
        mw = MainWindow(db, session, dict(DEFAULT_CONFIG))
        valid_keys = {k for k, _ in mw.nav_items_config}

        self.assertGreaterEqual(len(TUTORIAL_STEPS), 8)
        for step in TUTORIAL_STEPS:
            self.assertIn(step.page_key, valid_keys)
            self.assertTrue(step.title.strip())
            self.assertTrue(step.description.strip())

    def test_filter_steps_excludes_registration_when_disabled(self):
        steps = InteractiveTutorialManager.filter_steps(
            TUTORIAL_STEPS,
            visible_pages=None,
            user_role="STAFF",
            config={"show_course_registration": False},
        )
        page_keys = {s.page_key for s in steps}
        self.assertNotIn("registration", page_keys)
        self.assertIn("dashboard", page_keys)

    def test_filter_steps_includes_registration_when_enabled(self):
        steps = InteractiveTutorialManager.filter_steps(
            TUTORIAL_STEPS,
            visible_pages=None,
            user_role="STAFF",
            config={"show_course_registration": True},
        )
        page_keys = {s.page_key for s in steps}
        self.assertIn("registration", page_keys)

    def test_filter_steps_respects_visible_pages_and_roles(self):
        custom_steps = [
            TutorialStep("s1", "dashboard", "หน้าหลัก", "1", "desc"),
            TutorialStep("s2", "pos", "POS", "2", "desc", allowed_roles=("ADMIN",)),
            TutorialStep("s3", "settings", "ตั้งค่า", "3", "desc"),
        ]
        staff_steps = InteractiveTutorialManager.filter_steps(
            custom_steps,
            visible_pages=["dashboard", "pos"],
            user_role="STAFF",
        )
        self.assertEqual([s.step_id for s in staff_steps], ["s1"])

        admin_steps = InteractiveTutorialManager.filter_steps(
            custom_steps,
            visible_pages=["dashboard", "pos"],
            user_role="ADMIN",
        )
        self.assertEqual([s.step_id for s in admin_steps], ["s1", "s2"])


class TestTutorialWelcomeDialog(BaseTutorialTestCase):
    def test_welcome_dialog_start_callback(self):
        started = []
        cfg = {"tutorial_completed": False}
        dlg = TutorialWelcomeDialog(
            parent=self.root,
            on_start=lambda: started.append(True),
            config=cfg,
        )
        dlg._handle_start()
        self.assertEqual(dlg.result, "start")
        self.assertEqual(started, [True])
        self.assertFalse(cfg["tutorial_completed"])

    def test_welcome_dialog_skip_persists_when_dont_ask_checked(self):
        skipped = []
        cfg = {"tutorial_completed": False}
        dlg = TutorialWelcomeDialog(
            parent=self.root,
            on_skip=lambda dont_ask: skipped.append(dont_ask),
            config=cfg,
        )
        dlg.dont_ask_var.set(True)
        dlg._handle_skip()
        self.assertEqual(dlg.result, "skip")
        self.assertEqual(skipped, [True])
        self.assertTrue(cfg["tutorial_completed"])

    def test_welcome_dialog_skip_does_not_persist_when_unchecked(self):
        skipped = []
        cfg = {"tutorial_completed": False}
        dlg = TutorialWelcomeDialog(
            parent=self.root,
            on_skip=lambda dont_ask: skipped.append(dont_ask),
            config=cfg,
        )
        dlg.dont_ask_var.set(False)
        dlg._handle_skip()
        self.assertEqual(dlg.result, "skip")
        self.assertEqual(skipped, [False])
        self.assertFalse(cfg["tutorial_completed"])


class TestInteractiveTutorialManager(BaseTutorialTestCase):
    def _sample_steps(self):
        return [
            TutorialStep("step1", "dashboard", "ภาพรวม", "หัวข้อ 1", "คำอธิบาย 1", target_attr="box1"),
            TutorialStep("step2", "dashboard", "ภาพรวม", "หัวข้อ 2", "คำอธิบาย 2"),
            TutorialStep("step3", "pos", "POS", "หัวข้อ 3", "คำอธิบาย 3"),
        ]

    def test_navigation_prev_next_jump_and_complete_persistence(self):
        navigated = []
        completed = []
        cfg = {"tutorial_completed": False}

        box1 = ctk.CTkFrame(self.root, width=120, height=40)
        box1.pack()
        self.root.update_idletasks()

        mgr = InteractiveTutorialManager(
            root=self.root,
            navigate_callback=navigated.append,
            config=cfg,
            on_complete=lambda: completed.append(True),
            widget_resolver=lambda step: box1 if step.target_attr == "box1" else None,
            steps=self._sample_steps(),
        )

        self.assertTrue(mgr.start())
        self.assertTrue(mgr.is_active)
        self.assertIsInstance(mgr.overlay, TutorialOverlay)
        self.assertEqual(mgr.current_index, 0)
        self.assertEqual(navigated, ["dashboard"])

        # Step 2 is also on dashboard -> should not re-trigger navigate_callback
        mgr.next_step()
        self.assertEqual(mgr.current_index, 1)
        self.assertEqual(navigated, ["dashboard"])

        # Step 3 switches to pos
        mgr.next_step()
        self.assertEqual(mgr.current_index, 2)
        self.assertEqual(navigated, ["dashboard", "pos"])

        # Back to step 2 switches back to dashboard
        mgr.prev_step()
        self.assertEqual(mgr.current_index, 1)
        self.assertEqual(navigated, ["dashboard", "pos", "dashboard"])

        # Jump to last step and complete via next_step()
        mgr.go_to_step(2)
        self.assertEqual(mgr.current_index, 2)
        mgr.next_step()

        self.assertFalse(mgr.is_active)
        self.assertIsNone(mgr.overlay)
        self.assertTrue(cfg["tutorial_completed"])
        self.assertEqual(completed, [True])

    def test_skip_persists_completion_and_tears_down_overlay(self):
        skipped = []
        cfg = {"tutorial_completed": False}
        mgr = InteractiveTutorialManager(
            root=self.root,
            config=cfg,
            on_skip=lambda: skipped.append(True),
            steps=self._sample_steps(),
        )
        mgr.start()
        mgr.skip()

        self.assertFalse(mgr.is_active)
        self.assertIsNone(mgr.overlay)
        self.assertTrue(cfg["tutorial_completed"])
        self.assertEqual(skipped, [True])


class TestMainWindowTutorialIntegration(BaseTutorialTestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, "test_tutorial_main.db")
        self.db = DatabaseManager(self.db_path)
        self.session = UserSession(1, "admin", "ผู้ดูแลระบบ", "ADMIN", "การเงิน")
        self.config = dict(DEFAULT_CONFIG)
        self.config["auto_backup_on_exit"] = False
        self.config["tutorial_completed"] = False

    def _build_minimal_main_window(self):
        main = MainWindow(self.db, self.session, self.config)
        main.root = self.root
        main.nav_buttons = {}
        for k, full in main.nav_items_config:
            btn = ctk.CTkButton(self.root, text=full)
            main.nav_buttons[k] = btn
        main.content_area = ctk.CTkFrame(self.root)
        main.content_area.pack(fill="both", expand=True)
        return main

    def test_visible_tutorial_pages_respects_show_course_registration(self):
        main = self._build_minimal_main_window()
        self.config["show_course_registration"] = False
        self.assertNotIn("registration", main.get_visible_tutorial_pages())

        self.config["show_course_registration"] = True
        self.assertIn("registration", main.get_visible_tutorial_pages())

    def test_start_tutorial_navigates_main_window_and_saves_config(self):
        main = self._build_minimal_main_window()
        main.switch_tab("dashboard")

        with patch("ui.main_window.save_config", return_value=True) as mock_save:
            started = main.start_tutorial()
            self.assertTrue(started)
            self.assertIsNotNone(main.tutorial_manager)
            self.assertTrue(main.tutorial_manager.is_active)
            self.assertEqual(main.current_tab_key, "dashboard")

            # Advance until we reach POS page
            while main.tutorial_manager.current_step.page_key == "dashboard":
                main.tutorial_manager.next_step()

            self.assertEqual(main.current_tab_key, "pos")

            # Skip tutorial -> persists tutorial_completed=True
            main.tutorial_manager.skip()
            self.assertTrue(self.config["tutorial_completed"])
            mock_save.assert_called_once_with(self.config)

    def test_maybe_show_tutorial_welcome_only_when_not_completed(self):
        main = self._build_minimal_main_window()

        self.config["tutorial_completed"] = True
        main._maybe_show_tutorial_welcome()
        self.assertIsNone(main._welcome_dialog)

        self.config["tutorial_completed"] = False
        dlg = main.show_tutorial_welcome()
        self.assertIsInstance(dlg, TutorialWelcomeDialog)
        dlg._handle_start()
        self.assertIsNotNone(main.tutorial_manager)
        self.assertTrue(main.tutorial_manager.is_active)
        main.tutorial_manager.stop()

    def test_first_launch_schedules_after_450_and_sidebar_button_toggles(self):
        main = MainWindow(self.db, self.session, self.config)
        scheduled = []

        fake_root = MagicMock()
        fake_root.after.side_effect = lambda ms, fn: scheduled.append((ms, fn))

        with (
            patch("ui.main_window.ctk.set_appearance_mode"),
            patch("ui.main_window.ctk.set_default_color_theme"),
            patch("ui.main_window.ctk.StringVar", return_value=MagicMock()),
            patch("ui.main_window.SearchableDropdown", return_value=MagicMock()),
            patch("ui.main_window.ctk.CTk", return_value=fake_root),
            patch("ui.main_window.ctk.CTkFrame", return_value=MagicMock()),
            patch("ui.main_window.ctk.CTkScrollableFrame", return_value=MagicMock()),
            patch("ui.main_window.ctk.CTkLabel", side_effect=lambda *a, **kw: MagicMock()),
            patch("ui.main_window.ctk.CTkButton", side_effect=lambda *a, **kw: MagicMock()),
            patch("ui.main_window.IconManager.set_window_icon"),
            patch("ui.main_window.IconManager.get_icon", return_value=None),
            patch.object(main, "_build_top_header"),
            patch.object(main, "switch_tab"),
            patch.object(main, "refresh_nav_visibility"),
        ):
            main.show()
            main.toggle_sidebar()
            main.btn_tutorial.configure.assert_called_with(text="❓")
            main.toggle_sidebar()
            main.btn_tutorial.configure.assert_called_with(text="❓ สอนใช้งาน")

        self.assertTrue(
            any(ms == 450 and fn == main._maybe_show_tutorial_welcome for ms, fn in scheduled),
            f"Expected after(450, _maybe_show_tutorial_welcome), got {scheduled}",
        )
        self.assertIsNotNone(main.btn_tutorial)


if __name__ == "__main__":
    unittest.main()
