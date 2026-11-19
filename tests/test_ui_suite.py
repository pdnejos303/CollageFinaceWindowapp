"""
Comprehensive Automated Test Suite for SPVC College Finance GodApp UI Subsystem.
Tests all UI components, views, dialogs, forms, events, callbacks, and edge cases.
Uses a stable shared Tk root and mocked modal waits for reliable, fast CI execution.
"""

import os
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

# Ensure proper TCL/TK path on Windows
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

from core.database import DatabaseManager
from core.models import UserSession
from ui.theme import AppTheme
from ui.icon_manager import IconManager
from ui.notifications import (
    ModernDialog,
    ModernToast,
    ModernMessageBox,
    showinfo,
    showwarning,
    showerror,
    askyesno,
    askokcancel,
    askquestion,
    askretrycancel,
    show_toast,
    highlight_missing_field,
    patch_tkinter_messagebox,
)
from ui.components import (
    SearchableDropdown,
    StatCard,
    SectionHeader,
    DataTable,
    open_document_file,
)
from ui.login_window import LoginWindow
from ui.dashboard_view import DashboardView
from ui.batch_class_view import BatchClassReceiptView
from ui.views import (
    SmartReceiptPOSView,
    StudentAndFeeCatalogView,
    ReceiptRegisterView,
    SettingsAndBackupView,
    CourseRegistrationView,
)
from ui.main_window import MainWindow


_MODULE_ROOT = None

def get_module_root():
    global _MODULE_ROOT
    if _MODULE_ROOT is None or not _MODULE_ROOT.winfo_exists():
        _MODULE_ROOT = ctk.CTk()
        _MODULE_ROOT.withdraw()
    return _MODULE_ROOT


class BaseUITestCase(unittest.TestCase):
    """Shared base TestCase that uses a single persistent module-level Tk root."""

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


class TestAppTheme(unittest.TestCase):
    """Audits the unified corporate financial design system palette and utilities."""

    def test_theme_constants_and_status_formatting(self):
        self.assertEqual(AppTheme.RADIUS, 0)
        self.assertEqual(AppTheme.FONT_FAMILY, "Tahoma")
        self.assertTrue(AppTheme.COLOR_SIDEBAR.startswith("#"))
        self.assertTrue(AppTheme.COLOR_PRIMARY_BLUE.startswith("#"))

        # Test Thai status map
        self.assertEqual(AppTheme.format_status("PAID"), "✔ ชำระครบ")
        self.assertEqual(AppTheme.format_status("PARTIAL"), "⏳ ค้างบางส่วน")
        self.assertEqual(AppTheme.format_status("UNPAID"), "✖ ค้างชำระ")
        self.assertEqual(AppTheme.format_status("ISSUED"), "✔ ปกติ")
        self.assertEqual(AppTheme.format_status("VOIDED"), "🚫 ยกเลิกแล้ว")
        self.assertEqual(AppTheme.format_status("UNKNOWN_XYZ"), "UNKNOWN_XYZ")

    def test_apply_sharp_theme(self):
        AppTheme.apply_sharp_theme()


class TestIconManager(BaseUITestCase):
    """Audits vector PNG icons, cache mechanism, and DPI window icon setters."""

    def test_icon_generation_and_caching(self):
        IconManager.ensure_icons()
        self.assertTrue(os.path.exists(IconManager.ICONS_DIR))

        # Retrieve various icon keys and states
        for key in ["dashboard", "pos", "batch_pos", "registration", "students", "receipts", "settings", "exit", "brand"]:
            img_normal = IconManager.get_icon(key, state="normal", size=(20, 20))
            self.assertIsNotNone(img_normal)
            img_active = IconManager.get_icon(key, state="active", size=(20, 20))
            self.assertIsNotNone(img_active)

        # Test cache hit
        dash_icon = IconManager.get_icon("dashboard", state="normal", size=(20, 20))
        cached = IconManager.get_icon("dashboard", state="normal", size=(20, 20))
        self.assertIs(cached, dash_icon)

    def test_set_window_icon_safely(self):
        if not self.root:
            self.skipTest("Tkinter display not available")
        # Should execute safely without uncaught exception
        IconManager.set_window_icon(self.root)


class TestModernNotifications(BaseUITestCase):
    """Audits modern modal dialogs, toasts, field highlighting, and drop-in messagebox functions."""

    def test_badge_icons_exist(self):
        for kind in ["success", "warning", "error", "question", "info"]:
            icon_file = os.path.join(ModernDialog.ICONS_DIR, f"{kind}.png")
            self.assertTrue(os.path.exists(icon_file), f"Badge icon missing: {kind}")

    def test_modern_dialog_instantiation(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        with patch.object(ModernDialog, "wait_window"):
            dlg = ModernDialog(
                title="สำเร็จ",
                message="บันทึกข้อมูลเรียบร้อยแล้ว",
                kind="success",
                parent=self.root,
            )
            self.assertEqual(dlg.kind, "success")
            self.assertEqual(dlg.result, None)
            dlg._on_confirm_key()
            self.assertEqual(dlg.result, "ok")

    def test_modern_dialog_askyesno(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        # Confirm action
        with patch.object(ModernDialog, "wait_window"):
            dlg = ModernDialog(
                title="ยืนยัน",
                message="ต้องการลบข้อมูลหรือไม่?",
                kind="question",
                parent=self.root,
                buttons=[("ใช่ / ยืนยัน", True, "primary"), ("ไม่ / ยกเลิก", False, "secondary")],
                default_return=False,
            )
            dlg._on_confirm_key()
            self.assertTrue(dlg.result)

            # Cancel / Escape action
            dlg2 = ModernDialog(
                title="ยืนยัน",
                message="ต้องการลบข้อมูลหรือไม่?",
                kind="question",
                parent=self.root,
                buttons=[("ใช่ / ยืนยัน", True, "primary"), ("ไม่ / ยกเลิก", False, "secondary")],
                default_return=False,
            )
            dlg2._on_close()
            self.assertFalse(dlg2.result)

    def test_modern_dialog_convenience_functions(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        with patch.object(ModernDialog, "wait_window"):
            self.assertEqual(showinfo("ข้อมูล", "ข้อความ", parent=self.root), "ok")
            self.assertEqual(showwarning("แจ้งเตือน", "คำเตือน", parent=self.root), "ok")
            self.assertEqual(showerror("ผิดพลาด", "ข้อผิดพลาด", parent=self.root), "ok")

    def test_modern_toast(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        toast = ModernToast(parent=self.root, title="แจ้งเตือน", message="ข้อความทดสอบ", kind="success", duration_ms=100)
        self.assertIsNotNone(toast)
        toast._safe_destroy()

    def test_highlight_missing_field_and_auto_reset(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        ent = ctk.CTkEntry(self.root)
        ent.pack()
        self.root.update()
        orig_border = ent.cget("border_color")

        highlight_missing_field(ent, "กรุณากรอกข้อมูล", show_toast_msg=False)
        self.assertEqual(ent.cget("border_color"), "#EF4444")

        # Test programmatic reset
        ent._reset_missing_highlight()
        self.assertEqual(ent.cget("border_color"), orig_border)

    def test_patch_tkinter_messagebox(self):
        patch_tkinter_messagebox()
        import tkinter.messagebox as mb
        self.assertEqual(mb.showinfo, showinfo)
        self.assertEqual(mb.showwarning, showwarning)
        self.assertEqual(mb.showerror, showerror)
        self.assertEqual(mb.askyesno, askyesno)


class TestUIComponents(BaseUITestCase):
    """Audits reusable widgets: SearchableDropdown, StatCard, SectionHeader, DataTable, open_document_file."""

    def test_stat_card_updates(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        card = StatCard(self.root, title="ยอดรวม", value="฿0.00", subtitle="0 รายการ")
        card.pack()
        self.assertEqual(card.lbl_value.cget("text"), "฿0.00")

        card.set_value("฿1,500.00", subtitle="5 รายการ")
        self.assertEqual(card.lbl_value.cget("text"), "฿1,500.00")
        self.assertEqual(card.lbl_subtitle.cget("text"), "5 รายการ")

    def test_section_header_and_action_buttons(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        hdr = SectionHeader(self.root, title="หัวข้อหลัก", subtitle="คำอธิบาย")
        hdr.pack()

        clicked = []
        btn = hdr.add_action_button("ปุ่มกระทำ", command=lambda: clicked.append(True))
        self.assertIsNotNone(btn)
        btn.invoke()
        self.assertEqual(len(clicked), 1)

    def test_searchable_dropdown_filtering_and_selection(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        var = ctk.StringVar(value="ตัวเลือก 1")
        items = ["ตัวเลือก 1", "ตัวเลือก 2", "ตัวเลือก 3", "ห้อง ปวช.1/1"]
        called_val = []

        dd = SearchableDropdown(
            self.root,
            values=items,
            variable=var,
            command=lambda v: called_val.append(v),
        )
        dd.pack()
        self.root.update()

        # Check initial value
        self.assertEqual(dd.get(), "ตัวเลือก 1")

        # Test set programmatic
        dd.set("ตัวเลือก 2")
        self.assertEqual(dd.get(), "ตัวเลือก 2")
        self.assertEqual(var.get(), "ตัวเลือก 2")

        # Test configure values
        dd.configure(values=["A", "B", "C"])
        self.assertEqual(dd.cget("values"), ["A", "B", "C"])

        # Test live search filter execution
        dd._set_entry_text("B")
        dd._perform_live_filter()
        self.assertIn("B", dd._filtered_values)
        dd._close_popup()

    def test_data_table_operations_and_pagination(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        cols = [("id", "ID", 60, "center"), ("name", "ชื่อ", 150, "w"), ("amt", "ยอด", 90, "e")]
        table = DataTable(self.root, columns=cols, page_size=10)
        table.pack()

        test_rows = [([str(i), f"ชื่อ {i}", f"฿{i * 100:.2f}"], "PAID") for i in range(1, 16)]
        table.set_rows(test_rows)

        # Check row rendering on page 1 (size=10, total=15)
        children = table.tree.get_children()
        self.assertEqual(len(children), 10)
        self.assertEqual(table.total_count, 15)
        self.assertEqual(table.total_pages, 2)
        self.assertEqual(table.current_page, 1)

        # Test Next Page
        table._go_next_page()
        self.assertEqual(table.current_page, 2)
        children_p2 = table.tree.get_children()
        self.assertEqual(len(children_p2), 5)

        # Test Prev Page
        table._go_prev_page()
        self.assertEqual(table.current_page, 1)

        # Test Last & First Page
        table._go_last_page()
        self.assertEqual(table.current_page, 2)
        table._go_first_page()
        self.assertEqual(table.current_page, 1)

        # Test Selection
        cur_children = table.tree.get_children()
        table.tree.selection_set(cur_children[0])
        val = table.get_selected_values()
        self.assertEqual(val[0], "1")
        self.assertEqual(table.get_selected_row(), val)  # Alias test
        self.assertEqual(table.get_selected_count(), 1)

        # Test Select All (Ctrl+A)
        table._on_select_all()
        self.assertEqual(table.get_selected_count(), 10)

    def test_open_document_file(self):
        # 1. Missing file -> returns False
        with patch("tkinter.messagebox.showerror"):
            res = open_document_file("non_existent_file_xyz_123.pdf")
            self.assertFalse(res)

        # 2. Existing file -> mocked startfile
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(b"%PDF-dummy")
            tmp_path = tmp.name

        try:
            with patch("os.startfile", create=True) as mock_startfile:
                res = open_document_file(tmp_path)
                self.assertTrue(res)
                mock_startfile.assert_called_once()
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)


class TestLoginWindow(unittest.TestCase):
    """Audits the authentication window configuration and session handling."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, "test_login.db")
        self.db = DatabaseManager(self.db_path)

    def test_login_window_init(self):
        login = LoginWindow(self.db, "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")
        self.assertIsNotNone(login.db)
        self.assertEqual(login.college_name, "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")
        self.assertIsNone(login.session)


class TestDashboardView(BaseUITestCase):
    """Audits financial metrics, filter dropdowns, and Excel export in DashboardView."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, "test_dashboard.db")
        self.db = DatabaseManager(self.db_path)
        self.session = UserSession(1, "admin", "ผู้ดูแลระบบ", "ADMIN", "การเงิน")
        self.config = {"college_name": "วิทยาลัยอาชีวศึกษาสุพรรณบุรี", "current_semester": "1/2569"}

    def test_dashboard_view_filters_and_refresh(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        nav_calls = []
        view = DashboardView(
            self.root,
            db=self.db,
            session=self.session,
            config=self.config,
            on_navigate=lambda k: nav_calls.append(k),
        )
        view.pack()

        # Check KPI cards exist and populated
        self.assertIsNotNone(view.card_today_income)
        self.assertIsNotNone(view.card_total_income)
        self.assertIsNotNone(view.card_outstanding)

        # Test filter selection
        view.selected_level_var.set("ปวช.1")
        view.refresh_dashboard()
        self.assertIn("ปวช.1", view.lbl_filter_scope.cget("text"))

        # Test reset filters
        view.reset_filters()
        self.assertEqual(view.selected_level_var.get(), "ทั้งหมด")

        # Test navigation shortcut
        view._navigate("pos")
        self.assertEqual(nav_calls, ["pos"])

    def test_dashboard_export_excel_flow(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        view = DashboardView(self.root, db=self.db, session=self.session, config=self.config)
        view.pack()

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            save_path = tmp.name

        try:
            with patch("tkinter.filedialog.asksaveasfilename", return_value=save_path), \
                 patch("tkinter.messagebox.showinfo"), \
                 patch("ui.dashboard_view.open_document_file", return_value=True):
                view.export_excel()
                self.assertTrue(os.path.exists(save_path))
                self.assertGreater(os.path.getsize(save_path), 0)
        finally:
            if os.path.exists(save_path):
                os.remove(save_path)


class TestBatchClassReceiptView(BaseUITestCase):
    """Audits classroom batch studio, uniform fee additions, removal with get_selected_values, and issuing."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, "test_batch.db")
        self.db = DatabaseManager(self.db_path)
        self.session = UserSession(1, "finance", "คุณกนกวรรณ", "FINANCE", "การเงิน")
        self.config = {
            "college_name": "วิทยาลัยอาชีวศึกษาสุพรรณบุรี",
            "receipt_book_no": "7006",
            "current_semester": "1/2569",
            "receipt_no_start": "700605400",
            "receipt_no_end": "700605500",
            "default_export_folder": self.temp_dir,
        }

    def test_batch_fee_management_and_removal(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        view = BatchClassReceiptView(self.root, db=self.db, session=self.session, config=self.config)
        view.pack()

        # 1. Preset curriculum fee load
        view.curriculum_var.set("หลักสูตรในระบบ-ปวช")
        view.load_preset_curriculum_fees()
        self.assertGreater(len(view.classroom_fees), 0)
        initial_count = len(view.classroom_fees)

        # 2. Add custom item
        view.fee_item_var.set("ค่าบำรุงห้องปฏิบัติการพิเศษ")
        view.unit_price_var.set("250.00")
        view.qty_var.set("1")
        view.add_fee_item()
        self.assertEqual(len(view.classroom_fees), initial_count + 1)

        # 3. Test removal using table selection (verifying get_selected_values / get_selected_row fix!)
        children = view.tbl_fees.tree.get_children()
        self.assertTrue(len(children) > 0)
        view.tbl_fees.tree.selection_set(children[-1])
        view.remove_selected_fee()
        self.assertEqual(len(view.classroom_fees), initial_count)

        # 4. Clear all fees
        view.clear_all_fees()
        self.assertEqual(len(view.classroom_fees), 0)

    def test_batch_student_roster_loading(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        view = BatchClassReceiptView(self.root, db=self.db, session=self.session, config=self.config)
        view.pack()

        # Load from system room
        view.system_room_var.set("ปวช.1/1")
        with patch("tkinter.messagebox.showwarning"):
            view.load_students_from_system()

        # Clear students
        view.clear_students()
        self.assertEqual(len(view.classroom_students), 0)
        self.assertEqual(view.lbl_student_count.cget("text"), "👥 พบ 0 คน")


class TestSmartReceiptPOSView(BaseUITestCase):
    """Audits the main POS view: cart calculations, student lookup, adding/removing line items, reset."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, "test_pos.db")
        self.db = DatabaseManager(self.db_path)
        self.session = UserSession(1, "admin", "ผู้ดูแลระบบ", "ADMIN", "การเงิน")
        self.config = {
            "college_name": "วิทยาลัยอาชีวศึกษาสุพรรณบุรี",
            "receipt_book_no": "7006",
            "current_semester": "1/2569",
            "receipt_prefix": "",
            "receipt_no_start": "700605433",
            "receipt_no_end": "700999999",
            "default_export_folder": self.temp_dir,
        }

    def test_pos_cart_workflow(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        webhook_mock = MagicMock()
        view = SmartReceiptPOSView(
            self.root,
            db=self.db,
            session=self.session,
            config=self.config,
            webhook=webhook_mock,
        )
        view.pack()

        # 1. Lookup student
        view.entry_stu_code.insert(0, "692090101")
        view._lookup_student_by_code_entry()
        self.assertTrue(len(view.entry_stu_name.get()) > 0)

        # 2. Add Fee to Cart
        view.curriculum_var.set("หลักสูตรในระบบ-ปวช")
        view.on_change_curriculum("หลักสูตรในระบบ-ปวช")
        cat_items = view.db.get_fee_catalog("หลักสูตรในระบบ-ปวช")
        if cat_items:
            first_name = cat_items[0]["fee_name"]
            view.fee_item_var.set(first_name)
            view.on_select_fee_item(first_name)
            view.add_selected_fee_to_cart()

            self.assertEqual(len(view.cart_items), 1)
            tot = sum(float(i['line_total']) for i in view.cart_items)
            self.assertGreater(tot, 0.0)

            # 3. Remove item from cart
            children = view.table.tree.get_children()
            if children:
                view.table.tree.selection_set(children[0])
                view.remove_selected_item()
                self.assertEqual(len(view.cart_items), 0)

        # 4. Reset form
        view.reset_pos_form()
        self.assertEqual(len(view.cart_items), 0)
        self.assertEqual(view.entry_stu_code.get(), "")

    def test_issue_receipt_and_export_sanitizes_name_and_saves(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        view = SmartReceiptPOSView(
            self.root,
            db=self.db,
            session=self.session,
            config=self.config,
            webhook=MagicMock(),
        )
        view.pack()

        view.entry_stu_name.insert(0, "นางสาว ทดสอบ / ระบบ")
        view.curriculum_var.set("หลักสูตรในระบบ-ปวช")
        view.on_change_curriculum("หลักสูตรในระบบ-ปวช")
        cat_items = view.db.get_fee_catalog("หลักสูตรในระบบ-ปวช")
        self.assertTrue(len(cat_items) > 0)
        first_name = cat_items[0]["fee_name"]
        view.fee_item_var.set(first_name)
        view.on_select_fee_item(first_name)
        view.add_selected_fee_to_cart()
        self.assertEqual(len(view.cart_items), 1)

        out_pdf = os.path.join(self.temp_dir, "receipt_test.pdf")
        with (
            patch("ui.views.filedialog.asksaveasfilename", return_value=out_pdf) as mock_save_dlg,
            patch("ui.views.save_config"),
            patch("ui.views.open_document_file", return_value=True),
            patch("ui.views.messagebox.showinfo") as mock_info,
            patch("ui.views.messagebox.showerror") as mock_err,
        ):
            view.issue_receipt_and_export()

        mock_err.assert_not_called()
        mock_info.assert_called_once()
        mock_save_dlg.assert_called_once()
        initial_file = mock_save_dlg.call_args.kwargs.get("initialfile", "")
        self.assertNotIn("/", initial_file)
        self.assertNotIn(" ", initial_file)
        self.assertTrue(os.path.exists(out_pdf))
        self.assertEqual(len(view.cart_items), 0)

    def test_save_receipt_only_and_export_prompt(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        view = SmartReceiptPOSView(
            self.root,
            db=self.db,
            session=self.session,
            config=self.config,
            webhook=MagicMock(),
        )
        view.pack()

        view.entry_stu_name.insert(0, "นาย สมชาย / ใจดี")
        view.curriculum_var.set("หลักสูตรในระบบ-ปวช")
        view.on_change_curriculum("หลักสูตรในระบบ-ปวช")
        cat_items = view.db.get_fee_catalog("หลักสูตรในระบบ-ปวช")
        self.assertTrue(len(cat_items) > 0)
        first_name = cat_items[0]["fee_name"]
        view.fee_item_var.set(first_name)
        view.on_select_fee_item(first_name)
        view.add_selected_fee_to_cart()

        out_pdf = os.path.join(self.temp_dir, "saved_only_test.pdf")
        with (
            patch("ui.views.messagebox.askyesno", return_value=True),
            patch("ui.views.filedialog.asksaveasfilename", return_value=out_pdf) as mock_save_dlg,
            patch("ui.views.save_config"),
            patch("ui.views.open_document_file", return_value=True),
            patch("ui.views.messagebox.showinfo"),
            patch("ui.views.messagebox.showerror") as mock_err,
        ):
            view.save_receipt_only()

        mock_err.assert_not_called()
        mock_save_dlg.assert_called_once()
        initial_file = mock_save_dlg.call_args.kwargs.get("initialfile", "")
        self.assertNotIn("/", initial_file)
        self.assertNotIn(" ", initial_file)
        self.assertTrue(os.path.exists(out_pdf))
        self.assertEqual(len(view.cart_items), 0)


class TestStudentAndFeeCatalogView(BaseUITestCase):
    """Audits StudentAndFeeCatalogView regex class-year extraction."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, "test_students.db")
        self.db = DatabaseManager(self.db_path)
        self.session = UserSession(1, "admin", "ผู้ดูแลระบบ", "ADMIN", "การเงิน")
        self.config = {"college_name": "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"}

    def test_on_form_cls_change_extracts_year_with_regex(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        view = StudentAndFeeCatalogView(self.root, db=self.db, session=self.session, config=self.config)
        view.pack()
        view._on_form_cls_change("ปวช.2/1")
        self.assertEqual(view.form_yr_var.get(), "2")


class TestReceiptRegisterView(BaseUITestCase):
    """Audits the receipt history register: chips filtering, table selection, reprint, void."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, "test_register.db")
        self.db = DatabaseManager(self.db_path)
        self.session = UserSession(1, "admin", "ผู้ดูแลระบบ", "ADMIN", "การเงิน")
        self.config = {"college_name": "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"}

    def test_receipt_register_chips_and_selection(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        view = ReceiptRegisterView(self.root, db=self.db, session=self.session, config=self.config)
        view.pack()

        # Preset chips
        view._on_preset_click("TODAY")
        self.assertEqual(view.period_mode_var.get(), "วันนี้")

        view._on_preset_click("ALL")
        self.assertEqual(view.period_mode_var.get(), "ทั้งหมด (ทุกวัน)")

        # Clear selection state
        view.table.tree.selection_set([])
        view.on_select_row()
        self.assertEqual(view.lbl_selection.cget("text"), "")

    def test_audit_log_filtering_and_subtab_rendering(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        self.db.add_option("CLASS_LEVEL", "ปวช.9/9")
        action_types = self.db.get_audit_action_types()
        self.assertIn("OPTION_ADD", action_types)

        filtered = self.db.get_audit_logs(search="ปวช.9/9", action_filter="OPTION_ADD")
        self.assertGreaterEqual(len(filtered), 1)
        self.assertEqual(filtered[0]["action"], "OPTION_ADD")

        today_str = filtered[0]["timestamp"][:10]
        date_filtered = self.db.get_audit_logs(date_from=today_str, date_to=today_str)
        self.assertGreaterEqual(len(date_filtered), 1)

        view = ReceiptRegisterView(self.root, db=self.db, session=self.session, config=self.config)
        view.pack()
        self.assertFalse(view._audit_panel_built)

        view.switch_subtab("audit")
        self.assertTrue(view._audit_panel_built)
        self.assertEqual(view.active_subtab, "audit")
        self.assertGreaterEqual(len(view.audit_table.tree.get_children()), 1)

        view.audit_action_var.set("OPTION_ADD")
        view.ent_audit_search.insert(0, "ปวช.9/9")
        view.refresh_audit_logs()
        self.assertEqual(len(view.audit_table.tree.get_children()), len(filtered))

        view.switch_subtab("receipts")
        self.assertEqual(view.active_subtab, "receipts")


class TestSettingsAndBackupView(BaseUITestCase):
    """Audits system configuration, dropdown option CRUD, and database backup."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, "test_settings.db")
        self.db = DatabaseManager(self.db_path)
        self.session = UserSession(1, "admin", "ผู้ดูแลระบบ", "ADMIN", "การเงิน")
        self.config = {
            "college_name": "วิทยาลัยอาชีวศึกษาสุพรรณบุรี",
            "receipt_book_no": "7006",
            "current_semester": "1/2569",
            "show_course_registration": False,
        }

    def test_settings_save_and_toggle_registration(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        updated = []
        with patch("ui.views.save_config", return_value=True):
            view = SettingsAndBackupView(
                self.root,
                db=self.db,
                session=self.session,
                config=self.config,
                on_config_updated=lambda: updated.append(True),
            )
            view.pack()

            # Toggle course registration
            view.show_registration_var.set(True)
            view.on_toggle_course_registration()
            self.assertTrue(self.config["show_course_registration"])
            self.assertEqual(len(updated), 1)

            # Save general config
            with patch("tkinter.messagebox.showinfo"):
                view.ent_college.delete(0, "end")
                view.ent_college.insert(0, "วิทยาลัยอาชีวศึกษาทดสอบ")
                view.save_general_config()
                self.assertEqual(self.config["college_name"], "วิทยาลัยอาชีวศึกษาทดสอบ")

    def test_dropdown_options_management(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        with patch("ui.views.save_config", return_value=True):
            view = SettingsAndBackupView(self.root, db=self.db, session=self.session, config=self.config)
            view.pack()
            view.switch_settings_tab("dropdowns")

            # Add option
            view.ent_new_opt.insert(0, "ห้องทดสอบพิเศษ")
            with patch("tkinter.messagebox.showinfo"):
                view.add_dropdown_option()
            cat_code = view.category_labels[view.opt_cat_var.get()]
            opts = self.db.get_options(cat_code)
            self.assertIn("ห้องทดสอบพิเศษ", opts)


class TestCourseRegistrationView(BaseUITestCase):
    """Audits course registration (อศท.28) catalog, adding courses, credit calculations."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, "test_course.db")
        self.db = DatabaseManager(self.db_path)
        self.session = UserSession(1, "admin", "ผู้ดูแลระบบ", "ADMIN", "การเงิน")
        self.config = {"college_name": "วิทยาลัยอาชีวศึกษาสุพรรณบุรี", "current_semester": "1/2569"}

    def test_course_registration_workflow(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        view = CourseRegistrationView(self.root, db=self.db, session=self.session, config=self.config)
        view.pack()

        # Add course manually
        initial_count = len(view.registered_courses)
        view.ent_c_code.insert(0, "30201-0001")
        view.ent_c_name.insert(0, "การบัญชีการเงิน")
        view.ent_c_t.insert(0, "2")
        view.ent_c_p.insert(0, "2")
        view.ent_c_cr.insert(0, "3")
        view.add_course_to_table()

        self.assertEqual(len(view.registered_courses), initial_count + 1)
        self.assertIn("นก.", view.lbl_totals.cget("text"))

        # Remove course
        children = view.table.tree.get_children()
        if children:
            view.table.tree.selection_set(children[-1])
        view.remove_selected_course()
        self.assertEqual(len(view.registered_courses), initial_count)


class TestMainWindow(BaseUITestCase):
    """Audits the primary single-root application window, sidebar collapsing, and error catcher."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, "test_main.db")
        self.db = DatabaseManager(self.db_path)
        self.session = UserSession(1, "admin", "ผู้ดูแลระบบ", "ADMIN", "การเงิน")
        self.config = {
            "college_name": "วิทยาลัยอาชีวศึกษาสุพรรณบุรี",
            "current_semester": "1/2569",
            "show_course_registration": True,
            "auto_backup_on_exit": False,
        }

    def test_main_window_structure_and_sidebar_toggle(self):
        if not self.root:
            self.skipTest("Tkinter display not available")

        main = MainWindow(self.db, self.session, self.config)
        main.root = self.root
        main.root.report_callback_exception = main._on_tkinter_callback_exception

        # Test tab switching
        main.nav_buttons = {}
        for k, full in main.nav_items_config:
            btn = ctk.CTkButton(self.root, text=full)
            main.nav_buttons[k] = btn

        main.content_area = ctk.CTkFrame(self.root)
        main.content_area.pack()

        main.switch_tab("dashboard")
        self.assertEqual(main.current_tab_key, "dashboard")
        self.assertIsInstance(main.current_view, DashboardView)

        main.switch_tab("pos")
        self.assertEqual(main.current_tab_key, "pos")
        self.assertIsInstance(main.current_view, SmartReceiptPOSView)

        # Test Error Catcher doesn't crash
        with patch("tkinter.messagebox.showerror"):
            try:
                raise ValueError("Test Synthetic System Error")
            except ValueError as ve:
                main._on_tkinter_callback_exception(type(ve), ve, ve.__traceback__)


if __name__ == "__main__":
    unittest.main()
