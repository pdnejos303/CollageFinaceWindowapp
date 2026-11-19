import csv
import datetime
import math
import os
import shutil
import sqlite3
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from core.models import UserSession, StudentRecord, TransactionRecord
from core.security import SecurityManager, ROLE_PERMISSIONS
from core.config import (
    DEFAULT_CONFIG,
    normalize_template_style,
    thai_baht_text,
    format_thai_date_short,
    format_thai_date_padded,
    format_thai_date_display,
    format_receipt_date_thai,
    parse_date_robust,
    split_receipt_prefix_and_number,
    format_next_receipt_number,
    combine_receipt_no,
    load_config,
    save_config,
    backup_database,
    export_full_zip_backup,
    restore_database,
)
from core.database import DatabaseManager, derive_year_level_and_education_level


class TestCoreModels(unittest.TestCase):
    """Tests for core/models.py dataclasses and properties."""

    def test_user_session(self):
        user = UserSession(
            user_id=1,
            username="admin",
            full_name="ผู้ดูแลระบบ",
            role="ADMIN",
            department="งานการเงิน",
        )
        self.assertEqual(user.user_id, 1)
        self.assertEqual(user.username, "admin")
        self.assertEqual(user.role, "ADMIN")

    def test_student_record_balance(self):
        stu = StudentRecord(
            student_code="692020101",
            full_name="สมชาย ใจดี",
            faculty="บริหารธุรกิจ",
            major="การบัญชี",
            year_level=1,
            scholarship_type="ปกติ",
            total_fee=950.0,
            paid_amount=300.0,
            status="PARTIAL",
        )
        self.assertEqual(stu.balance, 650.0)

        # Paid in full
        stu.paid_amount = 950.0
        self.assertEqual(stu.balance, 0.0)

        # Overpaid -> balance should remain 0.0
        stu.paid_amount = 1200.0
        self.assertEqual(stu.balance, 0.0)

        # None values safety
        stu.total_fee = None
        stu.paid_amount = None
        self.assertEqual(stu.balance, 0.0)

    def test_transaction_record(self):
        tx = TransactionRecord(
            tx_code="RC-001",
            tx_date="2026-11-17",
            tx_type="INCOME",
            category="ค่าเทอม",
            title="ชำระค่าธรรมเนียม",
            amount=950.0,
            payment_method="เงินสด",
            ref_code="692020101",
            recorded_by="admin",
        )
        self.assertEqual(tx.status, "COMPLETED")
        self.assertEqual(tx.amount, 950.0)


class TestSecurityManager(unittest.TestCase):
    """Tests for core/security.py PBKDF2 hashing, RBAC, and PDPA masking."""

    def test_password_hash_and_verify(self):
        pwd = "SecretPassword123"
        hashed = SecurityManager.hash_password(pwd)
        self.assertIn("$", hashed)
        self.assertTrue(SecurityManager.verify_password(pwd, hashed))
        self.assertFalse(SecurityManager.verify_password("WrongPassword", hashed))

    def test_password_verify_edge_cases(self):
        # Empty and None inputs
        self.assertFalse(SecurityManager.verify_password("", "some_hash"))
        self.assertFalse(SecurityManager.verify_password(None, "some_hash"))
        self.assertFalse(SecurityManager.verify_password("pwd", ""))
        self.assertFalse(SecurityManager.verify_password("pwd", None))
        self.assertFalse(SecurityManager.verify_password("pwd", "nohashdollar"))

    def test_rbac_permissions(self):
        # ADMIN permissions
        self.assertTrue(SecurityManager.has_permission("ADMIN", "dashboard.view"))
        self.assertTrue(SecurityManager.has_permission("ADMIN", "settings.manage"))
        self.assertTrue(SecurityManager.has_permission("admin", "ledger.void"))  # case-insensitive

        # FINANCE permissions
        self.assertTrue(SecurityManager.has_permission("FINANCE", "students.pay"))
        self.assertFalse(SecurityManager.has_permission("FINANCE", "settings.manage"))
        self.assertFalse(SecurityManager.has_permission("FINANCE", "ledger.void"))

        # AUDITOR permissions
        self.assertTrue(SecurityManager.has_permission("AUDITOR", "audit.view"))
        self.assertFalse(SecurityManager.has_permission("AUDITOR", "students.manage"))

        # Unknown role or None
        self.assertFalse(SecurityManager.has_permission("GUEST", "dashboard.view"))
        self.assertFalse(SecurityManager.has_permission("", "dashboard.view"))
        self.assertFalse(SecurityManager.has_permission(None, "dashboard.view"))

    def test_mask_citizen_id(self):
        # 13 digits
        self.assertEqual(
            SecurityManager.mask_citizen_id("1729900123456"),
            "1-7299-XXXXX-45-6",
        )
        # Formatted string with hyphens
        self.assertEqual(
            SecurityManager.mask_citizen_id("1-7299-00123-45-6"),
            "1-7299-XXXXX-45-6",
        )
        # Empty or None
        self.assertEqual(SecurityManager.mask_citizen_id(""), "")
        self.assertEqual(SecurityManager.mask_citizen_id(None), "")
        # Short digits
        self.assertEqual(SecurityManager.mask_citizen_id("12345678"), "123XXX78")
        self.assertEqual(SecurityManager.mask_citizen_id("123"), "XXX-XXXX-XXXX")

    def test_mask_bank_account(self):
        # >= 10 digits
        self.assertEqual(
            SecurityManager.mask_bank_account("1234567890"),
            "XXX-X-XX789-0",
        )
        # Empty or None
        self.assertEqual(SecurityManager.mask_bank_account(""), "")
        self.assertEqual(SecurityManager.mask_bank_account(None), "")
        # Short digits
        self.assertEqual(SecurityManager.mask_bank_account("12345"), "X2345")
        self.assertEqual(SecurityManager.mask_bank_account("12"), "XXX-X-XXXXX-X")

    def test_format_currency(self):
        self.assertEqual(SecurityManager.format_currency(125400.0), "฿125,400.00")
        self.assertEqual(SecurityManager.format_currency(0), "฿0.00")
        self.assertEqual(SecurityManager.format_currency("1500.5"), "฿1,500.50")
        self.assertEqual(SecurityManager.format_currency(None), "฿0.00")
        self.assertEqual(SecurityManager.format_currency("invalid"), "฿0.00")
        self.assertEqual(SecurityManager.format_currency(float("nan")), "฿0.00")
        self.assertEqual(SecurityManager.format_currency(float("inf")), "฿0.00")


class TestCoreConfig(unittest.TestCase):
    """Tests for core/config.py Thai formatting, date robust parsing, and config/backup."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_normalize_template_style(self):
        self.assertEqual(normalize_template_style("MODERN"), "MODERN")
        self.assertEqual(normalize_template_style("ทันสมัย"), "MODERN")
        self.assertEqual(normalize_template_style("CLASSIC"), "CLASSIC")
        self.assertEqual(normalize_template_style("ดั้งเดิม"), "CLASSIC")
        self.assertEqual(normalize_template_style(""), "CLASSIC")
        self.assertEqual(normalize_template_style(None), "CLASSIC")

    def test_thai_baht_text(self):
        self.assertEqual(thai_baht_text(0), "ศูนย์บาทถ้วน")
        self.assertEqual(thai_baht_text(-100), "ศูนย์บาทถ้วน")
        self.assertEqual(thai_baht_text("abc"), "ศูนย์บาทถ้วน")
        self.assertEqual(thai_baht_text(1), "หนึ่งบาทถ้วน")
        self.assertEqual(thai_baht_text(11), "สิบเอ็ดบาทถ้วน")
        self.assertEqual(thai_baht_text(21), "ยี่สิบเอ็ดบาทถ้วน")
        self.assertEqual(thai_baht_text(101), "หนึ่งร้อยเอ็ดบาทถ้วน")
        self.assertEqual(thai_baht_text(1001), "หนึ่งพันเอ็ดบาทถ้วน")
        self.assertEqual(thai_baht_text(950.0), "เก้าร้อยห้าสิบบาทถ้วน")
        self.assertEqual(thai_baht_text(2500.0), "สองพันห้าร้อยบาทถ้วน")
        # Million boundary test (verifying 'เอ็ด' fix)
        self.assertEqual(thai_baht_text(1000001), "หนึ่งล้านเอ็ดบาทถ้วน")
        self.assertEqual(thai_baht_text(20000000), "ยี่สิบล้านบาทถ้วน")
        self.assertEqual(thai_baht_text(20000001), "ยี่สิบล้านเอ็ดบาทถ้วน")
        # Satang
        self.assertEqual(thai_baht_text(10.50), "สิบบาทห้าสิบสตางค์")
        self.assertEqual(thai_baht_text(10.05), "สิบบาทห้าสตางค์")
        self.assertEqual(thai_baht_text(10.21), "สิบบาทยี่สิบเอ็ดสตางค์")

    def test_parse_date_robust(self):
        # ISO format
        self.assertEqual(parse_date_robust("2026-11-17"), datetime.date(2026, 11, 17))
        # ISO with T timestamp (verifying ISO-8601 fix)
        self.assertEqual(parse_date_robust("2026-11-17T14:30:00"), datetime.date(2026, 11, 17))
        self.assertEqual(parse_date_robust("2026-11-17 14:30:00"), datetime.date(2026, 11, 17))
        # Thai Buddhist DD/MM/YYYY
        self.assertEqual(parse_date_robust("17/11/2569"), datetime.date(2026, 11, 17))
        # Thai Buddhist 2-digit year
        self.assertEqual(parse_date_robust("17/11/69"), datetime.date(2026, 11, 17))
        # Date and datetime objects
        today = datetime.date.today()
        self.assertEqual(parse_date_robust(today), today)
        now = datetime.datetime.now()
        self.assertEqual(parse_date_robust(now), now.date())
        # Invalid / Empty
        self.assertIsNone(parse_date_robust(""))
        self.assertIsNone(parse_date_robust(None))
        self.assertIsNone(parse_date_robust("-"))
        self.assertIsNone(parse_date_robust("invalid-date"))

    def test_thai_date_formatting(self):
        dt = datetime.date(2026, 11, 17)
        self.assertEqual(format_thai_date_short(dt), "17/11/2569")
        self.assertEqual(format_thai_date_padded(dt), "17/11/2569")
        self.assertIn("17 พ.ย. 2569", format_thai_date_display(dt))
        self.assertEqual(format_receipt_date_thai("2026-11-17"), "17/11/2569")

    def test_receipt_number_helpers(self):
        self.assertEqual(split_receipt_prefix_and_number("RC-00125"), ("RC-", "00125"))
        self.assertEqual(split_receipt_prefix_and_number("700605433"), ("", "700605433"))
        self.assertEqual(split_receipt_prefix_and_number(""), ("", ""))
        self.assertEqual(split_receipt_prefix_and_number(None), ("", ""))

        self.assertEqual(format_next_receipt_number("00125"), "00126")
        self.assertEqual(format_next_receipt_number("99"), "100")
        self.assertEqual(format_next_receipt_number("700605433"), "700605434")

        self.assertEqual(combine_receipt_no("RC-", "0001"), "RC-0001")
        self.assertEqual(combine_receipt_no("RC-", "RC-0001"), "RC-0001")
        self.assertEqual(combine_receipt_no("", "700605433"), "700605433")

    def test_config_load_and_save(self):
        cfg = load_config(base_dir=self.temp_dir)
        self.assertEqual(cfg["college_name"], DEFAULT_CONFIG["college_name"])

        cfg["college_name"] = "วิทยาลัยทดสอบสุพรรณบุรี"
        cfg["receipt_book_no"] = "8888"
        save_config(cfg, base_dir=self.temp_dir)

        reloaded = load_config(base_dir=self.temp_dir)
        self.assertEqual(reloaded["college_name"], "วิทยาลัยทดสอบสุพรรณบุรี")
        self.assertEqual(reloaded["receipt_book_no"], "8888")

    def test_database_backup_and_restore(self):
        db_path = os.path.join(self.temp_dir, "test_orig.db")
        db = DatabaseManager(db_path)
        db.add_option("FACULTY", "คณะนวัตกรรม")

        # Backup
        backup_path = backup_database(db_path)
        self.assertTrue(os.path.exists(backup_path))
        self.assertGreater(os.path.getsize(backup_path), 0)

        # Full ZIP backup
        zip_path = export_full_zip_backup(db_path, evidence_dir=self.temp_dir)
        self.assertTrue(os.path.exists(zip_path))
        self.assertTrue(zip_path.endswith(".zip"))

        # Restore
        target_path = os.path.join(self.temp_dir, "test_restored.db")
        restore_database(backup_path, target_path)
        self.assertTrue(os.path.exists(target_path))

        db_restored = DatabaseManager(target_path)
        self.assertIn("คณะนวัตกรรม", db_restored.get_options("FACULTY"))

        # Non-sqlite file restore error
        bad_file = os.path.join(self.temp_dir, "bad.txt")
        with open(bad_file, "w") as f:
            f.write("not a sqlite database")
        with self.assertRaises(ValueError):
            restore_database(bad_file, target_path)


class TestDatabaseManagerCore(unittest.TestCase):
    """Comprehensive tests for core/database.py methods and transactions."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, "spvc_core_test.db")
        self.db = DatabaseManager(self.db_path)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_derive_year_level_and_education_level(self):
        self.assertEqual(derive_year_level_and_education_level("ปวช.1/1"), (1, "ปวช.1"))
        self.assertEqual(derive_year_level_and_education_level("ปวส.2/1"), (2, "ปวส.2"))
        self.assertEqual(derive_year_level_and_education_level("ทล.บ.1/1"), (1, "ทล.บ.1"))
        self.assertEqual(derive_year_level_and_education_level("", default_year=1, default_edu="ปวช.1"), (1, "ปวช.1"))
        self.assertEqual(derive_year_level_and_education_level(None, default_year=2, default_edu="ปวส.1"), (2, "ปวส.1"))

    def test_schema_idempotence_and_defaults(self):
        # Calling schema init multiple times should be safe
        self.db.initialize_schema()
        self.db.seed_defaults_if_empty()

        # Verify defaults seeded
        users = self.db.authenticate("admin", "admin123")
        self.assertIsNotNone(users)
        self.assertEqual(users["role"], "ADMIN")

        fin_user = self.db.authenticate("finance", "fin123")
        self.assertIsNotNone(fin_user)
        self.assertEqual(fin_user["role"], "FINANCE")

    def test_authenticate_safe_none_handling(self):
        self.assertIsNone(self.db.authenticate(None, None))
        self.assertIsNone(self.db.authenticate("", ""))
        self.assertIsNone(self.db.authenticate("admin", None))
        self.assertIsNone(self.db.authenticate(None, "admin123"))
        self.assertIsNone(self.db.authenticate("admin", "wrongpassword"))

    def test_master_options_crud(self):
        # Add option
        self.db.add_option("DEPARTMENT", "วิศวกรรมซอฟต์แวร์")
        opts = self.db.get_options("DEPARTMENT")
        self.assertIn("วิศวกรรมซอฟต์แวร์", opts)

        # Update option
        self.db.update_option("DEPARTMENT", "วิศวกรรมซอฟต์แวร์", "วิศวกรรมปัญญาประดิษฐ์")
        opts = self.db.get_options("DEPARTMENT")
        self.assertIn("วิศวกรรมปัญญาประดิษฐ์", opts)
        self.assertNotIn("วิศวกรรมซอฟต์แวร์", opts)

        # Delete option
        self.db.delete_option("DEPARTMENT", "วิศวกรรมปัญญาประดิษฐ์")
        opts = self.db.get_options("DEPARTMENT")
        self.assertNotIn("วิศวกรรมปัญญาประดิษฐ์", opts)

        # Auto generated TERMS
        terms = self.db.get_options("TERM")
        self.assertGreater(len(terms), 3)

    def test_fee_catalog_management(self):
        pvc_fees = self.db.get_fee_catalog(curriculum_filter="หลักสูตรในระบบ-ปวช")
        self.assertGreaterEqual(len(pvc_fees), 5)

        # Upsert new item
        self.db.upsert_fee_item(
            fee_code="SPVC-CUSTOM-01",
            fee_name="ค่าวัสดุโครงงานนวัตกรรม",
            calc_type="FIXED",
            default_price=450.0,
            curriculum_name="หลักสูตรในระบบ-ปวช",
        )
        catalog = self.db.get_fee_catalog()
        codes = [f["fee_code"] for f in catalog]
        self.assertIn("SPVC-CUSTOM-01", codes)

        # Update item
        self.db.upsert_fee_item(
            fee_code="SPVC-CUSTOM-01",
            fee_name="ค่าวัสดุโครงงานนวัตกรรม (แก้ไข)",
            calc_type="FIXED",
            default_price=500.0,
            curriculum_name="หลักสูตรในระบบ-ปวช",
        )
        updated_item = next(f for f in self.db.get_fee_catalog() if f["fee_code"] == "SPVC-CUSTOM-01")
        self.assertEqual(updated_item["default_price"], 500.0)
        self.assertEqual(updated_item["fee_name"], "ค่าวัสดุโครงงานนวัตกรรม (แก้ไข)")

        # Delete item
        self.db.delete_fee_item("SPVC-CUSTOM-01")
        catalog_after = self.db.get_fee_catalog()
        self.assertNotIn("SPVC-CUSTOM-01", [f["fee_code"] for f in catalog_after])

    def test_student_crud(self):
        # Create student
        self.db.add_student(
            student_code="692099999",
            full_name="นายวิชัย เทคโนโลยี",
            citizen_id="1729900999888",
            faculty="เทคโนโลยีสารสนเทศ",
            major="เทคโนโลยีสารสนเทศ",
            class_group="ปวช.1/1",
            total_fee=1200.0,
            phone="0819998888",
        )
        stu = self.db.get_student_by_code("692099999")
        self.assertIsNotNone(stu)
        self.assertEqual(stu["full_name"], "นายวิชัย เทคโนโลยี")
        self.assertEqual(stu["balance"], 1200.0)

        # Update student
        self.db.update_student(
            student_code="692099999",
            full_name="นายวิชัย เทคโนโลยี (ปรับปรุง)",
            faculty="เทคโนโลยีสารสนเทศ",
            major="เทคโนโลยีสารสนเทศ",
            class_group="ปวช.1/2",
        )
        stu_updated = self.db.get_student_by_code("692099999")
        self.assertEqual(stu_updated["full_name"], "นายวิชัย เทคโนโลยี (ปรับปรุง)")
        self.assertEqual(stu_updated["class_group"], "ปวช.1/2")

        # Count & Search with safe None handling
        self.assertGreaterEqual(self.db.count_students(search=None), 1)
        self.assertEqual(self.db.count_students(search="วิชัย"), 1)

        # Batch lookup
        batch = self.db.get_students_by_codes(["692099999", "692020101"])
        self.assertEqual(len(batch), 2)

        # Delete student
        self.db.delete_student("692099999")
        self.assertIsNone(self.db.get_student_by_code("692099999"))

    def test_student_bulk_import_csv(self):
        csv_file = os.path.join(self.temp_dir, "students_import.csv")
        with open(csv_file, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["student_code", "full_name", "citizen_id", "faculty", "major", "year", "scholarship", "fee", "phone", "edu_level", "class_group"])
            writer.writerow(["692090901", "นายมานะ เรียนดี", "1729900888111", "บริหารธุรกิจ", "การบัญชี", 1, "ปกติ", 950.0, "081-111-2222", "ปวช.1", "ปวช.1/1"])
            writer.writerow(["692090902", "นางสาวมานี รักเรียน", "1729900888222", "บริหารธุรกิจ", "การบัญชี", 1, "ปกติ", 950.0, "081-222-3333", "ปวช.1", "ปวช.1/1"])

        res = self.db.bulk_import_students(csv_file, actor="admin", role="ADMIN")
        self.assertEqual(res["inserted"], 2)
        self.assertEqual(res["total"], 2)

        # Verify students exist
        stu1 = self.db.get_student_by_code("692090901")
        self.assertIsNotNone(stu1)
        self.assertEqual(stu1["full_name"], "นายมานะ เรียนดี")

    def test_student_bulk_import_cp874_fallback(self):
        # Test Thai Windows CP874 encoded CSV
        csv_cp874 = os.path.join(self.temp_dir, "students_cp874.csv")
        with open(csv_cp874, "w", encoding="cp874", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["student_code", "full_name", "citizen_id", "faculty", "major", "year", "scholarship", "fee", "phone", "edu_level", "class_group"])
            writer.writerow(["692090903", "นายสมศักดิ์ วินโดวส์ไทย", "1729900888333", "บริหารธุรกิจ", "การบัญชี", 1, "ปกติ", 950.0, "081-333-4444", "ปวช.1", "ปวช.1/1"])

        res = self.db.bulk_import_students(csv_cp874, actor="admin", role="ADMIN")
        self.assertEqual(res["inserted"], 1)
        stu = self.db.get_student_by_code("692090903")
        self.assertIsNotNone(stu)
        self.assertEqual(stu["full_name"], "นายสมศักดิ์ วินโดวส์ไทย")

    def test_receipt_issuing_and_querying(self):
        # Test flexible receipt issue
        items = [
            {"fee_code": "SPVC-01", "item_name": "ค่าคู่มือนักเรียนใหม่", "quantity": 1, "unit_price": 100.0, "discount": 0.0},
            {"fee_code": "SPVC-03", "item_name": "ค่าบำรุงห้องสมุด", "quantity": 1, "unit_price": 50.0, "discount": 0.0},
            {"fee_code": "SPVC-05", "item_name": "เงินรายได้สถานศึกษา", "quantity": 1, "unit_price": 300.0, "discount": 50.0},
        ]
        rcp = self.db.issue_flexible_receipt(
            student_code="692020102",  # Unpaid seeded student
            line_items=items,
            payment_method="โอนบัญชี (Transfer)",
            actor="finance",
            cashier_name="คุณกนกวรรณ",
            receipt_date_override=None,  # Verify None-safe handling
            receipt_no_override=None,
        )
        self.assertIsNotNone(rcp)
        self.assertEqual(rcp["subtotal_amount"], 450.0)
        self.assertEqual(rcp["discount_amount"], 50.0)
        self.assertEqual(rcp["net_total_amount"], 400.0)
        self.assertEqual(len(rcp["items"]), 3)

        # Check student balance updated
        stu = self.db.get_student_by_code("692020102")
        self.assertEqual(stu["paid_amount"], 400.0)
        self.assertEqual(stu["status"], "PARTIAL")

        # Query receipt
        rcp_fetched = self.db.get_receipt_by_no(rcp["receipt_no"])
        self.assertIsNotNone(rcp_fetched)
        self.assertEqual(rcp_fetched["receipt_no"], rcp["receipt_no"])

        # Stats query with search=None
        stats = self.db.get_receipts_stats(search=None)
        self.assertGreaterEqual(stats["valid_count"], 1)
        self.assertGreaterEqual(stats["total_amount"], 400.0)

    def test_issue_batch_class_receipts(self):
        students = [
            {"student_code": "692020101", "full_name": "นายธนภัทร ศรีสุพรรณ"},
            {"student_code": "692020103", "full_name": "นางสาวกมลชนก สุพรรณภูมิ"},
        ]
        items = [
            {"fee_code": "SPVC-06", "item_name": "ค่ารักษาสภาพแวดล้อม", "quantity": 1, "unit_price": 200.0, "discount": 0.0},
        ]
        issued = self.db.issue_batch_class_receipts(
            students_list=students,
            line_items=items,
            payment_method="เงินสด (Cash)",
            actor="admin",
            cashier_name="ครูผู้รับเงิน",
            receipt_date_override=None,  # Test None-safe override
        )
        self.assertEqual(len(issued), 2)
        self.assertEqual(issued[0]["net_total_amount"], 200.0)
        self.assertEqual(issued[1]["net_total_amount"], 200.0)

    def test_void_receipt(self):
        items = [{"fee_code": "SPVC-TEST", "item_name": "ค่าทดสอบ", "quantity": 1, "unit_price": 500.0, "discount": 0.0}]
        rcp = self.db.issue_flexible_receipt(
            student_code="692020101",
            line_items=items,
            payment_method="เงินสด",
            actor="finance",
            cashier_name="คุณกนกวรรณ",
        )
        r_no = rcp["receipt_no"]

        # Void receipt
        self.db.void_receipt(r_no, reason="ทดสอบการยกเลิกใบเสร็จ", actor="admin", role="ADMIN")

        voided_rcp = self.db.get_receipt_by_no(r_no)
        self.assertEqual(voided_rcp["status"], "VOIDED")
        self.assertEqual(voided_rcp["void_reason"], "ทดสอบการยกเลิกใบเสร็จ")

        # Voiding again should raise ValueError
        with self.assertRaises(ValueError):
            self.db.void_receipt(r_no, reason="ยกเลิกซ้ำ", actor="admin", role="ADMIN")

        # Voiding with empty reason or None should raise ValueError
        with self.assertRaises(ValueError):
            self.db.void_receipt(r_no, reason="", actor="admin", role="ADMIN")
        with self.assertRaises(ValueError):
            self.db.void_receipt(None, reason="เหตุผล", actor="admin", role="ADMIN")

    def test_audit_hash_chain_integrity(self):
        self.db.log_audit("admin", "ADMIN", "TEST_ACTION_1", "TARGET_1", "Detail 1")
        self.db.log_audit("admin", "ADMIN", "TEST_ACTION_2", "TARGET_2", "Detail 2")

        is_valid, count, msg = self.db.verify_audit_integrity()
        self.assertTrue(is_valid)
        self.assertGreaterEqual(count, 2)
        self.assertIn("100%", msg)

        # Verify tamper detection: alter one row directly
        with self.db._connect() as conn:
            conn.execute("UPDATE audit_logs SET detail = 'TAMPERED' WHERE id = 1")
            conn.commit()

        tampered_valid, _, tamper_msg = self.db.verify_audit_integrity()
        self.assertFalse(tampered_valid)
        self.assertTrue("Chain broken" in tamper_msg or "mismatch" in tamper_msg)

    def test_dashboard_analytics_and_summaries(self):
        metadata = self.db.get_dashboard_filter_metadata()
        self.assertIn("terms", metadata)
        self.assertIn("levels", metadata)
        self.assertIn("departments", metadata)

        analytics = self.db.get_dashboard_analytics(
            term_filter="ALL",
            period_filter="ALL",
            level_filter="ALL",
            department_filter="ALL",
        )
        self.assertIn("total_income", analytics)
        self.assertIn("collection_rate_percent", analytics)
        self.assertIn("by_payment_method", analytics)
        self.assertIn("top_fee_items", analytics)

        summary = self.db.get_dashboard_summary()
        self.assertIn("total_income", summary)
        self.assertIn("outstanding_receivables", summary)

        daily = self.db.get_daily_closing_summary()
        self.assertIn("income_total", daily)
        self.assertIn("income_by_method", daily)

    def test_course_registration_catalog_and_save(self):
        # Catalog lookup
        courses = self.db.get_course_catalog()
        self.assertGreaterEqual(len(courses), 5)

        # Upsert course
        self.db.upsert_course_catalog(
            course_code="30000-9999",
            course_name="การประยุกต์ใช้ AI ในงานบัญชี",
            theory_hours=2,
            practice_hours=2,
            credits=3,
            major="การบัญชี",
            class_level="ปวส.1",
        )
        all_courses = self.db.get_course_catalog()
        self.assertIn("30000-9999", [c["course_code"] for c in all_courses])

        # Student registered courses
        reg = self.db.get_student_registered_courses("69219100049")
        self.assertGreaterEqual(len(reg), 1)

        # Save new student registrations
        new_reg_courses = [
            {"course_code": "30000-9999", "course_name": "การประยุกต์ใช้ AI ในงานบัญชี", "theory_hours": 2, "practice_hours": 2, "credits": 3}
        ]
        self.db.save_student_registered_courses(
            student_code="692020101",
            term_code="2/2569",
            courses_list=new_reg_courses,
        )
        saved = self.db.get_student_registered_courses("692020101")
        self.assertEqual(len(saved), 1)
        self.assertEqual(saved[0]["course_code"], "30000-9999")

        # Safe None handling on student code
        self.assertEqual(self.db.get_student_registered_courses(None), [])


if __name__ == "__main__":
    unittest.main()
