import io
import os
import shutil
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch
import urllib.error

# Ensure root directory is on sys.path
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import docx
import openpyxl
from exporters.base import BaseFinanceExporter, safe_float, safe_int, safe_str
from exporters.docx_exporter import DocxReceiptExporter
from exporters.excel_exporter import ExcelFinanceExporter
from exporters.pdf_exporter import PdfReceiptExporter, PdfStatementExporter
from exporters.webhook_exporter import FinanceWebhookDispatcher


class DummyExporter(BaseFinanceExporter):
    def export(self, data, metadata=None):
        return self.save_path


class TestBaseFinanceExporter(unittest.TestCase):
    """Unit tests for BaseFinanceExporter and safe parsing utilities."""

    def test_abstract_instantiation_fails(self):
        with self.assertRaises(TypeError):
            BaseFinanceExporter("test.pdf")

    def test_empty_or_none_save_path_raises_value_error(self):
        with self.assertRaises(ValueError):
            DummyExporter("")
        with self.assertRaises(ValueError):
            DummyExporter("   ")
        with self.assertRaises(ValueError):
            DummyExporter(None)

    def test_dummy_exporter_success(self):
        exp = DummyExporter("my_report.pdf", college_name="  วิทยาลัยเทคนิค  ")
        self.assertEqual(exp.save_path, "my_report.pdf")
        self.assertEqual(exp.college_name, "วิทยาลัยเทคนิค")
        self.assertEqual(exp.export({}), "my_report.pdf")

    def test_safe_float_conversions(self):
        self.assertEqual(safe_float(123), 123.0)
        self.assertEqual(safe_float(123.45), 123.45)
        self.assertEqual(safe_float("456.78"), 456.78)
        self.assertEqual(safe_float("-99.5"), -99.5)
        self.assertEqual(safe_float(None), 0.0)
        self.assertEqual(safe_float("", default=10.0), 10.0)
        self.assertEqual(safe_float("invalid", default=-1.0), -1.0)
        self.assertEqual(safe_float([], default=0.0), 0.0)

    def test_safe_int_conversions(self):
        self.assertEqual(safe_int(100), 100)
        self.assertEqual(safe_int(123.9), 123)
        self.assertEqual(safe_int("456"), 456)
        self.assertEqual(safe_int("789.2"), 789)
        self.assertEqual(safe_int(None), 0)
        self.assertEqual(safe_int("", default=5), 5)
        self.assertEqual(safe_int("abc", default=99), 99)
        self.assertEqual(safe_int({}, default=0), 0)

    def test_safe_str_conversions(self):
        self.assertEqual(safe_str("  hello  "), "hello")
        self.assertEqual(safe_str(None, default="N/A"), "N/A")
        self.assertEqual(safe_str(123), "123")
        self.assertEqual(safe_str("ภาษาไทย"), "ภาษาไทย")


class TestFinanceWebhookDispatcher(unittest.TestCase):
    """Unit tests for FinanceWebhookDispatcher."""

    def test_disabled_dispatcher_does_nothing(self):
        dispatcher = FinanceWebhookDispatcher("https://example.com/webhook", enabled=False)
        self.assertFalse(dispatcher.enabled)
        res = dispatcher.notify_event_async("TEST", "Title", 1000.0, "admin")
        self.assertIsNone(res)

    def test_empty_url_is_disabled(self):
        dispatcher = FinanceWebhookDispatcher("", enabled=True)
        self.assertFalse(dispatcher.enabled)

    def test_threshold_and_high_value_flag(self):
        dispatcher = FinanceWebhookDispatcher(
            "https://httpbin.org/post",
            enabled=True,
            alert_threshold=50000.0,
        )
        self.assertTrue(dispatcher.enabled)
        self.assertEqual(dispatcher.alert_threshold, 50000.0)

    @patch("urllib.request.urlopen")
    def test_notify_event_async_success(self, mock_urlopen):
        mock_resp = MagicMock()
        mock_resp.getcode.return_value = 200
        mock_resp.__enter__.return_value = mock_resp
        mock_urlopen.return_value = mock_resp

        dispatcher = FinanceWebhookDispatcher(
            "https://example.com/webhook",
            enabled=True,
            alert_threshold=10000.0,
        )

        callback_results = []
        def on_complete(ok, msg):
            callback_results.append((ok, msg))

        thread = dispatcher.notify_event_async(
            event_type="PAYMENT_RECEIVED",
            title="ชำระค่าเทอม ปวช.1",
            amount=15000.0,
            actor="cashier_01",
            ref_code="RCP-2026-001",
            extra={"semester": "2/2569"},
            callback=on_complete,
        )

        self.assertIsNotNone(thread)
        thread.join(timeout=3.0)

        self.assertEqual(len(callback_results), 1)
        self.assertTrue(callback_results[0][0])
        self.assertEqual(callback_results[0][1], "HTTP 200")
        mock_urlopen.assert_called_once()

    @patch("urllib.request.urlopen")
    def test_notify_event_async_http_error(self, mock_urlopen):
        err = urllib.error.HTTPError(
            url="https://example.com/webhook",
            code=500,
            msg="Internal Server Error",
            hdrs={},
            fp=io.BytesIO(b"error"),
        )
        mock_urlopen.side_effect = err

        dispatcher = FinanceWebhookDispatcher(
            "https://example.com/webhook",
            enabled=True,
        )

        callback_results = []
        def on_complete(ok, msg):
            callback_results.append((ok, msg))

        thread = dispatcher.notify_event_async(
            event_type="TEST",
            title="Test Error",
            amount=500.0,
            actor="test_user",
            callback=on_complete,
        )
        self.assertIsNotNone(thread)
        thread.join(timeout=3.0)

        self.assertEqual(len(callback_results), 1)
        self.assertFalse(callback_results[0][0])
        self.assertIn("500", callback_results[0][1])

    @patch("urllib.request.urlopen")
    def test_notify_event_async_url_error(self, mock_urlopen):
        mock_urlopen.side_effect = urllib.error.URLError("Connection refused")

        dispatcher = FinanceWebhookDispatcher(
            "https://invalid-host-999.org",
            enabled=True,
        )

        callback_results = []
        def on_complete(ok, msg):
            callback_results.append((ok, msg))

        thread = dispatcher.notify_event_async(
            event_type="TEST",
            title="Test URLError",
            amount="invalid_number",  # Safe conversion to 0.0
            actor=None,
            callback=on_complete,
        )
        self.assertIsNotNone(thread)
        thread.join(timeout=3.0)

        self.assertEqual(len(callback_results), 1)
        self.assertFalse(callback_results[0][0])
        self.assertIn("Connection refused", callback_results[0][1])

    @patch("urllib.request.urlopen")
    def test_broken_callback_does_not_crash_thread(self, mock_urlopen):
        mock_resp = MagicMock()
        mock_resp.getcode.return_value = 200
        mock_resp.__enter__.return_value = mock_resp
        mock_urlopen.return_value = mock_resp

        dispatcher = FinanceWebhookDispatcher("https://example.com/webhook", enabled=True)

        def faulty_callback(ok, msg):
            raise RuntimeError("Boom!")

        thread = dispatcher.notify_event_async(
            event_type="TEST",
            title="Faulty Callback",
            amount=100.0,
            actor="admin",
            callback=faulty_callback,
        )
        self.assertIsNotNone(thread)
        thread.join(timeout=3.0)
        self.assertFalse(thread.is_alive())


class TestExcelFinanceExporter(unittest.TestCase):
    """Unit tests for ExcelFinanceExporter."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="test_excel_")

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_standard_multi_sheet_export(self):
        save_path = os.path.join(self.test_dir, "financial_report.xlsx")
        exporter = ExcelFinanceExporter(save_path, college_name="วิทยาลัยอาชีวศึกษาสุพรรณบุรี")

        data = {
            "summary": {
                "total_income": 450000.0,
                "total_expense": 120000.0,
                "net_balance": 330000.0,
            },
            "transactions": [
                {
                    "tx_code": "TX-001",
                    "tx_date": "2026-11-17 10:00",
                    "tx_type": "INCOME",
                    "category": "ค่าธรรมเนียมการศึกษา",
                    "title": "ค่าบำรุงการศึกษา ภาคเรียนที่ 2/2569",
                    "amount": 2500.0,
                    "payment_method": "เงินสด (Cash)",
                    "ref_code": "RCP-001",
                    "recorded_by": "เจ้าหน้าที่การเงิน",
                    "status": "COMPLETED",
                },
                {
                    "tx_code": "TX-002",
                    "tx_date": "2026-11-17 11:30",
                    "tx_type": "EXPENSE",
                    "category": "ค่าวัสดุฝึก",
                    "title": "ซื้ออุปกรณ์คอมพิวเตอร์และสื่อการสอน",
                    "amount": 15000.0,
                    "payment_method": "โอนเงินผ่านธนาคาร (Bank Transfer)",
                    "ref_code": "PO-2026-09",
                    "recorded_by": "หัวหน้าแผนก",
                    "status": "COMPLETED",
                },
            ],
            "students": [
                {
                    "student_code": "69219100049",
                    "full_name": "นางสาวปรันดา เจียศิริพันธ์",
                    "education_level": "ปวช.1",
                    "class_group": "ปวช. 1/2",
                    "faculty": "พาณิชยกรรม/บริหารธุรกิจ",
                    "major": "เทคโนโลยีธุรกิจดิจิทัล",
                    "scholarship_type": "ปกติ",
                    "total_fee": 3500.0,
                    "paid_amount": 3500.0,
                    "status": "PAID",
                },
                {
                    "student_code": "69219100050",
                    "full_name": "นายสมชาย ใจดี",
                    "education_level": "ปวส.1",
                    "class_group": "ปวส. 1/1",
                    "faculty": "ช่างอุตสาหกรรม",
                    "major": "ช่างยนต์",
                    "scholarship_type": "กู้ยืม กยศ.",
                    "total_fee": 5000.0,
                    "paid_amount": 2000.0,
                    "status": "PARTIAL",
                },
            ],
            "budgets": [
                {
                    "dept_code": "D01",
                    "dept_name": "แผนกวิชาเทคโนโลยีธุรกิจดิจิทัล",
                    "fiscal_year": "2569",
                    "allocated_budget": 100000.0,
                    "spent_budget": 45000.0,
                }
            ],
        }

        ret = exporter.export(data)
        self.assertEqual(ret, save_path)
        self.assertTrue(os.path.exists(save_path))
        self.assertGreater(os.path.getsize(save_path), 0)

        wb = openpyxl.load_workbook(save_path)
        self.assertIn("General Ledger", wb.sheetnames)
        self.assertIn("Student Tuition", wb.sheetnames)
        self.assertIn("Department Budgets", wb.sheetnames)

        ws_ledger = wb["General Ledger"]
        self.assertIn("สมุดบัญชีแยกประเภท", ws_ledger.cell(row=1, column=1).value)
        self.assertEqual(ws_ledger.cell(row=5, column=2).value, "TX-001")
        self.assertEqual(ws_ledger.cell(row=5, column=7).value, 2500.0)

        ws_stu = wb["Student Tuition"]
        self.assertEqual(ws_stu.cell(row=4, column=1).value, "69219100049")
        self.assertEqual(ws_stu.cell(row=4, column=11).value, "PAID")
        wb.close()

    def test_export_with_none_values_and_empty_data(self):
        save_path = os.path.join(self.test_dir, "empty_report.xlsx")
        exporter = ExcelFinanceExporter(save_path)

        data = {
            "summary": {
                "total_income": None,
                "total_expense": None,
                "net_balance": None,
            },
            "transactions": [
                {
                    "tx_code": None,
                    "amount": None,
                    "status": None,
                }
            ],
            "students": [
                {
                    "student_code": None,
                    "total_fee": None,
                    "paid_amount": None,
                    "status": None,
                }
            ],
            "budgets": [
                {
                    "allocated_budget": None,
                    "spent_budget": None,
                }
            ],
        }

        ret = exporter.export(data)
        self.assertTrue(os.path.exists(ret))
        wb = openpyxl.load_workbook(ret)
        ws_ledger = wb["General Ledger"]
        # Ensure summary string formatted cleanly without crashing on None
        self.assertIn("รายรับรวม: ฿0.00", ws_ledger.cell(row=2, column=1).value)
        wb.close()

    def test_high_volume_streaming_mode(self):
        save_path = os.path.join(self.test_dir, "streaming_ledger.xlsx")
        exporter = ExcelFinanceExporter(save_path)

        # Trigger > 5000 rows streaming mode
        mock_txs = [
            {
                "tx_code": f"TX-{i:05d}",
                "tx_date": "2026-11-17",
                "tx_type": "INCOME",
                "category": "ค่าธรรมเนียม",
                "title": f"รายการที่ {i}",
                "amount": 100.0,
                "payment_method": "เงินสด",
                "ref_code": "",
                "recorded_by": "ระบบ",
                "status": "COMPLETED",
            }
            for i in range(5005)
        ]

        data = {
            "summary": {"total_income": 500500.0, "total_expense": 0.0, "net_balance": 500500.0},
            "transactions": mock_txs,
            "students": [],
            "budgets": [],
        }

        ret = exporter.export(data)
        self.assertTrue(os.path.exists(ret))
        self.assertGreater(os.path.getsize(ret), 10000)

        wb = openpyxl.load_workbook(ret, read_only=True)
        self.assertIn("General Ledger", wb.sheetnames)
        ws = wb["General Ledger"]
        rows = list(ws.iter_rows(values_only=True))
        self.assertGreater(len(rows), 5000)
        wb.close()

    def test_export_dashboard_analytics(self):
        save_path = os.path.join(self.test_dir, "dashboard_analytics.xlsx")
        exporter = ExcelFinanceExporter(save_path)

        analytics = {
            "term_filter": "2/2569",
            "period_label": "พฤศจิกายน 2569",
            "level_label": "ปวช.1",
            "department_label": "เทคโนโลยีธุรกิจดิจิทัล",
            "total_income": 1250000.0,
            "month_income": 450000.0,
            "today_income": 35000.0,
            "receipt_count": 250,
            "void_receipt_count": 3,
            "total_tuition_expected": 1500000.0,
            "outstanding_receivables": 250000.0,
            "student_count": 300,
            "paid_student_count": 250,
            "unpaid_student_count": 50,
            "collection_rate_percent": 83.33,
            "by_payment_method": [
                {"name": "เงินสด (Cash)", "count": 150, "amount": 750000.0, "percent": 60.0},
                {"name": "โอนเงินผ่านธนาคาร", "count": 100, "amount": 500000.0, "percent": 40.0},
            ],
            "by_education_level": [
                {"level": "ปวช.1", "count": 150, "amount": 600000.0, "percent": 48.0},
                {"level": "ปวช.2", "count": 100, "amount": 650000.0, "percent": 52.0},
            ],
            "top_fee_items": [
                {"item_name": "ค่าบำรุงการศึกษา", "quantity": 250, "total": 625000.0, "percent": 50.0},
                {"item_name": "ค่าอุปกรณ์การเรียน", "quantity": 250, "total": 250000.0, "percent": 20.0},
            ],
            "recent_receipts": [
                {
                    "receipt_no": "RCP-2026-0001",
                    "receipt_date": "2026-11-17 09:30",
                    "term_code": "2/2569",
                    "student_code": "69219100049",
                    "snapshot_student_name": "นางสาวปรันดา เจียศิริพันธ์",
                    "snapshot_education_level": "ปวช.1",
                    "snapshot_major": "เทคโนโลยีธุรกิจดิจิทัล",
                    "payment_method": "เงินสด",
                    "net_total_amount": 2500.0,
                    "cashier_name": "เจ้าหน้าที่การเงิน",
                    "status": "ISSUED",
                }
            ],
        }

        ret = exporter.export_dashboard(analytics, metadata={"cashier_name": "คุณวิภาวี"})
        self.assertTrue(os.path.exists(ret))
        wb = openpyxl.load_workbook(ret)
        self.assertIn("Dashboard Summary", wb.sheetnames)
        self.assertIn("Recent Receipts", wb.sheetnames)
        wb.close()

    def test_export_dashboard_empty_analytics(self):
        save_path = os.path.join(self.test_dir, "empty_dashboard.xlsx")
        exporter = ExcelFinanceExporter(save_path)
        ret = exporter.export_dashboard({})
        self.assertTrue(os.path.exists(ret))


class TestDocxReceiptExporter(unittest.TestCase):
    """Unit tests for DocxReceiptExporter."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="test_docx_")
        self.base_data = {
            "receipt_no": "700605308",
            "receipt_date": "2026-10-30",
            "term_code": "2/2569",
            "student_code": "69219100049",
            "full_name": "นางสาวปรันดา เจียศิริพันธ์",
            "class_group": "ปวช. 1/2",
            "education_level": "ปวช.1",
            "faculty": "พาณิชยกรรม/บริหารธุรกิจ",
            "major": "เทคโนโลยีธุรกิจดิจิทัล",
            "cashier_position": "ครู",
            "net_total_amount": 2500.0,
            "items": [
                {"line_no": 1, "item_name": "ค่าบำรุงการศึกษา", "quantity": 1, "unit_price": 2000.0, "line_total": 2000.0},
                {"line_no": 2, "item_name": "ค่าอุปกรณ์การเรียน", "quantity": 2, "unit_price": 250.0, "line_total": 500.0},
            ],
        }

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_export_classic_receipt(self):
        save_path = os.path.join(self.test_dir, "receipt_classic.docx")
        exporter = DocxReceiptExporter(save_path, college_name="วิทยาลัยอาชีวศึกษาสุพรรณบุรี")
        ret = exporter.export(self.base_data, metadata={"bill_template_style": "CLASSIC"})
        self.assertTrue(os.path.exists(ret))
        self.assertGreater(os.path.getsize(ret), 0)

        doc = docx.Document(ret)
        full_text = " ".join(p.text for p in doc.paragraphs)
        for t in doc.tables:
            for row in t.rows:
                for c in row.cells:
                    full_text += " " + c.text

        self.assertIn("ใบเสร็จรับเงิน", full_text)
        self.assertIn("วิทยาลัยอาชีวศึกษาสุพรรณบุรี", full_text)
        self.assertIn("ปรันดา เจียศิริพันธ์", full_text)
        self.assertIn("69219100049", full_text)
        self.assertIn("สองพันห้าร้อยบาทถ้วน", full_text)

    def test_export_modern_receipt(self):
        save_path = os.path.join(self.test_dir, "receipt_modern.docx")
        exporter = DocxReceiptExporter(save_path)
        payload = dict(self.base_data)
        payload["template_style"] = "MODERN"
        ret = exporter.export(payload)
        self.assertTrue(os.path.exists(ret))
        doc = docx.Document(ret)
        self.assertGreater(len(doc.tables), 2)

    def test_export_voided_receipt(self):
        save_path = os.path.join(self.test_dir, "receipt_voided.docx")
        exporter = DocxReceiptExporter(save_path)
        payload = dict(self.base_data)
        payload["status"] = "VOIDED"
        payload["template_style"] = "MODERN"
        ret = exporter.export(payload)
        doc = docx.Document(ret)
        full_text = " ".join(p.text for p in doc.paragraphs)
        self.assertIn("VOIDED", full_text)

    def test_export_course_registration_card_classic_and_modern(self):
        reg_data = {
            "report_type": "COURSE_REGISTRATION",
            "student_code": "69219100049",
            "full_name": "นางสาวปรันดา เจียศิริพันธ์",
            "class_group": "ปวช. 1/2",
            "fiscal_year": "2569",
            "major": "เทคโนโลยีธุรกิจดิจิทัล",
            "subject_type": "อุตสาหกรรมดิจิทัลและเทคโนโลยีสารสนเทศ",
            "career_group": "ธุรกิจดิจิทัลและพาณิชย์อิเล็กทรอนิกส์",
            "courses": [
                {"course_code": "20000-1101", "course_name": "ภาษาไทยพื้นฐาน", "theory_hours": 2, "practice_hours": 0, "credits": 2},
                {"course_code": "20204-2001", "course_name": "การเขียนโปรแกรมคอมพิวเตอร์", "theory_hours": 2, "practice_hours": 2, "credits": 3},
            ],
            "fee_items": [
                {"item_name": "ค่าใช้บริการอินเตอร์เน็ต", "line_total": 300.0},
                {"item_name": "ค่ารักษาสภาพแวดล้อม", "line_total": 330.0},
            ],
        }

        # Classic
        p_classic = os.path.join(self.test_dir, "ost28_classic.docx")
        DocxReceiptExporter(p_classic).export(reg_data, metadata={"bill_template_style": "CLASSIC"})
        self.assertTrue(os.path.exists(p_classic))

        # Modern
        p_modern = os.path.join(self.test_dir, "ost28_modern.docx")
        reg_modern = dict(reg_data)
        reg_modern["template_style"] = "MODERN"
        DocxReceiptExporter(p_modern).export(reg_modern)
        self.assertTrue(os.path.exists(p_modern))

    def test_export_daily_closing_report(self):
        closing_data = {
            "report_type": "DAILY_CLOSING",
            "date": "2026-11-17",
            "tx_count": 12,
            "income_total": 35400.0,
            "income_by_method": {
                "เงินสด (Cash)": 20400.0,
                "โอนเงิน (QR PromptPay)": 15000.0,
            },
            "income_by_fee_item": {
                "ค่าบำรุงการศึกษา": 30000.0,
                "ค่าคู่มือนักเรียนใหม่": 5400.0,
            },
            "cashier_name": "นางสาวกนกวรรณ รับชำระ",
            "cashier_position": "เจ้าหน้าที่การเงิน",
        }
        save_path = os.path.join(self.test_dir, "daily_closing.docx")
        DocxReceiptExporter(save_path).export(closing_data)
        self.assertTrue(os.path.exists(save_path))
        doc = docx.Document(save_path)
        full_text = " ".join(p.text for p in doc.paragraphs)
        for t in doc.tables:
            for row in t.rows:
                for c in row.cells:
                    full_text += " " + c.text
        self.assertIn("รายงานสรุปการรับเงินค่าเทอมประจำวัน", full_text)
        self.assertIn("กนกวรรณ รับชำระ", full_text)

    def test_export_batch_receipts(self):
        save_path = os.path.join(self.test_dir, "batch_receipts.docx")
        exporter = DocxReceiptExporter(save_path)

        receipts = [
            dict(self.base_data, receipt_no="RCP-001"),
            dict(self.base_data, receipt_no="RCP-002", full_name="นายทดสอบ สอง"),
        ]
        ret = exporter.export_batch(receipts)
        self.assertTrue(os.path.exists(ret))
        doc = docx.Document(ret)
        full_text = " ".join(p.text for p in doc.paragraphs)
        for t in doc.tables:
            for row in t.rows:
                for c in row.cells:
                    full_text += " " + c.text
        self.assertIn("RCP-001", full_text)
        self.assertIn("RCP-002", full_text)

    def test_export_batch_empty_raises_value_error(self):
        exporter = DocxReceiptExporter(os.path.join(self.test_dir, "dummy.docx"))
        with self.assertRaises(ValueError):
            exporter.export_batch([])
        with self.assertRaises(ValueError):
            exporter.export_batch(None)

    def test_export_with_none_values_and_empty_dict(self):
        save_path = os.path.join(self.test_dir, "empty_docx.docx")
        exporter = DocxReceiptExporter(save_path)
        ret = exporter.export({})
        self.assertTrue(os.path.exists(ret))

        # Test items with None numeric values
        save_path_none = os.path.join(self.test_dir, "none_values_docx.docx")
        ret_none = exporter.export({
            "net_total_amount": None,
            "amount": None,
            "items": [
                {"item_name": None, "quantity": None, "unit_price": None, "line_total": None}
            ]
        })
        self.assertTrue(os.path.exists(ret_none))


class TestPdfReceiptExporter(unittest.TestCase):
    """Unit tests for PdfReceiptExporter."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="test_pdf_")
        self.base_data = {
            "receipt_no": "700605308",
            "receipt_date": "2026-10-30",
            "term_code": "2/2569",
            "student_code": "69219100049",
            "full_name": "นางสาวปรันดา เจียศิริพันธ์",
            "class_group": "ปวช. 1/2",
            "education_level": "ปวช.1",
            "faculty": "พาณิชยกรรม/บริหารธุรกิจ",
            "major": "เทคโนโลยีธุรกิจดิจิทัล",
            "cashier_position": "ครู",
            "net_total_amount": 2500.0,
            "items": [
                {"line_no": 1, "item_name": "ค่าบำรุงการศึกษา", "quantity": 1, "unit_price": 2000.0, "line_total": 2000.0},
                {"line_no": 2, "item_name": "ค่าวัสดุและอุปกรณ์", "quantity": 1, "unit_price": 500.0, "line_total": 500.0},
            ],
        }

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_export_classic_pdf_receipt(self):
        save_path = os.path.join(self.test_dir, "receipt_classic.pdf")
        exporter = PdfReceiptExporter(save_path, college_name="วิทยาลัยอาชีวศึกษาสุพรรณบุรี")
        ret = exporter.export(self.base_data, metadata={"bill_template_style": "CLASSIC"})
        self.assertTrue(os.path.exists(ret))
        self.assertGreater(os.path.getsize(ret), 0)

        with open(ret, "rb") as f:
            header = f.read(5)
            self.assertEqual(header, b"%PDF-")

    def test_export_modern_pdf_receipt(self):
        save_path = os.path.join(self.test_dir, "receipt_modern.pdf")
        exporter = PdfReceiptExporter(save_path)
        payload = dict(self.base_data)
        payload["template_style"] = "MODERN"
        ret = exporter.export(payload)
        self.assertTrue(os.path.exists(ret))

        with open(ret, "rb") as f:
            header = f.read(5)
            self.assertEqual(header, b"%PDF-")

    def test_export_voided_pdf_receipt(self):
        save_path = os.path.join(self.test_dir, "receipt_voided.pdf")
        exporter = PdfReceiptExporter(save_path)
        payload = dict(self.base_data)
        payload["status"] = "VOIDED"
        payload["template_style"] = "MODERN"
        ret = exporter.export(payload)
        self.assertTrue(os.path.exists(ret))

    def test_export_course_registration_card_classic_and_modern(self):
        reg_data = {
            "report_type": "COURSE_REGISTRATION",
            "student_code": "69219100049",
            "full_name": "นางสาวปรันดา เจียศิริพันธ์",
            "class_group": "ปวช. 1/2",
            "fiscal_year": "2569",
            "major": "เทคโนโลยีธุรกิจดิจิทัล",
            "subject_type": "อุตสาหกรรมดิจิทัลและเทคโนโลยีสารสนเทศ",
            "career_group": "ธุรกิจดิจิทัลและพาณิชย์อิเล็กทรอนิกส์",
            "courses": [
                {"course_code": "20000-1101", "course_name": "ภาษาไทยพื้นฐาน", "theory_hours": 2, "practice_hours": 0, "credits": 2},
                {"course_code": "20204-2001", "course_name": "การเขียนโปรแกรมคอมพิวเตอร์", "theory_hours": 2, "practice_hours": 2, "credits": 3},
            ],
            "fee_items": [
                {"item_name": "ค่าใช้บริการอินเตอร์เน็ต", "line_total": 300.0},
                {"item_name": "ค่ารักษาสภาพแวดล้อม", "line_total": 330.0},
            ],
        }

        # Classic PDF
        p_classic = os.path.join(self.test_dir, "ost28_classic.pdf")
        PdfReceiptExporter(p_classic).export(reg_data, metadata={"bill_template_style": "CLASSIC"})
        self.assertTrue(os.path.exists(p_classic))

        # Modern PDF
        p_modern = os.path.join(self.test_dir, "ost28_modern.pdf")
        reg_modern = dict(reg_data)
        reg_modern["template_style"] = "MODERN"
        PdfReceiptExporter(p_modern).export(reg_modern)
        self.assertTrue(os.path.exists(p_modern))

    def test_export_daily_closing_pdf(self):
        closing_data = {
            "report_type": "DAILY_CLOSING",
            "date": "2026-11-17",
            "tx_count": 8,
            "income_total": 24000.0,
            "income_by_method": {
                "เงินสด": 14000.0,
                "พร้อมเพย์": 10000.0,
            },
            "income_by_fee_item": {
                "ค่าธรรมเนียมการศึกษา": 24000.0,
            },
            "cashier_name": "นางสาวกนกวรรณ รับชำระ",
            "cashier_position": "เจ้าหน้าที่การเงิน",
        }
        save_path = os.path.join(self.test_dir, "daily_closing.pdf")
        PdfReceiptExporter(save_path).export(closing_data)
        self.assertTrue(os.path.exists(save_path))

    def test_export_batch_pdf(self):
        save_path = os.path.join(self.test_dir, "batch_receipts.pdf")
        exporter = PdfReceiptExporter(save_path)

        receipts = [
            dict(self.base_data, receipt_no="RCP-001"),
            dict(self.base_data, receipt_no="RCP-002", full_name="นายสมศักดิ์ รักเรียน"),
        ]
        ret = exporter.export_batch(receipts)
        self.assertTrue(os.path.exists(ret))

    def test_export_batch_empty_raises_value_error(self):
        exporter = PdfReceiptExporter(os.path.join(self.test_dir, "dummy.pdf"))
        with self.assertRaises(ValueError):
            exporter.export_batch([])
        with self.assertRaises(ValueError):
            exporter.export_batch(None)

    def test_backward_compatible_statement_alias(self):
        save_path = os.path.join(self.test_dir, "alias_test.pdf")
        exporter = PdfStatementExporter(save_path)
        ret = exporter.export(self.base_data)
        self.assertTrue(os.path.exists(ret))

    def test_edge_cases_empty_dict_and_none_values(self):
        save_path = os.path.join(self.test_dir, "empty_pdf.pdf")
        exporter = PdfReceiptExporter(save_path)
        ret = exporter.export({})
        self.assertTrue(os.path.exists(ret))

        save_path_none = os.path.join(self.test_dir, "none_pdf.pdf")
        ret_none = exporter.export({
            "net_total_amount": None,
            "amount": None,
            "items": [
                {"item_name": None, "quantity": None, "unit_price": None, "line_total": None}
            ]
        })
        self.assertTrue(os.path.exists(ret_none))

    def test_many_items_and_long_thai_text(self):
        save_path = os.path.join(self.test_dir, "many_items.pdf")
        exporter = PdfReceiptExporter(save_path)

        many_items = [
            {
                "line_no": i,
                "item_name": f"รายการค่าธรรมเนียมและค่าบำรุงการศึกษาพิเศษประจำปีการศึกษา 2569 ลำดับที่ {i} ชื่อยาวมากเพื่อทดสอบการตัดคำและไม่ให้ทับซ้อนคอลัมน์",
                "quantity": 1,
                "unit_price": 100.0,
                "line_total": 100.0,
            }
            for i in range(1, 18)
        ]

        payload = dict(self.base_data)
        payload["items"] = many_items
        payload["net_total_amount"] = 1700.0
        payload["template_style"] = "MODERN"
        payload["remark"] = "หมายเหตุยาวพิเศษสำหรับการทดสอบความถูกต้องของการจัดวางเลย์เอาต์ท้ายเอกสาร"

        ret = exporter.export(payload)
        self.assertTrue(os.path.exists(ret))
        self.assertGreater(os.path.getsize(ret), 0)


if __name__ == "__main__":
    unittest.main()
