import csv
import os
import sys
import tempfile
import unittest
from PIL import Image

sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from core.config import backup_database, export_full_zip_backup, restore_database, thai_baht_text
from core.database import DatabaseManager
from exporters.docx_exporter import DocxReceiptExporter
from exporters.pdf_exporter import PdfReceiptExporter


class TestSPVCSmartReceiptSystem(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, "test_spvc_receipt.db")
        self.db = DatabaseManager(self.db_path)

        self.dummy_slip = os.path.join(self.temp_dir, "slip_test.png")
        img = Image.new("RGB", (300, 400), color="green")
        img.save(self.dummy_slip)

    def test_thai_baht_text_conversion(self):
        self.assertEqual(thai_baht_text(950.0), "เก้าร้อยห้าสิบบาทถ้วน")
        self.assertEqual(thai_baht_text(2500.0), "สองพันห้าร้อยบาทถ้วน")
        self.assertEqual(thai_baht_text(121.50), "หนึ่งร้อยยี่สิบเอ็ดบาทห้าสิบสตางค์")

    def test_dropdown_master_options_and_fee_catalog_from_ex(self):
        curricula = self.db.get_options("CURRICULUM")
        self.assertIn("ค่าวัสดุฝึกสำเร็จรูป", curricula)
        self.assertIn("หลักสูตรในระบบ-ปวช", curricula)
        self.assertIn("หลักสูตรในระบบ-ปวส", curricula)
        self.assertIn("หลักสูตรปริญญาตรี", curricula)
        self.assertIn("หลักสูตรระยะสั้น", curricula)
        self.assertIn("อื่นๆ", curricula)

        # Add & delete custom dropdown option
        self.db.add_option("CLASS_LEVEL", "ปวช.1/9")
        self.assertIn("ปวช.1/9", self.db.get_options("CLASS_LEVEL"))
        self.db.delete_option("CLASS_LEVEL", "ปวช.1/9")
        self.assertNotIn("ปวช.1/9", self.db.get_options("CLASS_LEVEL"))

        # Check EX fee items
        pvc_fees = self.db.get_fee_catalog(curriculum_filter="หลักสูตรในระบบ-ปวช")
        self.assertGreaterEqual(len(pvc_fees), 6)

    def test_flexible_multi_line_receipt_issuing_and_docx(self):
        items = [
            {"fee_code": "SPVC-01", "item_name": "ค่าคู่มือนักเรียนนักศึกษาใหม่", "quantity": 1, "unit_price": 100.0, "discount": 0.0},
            {"fee_code": "SPVC-05", "item_name": "เงินรายได้สถานศึกษา", "quantity": 1, "unit_price": 300.0, "discount": 0.0},
        ]
        rcp = self.db.issue_flexible_receipt(
            student_code="692090101",
            line_items=items,
            payment_method="เงินสด (Cash)",
            actor="finance",
            cashier_name="คุณกนกวรรณ รับชำระค่าเทอม",
            role="FINANCE",
            evidence_path=self.dummy_slip,
            remark="ทดสอบใบเสร็จหลายรายการย่อย",
        )
        self.assertEqual(rcp["net_total_amount"], 400.0)
        self.assertEqual(len(rcp["items"]), 2)

        out_docx = os.path.join(self.temp_dir, "spvc_receipt.docx")
        DocxReceiptExporter(out_docx, college_name="วิทยาลัยอาชีวศึกษาสุพรรณบุรี").export(rcp)
        self.assertTrue(os.path.exists(out_docx))
        self.assertGreater(os.path.getsize(out_docx), 0)

        out_pdf = os.path.join(self.temp_dir, "spvc_receipt.pdf")
        PdfReceiptExporter(out_pdf, college_name="วิทยาลัยอาชีวศึกษาสุพรรณบุรี").export(rcp)
        self.assertTrue(os.path.exists(out_pdf))
        self.assertGreater(os.path.getsize(out_pdf), 0)

    def test_batch_classroom_receipt_issuing_and_batch_docx(self):
        csv_path = os.path.join(self.temp_dir, "room_pvc1_8.csv")
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["รหัสนักศึกษา", "ชื่อ-นามสกุล", "บัตรประชาชน", "คณะ", "สาขา", "ชั้นปี", "สิทธิ", "ค่าเทอม", "โทร", "ระดับ", "ห้อง"])
            writer.writerow(["69901", "นายเอก สุพรรณ", "1729900001", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี", "1", "ปกติ", "950", "", "ปวช.1", "ปวช.1/8"])
            writer.writerow(["69902", "น.ส.โท สุพรรณ", "1729900002", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี", "1", "ปกติ", "950", "", "ปวช.1", "ปวช.1/8"])
            writer.writerow(["69903", "นายตรี สุพรรณ", "1729900003", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี", "1", "ปกติ", "950", "", "ปวช.1", "ปวช.1/8"])

        imp = self.db.bulk_import_students(csv_path, actor="admin", role="ADMIN")
        self.assertEqual(len(imp["student_codes"]), 3)

        students_in_room = self.db.get_students(level_filter="ปวช.1/8")
        self.assertEqual(len(students_in_room), 3)

        items = [
            {"fee_code": "SPVC-05", "item_name": "เงินรายได้สถานศึกษา", "quantity": 1, "unit_price": 300.0, "discount": 0.0}
        ]
        batch_rcps = self.db.issue_batch_class_receipts(
            students_list=students_in_room,
            line_items=items,
            payment_method="เงินสด (Cash)",
            actor="admin",
            cashier_name="ครูการเงิน",
            start_no="700605500",
            end_no="700611000",
        )
        self.assertEqual(len(batch_rcps), 3)
        self.assertEqual(batch_rcps[0]["receipt_no"], "700605500")
        self.assertEqual(batch_rcps[1]["receipt_no"], "700605501")
        self.assertEqual(batch_rcps[2]["receipt_no"], "700605502")

        batch_docx = os.path.join(self.temp_dir, "batch_classroom.docx")
        DocxReceiptExporter(batch_docx, college_name="วิทยาลัยอาชีวศึกษาสุพรรณบุรี").export_batch(batch_rcps)
        self.assertTrue(os.path.exists(batch_docx))
        self.assertGreater(os.path.getsize(batch_docx), 0)

        batch_pdf = os.path.join(self.temp_dir, "batch_classroom.pdf")
        PdfReceiptExporter(batch_pdf, college_name="วิทยาลัยอาชีวศึกษาสุพรรณบุรี").export_batch(batch_rcps)
        self.assertTrue(os.path.exists(batch_pdf))
        self.assertGreater(os.path.getsize(batch_pdf), 0)

    def test_void_receipt_reverses_student_balance(self):
        rcp = self.db.record_tuition_payment(
            student_code="693020201",
            pay_amount=2000.0,
            payment_method="เงินสด (Cash)",
            actor="finance",
            role="FINANCE",
        )
        self.assertEqual(rcp["balance"], 850.0)

        self.db.void_receipt(rcp["receipt_no"], "คีย์ผิดคน", "admin", "ADMIN")
        stus = self.db.get_students(search="693020201")
        self.assertEqual(stus[0]["paid_amount"], 0.0)
        self.assertEqual(stus[0]["balance"], 2850.0)
        self.assertEqual(stus[0]["status"], "UNPAID")

    def test_hot_backup_zip_and_safe_restore(self):
        bk = backup_database(self.db_path, self.temp_dir)
        self.assertTrue(os.path.exists(bk))

        zip_p = export_full_zip_backup(self.db_path, self.temp_dir, self.temp_dir)
        self.assertTrue(os.path.exists(zip_p))

        restored = restore_database(bk, self.db_path)
        self.assertEqual(restored, self.db_path)


if __name__ == "__main__":
    unittest.main()
