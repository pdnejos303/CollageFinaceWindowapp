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

    def test_high_volume_pagination_and_no_n_plus_1_queries(self):
        csv_path = os.path.join(self.temp_dir, "bulk_250.csv")
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["รหัสนักศึกษา", "ชื่อ-นามสกุล", "บัตรประชาชน", "คณะ", "สาขา", "ชั้นปี", "สิทธิ", "ค่าเทอม", "โทร", "ระดับ", "ห้อง"])
            for i in range(1, 251):
                writer.writerow([f"6988{i:04d}", f"นักศึกษาคนที่ {i}", f"17299{i:08d}", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี", "1", "ปกติ", "950", "", "ปวช.1", "ปวช.1/99"])

        imp = self.db.bulk_import_students(csv_path, actor="admin", role="ADMIN")
        self.assertEqual(imp["total"], 250)

        total_in_room = self.db.count_students(level_filter="ปวช.1/99")
        self.assertEqual(total_in_room, 250)

        page1 = self.db.get_students(level_filter="ปวช.1/99", limit=100, offset=0)
        page2 = self.db.get_students(level_filter="ปวช.1/99", limit=100, offset=100)
        page3 = self.db.get_students(level_filter="ปวช.1/99", limit=100, offset=200)
        self.assertEqual(len(page1), 100)
        self.assertEqual(len(page2), 100)
        self.assertEqual(len(page3), 50)

        stu_single = self.db.get_student_by_code("69880125")
        self.assertIsNotNone(stu_single)
        self.assertEqual(stu_single["full_name"], "นักศึกษาคนที่ 125")

        by_codes = self.db.get_students_by_codes(imp["student_codes"])
        self.assertEqual(len(by_codes), 250)

        items = [
            {"fee_code": "SPVC-05", "item_name": "เงินรายได้สถานศึกษา", "quantity": 1, "unit_price": 300.0, "discount": 0.0}
        ]
        batch_rcps = self.db.issue_batch_class_receipts(
            students_list=by_codes,
            line_items=items,
            payment_method="เงินสด (Cash)",
            actor="admin",
            cashier_name="ครูการเงิน",
            start_no="700606000",
            end_no="700999999",
        )
        self.assertEqual(len(batch_rcps), 250)
        self.assertEqual(batch_rcps[0]["receipt_no"], "700606000")
        self.assertEqual(batch_rcps[-1]["receipt_no"], "700606249")
        self.assertEqual(self.db.get_next_receipt_no(start_no="700605433", end_no="700999999"), "700606250")

        stats = self.db.get_receipts_stats(search="นักศึกษาคนที่")
        self.assertEqual(stats["total_count"], 250)
        self.assertEqual(stats["valid_count"], 250)
        self.assertEqual(stats["total_amount"], 250 * 300.0)

        rcp_page1 = self.db.get_receipts(search="นักศึกษาคนที่", limit=100, offset=0)
        self.assertEqual(len(rcp_page1), 100)
        self.assertEqual(len(rcp_page1[0]["items"]), 1)

        rcp_exact = self.db.get_receipt_by_no("700606249")
        self.assertIsNotNone(rcp_exact)
        self.assertEqual(rcp_exact["student_code"], "69880250")

    def test_course_registration_card_ost28_export(self):
        catalog = self.db.get_course_catalog()
        self.assertGreaterEqual(len(catalog), 11)

        courses = self.db.get_student_registered_courses("69219100049", "1/2569")
        self.assertEqual(len(courses), 11)
        self.assertEqual(sum(int(c["theory_hours"]) for c in courses), 10)
        self.assertEqual(sum(int(c["practice_hours"]) for c in courses), 20)
        self.assertEqual(sum(int(c["credits"]) for c in courses), 19)

        stu = self.db.get_student_by_code("69219100049")
        rcp = self.db.get_receipt_by_no("700605308")
        reg_payload = {
            "report_type": "COURSE_REGISTRATION",
            "student": stu,
            "academic_year": "1/2569",
            "courses": courses,
            "fee_items": rcp["items"],
            "total_amount": rcp["net_total_amount"],
            "baht_text": rcp["amount_text_th"],
        }
        out_pdf = os.path.join(self.temp_dir, "ost28.pdf")
        out_docx = os.path.join(self.temp_dir, "ost28.docx")
        PdfReceiptExporter(out_pdf).export(reg_payload)
        DocxReceiptExporter(out_docx).export(reg_payload)
        self.assertTrue(os.path.exists(out_pdf))
        self.assertTrue(os.path.exists(out_docx))


if __name__ == "__main__":
    unittest.main()
