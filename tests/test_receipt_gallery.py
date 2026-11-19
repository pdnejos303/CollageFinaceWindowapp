import os
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from core.database import DatabaseManager
from core.receipt_gallery import (
    SUPPORTED_IMAGE_EXTS,
    ReceiptImageRecord,
    copy_receipt_image,
    delete_receipt_image,
    extract_receipt_no_from_filename,
    filter_receipt_images,
    get_gallery_filter_options,
    migrate_receipt_folder,
    move_receipt_image,
    rename_receipt_image,
    scan_receipt_images,
)


def _create_dummy_file(path: str, content: bytes = b"PNGDATA", mtime: float = None) -> str:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as f:
        f.write(content)
    if mtime is not None:
        os.utime(path, (mtime, mtime))
    return path


class TestReceiptGallery(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.root = self.tmpdir.name
        self.db_path = os.path.join(self.root, "test_gallery.db")
        self.db = DatabaseManager(self.db_path)

    def tearDown(self):
        self.tmpdir.cleanup()

    def _insert_receipt(
        self,
        receipt_no: str,
        student_name: str = "สมชาย ใจดี",
        class_group: str = "ปวช.1/1",
        education_level: str = "ปวช.1",
        receipt_date: str = "2026-11-19",
        net_total: float = 1500.0,
        status: str = "ISSUED",
    ) -> None:
        with self.db._connect() as conn:
            conn.execute(
                """
                INSERT INTO receipts (
                    receipt_no, receipt_date, student_code, snapshot_student_name,
                    snapshot_education_level, snapshot_class_group,
                    net_total_amount, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    receipt_no,
                    receipt_date,
                    "692020101",
                    student_name,
                    education_level,
                    class_group,
                    net_total,
                    status,
                ),
            )
            conn.commit()

    def test_supported_exts_constant(self):
        self.assertEqual(
            SUPPORTED_IMAGE_EXTS,
            (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"),
        )

    def test_extract_receipt_no_from_filename(self):
        self.assertEqual(
            extract_receipt_no_from_filename("Receipt_RCP-2569-0001_สมชาย.png"),
            "RCP-2569-0001",
        )
        self.assertEqual(
            extract_receipt_no_from_filename("Reprint_Receipt_7006-0042.jpg"),
            "7006-0042",
        )
        self.assertEqual(
            extract_receipt_no_from_filename("RCP-2569-0099_test.webp"),
            "RCP-2569-0099",
        )
        self.assertEqual(
            extract_receipt_no_from_filename("25690012_somchai.png"),
            "25690012",
        )
        self.assertEqual(extract_receipt_no_from_filename("random_photo.png"), "")
        self.assertEqual(extract_receipt_no_from_filename(""), "")

    def test_scan_receipt_images_nonexistent_and_recursive(self):
        missing = os.path.join(self.root, "does_not_exist")
        self.assertEqual(scan_receipt_images(missing, db=self.db), [])

        gallery = os.path.join(self.root, "gallery")
        _create_dummy_file(os.path.join(gallery, "Receipt_RCP-2569-0001_a.png"), mtime=1000.0)
        _create_dummy_file(os.path.join(gallery, "sub", "Receipt_RCP-2569-0002_b.jpg"), mtime=2000.0)
        _create_dummy_file(os.path.join(gallery, "Receipt_RCP-2569-0003_c.pdf"), mtime=3000.0)

        self._insert_receipt(
            "RCP-2569-0001",
            student_name="สมชาย ใจดี",
            class_group="ปวช.1/1",
            education_level="ปวช.1",
            receipt_date="2026-11-18",
            net_total=1200.0,
            status="ISSUED",
        )
        self._insert_receipt(
            "RCP-2569-0002",
            student_name="สมหญิง รักเรียน",
            class_group="ปวส.2/1",
            education_level="ปวส.2",
            receipt_date="2026-11-19",
            net_total=2500.0,
            status="VOIDED",
        )

        # Recursive scan (default): newest first, ignores .pdf
        recs = scan_receipt_images(gallery, db=self.db, recursive=True)
        self.assertEqual(len(recs), 2)
        self.assertEqual(recs[0].receipt_no, "RCP-2569-0002")
        self.assertEqual(recs[0].student_name, "สมหญิง รักเรียน")
        self.assertEqual(recs[0].class_group, "ปวส.2/1")
        self.assertEqual(recs[0].year_level, "ปวส.2")
        self.assertEqual(recs[0].net_total_amount, 2500.0)
        self.assertEqual(recs[0].status, "VOIDED")
        self.assertEqual(recs[0].rel_path, os.path.join("sub", "Receipt_RCP-2569-0002_b.jpg"))

        self.assertEqual(recs[1].receipt_no, "RCP-2569-0001")
        self.assertEqual(recs[1].student_name, "สมชาย ใจดี")

        # Non-recursive scan: only top-level image
        top_only = scan_receipt_images(gallery, db=self.db, recursive=False)
        self.assertEqual(len(top_only), 1)
        self.assertEqual(top_only[0].receipt_no, "RCP-2569-0001")

    def test_scan_receipt_images_handles_stat_error(self):
        gallery = os.path.join(self.root, "gallery_err")
        _create_dummy_file(os.path.join(gallery, "Receipt_RCP-001.png"))
        with patch("core.receipt_gallery.os.stat", side_effect=PermissionError("denied")):
            self.assertEqual(scan_receipt_images(gallery, db=None), [])

    def test_filter_and_facets(self):
        records = [
            ReceiptImageRecord(
                file_path="/a/1.png",
                filename="Receipt_RCP-001_a.png",
                folder="/a",
                rel_path="sub1/Receipt_RCP-001_a.png",
                ext=".png",
                size_bytes=100,
                modified_ts=100.0,
                modified_date_str="2026-11-18",
                receipt_no="RCP-001",
                student_name="สมชาย ใจดี",
                class_group="ปวช.1/1",
                year_level="ปวช.1",
                receipt_date="2026-11-18",
                net_total_amount=500.0,
                status="ISSUED",
            ),
            ReceiptImageRecord(
                file_path="/a/2.png",
                filename="Receipt_RCP-002_b.png",
                folder="/a",
                rel_path="Receipt_RCP-002_b.png",
                ext=".png",
                size_bytes=100,
                modified_ts=200.0,
                modified_date_str="2026-11-19",
                receipt_no="RCP-002",
                student_name="มานี มีตา",
                class_group="ปวส.1/2",
                year_level="ปวส.1",
                receipt_date="2026-11-19",
                net_total_amount=900.0,
                status="ISSUED",
            ),
        ]

        facets = get_gallery_filter_options(records)
        self.assertEqual(facets["class_groups"], ["ปวช.1/1", "ปวส.1/2"])
        self.assertEqual(facets["year_levels"], ["ปวช.1", "ปวส.1"])

        self.assertEqual(len(filter_receipt_images(records, query="มานี")), 1)
        self.assertEqual(len(filter_receipt_images(records, query="sub1")), 1)
        self.assertEqual(len(filter_receipt_images(records, query="rcp-001")), 1)
        self.assertEqual(len(filter_receipt_images(records, class_group="ปวช.1/1")), 1)
        self.assertEqual(len(filter_receipt_images(records, year_level="ปวส.1")), 1)
        self.assertEqual(len(filter_receipt_images(records, date_str="2026-11-19")), 1)
        self.assertEqual(len(filter_receipt_images(records, date_str="2026-11")), 2)

    def test_copy_receipt_image(self):
        src = _create_dummy_file(os.path.join(self.root, "src", "Receipt_1.png"), b"ONE")
        dest_dir = os.path.join(self.root, "dest")
        dest1 = copy_receipt_image(src, os.path.join(dest_dir, "Receipt_1.png"), overwrite=False)
        self.assertTrue(os.path.isfile(dest1))

        # Collision without overwrite -> auto-renames
        src2 = _create_dummy_file(os.path.join(self.root, "src", "Receipt_1_new.png"), b"TWO")
        dest2 = copy_receipt_image(src2, os.path.join(dest_dir, "Receipt_1.png"), overwrite=False)
        self.assertTrue(dest2.endswith("Receipt_1_1.png"))
        with open(dest2, "rb") as f:
            self.assertEqual(f.read(), b"TWO")

        # Collision with overwrite -> replaces
        dest3 = copy_receipt_image(src2, os.path.join(dest_dir, "Receipt_1.png"), overwrite=True)
        self.assertEqual(os.path.realpath(dest3), os.path.realpath(dest1))
        with open(dest1, "rb") as f:
            self.assertEqual(f.read(), b"TWO")

        # Same-file copy with overwrite=True is a safe no-op
        self.assertEqual(copy_receipt_image(src, src, overwrite=True), os.path.realpath(src))

        with self.assertRaises(FileNotFoundError):
            copy_receipt_image(os.path.join(self.root, "missing.png"), dest1)

    def test_move_receipt_image_and_traversal_guard(self):
        allowed = os.path.join(self.root, "allowed")
        outside = os.path.join(self.root, "outside")
        os.makedirs(outside, exist_ok=True)

        src = _create_dummy_file(os.path.join(allowed, "a", "Receipt_M1.png"), b"MOVE")
        dest_dir = os.path.join(allowed, "b")

        moved = move_receipt_image(src, dest_dir, allowed_root=allowed)
        self.assertTrue(os.path.isfile(moved))
        self.assertFalse(os.path.exists(src))

        # Moving to same dir is a safe no-op
        self.assertEqual(move_receipt_image(moved, dest_dir, allowed_root=allowed), moved)

        # Path traversal blocked
        with self.assertRaises(ValueError):
            move_receipt_image(moved, outside, allowed_root=allowed)
        with self.assertRaises(ValueError):
            move_receipt_image(moved, os.path.join(allowed, "..", "outside"), allowed_root=allowed)

    def test_rename_receipt_image(self):
        allowed = os.path.join(self.root, "rename_root")
        src = _create_dummy_file(os.path.join(allowed, "old_name.png"), b"IMG")

        renamed = rename_receipt_image(src, "new/stem:name*", allowed_root=allowed)
        self.assertTrue(os.path.isfile(renamed))
        self.assertFalse(os.path.exists(src))
        self.assertEqual(os.path.basename(renamed), "new_stem_name_.png")

        with self.assertRaises(ValueError):
            rename_receipt_image(renamed, "///...", allowed_root=allowed)

        outside_file = _create_dummy_file(os.path.join(self.root, "outside.png"))
        with self.assertRaises(ValueError):
            rename_receipt_image(outside_file, "hacked", allowed_root=allowed)

    def test_delete_receipt_image(self):
        allowed = os.path.join(self.root, "del_root")
        img = _create_dummy_file(os.path.join(allowed, "Receipt_D1.png"))
        txt = _create_dummy_file(os.path.join(allowed, "notes.txt"), b"hello")
        outside = _create_dummy_file(os.path.join(self.root, "outside_del.png"))

        self.assertTrue(delete_receipt_image(img, allowed_root=allowed))
        self.assertFalse(os.path.exists(img))
        self.assertFalse(delete_receipt_image(img, allowed_root=allowed))

        with self.assertRaises(ValueError):
            delete_receipt_image(txt, allowed_root=allowed)

        with self.assertRaises(ValueError):
            delete_receipt_image(outside, allowed_root=allowed)

    def test_migrate_receipt_folder(self):
        src_dir = os.path.join(self.root, "old_exports")
        dest_dir = os.path.join(self.root, "new_exports")
        _create_dummy_file(os.path.join(src_dir, "Receipt_1.png"), b"1")
        _create_dummy_file(os.path.join(src_dir, "Receipt_2.jpg"), b"2")
        _create_dummy_file(os.path.join(src_dir, "ignore.pdf"), b"pdf")

        # Same-dir skipped
        same = migrate_receipt_folder(src_dir, src_dir, move_files=False)
        self.assertEqual(same, {"processed": 0, "skipped": 0, "errors": []})

        # Copy mode
        res_copy = migrate_receipt_folder(src_dir, dest_dir, move_files=False)
        self.assertEqual(res_copy["processed"], 2)
        self.assertEqual(res_copy["skipped"], 1)
        self.assertEqual(res_copy["errors"], [])
        self.assertTrue(os.path.isfile(os.path.join(src_dir, "Receipt_1.png")))
        self.assertTrue(os.path.isfile(os.path.join(dest_dir, "Receipt_1.png")))

        # Move mode into third folder
        dest_dir2 = os.path.join(self.root, "final_exports")
        res_move = migrate_receipt_folder(src_dir, dest_dir2, move_files=True)
        self.assertEqual(res_move["processed"], 2)
        self.assertEqual(res_move["skipped"], 1)
        self.assertEqual(res_move["errors"], [])
        self.assertFalse(os.path.exists(os.path.join(src_dir, "Receipt_1.png")))
        self.assertTrue(os.path.isfile(os.path.join(dest_dir2, "Receipt_1.png")))


if __name__ == "__main__":
    unittest.main()
