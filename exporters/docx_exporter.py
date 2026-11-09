import os
from typing import Any, Dict, List
from core.config import thai_baht_text
from exporters.base import BaseFinanceExporter

try:
    import docx
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except ImportError:
    docx = None


def set_cell_background(cell, fill_hex: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex.replace("#", ""))
    tcPr.append(shd)


class DocxReceiptExporter(BaseFinanceExporter):
    """Official Suphanburi Vocational College (SPVC) Receipt, Batch Classroom Receipts & Daily Report Exporter (.docx)."""

    def export(self, data: Dict[str, Any], metadata: Dict[str, Any] = None) -> str:
        if docx is None:
            raise RuntimeError("Module 'python-docx' is not installed.")

        os.makedirs(os.path.dirname(os.path.abspath(self.save_path)), exist_ok=True)
        meta = metadata or {}
        doc = Document()

        self._render_header_banner(doc, data, meta)

        if data.get("report_type") == "DAILY_CLOSING":
            self._render_daily_closing(doc, data)
        else:
            self._render_receipt_with_items(doc, data)

        self._render_signature(doc, data, meta)
        doc.save(self.save_path)
        return self.save_path

    def export_batch(self, receipts_list: List[Dict[str, Any]], metadata: Dict[str, Any] = None) -> str:
        """Export multiple student receipts (e.g. entire classroom) into a single combined multi-page .docx file."""
        if docx is None:
            raise RuntimeError("Module 'python-docx' is not installed.")
        if not receipts_list:
            raise ValueError("ไม่มีข้อมูลใบเสร็จสำหรับพิมพ์")

        os.makedirs(os.path.dirname(os.path.abspath(self.save_path)), exist_ok=True)
        meta = metadata or {}
        doc = Document()

        for idx, rcp_data in enumerate(receipts_list):
            if idx > 0:
                doc.add_page_break()
            self._render_header_banner(doc, rcp_data, meta)
            self._render_receipt_with_items(doc, rcp_data)
            self._render_signature(doc, rcp_data, meta)

        doc.save(self.save_path)
        return self.save_path

    def _render_header_banner(self, doc, data: Dict[str, Any], meta: Dict[str, Any]):
        h_para = doc.add_paragraph()
        h_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        h_run = h_para.add_run(self.college_name or meta.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"))
        h_run.font.name = "Tahoma"
        h_run.font.size = Pt(16)
        h_run.font.bold = True
        h_run.font.color.rgb = RGBColor(15, 23, 42)

        sub_para = doc.add_paragraph()
        sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        addr = meta.get("college_address", "455 ถ.พระพันวษา ต.ท่าพี่เลี้ยง อ.เมืองสุพรรณบุรี จ.สุพรรณบุรี 72000")
        tax_id = meta.get("college_tax_id", "0994000159721")
        sub_run = sub_para.add_run(f"{addr}\nสำนักงานคณะกรรมการการอาชีวศึกษา • รหัสหน่วยงาน/เลขประจำตัวผู้เสียภาษี: {tax_id}")
        sub_run.font.name = "Tahoma"
        sub_run.font.size = Pt(9.5)
        sub_run.font.color.rgb = RGBColor(71, 85, 105)

        doc_title = data.get("doc_title", "ใบเสร็จรับเงินค่าธรรมเนียมการศึกษา / OFFICIAL RECEIPT")
        t_table = doc.add_table(rows=1, cols=1)
        t_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_cell = t_table.cell(0, 0)
        t_cell.width = Inches(6.4)
        set_cell_background(t_cell, "0F172A")
        tp = t_cell.paragraphs[0]
        tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tr = tp.add_run(doc_title)
        tr.font.name = "Tahoma"
        tr.font.size = Pt(12.5)
        tr.font.bold = True
        tr.font.color.rgb = RGBColor(255, 255, 255)

        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    def _render_signature(self, doc, data: Dict[str, Any], meta: Dict[str, Any]):
        import re
        raw_recorder = data.get("recorded_by", data.get("cashier_name", "เจ้าหน้าที่การเงิน"))
        recorder = re.sub(r"\s*\(.*?\)\s*$", "", str(raw_recorder or "").strip()) or "เจ้าหน้าที่การเงิน"
        position = (data.get("cashier_position") or meta.get("cashier_position", "ครู")).strip()
        college = (self.college_name or meta.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")).strip()

        doc.add_paragraph().paragraph_format.space_after = Pt(8)
        sig_table = doc.add_table(rows=1, cols=2)
        sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        sig_cell = sig_table.cell(0, 1)
        sig_p = sig_cell.paragraphs[0]
        sig_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sig_run = sig_p.add_run(
            f"ลงชื่อ .................................................. ผู้รับเงิน\n"
            f"({recorder})\n"
            f"ตำแหน่ง {position}\n"
            f"{college}"
        )
        sig_run.font.name = "Tahoma"
        sig_run.font.size = Pt(10)
        sig_run.font.color.rgb = RGBColor(15, 23, 42)

    def _render_receipt_with_items(self, doc, data: Dict[str, Any]):
        info_table = doc.add_table(rows=4, cols=2)
        info_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        rcp_no = str(data.get("receipt_no", data.get("tx_code", "-")))
        book_no = str(data.get("book_no", "7006"))
        rcp_date = str(data.get("receipt_date", data.get("tx_date", "-")))
        term_code = str(data.get("term_code", "1/2569"))
        stu_name = str(data.get("full_name", data.get("snapshot_student_name", "-")))
        stu_code = str(data.get("student_code", "-"))
        cls_grp = str(data.get("class_group", data.get("snapshot_class_group", "-")))
        fac_name = str(data.get("snapshot_faculty", data.get("faculty_name", "")))
        dept = str(data.get("faculty", data.get("snapshot_department", "-")))
        pay_method = str(data.get("payment_method", "เงินสด (Cash)"))

        fac_dept_text = f"คณะ: {fac_name} | แผนกวิชา: {dept}" if fac_name else f"แผนกวิชา: {dept}"

        rows_data = [
            (f"เล่มที่: {book_no}   |   เลขที่ใบเสร็จรับเงิน: {rcp_no}", f"วันที่: {rcp_date}   |   ภาค/ปีการศึกษา: {term_code}"),
            (f"ได้รับเงินจาก: {stu_name}", f"รหัสประจำตัวนักศึกษา: {stu_code}"),
            (f"ชั้นปี / กลุ่มเรียน: {cls_grp}", fac_dept_text),
            (f"ช่องทางชำระเงิน: {pay_method}", f"สถานะใบเสร็จ: {data.get('receipt_status', data.get('status', 'ISSUED'))}"),
        ]

        for idx, (col_left, col_right) in enumerate(rows_data):
            c0 = info_table.cell(idx, 0)
            c1 = info_table.cell(idx, 1)
            set_cell_background(c0, "F8FAFC")
            set_cell_background(c1, "F8FAFC")
            r0 = c0.paragraphs[0].add_run(col_left)
            r0.font.name = "Tahoma"
            r0.font.size = Pt(9.5)
            r0.font.bold = idx in (0, 1)
            r1 = c1.paragraphs[0].add_run(col_right)
            r1.font.name = "Tahoma"
            r1.font.size = Pt(9.5)
            r1.font.bold = idx in (0, 1)

        doc.add_paragraph().paragraph_format.space_after = Pt(6)

        items: List[Dict[str, Any]] = data.get("items", [])
        if not items:
            amt_v = float(data.get("net_total_amount", data.get("amount", 0.0)))
            items = [
                {
                    "line_no": 1,
                    "item_name": data.get("title", "ชำระค่าธรรมเนียมการศึกษา"),
                    "quantity": 1,
                    "unit_price": amt_v,
                    "discount": 0.0,
                    "line_total": amt_v,
                }
            ]

        item_table = doc.add_table(rows=len(items) + 1, cols=5)
        item_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        headers = ["ลำดับ", "รายการ (Description)", "หน่วยกิต/จำนวน", "หน่วยละ", "จำนวนเงิน (บาท)"]
        for c_idx, h_text in enumerate(headers):
            cell = item_table.cell(0, c_idx)
            set_cell_background(cell, "1E293B")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 2 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(h_text)
            r.font.name = "Tahoma"
            r.font.size = Pt(9.5)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

        for r_idx, it in enumerate(items, start=1):
            row_cells = item_table.rows[r_idx].cells
            bg = "FFFFFF" if r_idx % 2 == 1 else "F8FAFC"
            vals = [
                str(it.get("line_no", r_idx)),
                str(it.get("item_name", "")),
                f"{float(it.get('quantity', 1)):,.0f}",
                f"{float(it.get('unit_price', 0)):,.2f}",
                f"{float(it.get('line_total', 0)):,.2f}",
            ]
            for c_idx, v_str in enumerate(vals):
                set_cell_background(row_cells[c_idx], bg)
                p = row_cells[c_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 2 else WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(v_str)
                r.font.name = "Tahoma"
                r.font.size = Pt(9.5)
                if c_idx == 4:
                    r.font.bold = True

        doc.add_paragraph().paragraph_format.space_after = Pt(6)

        net_val = float(data.get("net_total_amount", data.get("amount", 0.0)))
        baht_txt = data.get("amount_text_th") or thai_baht_text(net_val)

        amt_table = doc.add_table(rows=1, cols=2)
        amt_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        ac0, ac1 = amt_table.cell(0, 0), amt_table.cell(0, 1)
        set_cell_background(ac0, "DCFCE7")
        set_cell_background(ac1, "DCFCE7")

        ap0 = ac0.paragraphs[0]
        ar0 = ap0.add_run(f"ตัวอักษร: ({baht_txt})")
        ar0.font.name = "Tahoma"
        ar0.font.size = Pt(10.5)
        ar0.font.bold = True
        ar0.font.color.rgb = RGBColor(22, 101, 52)

        ap1 = ac1.paragraphs[0]
        ap1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        ar1 = ap1.add_run(f"รวมเป็นเงิน: ฿{net_val:,.2f} บาท")
        ar1.font.name = "Tahoma"
        ar1.font.size = Pt(12.5)
        ar1.font.bold = True
        ar1.font.color.rgb = RGBColor(21, 128, 61)

        if data.get("remark"):
            rem_p = doc.add_paragraph()
            rem_r = rem_p.add_run(f"หมายเหตุ: {data.get('remark')}")
            rem_r.font.name = "Tahoma"
            rem_r.font.size = Pt(9.5)

    def _render_daily_closing(self, doc, data: Dict[str, Any]):
        summary_table = doc.add_table(rows=3, cols=2)
        summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        inc_val = float(data.get("income_total", 0.0))
        s_rows = [
            ("วันที่ออกใบเสร็จ (Report Date):", str(data.get("date", "-"))),
            ("จำนวนใบเสร็จที่รับชำระทั้งหมด:", f"{data.get('tx_count', 0)} ฉบับ"),
            ("ยอดรวมเงินรับสุทธิประจำวัน:", f"฿{inc_val:,.2f} บาท ({thai_baht_text(inc_val)})"),
        ]
        for idx, (k, v) in enumerate(s_rows):
            c0, c1 = summary_table.cell(idx, 0), summary_table.cell(idx, 1)
            set_cell_background(c0, "F1F5F9")
            set_cell_background(c1, "FFFFFF")
            r0 = c0.paragraphs[0].add_run(k)
            r0.font.name = "Tahoma"
            r0.font.size = Pt(10)
            r0.font.bold = True
            r1 = c1.paragraphs[0].add_run(v)
            r1.font.name = "Tahoma"
            r1.font.size = Pt(10)
            r1.font.bold = idx == 2

        doc.add_paragraph().paragraph_format.space_after = Pt(6)

        by_method: Dict[str, float] = data.get("income_by_method", {})
        if by_method:
            mp = doc.add_paragraph()
            mr = mp.add_run("1. แยกยอดเงินรับตามช่องทางการชำระเงิน:")
            mr.font.name = "Tahoma"
            mr.font.size = Pt(10.5)
            mr.font.bold = True
            for method_name, amt_val in by_method.items():
                bp = doc.add_paragraph()
                bp.paragraph_format.space_before = Pt(1)
                bp.paragraph_format.space_after = Pt(1)
                br = bp.add_run(f"   • {method_name}:  ฿{float(amt_val):,.2f} บาท")
                br.font.name = "Tahoma"
                br.font.size = Pt(10)

        by_fee: Dict[str, float] = data.get("income_by_fee_item", {})
        if by_fee:
            fp = doc.add_paragraph()
            fp.paragraph_format.space_before = Pt(6)
            fr = fp.add_run("2. แยกยอดเงินรับตามหมวดรายการค่าธรรมเนียมในใบเสร็จ:")
            fr.font.name = "Tahoma"
            fr.font.size = Pt(10.5)
            fr.font.bold = True
            for fee_name, amt_val in by_fee.items():
                bp = doc.add_paragraph()
                bp.paragraph_format.space_before = Pt(1)
                bp.paragraph_format.space_after = Pt(1)
                br = bp.add_run(f"   • {fee_name}:  ฿{float(amt_val):,.2f} บาท")
                br.font.name = "Tahoma"
                br.font.size = Pt(10)
