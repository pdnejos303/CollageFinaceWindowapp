import os
import re
import warnings
from typing import Any, Dict, List
from core.config import thai_baht_text
from exporters.base import BaseFinanceExporter

try:
    from fpdf import FPDF
except ImportError:
    FPDF = None


def _setup_thai_fonts(pdf) -> tuple[str, bool]:
    """Registers Windows Tahoma Regular & Bold fonts for full Thai Unicode rendering."""
    font_family = "Helvetica"
    win_fonts_dir = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")
    tahoma_reg = os.path.join(win_fonts_dir, "tahoma.ttf")
    tahoma_bold = os.path.join(win_fonts_dir, "tahomabd.ttf")

    if os.path.exists(tahoma_reg) and os.path.exists(tahoma_bold):
        try:
            pdf.add_font("Tahoma", "", tahoma_reg)
            pdf.add_font("Tahoma", "B", tahoma_bold)
            return "Tahoma", True
        except Exception:
            pass
    return font_family, False


def _clean_signer_name(raw_name: str) -> str:
    """Removes trailing role parentheses like '(หัวหน้างานการเงิน)' so '(ชื่อ)' never has double parentheses."""
    cleaned = re.sub(r"\s*\(.*?\)\s*$", "", str(raw_name or "").strip())
    return cleaned or "เจ้าหน้าที่การเงิน"


class PdfReceiptExporter(BaseFinanceExporter):
    """
    Official Suphanburi Vocational College (SPVC) PDF Receipt,
    Batch Classroom Multi-Page PDF Receipts & Daily Closing PDF Report Exporter (.pdf).
    """

    def export(self, data: Dict[str, Any], metadata: Dict[str, Any] = None) -> str:
        if FPDF is None:
            raise RuntimeError("Module 'fpdf2' is not installed.")

        os.makedirs(os.path.dirname(os.path.abspath(self.save_path)), exist_ok=True)
        meta = metadata or {}

        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            pdf = FPDF(orientation="P", unit="mm", format="A4")
            pdf.set_margins(left=10, top=12, right=10)
            pdf.set_auto_page_break(auto=True, margin=14)
            font_family, _ = _setup_thai_fonts(pdf)

            pdf.add_page()
            self._render_header_banner(pdf, font_family, data, meta)

            if data.get("report_type") == "DAILY_CLOSING":
                self._render_daily_closing(pdf, font_family, data)
            else:
                self._render_receipt_with_items(pdf, font_family, data)

            self._render_signature(pdf, font_family, data, meta)
            pdf.output(self.save_path)

        return self.save_path

    def export_batch(self, receipts_list: List[Dict[str, Any]], metadata: Dict[str, Any] = None) -> str:
        """Export multiple student receipts (e.g. entire classroom) into a single multi-page PDF."""
        if FPDF is None:
            raise RuntimeError("Module 'fpdf2' is not installed.")
        if not receipts_list:
            raise ValueError("ไม่มีข้อมูลใบเสร็จสำหรับพิมพ์")

        os.makedirs(os.path.dirname(os.path.abspath(self.save_path)), exist_ok=True)
        meta = metadata or {}

        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            pdf = FPDF(orientation="P", unit="mm", format="A4")
            pdf.set_margins(left=10, top=12, right=10)
            pdf.set_auto_page_break(auto=True, margin=14)
            font_family, _ = _setup_thai_fonts(pdf)

            for rcp_data in receipts_list:
                pdf.add_page()
                self._render_header_banner(pdf, font_family, rcp_data, meta)
                self._render_receipt_with_items(pdf, font_family, rcp_data)
                self._render_signature(pdf, font_family, rcp_data, meta)

            pdf.output(self.save_path)

        return self.save_path

    def _render_header_banner(self, pdf, font_family: str, data: Dict[str, Any], meta: Dict[str, Any]):
        college = self.college_name or meta.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")
        addr = meta.get("college_address", "455 ถ.พระพันวษา ต.ท่าพี่เลี้ยง อ.เมืองสุพรรณบุรี จ.สุพรรณบุรี 72000")
        tax_id = meta.get("college_tax_id", "0994000159721")

        pdf.set_text_color(15, 23, 42)
        pdf.set_font(font_family, "B", 16)
        pdf.cell(190, 8, college, new_x="LMARGIN", new_y="NEXT", align="C")

        pdf.set_text_color(71, 85, 105)
        pdf.set_font(font_family, "", 9.5)
        pdf.cell(190, 5, addr, new_x="LMARGIN", new_y="NEXT", align="C")
        pdf.cell(
            190,
            5,
            f"สำนักงานคณะกรรมการการอาชีวศึกษา  •  รหัสหน่วยงาน/เลขประจำตัวผู้เสียภาษี: {tax_id}",
            new_x="LMARGIN",
            new_y="NEXT",
            align="C",
        )
        pdf.ln(3)

        doc_title = data.get("doc_title", "ใบเสร็จรับเงินค่าธรรมเนียมการศึกษา / OFFICIAL RECEIPT")
        pdf.set_fill_color(15, 23, 42)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font(font_family, "B", 12)
        pdf.cell(190, 9, doc_title, border=0, new_x="LMARGIN", new_y="NEXT", align="C", fill=True)
        pdf.ln(3)

    def _render_receipt_with_items(self, pdf, font_family: str, data: Dict[str, Any]):
        rcp_no = str(data.get("receipt_no", data.get("tx_code", "-")))
        book_no = str(data.get("book_no", "7006"))
        rcp_date = str(data.get("receipt_date", data.get("tx_date", "-")))[:10]
        term_code = str(data.get("term_code", "1/2569"))
        stu_name = str(data.get("full_name", data.get("snapshot_student_name", "-")))
        stu_code = str(data.get("student_code", "-"))
        cls_grp = str(data.get("class_group", data.get("snapshot_class_group", "-")))
        fac_name = str(data.get("snapshot_faculty", data.get("faculty_name", "")))
        dept = str(data.get("faculty", data.get("snapshot_department", "-")))
        pay_method = str(data.get("payment_method", "เงินสด (Cash)"))
        raw_status = str(data.get("receipt_status", data.get("status", "ISSUED"))).upper()
        status_th = "ปกติ (ชำระแล้ว)" if raw_status in ("ISSUED", "PAID", "COMPLETED") else "ยกเลิกแล้ว (VOID)"

        fac_dept_text = f"คณะ: {fac_name}  |  สาขา: {dept}" if fac_name else f"แผนกวิชา/สาขา: {dept}"

        info_rows = [
            (f"เล่มที่: {book_no}   |   เลขที่ใบเสร็จรับเงิน: {rcp_no}", f"วันที่: {rcp_date}   |   ภาค/ปีการศึกษา: {term_code}", True),
            (f"ได้รับเงินจาก: {stu_name}", f"รหัสประจำตัวนักศึกษา: {stu_code}", True),
            (f"ชั้นปี / กลุ่มเรียน: {cls_grp}", fac_dept_text, False),
            (f"ช่องทางชำระเงิน: {pay_method}", f"สถานะใบเสร็จ: {status_th}", False),
        ]

        pdf.set_draw_color(203, 213, 225)
        pdf.set_fill_color(248, 250, 252)
        pdf.set_text_color(15, 23, 42)

        for left_txt, right_txt, is_bold in info_rows:
            pdf.set_font(font_family, "B" if is_bold else "", 10)
            pdf.cell(95, 7.5, f"  {left_txt}", border=1, fill=True, align="L")
            pdf.cell(95, 7.5, f"  {right_txt}", border=1, new_x="LMARGIN", new_y="NEXT", fill=True, align="L")

        pdf.ln(4)

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

        # Item Table Header
        pdf.set_fill_color(30, 41, 59)
        pdf.set_text_color(255, 255, 255)
        pdf.set_draw_color(203, 213, 225)
        pdf.set_font(font_family, "B", 10)

        col_widths = [16, 94, 24, 26, 30]
        headers = ["ลำดับ", "รายการค่าธรรมเนียม (Description)", "จำนวน", "หน่วยละ", "จำนวนเงิน (บาท)"]
        aligns = ["C", "L", "C", "R", "R"]

        for idx, h_text in enumerate(headers):
            is_last = idx == len(headers) - 1
            pdf.cell(
                col_widths[idx],
                8,
                h_text,
                border=1,
                new_x="LMARGIN" if is_last else "RIGHT",
                new_y="NEXT" if is_last else "TOP",
                align=aligns[idx],
                fill=True,
            )

        # Item Rows
        pdf.set_text_color(15, 23, 42)
        for r_idx, it in enumerate(items, start=1):
            if r_idx % 2 == 1:
                pdf.set_fill_color(255, 255, 255)
            else:
                pdf.set_fill_color(248, 250, 252)

            row_vals = [
                str(it.get("line_no", r_idx)),
                f"  {it.get('item_name', '')}",
                f"{float(it.get('quantity', 1)):,.0f}",
                f"{float(it.get('unit_price', 0)):,.2f} ",
                f"{float(it.get('line_total', 0)):,.2f} ",
            ]
            for c_idx, val_str in enumerate(row_vals):
                is_last = c_idx == len(row_vals) - 1
                pdf.set_font(font_family, "B" if is_last else "", 10)
                pdf.cell(
                    col_widths[c_idx],
                    7.5,
                    val_str,
                    border=1,
                    new_x="LMARGIN" if is_last else "RIGHT",
                    new_y="NEXT" if is_last else "TOP",
                    align=aligns[c_idx],
                    fill=True,
                )

        pdf.ln(3)

        # Net Total Summary Box
        net_val = float(data.get("net_total_amount", data.get("amount", 0.0)))
        baht_txt = data.get("amount_text_th") or thai_baht_text(net_val)

        pdf.set_fill_color(239, 246, 255)
        pdf.set_draw_color(191, 219, 254)
        pdf.set_text_color(30, 58, 138)
        pdf.set_font(font_family, "B", 10.5)
        pdf.cell(115, 9, f"  ตัวอักษร: ({baht_txt})", border=1, fill=True, align="L")
        pdf.set_font(font_family, "B", 12)
        pdf.cell(
            75,
            9,
            f"รวมเป็นเงินสุทธิ: {net_val:,.2f} บาท  ",
            border=1,
            new_x="LMARGIN",
            new_y="NEXT",
            fill=True,
            align="R",
        )

        if data.get("remark"):
            pdf.ln(2)
            pdf.set_text_color(71, 85, 105)
            pdf.set_font(font_family, "", 9.5)
            pdf.cell(190, 6, f"หมายเหตุ: {data.get('remark')}", new_x="LMARGIN", new_y="NEXT")

    def _render_daily_closing(self, pdf, font_family: str, data: Dict[str, Any]):
        inc_val = float(data.get("income_total", 0.0))
        s_rows = [
            ("วันที่ออกใบเสร็จ (Report Date):", str(data.get("date", "-"))),
            ("จำนวนใบเสร็จที่รับชำระทั้งหมด:", f"{data.get('tx_count', 0)} ฉบับ"),
            ("ยอดรวมเงินรับสุทธิประจำวัน:", f"{inc_val:,.2f} บาท ({thai_baht_text(inc_val)})"),
        ]
        pdf.set_draw_color(203, 213, 225)
        for idx, (k, v) in enumerate(s_rows):
            pdf.set_fill_color(241, 245, 249)
            pdf.set_text_color(15, 23, 42)
            pdf.set_font(font_family, "B", 10.5)
            pdf.cell(75, 8, f"  {k}", border=1, fill=True)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_font(font_family, "B" if idx == 2 else "", 10.5)
            pdf.cell(115, 8, f"  {v}", border=1, new_x="LMARGIN", new_y="NEXT", fill=True)

        pdf.ln(5)
        by_method: Dict[str, float] = data.get("income_by_method", {})
        if by_method:
            pdf.set_font(font_family, "B", 11)
            pdf.cell(190, 7, "1. แยกยอดเงินรับตามช่องทางการชำระเงิน:", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font(font_family, "", 10)
            for method_name, amt_val in by_method.items():
                pdf.cell(190, 6.5, f"    - {method_name}:  {float(amt_val):,.2f} บาท", new_x="LMARGIN", new_y="NEXT")

        by_fee: Dict[str, float] = data.get("income_by_fee_item", {})
        if by_fee:
            pdf.ln(3)
            pdf.set_font(font_family, "B", 11)
            pdf.cell(190, 7, "2. แยกยอดเงินรับตามหมวดรายการค่าธรรมเนียมในใบเสร็จ:", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font(font_family, "", 10)
            for fee_name, amt_val in by_fee.items():
                pdf.cell(190, 6.5, f"    - {fee_name}:  {float(amt_val):,.2f} บาท", new_x="LMARGIN", new_y="NEXT")

    def _render_signature(self, pdf, font_family: str, data: Dict[str, Any], meta: Dict[str, Any]):
        """
        Renders an official centered signature block on the right half of the receipt,
        where the signer's name, position, and college name are perfectly centered
        underneath the signature line.
        """
        pdf.ln(12)
        raw_recorder = data.get("recorded_by", data.get("cashier_name", "เจ้าหน้าที่การเงิน"))
        recorder = _clean_signer_name(raw_recorder)
        position = (data.get("cashier_position") or meta.get("cashier_position", "ครู")).strip()
        college = (self.college_name or meta.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")).strip()

        # Right-half signature column: starts at X=105mm, width=95mm
        sig_x = 105.0
        sig_w = 95.0

        prefix_str = "ลงชื่อ "
        dots_str = ".................................................."
        suffix_str = " ผู้รับเงิน"
        full_sig_line = f"{prefix_str}{dots_str}{suffix_str}"

        pdf.set_text_color(15, 23, 42)
        pdf.set_font(font_family, "", 10.5)

        # Calculate the exact horizontal center of the dotted signature line
        w_prefix = pdf.get_string_width(prefix_str)
        w_dots = pdf.get_string_width(dots_str)
        w_total = pdf.get_string_width(full_sig_line)
        line_start_x = sig_x + (sig_w - w_total) / 2.0
        dots_center_x = line_start_x + w_prefix + (w_dots / 2.0)
        sub_box_x = dots_center_x - (sig_w / 2.0)

        pdf.set_x(sig_x)
        pdf.cell(sig_w, 6.5, full_sig_line, new_x="LMARGIN", new_y="NEXT", align="C")

        pdf.set_font(font_family, "B", 10.5)
        pdf.set_x(sub_box_x)
        pdf.cell(sig_w, 6.5, f"({recorder})", new_x="LMARGIN", new_y="NEXT", align="C")

        pdf.set_font(font_family, "", 10)
        pdf.set_x(sub_box_x)
        pdf.cell(sig_w, 6, f"ตำแหน่ง {position}", new_x="LMARGIN", new_y="NEXT", align="C")

        pdf.set_text_color(71, 85, 105)
        pdf.set_font(font_family, "", 9.5)
        pdf.set_x(sub_box_x)
        pdf.cell(sig_w, 6, college, new_x="LMARGIN", new_y="NEXT", align="C")


class PdfStatementExporter(PdfReceiptExporter):
    """Backward-compatible alias for PDF exporting."""
    pass
