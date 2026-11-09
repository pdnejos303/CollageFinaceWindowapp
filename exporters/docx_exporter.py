import os
import re
from typing import Any, Dict, List
from core.config import (
    format_receipt_date_thai,
    get_garuda_emblem_path,
    load_config,
    normalize_template_style,
    thai_baht_text,
)
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

DOCX_FONT_NAME = "TH Sarabun New"


def set_thai_font(run, size_pt: float, bold: bool = False, color_rgb: tuple = None):
    run.font.name = DOCX_FONT_NAME
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), DOCX_FONT_NAME)
    rFonts.set(qn("w:hAnsi"), DOCX_FONT_NAME)
    rFonts.set(qn("w:cs"), DOCX_FONT_NAME)
    rPr.append(rFonts)


def set_cell_background(cell, fill_hex: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex.replace("#", ""))
    tcPr.append(shd)


def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge, spec in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{edge}")
        if spec:
            el.set(qn("w:val"), spec.get("val", "single"))
            el.set(qn("w:sz"), str(spec.get("sz", 6)))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), spec.get("color", "000000"))
        else:
            el.set(qn("w:val"), "none")
        tcBorders.append(el)
    tcPr.append(tcBorders)


def _resolve_faculty_and_major(data: Dict[str, Any]) -> tuple[str, str]:
    faculty = (
        data.get("faculty")
        or data.get("snapshot_department")
        or data.get("snapshot_faculty")
        or "พาณิชยกรรม/บริหารธุรกิจ"
    )
    major = (
        data.get("major")
        or data.get("snapshot_major")
        or "เทคโนโลยีธุรกิจดิจิทัล"
    )
    return str(faculty).strip() or "-", str(major).strip() or "-"


class DocxReceiptExporter(BaseFinanceExporter):
    """Official Suphanburi Vocational College (SPVC) Receipt, Course Registration Card (อศท.28) & Daily Report Exporter (.docx)."""

    def export(self, data: Dict[str, Any], metadata: Dict[str, Any] = None) -> str:
        if docx is None:
            raise RuntimeError("Module 'python-docx' is not installed.")

        os.makedirs(os.path.dirname(os.path.abspath(self.save_path)), exist_ok=True)
        meta = metadata or {}
        doc = Document()
        self._configure_page_margins(doc)

        report_type = data.get("report_type", "RECEIPT")
        raw_style = data.get("template_style") or meta.get("bill_template_style") or load_config().get("bill_template_style", "CLASSIC")
        style = normalize_template_style(raw_style)

        if report_type == "DAILY_CLOSING":
            self._render_daily_report_header(doc, data, meta)
            self._render_daily_closing(doc, data)
            self._render_daily_signature(doc, data, meta)
        elif report_type == "COURSE_REGISTRATION":
            if style == "MODERN":
                self._render_course_registration_card_modern(doc, data, meta)
            else:
                self._render_course_registration_card(doc, data, meta)
        else:
            if style == "MODERN":
                self._render_official_ex_bill_modern(doc, data, meta)
            else:
                self._render_official_ex_bill(doc, data, meta)

        doc.save(self.save_path)
        return self.save_path

    def export_batch(self, receipts_list: List[Dict[str, Any]], metadata: Dict[str, Any] = None) -> str:
        if docx is None:
            raise RuntimeError("Module 'python-docx' is not installed.")
        if not receipts_list:
            raise ValueError("ไม่มีข้อมูลใบเสร็จสำหรับพิมพ์")

        os.makedirs(os.path.dirname(os.path.abspath(self.save_path)), exist_ok=True)
        meta = metadata or {}
        doc = Document()
        self._configure_page_margins(doc)

        for idx, rcp_data in enumerate(receipts_list):
            if idx > 0:
                doc.add_page_break()
            raw_style = rcp_data.get("template_style") or meta.get("bill_template_style") or load_config().get("bill_template_style", "CLASSIC")
            style = normalize_template_style(raw_style)
            if rcp_data.get("report_type") == "COURSE_REGISTRATION":
                if style == "MODERN":
                    self._render_course_registration_card_modern(doc, rcp_data, meta)
                else:
                    self._render_course_registration_card(doc, rcp_data, meta)
            else:
                if style == "MODERN":
                    self._render_official_ex_bill_modern(doc, rcp_data, meta)
                else:
                    self._render_official_ex_bill(doc, rcp_data, meta)

        doc.save(self.save_path)
        return self.save_path

    def _configure_page_margins(self, doc):
        for section in doc.sections:
            section.top_margin = Inches(0.45)
            section.bottom_margin = Inches(0.4)
            section.left_margin = Inches(0.6)
            section.right_margin = Inches(0.6)

    # ==========================================================================
    # 1. OFFICIAL RECEIPT MATCHING EXofBill (95b999ca-2a23-4412-bbee-a3b29b8aa4e6.jpeg)
    # ==========================================================================
    def _render_official_ex_bill(self, doc, data: Dict[str, Any], meta: Dict[str, Any]):
        college = (self.college_name or meta.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")).strip()
        rcp_no = str(data.get("receipt_no", data.get("tx_code", "-"))).strip()
        raw_date = data.get("receipt_date", data.get("tx_date", ""))
        rcp_date_th = format_receipt_date_thai(raw_date)
        term_code = str(data.get("term_code", meta.get("current_semester", "2/2569"))).strip()
        stu_name = str(data.get("full_name", data.get("snapshot_student_name", "-"))).strip()
        stu_code = str(data.get("student_code", "-")).strip()
        cls_grp = str(data.get("class_group", data.get("snapshot_class_group", data.get("education_level", "ปวช.1")))).strip()
        faculty_name, major_name = _resolve_faculty_and_major(data)
        position = (data.get("cashier_position") or meta.get("cashier_position", "ครู")).strip()
        garuda_path = get_garuda_emblem_path()

        # 1. Top Header 3-Column Table (Left Info | Center Garuda | Right Info)
        hdr_table = doc.add_table(rows=1, cols=3)
        hdr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        c_left, c_mid, c_right = hdr_table.rows[0].cells
        c_left.width = Inches(2.9)
        c_mid.width = Inches(1.2)
        c_right.width = Inches(2.9)
        for c in (c_left, c_mid, c_right):
            set_cell_borders(c)

        lp = c_left.paragraphs[0]
        lp.paragraph_format.space_after = Pt(2)
        r_title = lp.add_run("ใบเสร็จรับเงิน\n")
        set_thai_font(r_title, 18, bold=True)

        r_left_body = lp.add_run(
            f"ในราชการกรมอาชีวศึกษา\n"
            f"กระทรวงศึกษาธิการ\n"
            f"ได้รับเงินจาก     {stu_name}\n"
            f"ภาค/ปีการศึกษา     {term_code}\n"
            f"แผนกวิชา     {major_name}"
        )
        set_thai_font(r_left_body, 15, bold=False)

        mp = c_mid.paragraphs[0]
        mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if garuda_path and os.path.exists(garuda_path):
            try:
                mp.add_run().add_picture(garuda_path, width=Inches(0.95))
            except Exception:
                pass

        rp = c_right.paragraphs[0]
        rp.paragraph_format.space_after = Pt(2)
        r_right_body = rp.add_run(
            f"เลขที่     {rcp_no}\n"
            f"ที่ทำการ     {college}\n"
            f"วันที่     {rcp_date_th}\n"
            f"รหัสประจำตัว     {stu_code}\n"
            f"คณะ     {faculty_name}\n"
            f"ชั้น     {cls_grp}"
        )
        set_thai_font(r_right_body, 15, bold=False)

        doc.add_paragraph().paragraph_format.space_after = Pt(2)

        # 2. Items Table (รายการ | จำนวนเงิน) matching EXofBill
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

        min_blank_rows = max(0, 14 - len(items))
        total_rows = 1 + len(items) + min_blank_rows + 1
        item_table = doc.add_table(rows=total_rows, cols=2)
        item_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        line_spec = {"val": "single", "sz": 6, "color": "000000"}

        hc0, hc1 = item_table.rows[0].cells
        hc0.width = Inches(5.3)
        hc1.width = Inches(1.7)
        set_cell_borders(hc0, top=line_spec, bottom=line_spec)
        set_cell_borders(hc1, top=line_spec, bottom=line_spec)

        hp0 = hc0.paragraphs[0]
        hp0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hr0 = hp0.add_run("รายการ")
        set_thai_font(hr0, 15.5, bold=True)

        hp1 = hc1.paragraphs[0]
        hp1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hr1 = hp1.add_run("จำนวนเงิน")
        set_thai_font(hr1, 15.5, bold=True)

        for r_idx, it in enumerate(items, start=1):
            c0, c1 = item_table.rows[r_idx].cells
            c0.width = Inches(5.3)
            c1.width = Inches(1.7)
            set_cell_borders(c0)
            set_cell_borders(c1)

            iname = str(it.get("item_name", "")).strip()
            qty = float(it.get("quantity", 1.0))
            uprice = float(it.get("unit_price", 0.0))
            ltotal = float(it.get("line_total", 0.0))
            iname_display = f"{iname} ({qty:,.0f} หน่วย x {uprice:,.2f})" if qty > 1.0 else iname

            p0 = c0.paragraphs[0]
            p0.paragraph_format.space_before = Pt(2)
            p0.paragraph_format.space_after = Pt(2)
            r0 = p0.add_run(iname_display)
            set_thai_font(r0, 15, bold=False)

            p1 = c1.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p1.paragraph_format.space_before = Pt(2)
            p1.paragraph_format.space_after = Pt(2)
            r1 = p1.add_run(f"{ltotal:,.2f}")
            set_thai_font(r1, 15, bold=False)

        for b_idx in range(1 + len(items), total_rows - 1):
            bc0, bc1 = item_table.rows[b_idx].cells
            set_cell_borders(bc0)
            set_cell_borders(bc1)
            bp = bc0.paragraphs[0]
            r_blank = bp.add_run(" ")
            set_thai_font(r_blank, 14)
            bp.paragraph_format.space_after = Pt(3)

        # Total Row
        net_val = float(data.get("net_total_amount", data.get("amount", 0.0)))
        baht_txt = data.get("amount_text_th") or thai_baht_text(net_val)

        tc0, tc1 = item_table.rows[total_rows - 1].cells
        tc0.width = Inches(5.3)
        tc1.width = Inches(1.7)
        set_cell_borders(tc0, top=line_spec, bottom=line_spec)
        set_cell_borders(tc1, top=line_spec, bottom=line_spec)

        tp0 = tc0.paragraphs[0]
        tp0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tr0 = tp0.add_run(baht_txt)
        set_thai_font(tr0, 15, bold=False)

        tp1 = tc1.paragraphs[0]
        tp1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        tr1 = tp1.add_run(f"รวม        {net_val:,.2f}")
        set_thai_font(tr1, 15, bold=False)

        doc.add_paragraph().paragraph_format.space_after = Pt(4)

        # 3. Bottom Verification & Signature Table matching EXofBill
        ftr_table = doc.add_table(rows=1, cols=2)
        ftr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        fc0, fc1 = ftr_table.rows[0].cells
        fc0.width = Inches(3.8)
        fc1.width = Inches(3.2)
        set_cell_borders(fc0)
        set_cell_borders(fc1)

        fp0 = fc0.paragraphs[0]
        fp0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fr0 = fp0.add_run("(ตัวอักษร)\nไว้เป็นการถูกต้องแล้ว")
        set_thai_font(fr0, 15, bold=False)

        fp1 = fc1.paragraphs[0]
        fp1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fr1 = fp1.add_run(
            f"\nลงชื่อ .................................................. ผู้รับเงิน\n"
            f"(ตำแหน่ง)                     {position}"
        )
        set_thai_font(fr1, 15, bold=False)

    # ==========================================================================
    # 2. COURSE REGISTRATION CARD (บัตรลงทะเบียนรายวิชา อศท.28) MATCHING EXofBill
    # ==========================================================================
    def _render_course_registration_card(self, doc, data: Dict[str, Any], meta: Dict[str, Any]):
        college = (self.college_name or meta.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")).strip()
        stu_code = str(data.get("student_code", "69219100049")).strip()
        stu_name = str(data.get("full_name", "นางสาวปรันดา เจียศิริพันธ์")).strip()
        cls_grp = str(data.get("class_group", "ปวช. 1/2")).strip()
        fiscal_yr = str(data.get("fiscal_year", meta.get("fiscal_year", "2569"))).strip()
        major_name = str(data.get("major", "เทคโนโลยีธุรกิจดิจิทัล")).strip()
        subject_type = str(data.get("subject_type", "อุตสาหกรรมดิจิทัลและเทคโนโลยีสารสนเทศ")).strip()
        career_group = str(data.get("career_group", "ธุรกิจดิจิทัลและพาณิชย์อิเล็กทรอนิกส์")).strip()
        bank_acc = str(data.get("bank_account", "747-6-00234-2")).strip()
        bank_code = str(data.get("bank_code", "80215")).strip()
        ontime_dates = str(data.get("ontime_dates", "12-16 ตุลาคม 2569")).strip()
        late_dates = str(data.get("late_dates", "19-21 ตุลาคม 2569")).strip()

        top_p = doc.add_paragraph()
        top_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        top_p.paragraph_format.space_after = Pt(2)
        r_box = top_p.add_run("[ อศท.28 ]")
        set_thai_font(r_box, 13, bold=True)

        hdr_t = doc.add_table(rows=1, cols=2)
        hdr_t.alignment = WD_TABLE_ALIGNMENT.CENTER
        c0, c1 = hdr_t.rows[0].cells
        c0.width = Inches(3.4)
        c1.width = Inches(3.8)
        set_cell_borders(c0)
        set_cell_borders(c1)

        p0 = c0.paragraphs[0]
        r_ktb = p0.add_run("บมจ. ธนาคารกรุงไทย\n")
        set_thai_font(r_ktb, 16, bold=True)
        r_ktb_body = p0.add_run(
            f"KRUNG THAI BANK PUBLIC COMPANY LIMITED\n"
            f"ผู้รับ.............................................................................\n"
            f"เข้าบัญชี   {college}\n"
            f"เลขที่บัญชี {bank_acc}   Code. {bank_code}\n"
            f"*************รับเฉพาะเงินสด*************"
        )
        set_thai_font(r_ktb_body, 13.5, bold=False)

        p1 = c1.paragraphs[0]
        r_card_t = p1.add_run(f"บัตรลงทะเบียนรายวิชา\n{college}\n")
        set_thai_font(r_card_t, 16, bold=True)
        r_card_b = p1.add_run(
            f"รหัสนักเรียน นักศึกษา....{stu_code}...................................\n"
            f"ชื่อ-นามสกุล..........{stu_name}......................\n"
            f"ระดับชั้น........{cls_grp}.............ปีการศึกษา......{fiscal_yr}..........\n"
            f"สาขาวิชา.......{major_name}........................................."
        )
        set_thai_font(r_card_b, 13.5, bold=False)

        sub_p = doc.add_paragraph()
        sub_p.paragraph_format.space_before = Pt(2)
        sub_p.paragraph_format.space_after = Pt(2)
        r_sub = sub_p.add_run(
            f"ระดับชั้น {cls_grp}   ประเภทวิชา {subject_type}   กลุ่มอาชีพ {career_group}   สาขาวิชา {major_name}"
        )
        set_thai_font(r_sub, 12.5, bold=True)

        courses: List[Dict[str, Any]] = data.get("courses", [])
        c_table = doc.add_table(rows=17, cols=6)
        c_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        b_spec = {"val": "single", "sz": 4, "color": "000000"}

        headers = ["ลำดับที่", "รหัสวิชา", "ชื่อวิชา", "ท.", "ป.", "นก."]
        for ci, h_txt in enumerate(headers):
            cell = c_table.cell(0, ci)
            set_cell_borders(cell, top=b_spec, bottom=b_spec, left=b_spec, right=b_spec)
            cp = cell.paragraphs[0]
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_thai_font(cp.add_run(h_txt), 13, bold=True)

        tot_t, tot_p, tot_c = 0, 0, 0
        for idx in range(1, 16):
            row_cells = c_table.rows[idx].cells
            for cell in row_cells:
                set_cell_borders(cell, top=b_spec, bottom=b_spec, left=b_spec, right=b_spec)
            if idx <= len(courses):
                c_it = courses[idx - 1]
                t_v = int(float(c_it.get("theory_hours", 0)))
                p_v = int(float(c_it.get("practice_hours", 0)))
                c_v = int(float(c_it.get("credits", 0)))
                tot_t += t_v
                tot_p += p_v
                tot_c += c_v
                vals = [str(idx), str(c_it.get("course_code", "")), str(c_it.get("course_name", "")), str(t_v), str(p_v), str(c_v)]
            else:
                vals = [str(idx), "", "", "", "", ""]
            for ci, v_str in enumerate(vals):
                cp = row_cells[ci].paragraphs[0]
                cp.alignment = WD_ALIGN_PARAGRAPH.LEFT if ci == 2 else WD_ALIGN_PARAGRAPH.CENTER
                set_thai_font(cp.add_run(v_str), 12.5)

        tot_cells = c_table.rows[16].cells
        for cell in tot_cells:
            set_cell_borders(cell, top=b_spec, bottom=b_spec, left=b_spec, right=b_spec)
        set_thai_font(tot_cells[2].paragraphs[0].add_run("รวม"), 13, bold=True)
        tot_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for ci, tv in ((3, tot_t), (4, tot_p), (5, tot_c)):
            tot_cells[ci].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_thai_font(tot_cells[ci].paragraphs[0].add_run(str(tv)), 13, bold=True)

        fee_items: List[Dict[str, Any]] = data.get("fee_items") or data.get("items") or [
            {"item_name": "ค่าใช้บริการอินเตอร์เน็ต", "line_total": 300.0},
            {"item_name": "ค่าจ้างครูหรือผู้ที่มีความเชี่ยวชาญในสาขาเฉพาะ", "line_total": 500.0},
            {"item_name": "ค่ารักษาสภาพแวดล้อม", "line_total": 330.0},
        ]
        fee_total = sum(float(f.get("line_total", f.get("unit_price", 0.0))) for f in fee_items)

        fp = doc.add_paragraph()
        fp.paragraph_format.space_before = Pt(4)
        r_ft = fp.add_run(
            f"ค่าธรรมเนียมการศึกษา: "
            + ", ".join(f"{f.get('item_name')} ({float(f.get('line_total', 0)):,.0f} บาท)" for f in fee_items)
            + f"  |  รวมทั้งสิ้น {fee_total:,.0f} บาท\n"
            f"ลงทะเบียนตามกำหนด ภายในวันที่ {ontime_dates} โดยชำระเงินผ่านธนาคารกรุงไทยฯ ได้ทุกสาขาทั่วประเทศ ค่าธรรมเนียมชำระเงินผ่านธนาคาร 10.- บาท\n"
            f"ลงทะเบียนหลังกำหนด (ล่าช้า) ภายในวันที่ {late_dates}"
        )
        set_thai_font(r_ft, 13, bold=True)

        sig_p = doc.add_paragraph()
        sig_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_sig = sig_p.add_run(
            f"ลงชื่อ..........{stu_name}..........ผู้ลงทะเบียน\n"
            f"ลงชื่อ............................................................ครูที่ปรึกษา\n"
            f"ลงชื่อ.....................................................หัวหน้างานปกครอง\n"
            f"ชำระเงินแล้วตามใบเสร็จรับเงินเล่มที่................... เลขที่................... วันที่...................\n"
            f"ลงชื่อ...................................................เจ้าหน้าที่งานการเงิน"
        )
        set_thai_font(r_sig, 13)

        cut_p = doc.add_paragraph()
        r_cut = cut_p.add_run(
            f"ตัดส่วนนี้ _____________________________________________________________________________________\n"
            f"สำหรับธนาคาร                                                              วันที่............เดือน.......................พ.ศ................\n"
            f"รหัสนักเรียน นักศึกษา....{stu_code}.... ชื่อ-นามสกุล....{stu_name}....ระดับชั้น....{cls_grp}....สาขาวิชา....{major_name}....\n"
            f"จำนวนเงินเป็นตัวอักษร (............{thai_baht_text(fee_total)}............) ..........{fee_total:,.2f}..........บาท\n"
            f"ผู้นำฝาก...................................................เบอร์โทรศัพท์........................................ ลงชื่อ...................................................เจ้าหน้าที่ธนาคาร"
        )
        set_thai_font(r_cut, 13)

    # ==========================================================================
    # 1B. MODERN V2 UPGRADED RECEIPT (.DOCX)
    # ==========================================================================
    def _render_official_ex_bill_modern(self, doc, data: Dict[str, Any], meta: Dict[str, Any]):
        college = (self.college_name or meta.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")).strip()
        rcp_no = str(data.get("receipt_no", data.get("tx_code", "-"))).strip()
        raw_date = data.get("receipt_date", data.get("tx_date", ""))
        rcp_date_th = format_receipt_date_thai(raw_date)
        term_code = str(data.get("term_code", meta.get("current_semester", "2/2569"))).strip()
        stu_name = str(data.get("full_name", data.get("snapshot_student_name", "-"))).strip()
        stu_code = str(data.get("student_code", "-")).strip()
        cls_grp = str(data.get("class_group", data.get("snapshot_class_group", data.get("education_level", "ปวช.1")))).strip()
        faculty_name, major_name = _resolve_faculty_and_major(data)
        position = (data.get("cashier_position") or meta.get("cashier_position", "ครู")).strip()
        raw_status = str(data.get("receipt_status", data.get("status", "ISSUED"))).upper()

        # Header Table with Garuda in middle and Framed Document Box on right
        top_tbl = doc.add_table(rows=1, cols=3)
        top_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        c_left, c_mid, c_right = top_tbl.cell(0, 0), top_tbl.cell(0, 1), top_tbl.cell(0, 2)
        c_left.width = Inches(2.6)
        c_mid.width = Inches(1.5)
        c_right.width = Inches(2.9)
        for c in (c_left, c_mid):
            set_cell_borders(c)
        border_slate = {"val": "single", "sz": 6, "color": "CBD5E1"}
        set_cell_borders(c_right, top=border_slate, bottom=border_slate, left={"val": "single", "sz": 18, "color": "0F172A"}, right=border_slate)
        set_cell_background(c_right, "F8FAFC")

        lp = c_left.paragraphs[0]
        r_title = lp.add_run("ใบเสร็จรับเงิน\n")
        set_thai_font(r_title, 20, bold=True, color_rgb=(15, 23, 42))
        r_gov = lp.add_run("ในราชการกรมอาชีวศึกษา\nกระทรวงศึกษาธิการ")
        set_thai_font(r_gov, 15.5, color_rgb=(30, 41, 59))

        mp = c_mid.paragraphs[0]
        mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        garuda_path = get_garuda_emblem_path()
        if garuda_path and os.path.exists(garuda_path):
            try:
                mp.add_run().add_picture(garuda_path, width=Inches(0.95))
            except Exception:
                pass

        rp = c_right.paragraphs[0]
        r_meta = rp.add_run(
            f"เลขที่        {rcp_no}\n"
            f"ที่ทำการ    {college}\n"
            f"วันที่          {rcp_date_th}"
        )
        set_thai_font(r_meta, 15, bold=True, color_rgb=(15, 23, 42))

        # Modern Student Info Card Table
        info_tbl = doc.add_table(rows=3, cols=2)
        info_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        info_rows = [
            (f"ได้รับเงินจาก     {stu_name}", f"รหัสประจำตัว    {stu_code}"),
            (f"ภาค/ปีการศึกษา   {term_code}", f"คณะ         {faculty_name}"),
            (f"แผนกวิชา       {major_name}", f"ชั้น         {cls_grp}"),
        ]
        for r_idx, (l_txt, r_txt) in enumerate(info_rows):
            cl, cr = info_tbl.cell(r_idx, 0), info_tbl.cell(r_idx, 1)
            cl.width = Inches(4.1)
            cr.width = Inches(2.9)
            set_cell_background(cl, "F8FAFC")
            set_cell_background(cr, "F8FAFC")
            set_cell_borders(cl, top=border_slate, bottom=border_slate, left={"val": "single", "sz": 18, "color": "1E293B"}, right=border_slate)
            set_cell_borders(cr, top=border_slate, bottom=border_slate, left=border_slate, right=border_slate)
            rl = cl.paragraphs[0].add_run(l_txt)
            set_thai_font(rl, 15.5, bold=(r_idx == 0), color_rgb=(15, 23, 42))
            rr = cr.paragraphs[0].add_run(r_txt)
            set_thai_font(rr, 15.5, bold=(r_idx == 0), color_rgb=(15, 23, 42))

        doc.add_paragraph().paragraph_format.space_after = Pt(2)

        # Items Table with Dark Navy Header & Zebra Rows
        items: List[Dict[str, Any]] = data.get("items", [])
        if not items:
            amt_v = float(data.get("net_total_amount", data.get("amount", 0.0)))
            items = [{"line_no": 1, "item_name": data.get("title", "ชำระค่าธรรมเนียมการศึกษา"), "quantity": 1, "unit_price": amt_v, "line_total": amt_v}]

        filler_count = max(1, 12 - len(items))
        item_tbl = doc.add_table(rows=len(items) + filler_count + 2, cols=2)
        item_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        hdr_l, hdr_r = item_tbl.cell(0, 0), item_tbl.cell(0, 1)
        hdr_l.width = Inches(5.1)
        hdr_r.width = Inches(1.9)
        set_cell_background(hdr_l, "0F172A")
        set_cell_background(hdr_r, "0F172A")
        border_navy = {"val": "single", "sz": 8, "color": "0F172A"}
        set_cell_borders(hdr_l, top=border_navy, bottom=border_navy, left=border_navy, right=border_slate)
        set_cell_borders(hdr_r, top=border_navy, bottom=border_navy, left=border_slate, right=border_navy)

        hl_p = hdr_l.paragraphs[0]
        hl_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_thai_font(hl_p.add_run("รายการ"), 16, bold=True, color_rgb=(255, 255, 255))

        hr_p = hdr_r.paragraphs[0]
        hr_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_thai_font(hr_p.add_run("จำนวนเงิน"), 16, bold=True, color_rgb=(255, 255, 255))

        for idx, it in enumerate(items, start=1):
            cl, cr = item_tbl.cell(idx, 0), item_tbl.cell(idx, 1)
            cl.width = Inches(5.1)
            cr.width = Inches(1.9)
            bg_hex = "F8FAFC" if idx % 2 == 0 else "FFFFFF"
            set_cell_background(cl, bg_hex)
            set_cell_background(cr, bg_hex)
            set_cell_borders(cl, bottom=border_slate, left=border_slate, right=border_slate)
            set_cell_borders(cr, bottom=border_slate, left=border_slate, right=border_slate)

            iname = str(it.get("item_name", "")).strip()
            ltotal = float(it.get("line_total", 0.0))
            set_thai_font(cl.paragraphs[0].add_run(iname), 15.5, color_rgb=(15, 23, 42))
            cr.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_thai_font(cr.paragraphs[0].add_run(f"{ltotal:,.2f}"), 15.5, bold=True, color_rgb=(15, 23, 42))

        for f_i in range(filler_count):
            r_pos = len(items) + 1 + f_i
            cl, cr = item_tbl.cell(r_pos, 0), item_tbl.cell(r_pos, 1)
            set_cell_borders(cl, left=border_slate, right=border_slate)
            set_cell_borders(cr, left=border_slate, right=border_slate)
            set_thai_font(cl.paragraphs[0].add_run(" "), 12)

        net_val = float(data.get("net_total_amount", data.get("amount", 0.0)))
        baht_txt = data.get("amount_text_th") or thai_baht_text(net_val)
        tot_idx = len(items) + filler_count + 1
        tl, tr = item_tbl.cell(tot_idx, 0), item_tbl.cell(tot_idx, 1)
        set_cell_background(tl, "F1F5F9")
        set_cell_background(tr, "F1F5F9")
        set_cell_borders(tl, top=border_navy, bottom=border_navy, left=border_navy, right=border_slate)
        set_cell_borders(tr, top=border_navy, bottom=border_navy, left=border_slate, right=border_navy)

        tl.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_thai_font(tl.paragraphs[0].add_run(baht_txt), 15.5, bold=True, color_rgb=(15, 23, 42))
        tr.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_thai_font(tr.paragraphs[0].add_run(f"รวม        {net_val:,.2f}"), 16, bold=True, color_rgb=(15, 23, 42))

        if raw_status == "VOIDED":
            vp = doc.add_paragraph()
            vp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_thai_font(vp.add_run("*** ใบเสร็จรับเงินฉบับนี้ถูกยกเลิกแล้ว (VOIDED) ***"), 16, bold=True, color_rgb=(185, 28, 28))

        sig_tbl = doc.add_table(rows=1, cols=2)
        sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        sl, sr = sig_tbl.cell(0, 0), sig_tbl.cell(0, 1)
        set_cell_borders(sl)
        set_cell_background(sr, "F8FAFC")
        set_cell_borders(sr, top=border_slate, bottom=border_slate, left=border_slate, right=border_slate)

        sl.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_thai_font(sl.paragraphs[0].add_run("(ตัวอักษร)\nไว้เป็นการถูกต้องแล้ว"), 15.5, bold=True, color_rgb=(15, 23, 42))
        set_thai_font(
            sr.paragraphs[0].add_run(
                f"\nลงชื่อ .................................................. ผู้รับเงิน\n"
                f"(ตำแหน่ง)               {position}"
            ),
            15.5,
            color_rgb=(15, 23, 42),
        )

    # ==========================================================================
    # 2B. MODERN V2 UPGRADED COURSE REGISTRATION CARD (.DOCX)
    # ==========================================================================
    def _render_course_registration_card_modern(self, doc, data: Dict[str, Any], meta: Dict[str, Any]):
        stu_dict = data.get("student") if isinstance(data.get("student"), dict) else {}
        college = (self.college_name or meta.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")).strip()
        stu_code = str(data.get("student_code") or stu_dict.get("student_code") or "69219100049").strip()
        stu_name = str(data.get("full_name") or stu_dict.get("full_name") or "นางสาวปรันดา เจียศิริพันธ์").strip()
        cls_grp = str(data.get("class_group") or stu_dict.get("class_group") or "ปวช. 1/2").strip()
        fiscal_yr = str(data.get("fiscal_year") or meta.get("fiscal_year", "2569")).strip()
        major_name = str(data.get("major") or stu_dict.get("major") or "เทคโนโลยีธุรกิจดิจิทัล").strip()
        subject_type = str(data.get("subject_type", "อุตสาหกรรมดิจิทัลและเทคโนโลยีสารสนเทศ")).strip()
        career_group = str(data.get("career_group", "ธุรกิจดิจิทัลและพาณิชย์อิเล็กทรอนิกส์")).strip()
        bank_acc = str(data.get("bank_account", "747-6-00234-2")).strip()
        bank_code = str(data.get("bank_code", "80215")).strip()
        ontime_dates = str(data.get("ontime_dates", "12-16 ตุลาคม 2569")).strip()
        late_dates = str(data.get("late_dates", "19-21 ตุลาคม 2569")).strip()

        border_slate = {"val": "single", "sz": 6, "color": "CBD5E1"}
        border_navy = {"val": "single", "sz": 12, "color": "0F172A"}

        top_tbl = doc.add_table(rows=1, cols=2)
        top_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cl, cr = top_tbl.cell(0, 0), top_tbl.cell(0, 1)
        cl.width = Inches(3.4)
        cr.width = Inches(3.6)
        set_cell_background(cl, "F8FAFC")
        set_cell_background(cr, "FFFFFF")
        set_cell_borders(cl, top=border_navy, bottom=border_slate, left=border_slate, right=border_slate)
        set_cell_borders(cr, top=border_navy, bottom=border_slate, left=border_slate, right=border_slate)

        lp = cl.paragraphs[0]
        set_thai_font(lp.add_run("บมจ. ธนาคารกรุงไทย\n"), 16, bold=True, color_rgb=(15, 23, 42))
        set_thai_font(lp.add_run("KRUNG THAI BANK PUBLIC COMPANY LIMITED\n"), 9.5, bold=True, color_rgb=(71, 85, 105))
        set_thai_font(
            lp.add_run(
                f"ผู้รับ.............................................................................\n"
                f"เข้าบัญชี   {college}\n"
                f"เลขที่บัญชี {bank_acc}   Code. {bank_code}\n"
            ),
            13.5,
            color_rgb=(15, 23, 42),
        )
        set_thai_font(lp.add_run("*************รับเฉพาะเงินสด*************"), 13.5, bold=True, color_rgb=(15, 23, 42))

        rp = cr.paragraphs[0]
        set_thai_font(rp.add_run(f"[ อศท.28 ]   บัตรลงทะเบียนรายวิชา ({college})\n"), 15.5, bold=True, color_rgb=(15, 23, 42))
        set_thai_font(
            rp.add_run(
                f"รหัสนักเรียน นักศึกษา....{stu_code}......................................\n"
                f"ชื่อ-นามสกุล..........{stu_name}.........................\n"
                f"ระดับชั้น........{cls_grp}.............ปีการศึกษา......{fiscal_yr}.............\n"
                f"สาขาวิชา.......{major_name}.............................................."
            ),
            13.5,
            color_rgb=(15, 23, 42),
        )

        sub_p = doc.add_paragraph()
        set_thai_font(
            sub_p.add_run(f"ระดับชั้น {cls_grp}   ประเภทวิชา {subject_type}   กลุ่มอาชีพ {career_group}   สาขาวิชา {major_name}"),
            12.5,
            bold=True,
            color_rgb=(15, 23, 42),
        )

        courses: List[Dict[str, Any]] = data.get("courses", [])
        c_tbl = doc.add_table(rows=17, cols=6)
        c_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        headers = ["ลำดับที่", "รหัสวิชา", "ชื่อวิชา", "ท.", "ป.", "นก."]
        for c_i, h_txt in enumerate(headers):
            cell = c_tbl.cell(0, c_i)
            set_cell_background(cell, "0F172A")
            set_cell_borders(cell, top=border_navy, bottom=border_navy, left=border_slate, right=border_slate)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_thai_font(p.add_run(h_txt), 13, bold=True, color_rgb=(255, 255, 255))

        tot_t, tot_p, tot_c = 0, 0, 0
        for idx in range(1, 16):
            bg_hex = "F8FAFC" if idx % 2 == 0 else "FFFFFF"
            if idx <= len(courses):
                c_it = courses[idx - 1]
                t_v = int(float(c_it.get("theory_hours", 0)))
                p_v = int(float(c_it.get("practice_hours", 0)))
                cr_v = int(float(c_it.get("credits", 0)))
                tot_t += t_v
                tot_p += p_v
                tot_c += cr_v
                vals = [str(idx), str(c_it.get("course_code", "")), str(c_it.get("course_name", "")), str(t_v), str(p_v), str(cr_v)]
            else:
                vals = [str(idx), "", "", "", "", ""]
            for c_i, val in enumerate(vals):
                cell = c_tbl.cell(idx, c_i)
                set_cell_background(cell, bg_hex)
                set_cell_borders(cell, top=border_slate, bottom=border_slate, left=border_slate, right=border_slate)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_i == 2 else WD_ALIGN_PARAGRAPH.CENTER
                set_thai_font(p.add_run(val), 13, color_rgb=(15, 23, 42))

        c_tbl.cell(16, 0).merge(c_tbl.cell(16, 2))
        for c_i, val in ((0, "รวม"), (3, str(tot_t)), (4, str(tot_p)), (5, str(tot_c))):
            cell = c_tbl.cell(16, c_i)
            set_cell_background(cell, "E2E8F0")
            set_cell_borders(cell, top=border_navy, bottom=border_navy, left=border_slate, right=border_slate)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_thai_font(p.add_run(val), 13, bold=True, color_rgb=(15, 23, 42))

        fee_items: List[Dict[str, Any]] = data.get("fee_items") or data.get("items") or [
            {"item_name": "ค่าใช้บริการอินเตอร์เน็ต", "line_total": 300.0},
            {"item_name": "ค่าจ้างครูหรือผู้ที่มีความเชี่ยวชาญในสาขาเฉพาะ", "line_total": 500.0},
            {"item_name": "ค่ารักษาสภาพแวดล้อม", "line_total": 330.0},
        ]
        fee_total = sum(float(f.get("line_total", f.get("unit_price", 0.0))) for f in fee_items)

        fp = doc.add_paragraph()
        set_thai_font(
            fp.add_run(
                f"ค่าธรรมเนียมการศึกษา (รวม {fee_total:,.0f} บาท)\n"
                f"ลงทะเบียนตามกำหนด ภายในวันที่ {ontime_dates} โดยชำระเงินผ่านธนาคารกรุงไทยฯ ได้ทุกสาขาทั่วประเทศ ค่าธรรมเนียมชำระเงินผ่านธนาคาร 10.- บาท\n"
                f"ลงทะเบียนหลังกำหนด (ล่าช้า) ภายในวันที่ {late_dates}"
            ),
            12.5,
            bold=True,
            color_rgb=(15, 23, 42),
        )

        sig_p = doc.add_paragraph()
        sig_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_thai_font(
            sig_p.add_run(
                f"ลงชื่อ..........{stu_name}..........ผู้ลงทะเบียน\n"
                f"ลงชื่อ............................................................ครูที่ปรึกษา\n"
                f"ลงชื่อ.....................................................หัวหน้างานปกครอง\n"
                f"ชำระเงินแล้วตามใบเสร็จรับเงินเล่มที่................... เลขที่................... วันที่...................\n"
                f"ลงชื่อ...................................................เจ้าหน้าที่งานการเงิน"
            ),
            13,
            color_rgb=(15, 23, 42),
        )

        cut_p = doc.add_paragraph()
        set_thai_font(
            cut_p.add_run(
                f"ตัดส่วนนี้ _____________________________________________________________________________________\n"
                f"[ สำหรับธนาคาร ]                                                          วันที่............เดือน.......................พ.ศ................\n"
                f"รหัสนักเรียน นักศึกษา....{stu_code}.... ชื่อ-นามสกุล....{stu_name}....ระดับชั้น....{cls_grp}....สาขาวิชา....{major_name}....\n"
                f"จำนวนเงินเป็นตัวอักษร (............{thai_baht_text(fee_total)}............) ..........{fee_total:,.2f}..........บาท\n"
                f"ผู้นำฝาก...................................................เบอร์โทรศัพท์........................................ ลงชื่อ...................................................เจ้าหน้าที่ธนาคาร"
            ),
            13,
            color_rgb=(15, 23, 42),
        )

    # ==========================================================================
    # 3. DAILY CLOSING REPORT RENDERING
    # ==========================================================================
    def _render_daily_report_header(self, doc, data: Dict[str, Any], meta: Dict[str, Any]):
        h_para = doc.add_paragraph()
        h_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        h_run = h_para.add_run(self.college_name or meta.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"))
        set_thai_font(h_run, 20, bold=True, color_rgb=(15, 23, 42))

        sub_para = doc.add_paragraph()
        sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        addr = meta.get("college_address", "455 ถ.พระพันวษา ต.ท่าพี่เลี้ยง อ.เมืองสุพรรณบุรี จ.สุพรรณบุรี 72000")
        tax_id = meta.get("college_tax_id", "0994000159721")
        sub_run = sub_para.add_run(f"{addr}\nสำนักงานคณะกรรมการการอาชีวศึกษา • รหัสหน่วยงาน/เลขประจำตัวผู้เสียภาษี: {tax_id}")
        set_thai_font(sub_run, 13.5, bold=False, color_rgb=(71, 85, 105))

        doc_title = data.get("doc_title", "รายงานสรุปการรับเงินค่าเทอมประจำวัน")
        t_table = doc.add_table(rows=1, cols=1)
        t_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_cell = t_table.cell(0, 0)
        t_cell.width = Inches(6.4)
        set_cell_background(t_cell, "0F172A")
        tp = t_cell.paragraphs[0]
        tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tr = tp.add_run(doc_title)
        set_thai_font(tr, 16, bold=True, color_rgb=(255, 255, 255))

        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    def _render_daily_closing(self, doc, data: Dict[str, Any]):
        summary_table = doc.add_table(rows=3, cols=2)
        summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        inc_val = float(data.get("income_total", 0.0))
        s_rows = [
            ("วันที่ออกใบเสร็จ (Report Date):", format_receipt_date_thai(data.get("date", "-"))),
            ("จำนวนใบเสร็จที่รับชำระทั้งหมด:", f"{data.get('tx_count', 0)} ฉบับ"),
            ("ยอดรวมเงินรับสุทธิประจำวัน:", f"฿{inc_val:,.2f} บาท ({thai_baht_text(inc_val)})"),
        ]
        for idx, (k, v) in enumerate(s_rows):
            c0, c1 = summary_table.cell(idx, 0), summary_table.cell(idx, 1)
            set_cell_background(c0, "F1F5F9")
            set_cell_background(c1, "FFFFFF")
            r0 = c0.paragraphs[0].add_run(k)
            set_thai_font(r0, 14, bold=True)
            r1 = c1.paragraphs[0].add_run(v)
            set_thai_font(r1, 14, bold=(idx == 2))

        doc.add_paragraph().paragraph_format.space_after = Pt(6)

        by_method: Dict[str, float] = data.get("income_by_method", {})
        if by_method:
            mp = doc.add_paragraph()
            mr = mp.add_run("1. แยกยอดเงินรับตามช่องทางการชำระเงิน:")
            set_thai_font(mr, 15, bold=True)
            for method_name, amt_val in by_method.items():
                bp = doc.add_paragraph()
                bp.paragraph_format.space_before = Pt(1)
                bp.paragraph_format.space_after = Pt(1)
                br = bp.add_run(f"   • {method_name}:  ฿{float(amt_val):,.2f} บาท")
                set_thai_font(br, 14)

        by_fee: Dict[str, float] = data.get("income_by_fee_item", {})
        if by_fee:
            fp = doc.add_paragraph()
            fp.paragraph_format.space_before = Pt(6)
            fr = fp.add_run("2. แยกยอดเงินรับตามหมวดรายการค่าธรรมเนียมในใบเสร็จ:")
            set_thai_font(fr, 15, bold=True)
            for fee_name, amt_val in by_fee.items():
                bp = doc.add_paragraph()
                bp.paragraph_format.space_before = Pt(1)
                bp.paragraph_format.space_after = Pt(1)
                br = bp.add_run(f"   • {fee_name}:  ฿{float(amt_val):,.2f} บาท")
                set_thai_font(br, 14)

    def _render_daily_signature(self, doc, data: Dict[str, Any], meta: Dict[str, Any]):
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
            f"(ตำแหน่ง)   {position}\n"
            f"{college}"
        )
        set_thai_font(sig_run, 14, color_rgb=(15, 23, 42))
