import os
import re
import warnings
from typing import Any, Dict, List
from core.config import format_receipt_date_thai, get_base_dir, get_garuda_emblem_path, thai_baht_text
from exporters.base import BaseFinanceExporter

try:
    from fpdf import FPDF
except ImportError:
    FPDF = None


def _setup_thai_sarabun_fonts(pdf) -> tuple[str, bool]:
    """
    Registers Thai Sarabun font (THSarabunNew / THSarabun) for official government document rendering.
    Falls back to Windows Tahoma only if Sarabun is unavailable.
    """
    base_dir = get_base_dir()
    bundled_dir = os.path.join(base_dir, "data", "fonts")
    user_fonts_dir = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")
    win_fonts_dir = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")

    sarabun_candidates = [
        (os.path.join(bundled_dir, "THSarabunNew.ttf"), os.path.join(bundled_dir, "THSarabunNew Bold.ttf")),
        (os.path.join(bundled_dir, "THSarabun.ttf"), os.path.join(bundled_dir, "THSarabun Bold.ttf")),
        (os.path.join(user_fonts_dir, "THSarabunNew.ttf"), os.path.join(user_fonts_dir, "THSarabunNew Bold.ttf")),
        (os.path.join(user_fonts_dir, "THSarabun.ttf"), os.path.join(user_fonts_dir, "THSarabun Bold.ttf")),
        (os.path.join(win_fonts_dir, "THSarabunNew.ttf"), os.path.join(win_fonts_dir, "THSarabunNew Bold.ttf")),
        (os.path.join(win_fonts_dir, "THSarabun.ttf"), os.path.join(win_fonts_dir, "THSarabun Bold.ttf")),
    ]

    for reg_path, bold_path in sarabun_candidates:
        if os.path.exists(reg_path) and os.path.exists(bold_path):
            try:
                pdf.add_font("THSarabun", "", reg_path)
                pdf.add_font("THSarabun", "B", bold_path)
                return "THSarabun", True
            except Exception:
                pass

    tahoma_reg = os.path.join(win_fonts_dir, "tahoma.ttf")
    tahoma_bold = os.path.join(win_fonts_dir, "tahomabd.ttf")
    if os.path.exists(tahoma_reg) and os.path.exists(tahoma_bold):
        try:
            pdf.add_font("Tahoma", "", tahoma_reg)
            pdf.add_font("Tahoma", "B", tahoma_bold)
            return "Tahoma", True
        except Exception:
            pass
    return "Helvetica", False


def _clean_signer_name(raw_name: str) -> str:
    cleaned = re.sub(r"\s*\(.*?\)\s*$", "", str(raw_name or "").strip())
    return cleaned or "เจ้าหน้าที่การเงิน"


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


class PdfReceiptExporter(BaseFinanceExporter):
    """
    Official Suphanburi Vocational College (SPVC) PDF Exporter:
    1. Official Receipt (ใบเสร็จรับเงิน) matching EXofBill (with TH Sarabun font, high-res Garuda emblem, and bottom-anchored total/signature).
    2. Official Course Registration Card (บัตรลงทะเบียนรายวิชา อศท.28) matching EXofBill.
    3. Batch Classroom Multi-Page PDF Receipts & Daily Closing PDF Report.
    """

    def export(self, data: Dict[str, Any], metadata: Dict[str, Any] = None) -> str:
        if FPDF is None:
            raise RuntimeError("Module 'fpdf2' is not installed.")

        os.makedirs(os.path.dirname(os.path.abspath(self.save_path)), exist_ok=True)
        meta = metadata or {}

        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            pdf = FPDF(orientation="P", unit="mm", format="A4")
            pdf.set_margins(left=15, top=12, right=15)
            pdf.set_auto_page_break(auto=False)
            font_family, _ = _setup_thai_sarabun_fonts(pdf)

            pdf.add_page()
            report_type = data.get("report_type", "RECEIPT")
            if report_type == "DAILY_CLOSING":
                pdf.set_auto_page_break(auto=True, margin=14)
                self._render_daily_report_header(pdf, font_family, data, meta)
                self._render_daily_closing(pdf, font_family, data)
                self._render_daily_signature(pdf, font_family, data, meta)
            elif report_type == "COURSE_REGISTRATION":
                self._render_course_registration_card(pdf, font_family, data, meta)
            else:
                self._render_official_ex_bill(pdf, font_family, data, meta)

            pdf.output(self.save_path)

        return self.save_path

    def export_batch(self, receipts_list: List[Dict[str, Any]], metadata: Dict[str, Any] = None) -> str:
        if FPDF is None:
            raise RuntimeError("Module 'fpdf2' is not installed.")
        if not receipts_list:
            raise ValueError("ไม่มีข้อมูลใบเสร็จสำหรับพิมพ์")

        os.makedirs(os.path.dirname(os.path.abspath(self.save_path)), exist_ok=True)
        meta = metadata or {}

        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            pdf = FPDF(orientation="P", unit="mm", format="A4")
            pdf.set_margins(left=15, top=12, right=15)
            pdf.set_auto_page_break(auto=False)
            font_family, _ = _setup_thai_sarabun_fonts(pdf)

            for rcp_data in receipts_list:
                pdf.add_page()
                if rcp_data.get("report_type") == "COURSE_REGISTRATION":
                    self._render_course_registration_card(pdf, font_family, rcp_data, meta)
                else:
                    self._render_official_ex_bill(pdf, font_family, rcp_data, meta)

            pdf.output(self.save_path)

        return self.save_path

    # ==========================================================================
    # 1. OFFICIAL RECEIPT MATCHING EXofBill (95b999ca-2a23-4412-bbee-a3b29b8aa4e6.jpeg)
    # ==========================================================================
    def _render_official_ex_bill(self, pdf, font_family: str, data: Dict[str, Any], meta: Dict[str, Any]):
        is_sarabun = font_family == "THSarabun"
        sz_title = 20 if is_sarabun else 14.5
        sz_body = 15.5 if is_sarabun else 11
        sz_tbl_hdr = 16 if is_sarabun else 11.5

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

        # 1. Center Top High-Resolution Garuda Emblem
        garuda_path = get_garuda_emblem_path()
        if garuda_path and os.path.exists(garuda_path):
            try:
                pdf.image(garuda_path, x=92.5, y=10.5, w=25.0)
            except Exception:
                pass

        # 2. Top Header Two-Column Layout (Left & Right aligned with EXofBill)
        left_x = 16.0
        right_x = 134.0
        pdf.set_text_color(0, 0, 0)

        # Row 1: ใบเสร็จรับเงิน  |  เลขที่  700605308
        y1 = 14.0
        pdf.set_xy(left_x, y1)
        pdf.set_font(font_family, "B", sz_title)
        pdf.cell(70, 7.5, "ใบเสร็จรับเงิน", align="L")

        pdf.set_xy(right_x, y1)
        pdf.set_font(font_family, "", sz_body)
        pdf.cell(16, 7.5, "เลขที่", align="L")
        pdf.cell(45, 7.5, rcp_no, align="L")

        # Row 2: ในราชการกรมอาชีวศึกษา  |  ที่ทำการ  วิทยาลัยอาชีวศึกษาสุพรรณบุรี
        y2 = 22.5
        pdf.set_xy(left_x, y2)
        pdf.set_font(font_family, "", sz_body)
        pdf.cell(75, 7.0, "ในราชการกรมอาชีวศึกษา", align="L")

        pdf.set_xy(right_x, y2)
        pdf.cell(18, 7.0, "ที่ทำการ", align="L")
        pdf.cell(48, 7.0, college, align="L")

        # Row 3: กระทรวงศึกษาธิการ  |  วันที่  30/10/2569
        y3 = 30.0
        pdf.set_xy(left_x, y3)
        pdf.cell(75, 7.0, "กระทรวงศึกษาธิการ", align="L")

        pdf.set_xy(right_x, y3)
        pdf.cell(16, 7.0, "วันที่", align="L")
        pdf.cell(48, 7.0, rcp_date_th, align="L")

        # Row 4: ได้รับเงินจาก  นางสาวปรันดา เจียศิริพันธ์  |  รหัสประจำตัว  69219100049
        y4 = 38.0
        pdf.set_xy(left_x, y4)
        pdf.cell(26, 7.0, "ได้รับเงินจาก", align="L")
        pdf.cell(85, 7.0, stu_name, align="L")

        pdf.set_xy(right_x, y4)
        pdf.cell(25, 7.0, "รหัสประจำตัว", align="L")
        pdf.cell(40, 7.0, stu_code, align="L")

        # Row 5: ภาค/ปีการศึกษา  2/2569  |  คณะ  พาณิชยกรรม/บริหารธุรกิจ
        y5 = 46.0
        pdf.set_xy(left_x, y5)
        pdf.cell(31, 7.0, "ภาค/ปีการศึกษา", align="L")
        pdf.cell(60, 7.0, term_code, align="L")

        pdf.set_xy(right_x, y5)
        pdf.cell(14, 7.0, "คณะ", align="L")
        pdf.cell(52, 7.0, faculty_name, align="L")

        # Row 6: แผนกวิชา  เทคโนโลยีธุรกิจดิจิทัล  |  ชั้น  ปวช.1
        y6 = 54.0
        pdf.set_xy(left_x, y6)
        pdf.cell(21, 7.0, "แผนกวิชา", align="L")
        pdf.cell(85, 7.0, major_name, align="L")

        pdf.set_xy(right_x, y6)
        pdf.cell(12, 7.0, "ชั้น", align="L")
        pdf.cell(52, 7.0, cls_grp, align="L")

        # 3. Items Table Header (Top & bottom horizontal lines like EXofBill)
        hdr_top_y = 64.5
        hdr_btm_y = 73.0
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.35)
        pdf.line(16.0, hdr_top_y, 194.0, hdr_top_y)
        pdf.line(16.0, hdr_btm_y, 194.0, hdr_btm_y)

        pdf.set_font(font_family, "B", sz_tbl_hdr)
        pdf.set_xy(16.0, hdr_top_y + 1.0)
        pdf.cell(115.0, 6.8, "รายการ", align="C")
        pdf.set_xy(145.0, hdr_top_y + 1.0)
        pdf.cell(47.0, 6.8, "จำนวนเงิน", align="R")

        # 4. Items Rows
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

        cur_y = hdr_btm_y + 4.5
        pdf.set_font(font_family, "", sz_body)

        for it in items:
            iname = str(it.get("item_name", "")).strip()
            qty = float(it.get("quantity", 1.0))
            uprice = float(it.get("unit_price", 0.0))
            ltotal = float(it.get("line_total", 0.0))
            if qty > 1.0:
                iname_display = f"{iname} ({qty:,.0f} หน่วย x {uprice:,.2f})"
            else:
                iname_display = iname

            pdf.set_xy(16.0, cur_y)
            pdf.cell(132.0, 7.0, iname_display, align="L")
            pdf.set_xy(148.0, cur_y)
            pdf.cell(44.0, 7.0, f"{ltotal:,.2f}", align="R")
            cur_y += 9.0

        if raw_status == "VOIDED":
            pdf.set_text_color(185, 28, 28)
            pdf.set_font(font_family, "B", sz_tbl_hdr)
            pdf.set_xy(16.0, max(cur_y + 8.0, 155.0))
            pdf.cell(178.0, 8.0, "*** ใบเสร็จรับเงินฉบับนี้ถูกยกเลิกแล้ว (VOIDED) ***", align="C")
            pdf.set_text_color(0, 0, 0)

        if data.get("remark"):
            pdf.set_font(font_family, "", sz_body - 1)
            pdf.set_xy(16.0, max(cur_y + 4.0, 235.0))
            pdf.cell(175.0, 6.5, f"หมายเหตุ: {data.get('remark')}", align="L")

        # 5. Bottom Total Bar — Anchored near the bottom edge of A4 paper (Y=252mm) just like EXofBill!
        total_top_y = max(cur_y + 8.0, 252.0)
        total_btm_y = total_top_y + 9.0

        net_val = float(data.get("net_total_amount", data.get("amount", 0.0)))
        baht_txt = data.get("amount_text_th") or thai_baht_text(net_val)

        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.35)
        pdf.line(16.0, total_top_y, 194.0, total_top_y)
        pdf.line(16.0, total_btm_y, 194.0, total_btm_y)

        pdf.set_font(font_family, "", sz_body)
        pdf.set_xy(35.0, total_top_y + 1.2)
        pdf.cell(85.0, 6.8, baht_txt, align="C")

        pdf.set_xy(136.0, total_top_y + 1.2)
        pdf.cell(56.0, 6.8, f"รวม        {net_val:,.2f}", align="R")

        # 6. Footer Verification & Signature Section — Right near the bottom edge of A4 (Y=262 to 283mm)
        # Line 1 below Total bar: (ตัวอักษร) centered under baht_txt
        line1_y = total_btm_y + 5.5
        line2_y = total_btm_y + 13.0
        line3_y = total_btm_y + 20.5

        pdf.set_font(font_family, "", sz_body)
        pdf.set_xy(35.0, line1_y)
        pdf.cell(85.0, 6.8, "(ตัวอักษร)", align="C")

        # Line 2: ไว้เป็นการถูกต้องแล้ว (Left)  |  ลงชื่อ ....................................... ผู้รับเงิน (Right)
        pdf.set_xy(35.0, line2_y)
        pdf.cell(85.0, 6.8, "ไว้เป็นการถูกต้องแล้ว", align="C")

        sig_x = 122.0
        pdf.set_xy(sig_x, line2_y)
        pdf.cell(72.0, 6.8, "ลงชื่อ .................................................. ผู้รับเงิน", align="L")

        # Line 3: (ตำแหน่ง)               ครู (Right, near bottom margin of paper)
        pdf.set_xy(sig_x, line3_y)
        pdf.cell(22.0, 6.8, "(ตำแหน่ง)", align="L")
        pdf.set_xy(sig_x + 22.0, line3_y)
        pdf.cell(42.0, 6.8, position, align="C")

    # ==========================================================================
    # 2. COURSE REGISTRATION CARD (บัตรลงทะเบียนรายวิชา อศท.28) MATCHING EXofBill
    # ==========================================================================
    def _render_course_registration_card(self, pdf, font_family: str, data: Dict[str, Any], meta: Dict[str, Any]):
        """
        Renders the official SPVC Course Registration Card & Bank Payment Slip (บัตรลงทะเบียนรายวิชา อศท.28)
        matching EX/EXofBill/8e5da586-7e15-48c6-a674-e12b8f3b1b68.jpeg 100%.
        """
        is_sarabun = font_family == "THSarabun"
        sz_lg = 16.5 if is_sarabun else 12
        sz_md = 14 if is_sarabun else 10
        sz_sm = 13 if is_sarabun else 9

        college = (self.college_name or meta.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")).strip()
        stu_code = str(data.get("student_code", "69219100049")).strip()
        stu_name = str(data.get("full_name", "นางสาวปรันดา เจียศิริพันธ์")).strip()
        cls_grp = str(data.get("class_group", "ปวช. 1/2")).strip()
        edu_lvl = str(data.get("education_level", "ปวช.1")).strip()
        fiscal_yr = str(data.get("fiscal_year", meta.get("fiscal_year", "2569"))).strip()
        major_name = str(data.get("major", "เทคโนโลยีธุรกิจดิจิทัล")).strip()
        subject_type = str(data.get("subject_type", "อุตสาหกรรมดิจิทัลและเทคโนโลยีสารสนเทศ")).strip()
        career_group = str(data.get("career_group", "ธุรกิจดิจิทัลและพาณิชย์อิเล็กทรอนิกส์")).strip()
        bank_acc = str(data.get("bank_account", "747-6-00234-2")).strip()
        bank_code = str(data.get("bank_code", "80215")).strip()
        ontime_dates = str(data.get("ontime_dates", "12-16 ตุลาคม 2569")).strip()
        late_dates = str(data.get("late_dates", "19-21 ตุลาคม 2569")).strip()

        pdf.set_text_color(0, 0, 0)
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.3)

        # Top-Right Box: [ อศท.28 ]
        pdf.rect(175.0, 10.0, 20.0, 7.5)
        pdf.set_xy(175.0, 10.5)
        pdf.set_font(font_family, "B", sz_sm)
        pdf.cell(20.0, 6.5, "อศท.28", align="C")

        # Top Left: Krungthai Bank Info
        pdf.set_xy(15.0, 16.0)
        pdf.set_font(font_family, "B", sz_lg)
        pdf.cell(85.0, 6.0, "บมจ. ธนาคารกรุงไทย", align="L")
        pdf.set_xy(15.0, 21.0)
        pdf.set_font(font_family, "", 8 if not is_sarabun else 10.5)
        pdf.cell(85.0, 4.5, "KRUNG THAI BANK PUBLIC COMPANY LIMITED", align="L")

        pdf.set_font(font_family, "", sz_md)
        pdf.set_xy(15.0, 27.0)
        pdf.cell(85.0, 5.5, "ผู้รับ.............................................................................", align="L")
        pdf.set_xy(15.0, 32.8)
        pdf.cell(85.0, 5.5, f"เข้าบัญชี   {college}", align="L")
        pdf.set_xy(15.0, 38.6)
        pdf.cell(85.0, 5.5, f"เลขที่บัญชี {bank_acc}   Code. {bank_code}", align="L")
        pdf.set_xy(15.0, 44.4)
        pdf.set_font(font_family, "B", sz_md)
        pdf.cell(85.0, 5.5, "*************รับเฉพาะเงินสด*************", align="L")

        # Top Right: บัตรลงทะเบียนรายวิชา / วิทยาลัยอาชีวศึกษาสุพรรณบุรี
        pdf.set_xy(105.0, 16.0)
        pdf.set_font(font_family, "B", sz_lg)
        pdf.cell(90.0, 6.0, "บัตรลงทะเบียนรายวิชา", align="C")
        pdf.set_xy(105.0, 22.0)
        pdf.cell(90.0, 6.0, college, align="C")

        pdf.set_font(font_family, "", sz_md)
        pdf.set_xy(105.0, 28.5)
        pdf.cell(90.0, 5.5, f"รหัสนักเรียน นักศึกษา....{stu_code}......................................", align="L")
        pdf.set_xy(105.0, 34.2)
        pdf.cell(90.0, 5.5, f"ชื่อ-นามสกุล..........{stu_name}.........................", align="L")
        pdf.set_xy(105.0, 39.9)
        pdf.cell(90.0, 5.5, f"ระดับชั้น........{cls_grp}.............ปีการศึกษา......{fiscal_yr}.............", align="L")
        pdf.set_xy(105.0, 45.6)
        pdf.cell(90.0, 5.5, f"สาขาวิชา.......{major_name}..............................................", align="L")

        # Sub-header Bar above Course Table
        pdf.set_xy(15.0, 52.0)
        pdf.set_font(font_family, "B", 11.5 if is_sarabun else 8.5)
        sub_line = f"ระดับชั้น {cls_grp}   ประเภทวิชา {subject_type}   กลุ่มอาชีพ {career_group}   สาขาวิชา {major_name}"
        pdf.cell(180.0, 5.5, sub_line, align="L")

        # 15-Row Course Table
        # Columns: ลำดับที่ (14mm) | รหัสวิชา (28mm) | ชื่อวิชา (96mm) | ท. (14mm) | ป. (14mm) | นก. (14mm) = 180mm
        col_w = [14.0, 28.0, 96.0, 14.0, 14.0, 14.0]
        headers = ["ลำดับที่", "รหัสวิชา", "ชื่อวิชา", "ท.", "ป.", "นก."]
        tbl_x = 15.0
        tbl_y = 58.0
        row_h = 5.4

        pdf.set_xy(tbl_x, tbl_y)
        pdf.set_font(font_family, "B", sz_sm)
        for c_i, h_txt in enumerate(headers):
            pdf.cell(col_w[c_i], 6.0, h_txt, border=1, align="C")
        cur_y = tbl_y + 6.0

        courses: List[Dict[str, Any]] = data.get("courses", [])
        tot_t = 0
        tot_p = 0
        tot_c = 0

        pdf.set_font(font_family, "", sz_sm)
        for idx in range(1, 16):
            pdf.set_xy(tbl_x, cur_y)
            if idx <= len(courses):
                c_item = courses[idx - 1]
                ccode = str(c_item.get("course_code", "")).strip()
                cname = str(c_item.get("course_name", "")).strip()
                t_val = int(float(c_item.get("theory_hours", 0)))
                p_val = int(float(c_item.get("practice_hours", 0)))
                c_val = int(float(c_item.get("credits", 0)))
                tot_t += t_val
                tot_p += p_val
                tot_c += c_val
                pdf.cell(col_w[0], row_h, str(idx), border=1, align="C")
                pdf.cell(col_w[1], row_h, ccode, border=1, align="C")
                pdf.cell(col_w[2], row_h, f"  {cname}", border=1, align="L")
                pdf.cell(col_w[3], row_h, str(t_val), border=1, align="C")
                pdf.cell(col_w[4], row_h, str(p_val), border=1, align="C")
                pdf.cell(col_w[5], row_h, str(c_val), border=1, align="C")
            else:
                pdf.cell(col_w[0], row_h, str(idx), border=1, align="C")
                pdf.cell(col_w[1], row_h, "", border=1, align="C")
                pdf.cell(col_w[2], row_h, "", border=1, align="L")
                pdf.cell(col_w[3], row_h, "", border=1, align="C")
                pdf.cell(col_w[4], row_h, "", border=1, align="C")
                pdf.cell(col_w[5], row_h, "", border=1, align="C")
            cur_y += row_h

        # Course Table Total Row
        pdf.set_xy(tbl_x, cur_y)
        pdf.set_font(font_family, "B", sz_sm)
        pdf.cell(col_w[0] + col_w[1] + col_w[2], row_h + 0.4, "รวม", border=1, align="C")
        pdf.cell(col_w[3], row_h + 0.4, str(tot_t), border=1, align="C")
        pdf.cell(col_w[4], row_h + 0.4, str(tot_p), border=1, align="C")
        pdf.cell(col_w[5], row_h + 0.4, str(tot_c), border=1, align="C")
        cur_y += row_h + 2.5

        # Middle Section: ค่าธรรมเนียมการศึกษา Table
        pdf.set_xy(15.0, cur_y)
        pdf.set_font(font_family, "B", sz_md)
        pdf.cell(35.0, 5.5, "ค่าธรรมเนียมการศึกษา", align="L")
        # Underline ค่าธรรมเนียมการศึกษา
        pdf.line(15.0, cur_y + 5.2, 46.0, cur_y + 5.2)

        fee_items: List[Dict[str, Any]] = data.get("fee_items") or data.get("items") or [
            {"item_name": "ค่าใช้บริการอินเตอร์เน็ต", "line_total": 300.0},
            {"item_name": "ค่าจ้างครูหรือผู้ที่มีความเชี่ยวชาญในสาขาเฉพาะ", "line_total": 500.0},
            {"item_name": "ค่ารักษาสภาพแวดล้อม", "line_total": 330.0},
        ]
        fee_x = 38.0
        fee_y = cur_y + 6.0
        fee_total = 0.0
        pdf.set_font(font_family, "", sz_sm)
        for f_it in fee_items:
            fname = str(f_it.get("item_name", "")).strip()
            famt = float(f_it.get("line_total", f_it.get("unit_price", 0.0)))
            fee_total += famt
            pdf.set_xy(fee_x, fee_y)
            pdf.cell(86.0, 5.4, f"  {fname}", border=1, align="L")
            pdf.cell(24.0, 5.4, f"{famt:,.0f}  ", border=1, align="R")
            fee_y += 5.4

        pdf.set_xy(fee_x, fee_y)
        pdf.set_font(font_family, "B", sz_sm)
        pdf.cell(86.0, 5.6, "รวม", border=1, align="C")
        pdf.cell(24.0, 5.6, f"{fee_total:,.0f}  ", border=1, align="R")
        cur_y = fee_y + 7.5

        # Payment Dates Notice
        pdf.set_xy(15.0, cur_y)
        pdf.set_font(font_family, "B", 12.0 if is_sarabun else 8.5)
        pdf.cell(
            180.0,
            5.2,
            f"ลงทะเบียนตามกำหนด ภายในวันที่ {ontime_dates} โดยชำระเงินผ่านธนาคารกรุงไทยฯ ได้ทุกสาขาทั่วประเทศ ค่าธรรมเนียมชำระเงินผ่านธนาคาร 10.- บาท",
            align="L",
        )
        cur_y += 5.2
        pdf.set_xy(15.0, cur_y)
        pdf.cell(180.0, 5.2, f"ลงทะเบียนหลังกำหนด (ล่าช้า) ภายในวันที่ {late_dates}", align="L")
        cur_y += 6.5

        # Signatures Block on Right Half
        sig_x = 118.0
        pdf.set_font(font_family, "", sz_md)
        pdf.set_xy(sig_x, cur_y)
        pdf.cell(77.0, 5.5, f"ลงชื่อ..........{stu_name}..........ผู้ลงทะเบียน", align="R")
        cur_y += 5.5
        pdf.set_xy(sig_x, cur_y)
        pdf.cell(77.0, 5.5, "....................../....................../......................", align="C")
        cur_y += 5.8
        pdf.set_xy(sig_x, cur_y)
        pdf.cell(77.0, 5.5, "ลงชื่อ............................................................ครูที่ปรึกษา", align="R")
        cur_y += 5.8
        pdf.set_xy(sig_x, cur_y)
        pdf.cell(77.0, 5.5, "ลงชื่อ.....................................................หัวหน้างานปกครอง", align="R")
        cur_y += 5.8
        pdf.set_xy(sig_x, cur_y)
        pdf.cell(77.0, 5.5, "ชำระเงินแล้วตามใบเสร็จรับเงินเล่มที่...................................", align="R")
        cur_y += 5.8
        pdf.set_xy(sig_x, cur_y)
        pdf.cell(77.0, 5.5, "เลขที่................................วันที่......................................", align="R")
        cur_y += 5.8
        pdf.set_xy(sig_x, cur_y)
        pdf.cell(77.0, 5.5, "ลงชื่อ...................................................เจ้าหน้าที่งานการเงิน", align="R")

        # Bottom Tear-Off Section: ตัดส่วนนี้ — สำหรับธนาคาร
        cut_y = max(cur_y + 8.0, 250.0)
        pdf.set_xy(15.0, cut_y)
        pdf.set_font(font_family, "B", sz_md)
        pdf.cell(18.0, 5.5, "ตัดส่วนนี้", align="L")
        pdf.line(33.0, cut_y + 3.5, 195.0, cut_y + 3.5)

        bank_y = cut_y + 6.5
        pdf.set_xy(85.0, bank_y)
        pdf.set_font(font_family, "B", sz_md)
        pdf.cell(35.0, 5.5, "สำหรับธนาคาร", align="C")
        pdf.line(91.0, bank_y + 5.2, 114.0, bank_y + 5.2)

        pdf.set_font(font_family, "", sz_md)
        pdf.set_xy(130.0, bank_y)
        pdf.cell(65.0, 5.5, "วันที่............เดือน.......................พ.ศ................", align="R")

        bank_y += 7.5
        pdf.set_xy(15.0, bank_y)
        pdf.cell(
            180.0,
            6.0,
            f"รหัสนักเรียน นักศึกษา....{stu_code}.... ชื่อ-นามสกุล....{stu_name}....ระดับชั้น....{cls_grp}....สาขาวิชา....{major_name}....",
            align="L",
        )
        bank_y += 7.0
        baht_txt = thai_baht_text(fee_total)
        pdf.set_xy(30.0, bank_y)
        pdf.cell(
            165.0,
            6.0,
            f"จำนวนเงินเป็นตัวอักษร (............{baht_txt}............) ..........{fee_total:,.2f}..........บาท",
            align="L",
        )
        bank_y += 7.0
        pdf.set_xy(15.0, bank_y)
        pdf.cell(
            180.0,
            6.0,
            "ผู้นำฝาก...................................................เบอร์โทรศัพท์........................................ ลงชื่อ...................................................เจ้าหน้าที่ธนาคาร",
            align="L",
        )

    # ==========================================================================
    # 3. DAILY CLOSING REPORT RENDERING
    # ==========================================================================
    def _render_daily_report_header(self, pdf, font_family: str, data: Dict[str, Any], meta: Dict[str, Any]):
        is_sarabun = font_family == "THSarabun"
        college = self.college_name or meta.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")
        addr = meta.get("college_address", "455 ถ.พระพันวษา ต.ท่าพี่เลี้ยง อ.เมืองสุพรรณบุรี จ.สุพรรณบุรี 72000")
        tax_id = meta.get("college_tax_id", "0994000159721")

        pdf.set_text_color(15, 23, 42)
        pdf.set_font(font_family, "B", 20 if is_sarabun else 16)
        pdf.cell(180, 8, college, new_x="LMARGIN", new_y="NEXT", align="C")

        pdf.set_text_color(71, 85, 105)
        pdf.set_font(font_family, "", 13.5 if is_sarabun else 9.5)
        pdf.cell(180, 5.5, addr, new_x="LMARGIN", new_y="NEXT", align="C")
        pdf.cell(
            180,
            5.5,
            f"สำนักงานคณะกรรมการการอาชีวศึกษา  •  รหัสหน่วยงาน/เลขประจำตัวผู้เสียภาษี: {tax_id}",
            new_x="LMARGIN",
            new_y="NEXT",
            align="C",
        )
        pdf.ln(3)

        doc_title = data.get("doc_title", "รายงานสรุปการรับเงินค่าเทอมประจำวัน")
        pdf.set_fill_color(15, 23, 42)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font(font_family, "B", 16 if is_sarabun else 12)
        pdf.cell(180, 9, doc_title, border=0, new_x="LMARGIN", new_y="NEXT", align="C", fill=True)
        pdf.ln(3)

    def _render_daily_closing(self, pdf, font_family: str, data: Dict[str, Any]):
        is_sarabun = font_family == "THSarabun"
        sz = 15 if is_sarabun else 10.5
        inc_val = float(data.get("income_total", 0.0))
        s_rows = [
            ("วันที่ออกใบเสร็จ (Report Date):", format_receipt_date_thai(data.get("date", "-"))),
            ("จำนวนใบเสร็จที่รับชำระทั้งหมด:", f"{data.get('tx_count', 0)} ฉบับ"),
            ("ยอดรวมเงินรับสุทธิประจำวัน:", f"{inc_val:,.2f} บาท ({thai_baht_text(inc_val)})"),
        ]
        pdf.set_draw_color(203, 213, 225)
        for idx, (k, v) in enumerate(s_rows):
            pdf.set_fill_color(241, 245, 249)
            pdf.set_text_color(15, 23, 42)
            pdf.set_font(font_family, "B", sz)
            pdf.cell(70, 8, f"  {k}", border=1, fill=True)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_font(font_family, "B" if idx == 2 else "", sz)
            pdf.cell(110, 8, f"  {v}", border=1, new_x="LMARGIN", new_y="NEXT", fill=True)

        pdf.ln(5)
        by_method: Dict[str, float] = data.get("income_by_method", {})
        if by_method:
            pdf.set_font(font_family, "B", sz)
            pdf.cell(180, 7, "1. แยกยอดเงินรับตามช่องทางการชำระเงิน:", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font(font_family, "", sz)
            for method_name, amt_val in by_method.items():
                pdf.cell(180, 6.5, f"    - {method_name}:  {float(amt_val):,.2f} บาท", new_x="LMARGIN", new_y="NEXT")

        by_fee: Dict[str, float] = data.get("income_by_fee_item", {})
        if by_fee:
            pdf.ln(3)
            pdf.set_font(font_family, "B", sz)
            pdf.cell(180, 7, "2. แยกยอดเงินรับตามหมวดรายการค่าธรรมเนียมในใบเสร็จ:", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font(font_family, "", sz)
            for fee_name, amt_val in by_fee.items():
                pdf.cell(180, 6.5, f"    - {fee_name}:  {float(amt_val):,.2f} บาท", new_x="LMARGIN", new_y="NEXT")

    def _render_daily_signature(self, pdf, font_family: str, data: Dict[str, Any], meta: Dict[str, Any]):
        is_sarabun = font_family == "THSarabun"
        sz = 15 if is_sarabun else 10.5
        pdf.ln(14)
        raw_recorder = data.get("recorded_by", data.get("cashier_name", "เจ้าหน้าที่การเงิน"))
        recorder = _clean_signer_name(raw_recorder)
        position = (data.get("cashier_position") or meta.get("cashier_position", "ครู")).strip()
        college = (self.college_name or meta.get("college_name", "วิทยาลัยอาชีวศึกษาสุพรรณบุรี")).strip()

        sig_x = 105.0
        sig_w = 88.0
        pdf.set_text_color(15, 23, 42)
        pdf.set_font(font_family, "", sz)
        pdf.set_x(sig_x)
        pdf.cell(sig_w, 6.5, "ลงชื่อ .................................................. ผู้รับเงิน", new_x="LMARGIN", new_y="NEXT", align="C")
        pdf.set_font(font_family, "B", sz)
        pdf.set_x(sig_x)
        pdf.cell(sig_w, 6.5, f"({recorder})", new_x="LMARGIN", new_y="NEXT", align="C")
        pdf.set_font(font_family, "", sz)
        pdf.set_x(sig_x)
        pdf.cell(sig_w, 6, f"(ตำแหน่ง)   {position}", new_x="LMARGIN", new_y="NEXT", align="C")
        pdf.set_text_color(71, 85, 105)
        pdf.set_x(sig_x)
        pdf.cell(sig_w, 6, college, new_x="LMARGIN", new_y="NEXT", align="C")


class PdfStatementExporter(PdfReceiptExporter):
    """Backward-compatible alias for PDF exporting."""
    pass
