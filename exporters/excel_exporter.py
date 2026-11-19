import os
from typing import Any, Dict, List
from exporters.base import BaseFinanceExporter, safe_float, safe_int, safe_str

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
except ImportError:
    openpyxl = None


class ExcelFinanceExporter(BaseFinanceExporter):
    """Multi-sheet Corporate Excel (.xlsx) Exporter for College Financial Statements, Ledger, Students, & Budgets."""

    def export(self, data: Dict[str, Any], metadata: Dict[str, Any] = None) -> str:
        if openpyxl is None:
            raise RuntimeError("Module 'openpyxl' is not installed.")

        os.makedirs(os.path.dirname(os.path.abspath(self.save_path)), exist_ok=True)
        data = data if isinstance(data, dict) else {}
        transactions: List[Dict[str, Any]] = data.get("transactions", []) or []
        students: List[Dict[str, Any]] = data.get("students", []) or []
        budgets: List[Dict[str, Any]] = data.get("budgets", []) or []
        summary = data.get("summary", {}) or {}

        tot_inc = safe_float(summary.get("total_income"))
        tot_exp = safe_float(summary.get("total_expense"))
        net_bal = safe_float(summary.get("net_balance"))

        # High-volume streaming mode (write_only=True) for >5,000 rows to prevent RAM/CPU freeze at 100,000+ rows
        if len(transactions) + len(students) > 5000:
            wb = openpyxl.Workbook(write_only=True)
            ws_ledger = wb.create_sheet(title="General Ledger")
            ws_ledger.append([f"{self.college_name} - รายงานสมุดบัญชีแยกประเภท (General Ledger)"])
            ws_ledger.append([
                f"รายรับรวม: ฿{tot_inc:,.2f}  |  "
                f"รายจ่ายรวม: ฿{tot_exp:,.2f}  |  "
                f"ดุลการคลังสุทธิ: ฿{net_bal:,.2f}"
            ])
            ws_ledger.append([])
            ws_ledger.append(["#", "เลขที่เอกสาร", "วันที่-เวลา", "ประเภท", "หมวดหมู่", "รายการ", "จำนวนเงิน (บาท)", "ช่องทางชำระ", "อ้างอิง", "ผู้บันทึก", "สถานะ"])
            for r_idx, tx in enumerate(transactions, 1):
                ws_ledger.append([
                    r_idx,
                    safe_str(tx.get("tx_code")),
                    safe_str(tx.get("tx_date")),
                    safe_str(tx.get("tx_type") or "INCOME"),
                    safe_str(tx.get("category")),
                    safe_str(tx.get("title")),
                    safe_float(tx.get("amount")),
                    safe_str(tx.get("payment_method")),
                    safe_str(tx.get("ref_code")),
                    safe_str(tx.get("recorded_by")),
                    safe_str(tx.get("status") or "COMPLETED"),
                ])

            ws_stu = wb.create_sheet(title="Student Tuition")
            ws_stu.append([f"{self.college_name} - รายงานสถานะค่าเทอมและลูกหนี้การศึกษา (แยกตามระดับชั้น ปวช./ปวส./ป.ตรี และแผนกวิชา)"])
            ws_stu.append([])
            ws_stu.append(["รหัสนักศึกษา", "ชื่อ-นามสกุล", "ระดับชั้น", "กลุ่มเรียน/ห้อง", "แผนกวิชา", "สาขางาน", "ประเภททุน", "ค่าเทอมรวม", "ชำระแล้ว", "ค้างชำระ", "สถานะ"])
            for st in students:
                tot_fee = safe_float(st.get("total_fee"))
                paid_amt = safe_float(st.get("paid_amount"))
                bal = max(0.0, tot_fee - paid_amt)
                ws_stu.append([
                    safe_str(st.get("student_code")),
                    safe_str(st.get("full_name")),
                    safe_str(st.get("education_level") or "ปวช."),
                    safe_str(st.get("class_group") or "ปวช.1/1"),
                    safe_str(st.get("faculty")),
                    safe_str(st.get("major")),
                    safe_str(st.get("scholarship_type")),
                    tot_fee,
                    paid_amt,
                    bal,
                    safe_str(st.get("status") or "UNPAID"),
                ])

            ws_bgt = wb.create_sheet(title="Department Budgets")
            ws_bgt.append([f"{self.college_name} - สรุปงบประมาณประจำปีแยกตามคณะ/หน่วยงาน"])
            ws_bgt.append([])
            ws_bgt.append(["รหัสหน่วยงาน", "ชื่อคณะ / หน่วยงาน", "ปีงบประมาณ", "งบประมาณตั้งต้น", "เบิกจ่ายแล้ว", "งบคงเหลือ", "อัตราเบิกจ่าย (%)"])
            for b in budgets:
                alloc = safe_float(b.get("allocated_budget"))
                spent = safe_float(b.get("spent_budget"))
                rem = alloc - spent
                pct = (spent / alloc * 100.0) if alloc > 0 else 0.0
                ws_bgt.append([
                    safe_str(b.get("dept_code")),
                    safe_str(b.get("dept_name")),
                    safe_str(b.get("fiscal_year") or "2569"),
                    alloc,
                    spent,
                    rem,
                    f"{pct:.1f}%",
                ])

            wb.save(self.save_path)
            return self.save_path

        wb = openpyxl.Workbook()

        # Styles
        font_title = Font(name="Segoe UI", size=15, bold=True, color="111827")
        font_sub = Font(name="Segoe UI", size=10, color="6B7280")
        font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        font_data = Font(name="Segoe UI", size=10, color="111827")
        font_green = Font(name="Segoe UI", size=10, bold=True, color="16A34A")
        font_red = Font(name="Segoe UI", size=10, bold=True, color="DC2626")
        font_amber = Font(name="Segoe UI", size=10, bold=True, color="D97706")

        fill_header = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
        fill_pass = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
        fill_fail = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
        fill_warn = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

        thin_border = Border(
            left=Side(style="thin", color="E5E7EB"),
            right=Side(style="thin", color="E5E7EB"),
            top=Side(style="thin", color="E5E7EB"),
            bottom=Side(style="thin", color="E5E7EB"),
        )
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        align_right = Alignment(horizontal="right", vertical="center")

        # ------------------------------------------------------------------
        # Sheet 1: General Ledger (สมุดบัญชีรายรับ-รายจ่าย)
        # ------------------------------------------------------------------
        ws_ledger = wb.active
        ws_ledger.title = "General Ledger"

        ws_ledger.cell(row=1, column=1, value=f"{self.college_name} - รายงานสมุดบัญชีแยกประเภท (General Ledger)").font = font_title
        ws_ledger.cell(
            row=2,
            column=1,
            value=(
                f"รายรับรวม: ฿{tot_inc:,.2f}  |  "
                f"รายจ่ายรวม: ฿{tot_exp:,.2f}  |  "
                f"ดุลการคลังสุทธิ: ฿{net_bal:,.2f}"
            ),
        ).font = font_sub

        headers_tx = ["#", "เลขที่เอกสาร", "วันที่-เวลา", "ประเภท", "หมวดหมู่", "รายการ", "จำนวนเงิน (บาท)", "ช่องทางชำระ", "อ้างอิง", "ผู้บันทึก", "สถานะ"]
        for c_idx, h in enumerate(headers_tx, 1):
            c = ws_ledger.cell(row=4, column=c_idx, value=h)
            c.font = font_header
            c.fill = fill_header
            c.alignment = align_center

        for r_idx, tx in enumerate(transactions, 1):
            row_num = 4 + r_idx
            tx_type = safe_str(tx.get("tx_type") or "INCOME")
            status = safe_str(tx.get("status") or "COMPLETED")
            vals = [
                r_idx,
                safe_str(tx.get("tx_code")),
                safe_str(tx.get("tx_date")),
                tx_type,
                safe_str(tx.get("category")),
                safe_str(tx.get("title")),
                safe_float(tx.get("amount")),
                safe_str(tx.get("payment_method")),
                safe_str(tx.get("ref_code")),
                safe_str(tx.get("recorded_by")),
                status,
            ]
            for c_idx, val in enumerate(vals, 1):
                cell = ws_ledger.cell(row=row_num, column=c_idx, value=val)
                cell.border = thin_border
                cell.font = font_data
                if c_idx == 4:
                    cell.alignment = align_center
                    cell.fill = fill_pass if tx_type == "INCOME" else fill_fail
                    cell.font = font_green if tx_type == "INCOME" else font_red
                elif c_idx == 7:
                    cell.number_format = "#,##0.00"
                    cell.alignment = align_right
                    cell.font = font_green if tx_type == "INCOME" else font_red
                elif c_idx in (1, 2, 3, 8, 9, 10, 11):
                    cell.alignment = align_center
                else:
                    cell.alignment = align_left

        col_w_tx = {"A": 6, "B": 18, "C": 20, "D": 12, "E": 28, "F": 36, "G": 16, "H": 18, "I": 14, "J": 14, "K": 14}
        for col_l, w in col_w_tx.items():
            ws_ledger.column_dimensions[col_l].width = w

        # ------------------------------------------------------------------
        # Sheet 2: Student Tuition Status (ทะเบียนค่าเทอมนักศึกษา)
        # ------------------------------------------------------------------
        ws_stu = wb.create_sheet(title="Student Tuition")

        ws_stu.cell(row=1, column=1, value=f"{self.college_name} - รายงานสถานะค่าเทอมและลูกหนี้การศึกษา (แยกตามระดับชั้น ปวช./ปวส./ป.ตรี และแผนกวิชา)").font = font_title
        headers_stu = ["รหัสนักศึกษา", "ชื่อ-นามสกุล", "ระดับชั้น", "กลุ่มเรียน/ห้อง", "แผนกวิชา", "สาขางาน", "ประเภททุน", "ค่าเทอมรวม", "ชำระแล้ว", "ค้างชำระ", "สถานะ"]
        for c_idx, h in enumerate(headers_stu, 1):
            c = ws_stu.cell(row=3, column=c_idx, value=h)
            c.font = font_header
            c.fill = fill_header
            c.alignment = align_center

        for r_idx, st in enumerate(students, 1):
            row_num = 3 + r_idx
            st_status = safe_str(st.get("status") or "UNPAID")
            tot_fee = safe_float(st.get("total_fee"))
            paid_amt = safe_float(st.get("paid_amount"))
            bal = max(0.0, tot_fee - paid_amt)
            vals = [
                safe_str(st.get("student_code")),
                safe_str(st.get("full_name")),
                safe_str(st.get("education_level") or "ปวช."),
                safe_str(st.get("class_group") or "ปวช.1/1"),
                safe_str(st.get("faculty")),
                safe_str(st.get("major")),
                safe_str(st.get("scholarship_type")),
                tot_fee,
                paid_amt,
                bal,
                st_status,
            ]
            for c_idx, val in enumerate(vals, 1):
                cell = ws_stu.cell(row=row_num, column=c_idx, value=val)
                cell.border = thin_border
                cell.font = font_data
                if c_idx in (8, 9, 10):
                    cell.number_format = "#,##0.00"
                    cell.alignment = align_right
                elif c_idx == 11:
                    cell.alignment = align_center
                    if st_status == "PAID":
                        cell.fill = fill_pass
                        cell.font = font_green
                    elif st_status == "PARTIAL":
                        cell.fill = fill_warn
                        cell.font = font_amber
                    else:
                        cell.fill = fill_fail
                        cell.font = font_red
                elif c_idx in (1, 3, 4):
                    cell.alignment = align_center
                else:
                    cell.alignment = align_left

        col_w_stu = {"A": 15, "B": 24, "C": 14, "D": 18, "E": 26, "F": 26, "G": 18, "H": 15, "I": 15, "J": 15, "K": 14}
        for col_l, w in col_w_stu.items():
            ws_stu.column_dimensions[col_l].width = w

        # ------------------------------------------------------------------
        # Sheet 3: Department Budgets (งบประมาณคณะและหน่วยงาน)
        # ------------------------------------------------------------------
        ws_bgt = wb.create_sheet(title="Department Budgets")

        ws_bgt.cell(row=1, column=1, value=f"{self.college_name} - สรุปงบประมาณประจำปีแยกตามคณะ/หน่วยงาน").font = font_title
        headers_bgt = ["รหัสหน่วยงาน", "ชื่อคณะ / หน่วยงาน", "ปีงบประมาณ", "งบประมาณตั้งต้น", "เบิกจ่ายแล้ว", "งบคงเหลือ", "อัตราเบิกจ่าย (%)"]
        for c_idx, h in enumerate(headers_bgt, 1):
            c = ws_bgt.cell(row=3, column=c_idx, value=h)
            c.font = font_header
            c.fill = fill_header
            c.alignment = align_center

        for r_idx, b in enumerate(budgets, 1):
            row_num = 3 + r_idx
            alloc = safe_float(b.get("allocated_budget"))
            spent = safe_float(b.get("spent_budget"))
            rem = alloc - spent
            pct = (spent / alloc * 100.0) if alloc > 0 else 0.0
            vals = [
                safe_str(b.get("dept_code")),
                safe_str(b.get("dept_name")),
                safe_str(b.get("fiscal_year") or "2569"),
                alloc,
                spent,
                rem,
                f"{pct:.1f}%",
            ]
            for c_idx, val in enumerate(vals, 1):
                cell = ws_bgt.cell(row=row_num, column=c_idx, value=val)
                cell.border = thin_border
                cell.font = font_data
                if c_idx in (4, 5, 6):
                    cell.number_format = "#,##0.00"
                    cell.alignment = align_right
                elif c_idx in (1, 3, 7):
                    cell.alignment = align_center
                else:
                    cell.alignment = align_left

        col_w_bgt = {"A": 16, "B": 35, "C": 14, "D": 18, "E": 18, "F": 18, "G": 16}
        for col_l, w in col_w_bgt.items():
            ws_bgt.column_dimensions[col_l].width = w

        wb.save(self.save_path)
        return self.save_path

    def export_dashboard(self, analytics: Dict[str, Any], metadata: Dict[str, Any] = None) -> str:
        """Exports Executive Financial Dashboard Summary to a formatted multi-sheet Excel (.xlsx) file."""
        if openpyxl is None:
            raise RuntimeError("Module 'openpyxl' is not installed.")

        os.makedirs(os.path.dirname(os.path.abspath(self.save_path)), exist_ok=True)
        meta = metadata if isinstance(metadata, dict) else {}
        analytics = analytics if isinstance(analytics, dict) else {}
        wb = openpyxl.Workbook()

        # Styles
        font_title = Font(name="Segoe UI", size=15, bold=True, color="0F172A")
        font_sub = Font(name="Segoe UI", size=10, color="475569")
        font_sec_header = Font(name="Segoe UI", size=11, bold=True, color="1E3A8A")
        font_tbl_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        font_data = Font(name="Segoe UI", size=10, color="0F172A")
        font_data_bold = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
        font_green = Font(name="Segoe UI", size=10, bold=True, color="0F766E")
        font_red = Font(name="Segoe UI", size=10, bold=True, color="B91C1C")

        fill_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        fill_sec_bar = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        fill_pass = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
        fill_fail = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

        thin_border = Border(
            left=Side(style="thin", color="E2E8F0"),
            right=Side(style="thin", color="E2E8F0"),
            top=Side(style="thin", color="E2E8F0"),
            bottom=Side(style="thin", color="E2E8F0"),
        )
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        # ------------------------------------------------------------------
        # Sheet 1: Dashboard Overview (ภาพรวมสถิติการเงิน)
        # ------------------------------------------------------------------
        ws_dash = wb.active
        ws_dash.title = "Dashboard Summary"

        term_lbl = safe_str(analytics.get("term_filter") or "ทั้งหมด")
        period_lbl = safe_str(analytics.get("period_label") or "ทั้งหมด")
        level_lbl = safe_str(analytics.get("level_label") or "ทุกระดับชั้น")
        dept_lbl = safe_str(analytics.get("department_label") or "ทุกแผนกวิชา")
        cashier = safe_str(meta.get("cashier_name") or "เจ้าหน้าที่การเงิน")

        ws_dash.cell(row=1, column=1, value=f"{self.college_name} - สรุปภาพรวมสถิติการเงินและการรับชำระ (Executive Dashboard)").font = font_title
        ws_dash.cell(
            row=2,
            column=1,
            value=f"ภาคเรียน: {term_lbl}  |  ช่วงเวลา: {period_lbl}  |  ชั้นปี: {level_lbl}  |  แผนกวิชา: {dept_lbl}  |  ผู้พิมพ์: {cashier}",
        ).font = font_sub

        # KPI Block
        month_label = safe_str(analytics.get("month_label_thai") or "เดือนนี้")
        coll_rate = safe_float(analytics.get("collection_rate_percent"))
        stu_cnt = safe_int(analytics.get("student_count"))
        paid_cnt = safe_int(analytics.get("paid_student_count"))
        unpaid_cnt = safe_int(analytics.get("unpaid_student_count"))

        kpis = [
            ("ยอดจัดเก็บเงินรวมทั้งหมด", safe_float(analytics.get("total_income")), True),
            (f"ยอดรับชำระประจำเดือนนี้ ({month_label})", safe_float(analytics.get("month_income")), True),
            ("ยอดรับชำระวันนี้", safe_float(analytics.get("today_income")), True),
            ("จำนวนใบเสร็จที่ออก (ฉบับปกติ)", safe_int(analytics.get("receipt_count")), False),
            ("จำนวนใบเสร็จที่ยกเลิก (Void)", safe_int(analytics.get("void_receipt_count")), False),
            ("ยอดประมาณการค่าเทอมรวม", safe_float(analytics.get("total_tuition_expected")), True),
            ("ยอดค้างชำระค่าเทอม (AR)", safe_float(analytics.get("outstanding_receivables")), True),
            ("จำนวนนักศึกษาทั้งหมด", f"{stu_cnt} คน", False),
            ("สถานะการชำระเงิน", f"ชำระครบ {paid_cnt} คน / ค้างชำระ {unpaid_cnt} คน", False),
            ("อัตราการจัดเก็บค่าเทอม", f"{coll_rate:.1f}%", False),
        ]

        ws_dash.cell(row=4, column=1, value="📌 สรุปดัชนีชี้วัดหลักทางการเงิน (Key Financial Indicators)").font = font_sec_header
        for k_idx, (kpi_name, kpi_val, is_currency) in enumerate(kpis, 5):
            c_lbl = ws_dash.cell(row=k_idx, column=1, value=kpi_name)
            c_lbl.font = font_data_bold
            c_lbl.fill = fill_sec_bar
            c_lbl.border = thin_border

            c_val = ws_dash.cell(row=k_idx, column=2, value=kpi_val)
            c_val.font = font_data_bold
            c_val.border = thin_border
            if is_currency:
                c_val.number_format = "฿#,##0.00"
                c_val.alignment = align_right
                c_val.font = font_green if "ค้าง" not in kpi_name else font_red
            else:
                c_val.alignment = align_left

        curr_r = 5 + len(kpis) + 1

        # Payment Methods Table
        ws_dash.cell(row=curr_r, column=1, value="💳 ยอดรับชำระแยกตามช่องทางการเงิน").font = font_sec_header
        curr_r += 1
        pm_headers = ["ช่องทางการชำระเงิน", "จำนวนใบเสร็จ (ฉบับ)", "ยอดเงินรวม (บาท)", "สัดส่วน (%)"]
        for c_idx, h in enumerate(pm_headers, 1):
            c = ws_dash.cell(row=curr_r, column=c_idx, value=h)
            c.font = font_tbl_header
            c.fill = fill_header
            c.alignment = align_center

        methods = analytics.get("by_payment_method", []) or []
        if not methods:
            curr_r += 1
            ws_dash.cell(row=curr_r, column=1, value="ไม่มีข้อมูลการรับชำระ").font = font_data
        else:
            for m in methods:
                curr_r += 1
                m_name = safe_str(m.get("name") if isinstance(m, dict) else "")
                m_cnt = safe_int(m.get("count") if isinstance(m, dict) else 0)
                m_amt = safe_float(m.get("amount") if isinstance(m, dict) else 0.0)
                m_pct = safe_float(m.get("percent") if isinstance(m, dict) else 0.0)
                row_vals = [m_name, m_cnt, m_amt, f"{m_pct:.1f}%"]
                for c_idx, val in enumerate(row_vals, 1):
                    cell = ws_dash.cell(row=curr_r, column=c_idx, value=val)
                    cell.border = thin_border
                    cell.font = font_data
                    if c_idx == 1:
                        cell.alignment = align_left
                    elif c_idx == 2:
                        cell.alignment = align_center
                    elif c_idx == 3:
                        cell.number_format = "#,##0.00"
                        cell.alignment = align_right
                    else:
                        cell.alignment = align_center

        curr_r += 2

        # Education Levels Table
        ws_dash.cell(row=curr_r, column=1, value="🏫 ยอดรับชำระแยกตามระดับชั้น").font = font_sec_header
        curr_r += 1
        lvl_headers = ["ระดับชั้น", "จำนวนใบเสร็จ (ฉบับ)", "ยอดเงินรวม (บาท)", "สัดส่วน (%)"]
        for c_idx, h in enumerate(lvl_headers, 1):
            c = ws_dash.cell(row=curr_r, column=c_idx, value=h)
            c.font = font_tbl_header
            c.fill = fill_header
            c.alignment = align_center

        levels = analytics.get("by_education_level", []) or []
        if not levels:
            curr_r += 1
            ws_dash.cell(row=curr_r, column=1, value="ไม่มีข้อมูลระดับชั้น").font = font_data
        else:
            for l in levels:
                curr_r += 1
                l_lvl = safe_str(l.get("level") if isinstance(l, dict) else "")
                l_cnt = safe_int(l.get("count") if isinstance(l, dict) else 0)
                l_amt = safe_float(l.get("amount") if isinstance(l, dict) else 0.0)
                l_pct = safe_float(l.get("percent") if isinstance(l, dict) else 0.0)
                row_vals = [l_lvl, l_cnt, l_amt, f"{l_pct:.1f}%"]
                for c_idx, val in enumerate(row_vals, 1):
                    cell = ws_dash.cell(row=curr_r, column=c_idx, value=val)
                    cell.border = thin_border
                    cell.font = font_data
                    if c_idx == 1:
                        cell.alignment = align_left
                    elif c_idx == 2:
                        cell.alignment = align_center
                    elif c_idx == 3:
                        cell.number_format = "#,##0.00"
                        cell.alignment = align_right
                    else:
                        cell.alignment = align_center

        curr_r += 2

        # Top 5 Fee Items
        ws_dash.cell(row=curr_r, column=1, value="🏆 5 อันดับรายการค่าธรรมเนียมที่มียอดจัดเก็บสูงสุด").font = font_sec_header
        curr_r += 1
        fee_headers = ["อันดับ", "รายการค่าธรรมเนียม", "จำนวนครั้งที่ชำระ", "ยอดเงินรวม (บาท)", "สัดส่วน (%)"]
        for c_idx, h in enumerate(fee_headers, 1):
            c = ws_dash.cell(row=curr_r, column=c_idx, value=h)
            c.font = font_tbl_header
            c.fill = fill_header
            c.alignment = align_center

        top_fees = analytics.get("top_fee_items", []) or []
        if not top_fees:
            curr_r += 1
            ws_dash.cell(row=curr_r, column=1, value="ไม่มีข้อมูลค่าธรรมเนียม").font = font_data
        else:
            for rank_i, tf in enumerate(top_fees, 1):
                curr_r += 1
                tf_name = safe_str(tf.get("item_name") if isinstance(tf, dict) else "")
                tf_qty = safe_int(tf.get("quantity") if isinstance(tf, dict) else 0)
                tf_total = safe_float(tf.get("total") if isinstance(tf, dict) else 0.0)
                tf_pct = safe_float(tf.get("percent") if isinstance(tf, dict) else 0.0)
                row_vals = [rank_i, tf_name, tf_qty, tf_total, f"{tf_pct:.1f}%"]
                for c_idx, val in enumerate(row_vals, 1):
                    cell = ws_dash.cell(row=curr_r, column=c_idx, value=val)
                    cell.border = thin_border
                    cell.font = font_data
                    if c_idx in (1, 3, 5):
                        cell.alignment = align_center
                    elif c_idx == 2:
                        cell.alignment = align_left
                    elif c_idx == 4:
                        cell.number_format = "#,##0.00"
                        cell.alignment = align_right

        ws_dash.column_dimensions["A"].width = 38
        ws_dash.column_dimensions["B"].width = 42
        ws_dash.column_dimensions["C"].width = 24
        ws_dash.column_dimensions["D"].width = 20
        ws_dash.column_dimensions["E"].width = 16

        # ------------------------------------------------------------------
        # Sheet 2: Recent Receipts (ทะเบียนใบเสร็จล่าสุด)
        # ------------------------------------------------------------------
        ws_rec = wb.create_sheet(title="Recent Receipts")
        ws_rec.cell(row=1, column=1, value=f"{self.college_name} - รายการใบเสร็จรับเงินล่าสุด (Recent Receipts Activity)").font = font_title
        ws_rec.cell(row=2, column=1, value="รายการใบเสร็จรับเงิน 15 รายการล่าสุดในระบบ").font = font_sub

        rc_headers = [
            "#", "เลขที่ใบเสร็จ", "วันที่ออกใบเสร็จ", "ภาคเรียน", "รหัสนักศึกษา",
            "ชื่อ-นามสกุล", "ระดับชั้น", "แผนกวิชา / สาขา", "ช่องทางชำระเงิน",
            "ยอดเงินสุทธิ (บาท)", "ผู้รับเงิน", "สถานะ",
        ]
        for c_idx, h in enumerate(rc_headers, 1):
            c = ws_rec.cell(row=4, column=c_idx, value=h)
            c.font = font_tbl_header
            c.fill = fill_header
            c.alignment = align_center

        recent = analytics.get("recent_receipts", []) or []
        for r_idx, r in enumerate(recent, 1):
            row_num = 4 + r_idx
            st = safe_str(r.get("status") if isinstance(r, dict) else "ISSUED") or "ISSUED"
            vals = [
                r_idx,
                safe_str(r.get("receipt_no")),
                safe_str(r.get("receipt_date")),
                safe_str(r.get("term_code")),
                safe_str(r.get("student_code")),
                safe_str(r.get("snapshot_student_name")),
                safe_str(r.get("snapshot_education_level")),
                safe_str(r.get("snapshot_major") or r.get("snapshot_department")),
                safe_str(r.get("payment_method")),
                safe_float(r.get("net_total_amount")),
                safe_str(r.get("cashier_name")),
                "ปกติ" if st == "ISSUED" else "ยกเลิกแล้ว",
            ]
            for c_idx, val in enumerate(vals, 1):
                cell = ws_rec.cell(row=row_num, column=c_idx, value=val)
                cell.border = thin_border
                cell.font = font_data
                if c_idx == 10:
                    cell.number_format = "#,##0.00"
                    cell.alignment = align_right
                    cell.font = font_green if st == "ISSUED" else font_red
                elif c_idx == 12:
                    cell.alignment = align_center
                    cell.fill = fill_pass if st == "ISSUED" else fill_fail
                    cell.font = font_green if st == "ISSUED" else font_red
                elif c_idx in (1, 2, 3, 4, 5, 7, 9, 11):
                    cell.alignment = align_center
                else:
                    cell.alignment = align_left

        col_w_rec = {
            "A": 6, "B": 16, "C": 18, "D": 12, "E": 16,
            "F": 28, "G": 14, "H": 26, "I": 20, "J": 18, "K": 24, "L": 14,
        }
        for col_l, w in col_w_rec.items():
            ws_rec.column_dimensions[col_l].width = w

        wb.save(self.save_path)
        return self.save_path
