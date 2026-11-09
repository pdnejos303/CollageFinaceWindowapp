import os
from typing import Any, Dict, List
from exporters.base import BaseFinanceExporter

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
        transactions: List[Dict[str, Any]] = data.get("transactions", [])
        students: List[Dict[str, Any]] = data.get("students", [])
        budgets: List[Dict[str, Any]] = data.get("budgets", [])
        summary = data.get("summary", {})

        # High-volume streaming mode (write_only=True) for >5,000 rows to prevent RAM/CPU freeze at 100,000+ rows
        if len(transactions) + len(students) > 5000:
            wb = openpyxl.Workbook(write_only=True)
            ws_ledger = wb.create_sheet(title="General Ledger")
            ws_ledger.append([f"{self.college_name} - รายงานสมุดบัญชีแยกประเภท (General Ledger)"])
            ws_ledger.append([
                f"รายรับรวม: ฿{summary.get('total_income', 0):,.2f}  |  "
                f"รายจ่ายรวม: ฿{summary.get('total_expense', 0):,.2f}  |  "
                f"ดุลการคลังสุทธิ: ฿{summary.get('net_balance', 0):,.2f}"
            ])
            ws_ledger.append([])
            ws_ledger.append(["#", "เลขที่เอกสาร", "วันที่-เวลา", "ประเภท", "หมวดหมู่", "รายการ", "จำนวนเงิน (บาท)", "ช่องทางชำระ", "อ้างอิง", "ผู้บันทึก", "สถานะ"])
            for r_idx, tx in enumerate(transactions, 1):
                ws_ledger.append([
                    r_idx,
                    tx.get("tx_code", ""),
                    tx.get("tx_date", ""),
                    tx.get("tx_type", "INCOME"),
                    tx.get("category", ""),
                    tx.get("title", ""),
                    float(tx.get("amount", 0.0)),
                    tx.get("payment_method", ""),
                    tx.get("ref_code", ""),
                    tx.get("recorded_by", ""),
                    tx.get("status", "COMPLETED"),
                ])

            ws_stu = wb.create_sheet(title="Student Tuition")
            ws_stu.append([f"{self.college_name} - รายงานสถานะค่าเทอมและลูกหนี้การศึกษา (แยกตามระดับชั้น ปวช./ปวส./ป.ตรี และแผนกวิชา)"])
            ws_stu.append([])
            ws_stu.append(["รหัสนักศึกษา", "ชื่อ-นามสกุล", "ระดับชั้น", "กลุ่มเรียน/ห้อง", "แผนกวิชา", "สาขางาน", "ประเภททุน", "ค่าเทอมรวม", "ชำระแล้ว", "ค้างชำระ", "สถานะ"])
            for st in students:
                bal = max(0.0, float(st.get("total_fee", 0.0)) - float(st.get("paid_amount", 0.0)))
                ws_stu.append([
                    st.get("student_code", ""),
                    st.get("full_name", ""),
                    st.get("education_level", "ปวช."),
                    st.get("class_group", "ปวช.1/1"),
                    st.get("faculty", ""),
                    st.get("major", ""),
                    st.get("scholarship_type", ""),
                    float(st.get("total_fee", 0.0)),
                    float(st.get("paid_amount", 0.0)),
                    bal,
                    st.get("status", "UNPAID"),
                ])

            ws_bgt = wb.create_sheet(title="Department Budgets")
            ws_bgt.append([f"{self.college_name} - สรุปงบประมาณประจำปีแยกตามคณะ/หน่วยงาน"])
            ws_bgt.append([])
            ws_bgt.append(["รหัสหน่วยงาน", "ชื่อคณะ / หน่วยงาน", "ปีงบประมาณ", "งบประมาณตั้งต้น", "เบิกจ่ายแล้ว", "งบคงเหลือ", "อัตราเบิกจ่าย (%)"])
            for b in budgets:
                alloc = float(b.get("allocated_budget", 0.0))
                spent = float(b.get("spent_budget", 0.0))
                rem = alloc - spent
                pct = (spent / alloc * 100.0) if alloc > 0 else 0.0
                ws_bgt.append([b.get("dept_code", ""), b.get("dept_name", ""), b.get("fiscal_year", "2569"), alloc, spent, rem, f"{pct:.1f}%"])

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

        summary = data.get("summary", {})
        transactions: List[Dict[str, Any]] = data.get("transactions", [])

        ws_ledger.cell(row=1, column=1, value=f"{self.college_name} - รายงานสมุดบัญชีแยกประเภท (General Ledger)").font = font_title
        ws_ledger.cell(
            row=2,
            column=1,
            value=(
                f"รายรับรวม: ฿{summary.get('total_income', 0):,.2f}  |  "
                f"รายจ่ายรวม: ฿{summary.get('total_expense', 0):,.2f}  |  "
                f"ดุลการคลังสุทธิ: ฿{summary.get('net_balance', 0):,.2f}"
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
            tx_type = tx.get("tx_type", "INCOME")
            status = tx.get("status", "COMPLETED")
            vals = [
                r_idx,
                tx.get("tx_code", ""),
                tx.get("tx_date", ""),
                tx_type,
                tx.get("category", ""),
                tx.get("title", ""),
                float(tx.get("amount", 0.0)),
                tx.get("payment_method", ""),
                tx.get("ref_code", ""),
                tx.get("recorded_by", ""),
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
        students: List[Dict[str, Any]] = data.get("students", [])

        ws_stu.cell(row=1, column=1, value=f"{self.college_name} - รายงานสถานะค่าเทอมและลูกหนี้การศึกษา (แยกตามระดับชั้น ปวช./ปวส./ป.ตรี และแผนกวิชา)").font = font_title
        headers_stu = ["รหัสนักศึกษา", "ชื่อ-นามสกุล", "ระดับชั้น", "กลุ่มเรียน/ห้อง", "แผนกวิชา", "สาขางาน", "ประเภททุน", "ค่าเทอมรวม", "ชำระแล้ว", "ค้างชำระ", "สถานะ"]
        for c_idx, h in enumerate(headers_stu, 1):
            c = ws_stu.cell(row=3, column=c_idx, value=h)
            c.font = font_header
            c.fill = fill_header
            c.alignment = align_center

        for r_idx, st in enumerate(students, 1):
            row_num = 3 + r_idx
            st_status = st.get("status", "UNPAID")
            bal = max(0.0, float(st.get("total_fee", 0.0)) - float(st.get("paid_amount", 0.0)))
            vals = [
                st.get("student_code", ""),
                st.get("full_name", ""),
                st.get("education_level", "ปวช."),
                st.get("class_group", "ปวช.1/1"),
                st.get("faculty", ""),
                st.get("major", ""),
                st.get("scholarship_type", ""),
                float(st.get("total_fee", 0.0)),
                float(st.get("paid_amount", 0.0)),
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
        budgets: List[Dict[str, Any]] = data.get("budgets", [])

        ws_bgt.cell(row=1, column=1, value=f"{self.college_name} - สรุปงบประมาณประจำปีแยกตามคณะ/หน่วยงาน").font = font_title
        headers_bgt = ["รหัสหน่วยงาน", "ชื่อคณะ / หน่วยงาน", "ปีงบประมาณ", "งบประมาณตั้งต้น", "เบิกจ่ายแล้ว", "งบคงเหลือ", "อัตราเบิกจ่าย (%)"]
        for c_idx, h in enumerate(headers_bgt, 1):
            c = ws_bgt.cell(row=3, column=c_idx, value=h)
            c.font = font_header
            c.fill = fill_header
            c.alignment = align_center

        for r_idx, b in enumerate(budgets, 1):
            row_num = 3 + r_idx
            alloc = float(b.get("allocated_budget", 0.0))
            spent = float(b.get("spent_budget", 0.0))
            rem = alloc - spent
            pct = (spent / alloc * 100.0) if alloc > 0 else 0.0
            vals = [b.get("dept_code", ""), b.get("dept_name", ""), b.get("fiscal_year", "2569"), alloc, spent, rem, f"{pct:.1f}%"]
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
