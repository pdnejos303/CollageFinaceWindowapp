from dataclasses import dataclass
from typing import Optional


@dataclass
class UserSession:
    user_id: int
    username: str
    full_name: str
    role: str  # "ADMIN", "FINANCE", "AUDITOR"
    department: str


@dataclass
class StudentRecord:
    student_code: str
    full_name: str
    faculty: str
    major: str
    year_level: int
    scholarship_type: str
    total_fee: float
    paid_amount: float
    status: str  # "PAID", "PARTIAL", "UNPAID"
    phone: str = ""
    citizen_id: str = ""

    @property
    def balance(self) -> float:
        return max(0.0, round(self.total_fee - self.paid_amount, 2))


@dataclass
class TransactionRecord:
    tx_code: str
    tx_date: str
    tx_type: str        # "INCOME" or "EXPENSE"
    category: str       # e.g., "ค่าเทอม (Tuition)", "งบอุปกรณ์ (Equipment)", "เงินเดือน (Payroll)"
    title: str
    amount: float
    payment_method: str # "โอนบัญชี (Transfer)", "QR PromptPay", "เงินสด (Cash)", "เช็ค (Cheque)"
    ref_code: str       # Student Code or Department Code or Staff Code
    recorded_by: str
    status: str = "COMPLETED"  # "COMPLETED", "VOIDED"
    note: str = ""
